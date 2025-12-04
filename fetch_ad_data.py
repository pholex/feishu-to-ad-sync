#!/usr/bin/env python3
import csv
import subprocess
import os
import sys
import time
from dotenv import load_dotenv

# 获取脚本所在目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 确保output目录存在
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_output_path(filename):
    """获取output目录下文件的绝对路径"""
    return os.path.join(SCRIPT_DIR, 'output', filename)

def get_ps_path(filename):
    """获取powershell目录下文件的绝对路径"""
    return os.path.join(SCRIPT_DIR, 'powershell', filename)

# 加载环境变量
load_dotenv(os.path.join(SCRIPT_DIR, '.env'))

# 域控制器配置
DC_HOST = os.getenv("DC_HOST")
DC_USER = os.getenv("DC_USER")
DC_PASSWORD = os.getenv("DC_PASSWORD")
DC_DOMAIN = os.getenv("DC_DOMAIN")
DC_BASE_OU = os.getenv("DC_BASE_OU", "")
DC_EXCLUDE_OU = os.getenv("DC_EXCLUDE_OU", "")
DC_RESIGNED_OU = os.getenv("DC_RESIGNED_OU", "")

# SSH ControlMaster 配置
SSH_CONTROL_PATH = "/tmp/ssh-feishu-ad-sync-%r@%h:%p"
SSH_CONTROL_MASTER_INITIALIZED = False
DC_USER_HOME = None

def init_ssh_control_master():
    """初始化SSH ControlMaster连接"""
    global SSH_CONTROL_MASTER_INITIALIZED
    
    if SSH_CONTROL_MASTER_INITIALIZED:
        return
    
    cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlMaster=yes -o ControlPath={SSH_CONTROL_PATH} -o ControlPersist=5m -fN {DC_USER}@{DC_HOST}"
    try:
        subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
        SSH_CONTROL_MASTER_INITIALIZED = True
    except:
        pass

def init_dc_user_home():
    """初始化域控制器用户主目录路径"""
    global DC_USER_HOME
    
    if DC_USER_HOME:
        return DC_USER_HOME
    
    init_ssh_control_master()
    
    try:
        cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -Command \"echo $env:USERPROFILE\"'"
        result = run_ssh_with_retry(cmd, timeout=10)
        if result.returncode == 0:
            DC_USER_HOME = result.stdout.decode('utf-8').strip().replace('\\', '/')
            return DC_USER_HOME
    except:
        pass
    
    print(f"✗ 无法获取域控制器用户主目录")
    sys.exit(1)

def run_ssh_with_retry(cmd, max_retries=3, timeout=30):
    """执行SSH命令，带重试机制"""
    for attempt in range(max_retries):
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, timeout=timeout)
            if result.returncode == 0:
                return result
            if attempt < max_retries - 1:
                time.sleep(2)
        except subprocess.TimeoutExpired:
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
    return result

def run_scp_with_retry(cmd, max_retries=3):
    """执行SCP命令，带重试机制"""
    for attempt in range(max_retries):
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True)
            if result.returncode == 0:
                return result
            if attempt < max_retries - 1:
                time.sleep(2)
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
    return result

def upload_dc_config():
    """上传配置到域控制器"""
    config_content = f"""# AD Domain Configuration
$DC_BASE_OU = "{DC_BASE_OU}"
$DC_DOMAIN = "{DC_DOMAIN}"
$DC_EXCLUDE_OU = "{DC_EXCLUDE_OU}"
$DC_RESIGNED_OU = "{DC_RESIGNED_OU}"
"""
    
    config_file = get_output_path('dc_config.ps1')
    with open(config_file, 'w', encoding='utf-8-sig', newline='\r\n') as f:
        f.write(config_content)
    
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {config_file} {DC_USER}@{DC_HOST}:~/dc_config.ps1"
    run_scp_with_retry(scp_cmd)

def export_ad_users():
    """导出 AD 用户"""
    print("=" * 60)
    print("导出 AD 用户信息")
    print(f"目标 OU: {DC_BASE_OU}")
    if DC_EXCLUDE_OU:
        print(f"排除 OU: {DC_EXCLUDE_OU}")
    print("=" * 60)
    
    # 初始化
    init_ssh_control_master()
    init_dc_user_home()
    
    # 上传配置
    print("\n正在上传配置...")
    upload_dc_config()
    
    # 上传脚本
    ps_export = get_ps_path('export_users.ps1')
    if not os.path.exists(ps_export):
        print(f"✗ 找不到脚本: {ps_export}")
        sys.exit(1)
    
    print("正在上传脚本...")
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {ps_export} {DC_USER}@{DC_HOST}:~/ExportUsers.ps1"
    run_scp_with_retry(scp_cmd)
    
    # 执行导出
    print("正在执行导出...")
    ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -File {DC_USER_HOME}/ExportUsers.ps1'"
    run_ssh_with_retry(ssh_cmd)
    
    # 下载结果
    print("正在下载结果...")
    output_file = get_output_path('ad_users.csv')
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST}:~/ExportedUsers.csv {output_file}"
    result = run_scp_with_retry(scp_cmd)
    
    if result.returncode != 0:
        print("✗ 下载文件失败")
        sys.exit(1)
    
    # 统计用户数
    with open(output_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        user_count = sum(1 for row in reader) - 1
    
    print(f"✓ 已导出 {user_count} 个用户到 {output_file}")
    print("\n" + "=" * 60)
    print("完成")
    print("=" * 60)

if __name__ == "__main__":
    export_ad_users()
