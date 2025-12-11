#!/usr/bin/env python3
import csv
import subprocess
import os
import sys
import re
import time
from datetime import datetime
from dotenv import load_dotenv

# 修复青龙环境 - 确保QLAPI可用
ql_dir = os.getenv('QL_DIR')
if ql_dir and os.path.exists(ql_dir):
    preload_dir = os.path.join(ql_dir, 'shell', 'preload')
    if os.path.exists(preload_dir) and preload_dir not in sys.path:
        sys.path.insert(0, preload_dir)
        try:
            import client
            import builtins
            if not hasattr(builtins, 'QLAPI'):
                builtins.QLAPI = client.Client()
        except ImportError:
            pass  # 非青龙环境，忽略

# 获取脚本所在目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 确保output目录存在
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 路径辅助函数
def get_output_path(filename):
    """获取output目录下文件的绝对路径"""
    return os.path.join(SCRIPT_DIR, 'output', filename)

def get_ps_path(filename):
    """获取powershell目录下文件的绝对路径"""
    return os.path.join(SCRIPT_DIR, 'powershell', filename)

# 加载环境变量
load_dotenv(os.path.join(SCRIPT_DIR, '.env'))

# 检查 sshpass 是否安装
def check_sshpass():
    """检查 sshpass 是否已安装"""
    try:
        result = subprocess.run(['which', 'sshpass'], capture_output=True)
        if result.returncode != 0:
            print("✗ 错误: 未安装 sshpass")
            print("\n请根据您的系统安装 sshpass:")
            print("  macOS:        brew install sshpass")
            print("  Ubuntu/Debian: sudo apt-get install sshpass")
            print("  CentOS/RHEL:  sudo yum install sshpass")
            print("  Alpine:       apk add sshpass")
            sys.exit(1)
    except Exception as e:
        print(f"✗ 检查 sshpass 时出错: {e}")
        sys.exit(1)

check_sshpass()

# 域控制器配置
DC_HOST = os.getenv("DC_HOST")
DC_USER = os.getenv("DC_USER")
DC_PASSWORD = os.getenv("DC_PASSWORD")
DC_DOMAIN = os.getenv("DC_DOMAIN")
DC_BASE_OU = os.getenv("DC_BASE_OU", "")
DC_EXCLUDE_OU = os.getenv("DC_EXCLUDE_OU", "")
DC_RESIGNED_OU = os.getenv("DC_RESIGNED_OU", "")

# 远程用户主目录（延迟初始化）
DC_USER_HOME = None

# SSH ControlMaster 配置
SSH_CONTROL_PATH = "/tmp/ssh-feishu-ad-sync-%r@%h:%p"
SSH_CONTROL_MASTER_INITIALIZED = False

def init_ssh_control_master():
    """初始化SSH ControlMaster连接"""
    global SSH_CONTROL_MASTER_INITIALIZED
    
    if SSH_CONTROL_MASTER_INITIALIZED:
        return
    
    # 建立主连接（后台运行，保持5分钟）
    cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlMaster=yes -o ControlPath={SSH_CONTROL_PATH} -o ControlPersist=5m -fN {DC_USER}@{DC_HOST}"
    try:
        subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
        SSH_CONTROL_MASTER_INITIALIZED = True
    except:
        pass

def cleanup_ssh_control_master():
    """清理SSH ControlMaster连接"""
    cmd = f"ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} -O exit {DC_USER}@{DC_HOST} 2>/dev/null"
    subprocess.run(cmd, shell=True, capture_output=True)

def init_dc_user_home():
    """初始化域控制器用户主目录路径"""
    global DC_USER_HOME
    
    if DC_USER_HOME:
        return DC_USER_HOME
    
    # 初始化SSH主连接
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

def run_ssh_with_retry(cmd, max_retries=3, timeout=30, decode=True):
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

# 拼音例外映射（飞书拼音 -> AD拼音）
PINYIN_EXCEPTIONS = {}
exceptions_str = os.getenv("PINYIN_EXCEPTIONS", "")
if exceptions_str:
    for pair in exceptions_str.split(','):
        if '=' in pair:
            feishu_pinyin, ad_pinyin = pair.strip().split('=')
            PINYIN_EXCEPTIONS[feishu_pinyin.strip()] = ad_pinyin.strip()

# Dry-run 模式标志
DRY_RUN = False

# 自动确认标志
AUTO_YES = False

# Windows 编码（延迟初始化）
WINDOWS_ENCODING = None

def get_windows_encoding():
    """获取 Windows 系统编码"""
    global WINDOWS_ENCODING
    
    if WINDOWS_ENCODING:
        return WINDOWS_ENCODING
    
    try:
        # 获取 Windows 代码页
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -Command \"[System.Text.Encoding]::Default.CodePage\"'"
        result = run_ssh_with_retry(ssh_cmd, timeout=10)
        
        if result.returncode == 0:
            code_page = result.stdout.decode('utf-8', errors='ignore').strip()
            # 代码页映射
            encoding_map = {
                '936': 'gbk',      # 简体中文
                '950': 'big5',     # 繁体中文
                '65001': 'utf-8',  # UTF-8
                '1252': 'cp1252'   # 西欧
            }
            WINDOWS_ENCODING = encoding_map.get(code_page, 'gbk')
            print(f"检测到 Windows 编码: {WINDOWS_ENCODING} (代码页 {code_page})")
        else:
            # SSH 连接失败
            if 'Permission denied' in result.stderr or 'Permission denied' in result.stdout:
                print("\n✗ SSH 登录失败：密码错误或权限被拒绝")
                print(f"  主机: {DC_HOST}")
                print(f"  用户: {DC_USER}")
                print("  请检查 .env 文件中的 DC_PASSWORD 配置")
                sys.exit(1)
            else:
                WINDOWS_ENCODING = 'gbk'  # 默认使用 GBK
                print(f"无法检测编码，使用默认: {WINDOWS_ENCODING}")
    except Exception as e:
        print(f"\n✗ 连接域控制器失败: {e}")
        print(f"  主机: {DC_HOST}")
        print(f"  用户: {DC_USER}")
        print("  请检查网络连接和配置")
        sys.exit(1)
    
    return WINDOWS_ENCODING

def confirm(prompt, default=True):
    """交互式确认，默认为 y"""
    if DRY_RUN:
        # Dry-run 模式下不需要确认，直接返回 False（不执行）
        return False
    
    if AUTO_YES:
        # 自动确认模式，直接返回 True
        print(f"{prompt} [自动确认]")
        return True
    
    default_str = "y/n" if default else "y/n"
    default_hint = "（直接回车默认为 y）" if default else "（直接回车默认为 n）"
    response = input(f"{prompt} [{default_str}] {default_hint}: ").strip().lower()
    
    if response == "":
        return default
    return response in ['y', 'yes']

def check_dc_permissions():
    """检查域控制器连接和权限"""
    print("\n检查域控制器连接和权限...")
    
    # 初始化用户主目录
    init_dc_user_home()
    
    try:
        # 直接测试 AD 模块和权限
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -Command \"Import-Module ActiveDirectory; Get-ADDomain | Select-Object -ExpandProperty DNSRoot\"'"
        result = run_ssh_with_retry(ssh_cmd, timeout=30)
        
        if result.returncode == 0 and result.stdout.strip():
            domain = result.stdout.decode('utf-8', errors='ignore').strip()
            print(f"✓ 域控制器连接正常，当前域: {domain}")
            return True
        else:
            print(f"✗ 无法访问 Active Directory")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ 连接超时")
        return False
    except Exception as e:
        print(f"✗ 检查失败: {e}")
        return False

def get_ad_ou_count():
    """从AD实时获取OU数量（排除DC_EXCLUDE_OU和DC_BASE_OU本身）"""
    if DC_BASE_OU:
        base_filter = f"Get-ADOrganizationalUnit -Filter * -SearchBase '{DC_BASE_OU}'"
    else:
        base_filter = "Get-ADOrganizationalUnit -Filter *"
    
    # 添加排除逻辑
    if DC_EXCLUDE_OU:
        ps_script = f"{base_filter} | Where-Object {{$_.DistinguishedName -ne '{DC_BASE_OU}' -and $_.DistinguishedName -ne '{DC_EXCLUDE_OU}' -and $_.DistinguishedName -notlike '*,{DC_EXCLUDE_OU}'}} | Measure-Object | Select-Object -ExpandProperty Count"
    elif DC_BASE_OU:
        ps_script = f"{base_filter} | Where-Object {{$_.DistinguishedName -ne '{DC_BASE_OU}'}} | Measure-Object | Select-Object -ExpandProperty Count"
    else:
        ps_script = f"{base_filter} | Measure-Object | Select-Object -ExpandProperty Count"
    
    # Base64编码PowerShell命令（UTF-16LE）
    import base64
    encoded = base64.b64encode(ps_script.encode('utf-16-le')).decode('ascii')
    
    ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -EncodedCommand {encoded}'"
    result = run_ssh_with_retry(ssh_cmd, timeout=30)
    
    if result.returncode != 0:
        stderr = result.stderr.decode('utf-8', errors='replace') if result.stderr else ''
        raise Exception(f"获取AD OU数量失败: {stderr}")
    
    try:
        count = int(result.stdout.decode('utf-8').strip())
        return count
    except ValueError:
        stdout = result.stdout.decode('utf-8', errors='replace') if result.stdout else ''
        raise Exception(f"解析AD OU数量失败: {stdout}")

def get_ad_user_count():
    """从AD实时获取用户数量"""
    if DC_BASE_OU:
        filter_cmd = f"Get-ADUser -Filter {{Enabled -eq $true}} -SearchBase '{DC_BASE_OU}'"
    else:
        filter_cmd = "Get-ADUser -Filter {Enabled -eq $true}"
    
    if DC_EXCLUDE_OU:
        ps_script = f"{filter_cmd} | Where-Object {{$_.DistinguishedName -notlike '*{DC_EXCLUDE_OU}'}} | Measure-Object | Select-Object -ExpandProperty Count"
    else:
        ps_script = f"{filter_cmd} | Measure-Object | Select-Object -ExpandProperty Count"
    
    # Base64编码PowerShell命令（UTF-16LE）
    import base64
    encoded = base64.b64encode(ps_script.encode('utf-16-le')).decode('ascii')
    
    ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -EncodedCommand {encoded}'"
    result = run_ssh_with_retry(ssh_cmd, timeout=30)
    
    if result.returncode != 0:
        stderr = result.stderr.decode('utf-8', errors='replace') if result.stderr else ''
        raise Exception(f"获取AD用户数量失败: {stderr}")
    
    try:
        count = int(result.stdout.decode('utf-8').strip())
        return count
    except ValueError:
        stdout = result.stdout.decode('utf-8', errors='replace') if result.stdout else ''
        raise Exception(f"解析AD用户数量失败: {stdout}")

def get_existing_ad_departments():
    """从域控制器获取现有部门OU列表，返回完整路径集合"""
    print("正在获取 AD 域现有部门 OU...")
    
    # 根据配置决定是否限制 SearchBase
    if DC_BASE_OU:
        ps_script = f"""
Import-Module ActiveDirectory
$baseOU = "{DC_BASE_OU}"
Get-ADOrganizationalUnit -Filter * -SearchBase $baseOU | 
    Select-Object Name, DistinguishedName | 
    Export-Csv -Path "~/ExistingOUs.csv" -NoTypeInformation -Encoding UTF8
"""
        search_info = f"在 {DC_BASE_OU} 下"
    else:
        ps_script = """
Import-Module ActiveDirectory
Get-ADOrganizationalUnit -Filter * | 
    Select-Object Name, DistinguishedName | 
    Export-Csv -Path "~/ExistingOUs.csv" -NoTypeInformation -Encoding UTF8
"""
        search_info = "全域"
    
    # 写入临时脚本，使用 UTF-8 BOM 编码
    with open(get_output_path('temp_get_ous.ps1'), 'w', encoding='utf-8-sig') as f:
        f.write(ps_script)
    
    try:
        # 上传并执行
        scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {get_output_path('temp_get_ous.ps1')} {DC_USER}@{DC_HOST}:~/GetOUs.ps1"
        run_scp_with_retry(scp_cmd)
        
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -File {DC_USER_HOME}/GetOUs.ps1'"
        result = run_ssh_with_retry(ssh_cmd)
        
        if result.returncode != 0:
            print(f"执行脚本失败，尝试获取全部 OU...")
            # 如果失败，尝试获取所有 OU
            ps_script_all = """
Import-Module ActiveDirectory
Get-ADOrganizationalUnit -Filter * | 
    Select-Object Name, DistinguishedName | 
    Export-Csv -Path "~/ExistingOUs.csv" -NoTypeInformation -Encoding UTF8
"""
            with open(get_output_path('temp_get_ous.ps1'), 'w', encoding='utf-8-sig') as f:
                f.write(ps_script_all)
            
            scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {get_output_path('temp_get_ous.ps1')} {DC_USER}@{DC_HOST}:~/GetOUs.ps1"
            run_scp_with_retry(scp_cmd)
            
            ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -File {DC_USER_HOME}/GetOUs.ps1'"
            run_ssh_with_retry(ssh_cmd)
            search_info = "全域(降级)"
        
        # 下载结果
        scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST}:~/ExistingOUs.csv {get_output_path('ad_existing_ous.csv')}"
        result = run_scp_with_retry(scp_cmd)
        
        if result.returncode != 0:
            # 文件下载失败，可能是AD上没有OU，先验证数量
            print("⚠ 下载AD OU文件失败，验证AD上是否真的没有OU...")
            try:
                actual_count = get_ad_ou_count()
                if actual_count == 0:
                    print("✓ AD上确实没有OU，继续执行")
                    return {}
                else:
                    print(f"❌ AD上有 {actual_count} 个OU，但文件下载失败")
                    print("数据获取失败，同步失败！")
                    sys.exit(1)
            except Exception as e:
                print(f"❌ 无法验证AD OU数量: {e}")
                print("无法确认数据状态，同步失败！")
                sys.exit(1)
    finally:
        # 清理临时文件
        if os.path.exists(get_output_path('temp_get_ous.ps1')):
            os.remove(get_output_path('temp_get_ous.ps1'))
    
    # 读取现有 OU，返回 {名称: DN} 的字典
    existing_ous = {}
    base_ou_dn = DC_BASE_OU if DC_BASE_OU else ""
    
    try:
        # PowerShell 使用 UTF8 导出，直接使用 utf-8-sig
        with open(get_output_path('ad_existing_ous.csv'), 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row['Name']
                dn = row['DistinguishedName']
                # 排除基础 OU 本身
                if dn == base_ou_dn:
                    continue
                # 排除指定的 OU 及其所有子 OU
                if DC_EXCLUDE_OU and (dn == DC_EXCLUDE_OU or dn.endswith(',' + DC_EXCLUDE_OU)):
                    continue
                # 使用 DN 作为唯一标识
                existing_ous[dn] = name
        print(f"✓ 发现 {len(existing_ous)} 个现有部门 OU ({search_info})")
        
        # 验证数据完整性：从AD实时获取数量并对比
        print("正在验证OU数据完整性...")
        try:
            actual_count = get_ad_ou_count()
            csv_count = len(existing_ous)
            print(f"  AD实际OU数: {actual_count}")
            print(f"  CSV文件OU数: {csv_count}")
            
            if actual_count != csv_count:
                print(f"❌ OU数量不匹配！实际: {actual_count}, CSV: {csv_count}")
                print("数据可能不完整，同步失败！")
                sys.exit(1)
            print("✓ OU数据验证通过")
        except Exception as e:
            print(f"❌ 验证OU数据失败: {e}")
            print("无法确认数据完整性，同步失败！")
            sys.exit(1)
    except Exception as e:
        print(f"❌ 读取现有 OU 列表失败: {e}")
        sys.exit(1)
    
    return existing_ous

def get_existing_ad_users():
    """从域控制器获取现有用户列表，返回两个字典：
    1. {EmployeeNumber: {SamAccountName, DisplayName, EmailAddress}} - 有 Union ID 的用户
    2. {SamAccountName: {SamAccountName, DisplayName, EmailAddress}} - 没有 Union ID 的用户
    """
    print("正在获取..." if not DRY_RUN else "")
    
    # 使用本地脚本副本
    ps_export = get_ps_path('export_users.ps1')
    if not os.path.exists(ps_export):
        print(f"❌ 错误: 找不到脚本 {ps_export}")
        print("无法获取AD现有用户，同步失败！")
        sys.exit(1)
    
    try:
        # 上传脚本
        scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {ps_export} {DC_USER}@{DC_HOST}:~/ExportUsers.ps1"
        run_scp_with_retry(scp_cmd)
        
        # 执行导出
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -File {DC_USER_HOME}/ExportUsers.ps1'"
        run_ssh_with_retry(ssh_cmd)
        
        # 下载导出的用户列表
        scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST}:~/ExportedUsers.csv {get_output_path('ad_existing_users.csv')}"
        result = run_scp_with_retry(scp_cmd)
        
        if result.returncode != 0:
            # 文件下载失败，可能是AD上没有用户，先验证数量
            print("⚠ 下载AD用户文件失败，验证AD上是否真的没有用户...")
            try:
                actual_count = get_ad_user_count()
                if actual_count == 0:
                    print("✓ AD上确实没有用户，继续执行")
                    return {}, {}
                else:
                    print(f"❌ AD上有 {actual_count} 个用户，但文件下载失败")
                    print("数据获取失败，同步失败！")
                    sys.exit(1)
            except Exception as e:
                print(f"❌ 无法验证AD用户数量: {e}")
                print("无法确认数据状态，同步失败！")
                sys.exit(1)
    except Exception as e:
        print(f"❌ 连接域控制器失败: {e}")
        print("无法获取AD现有用户，同步失败！")
        sys.exit(1)
    
    # 读取现有用户的信息，使用 EmployeeNumber (Union ID) 作为键，没有则用 SamAccountName
    existing_users = {}
    users_without_union_id = {}  # 没有 Union ID 的用户，用 SamAccountName 作为键
    
    try:
        # PowerShell 使用 UTF8 导出，直接使用 utf-8-sig
        with open(get_output_path('ad_existing_users.csv'), 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                sam = row['SamAccountName']
                employee_number = row.get('EmployeeNumber', '').strip()
                user_info = {
                    'SamAccountName': sam,
                    'DisplayName': row.get('DisplayName', ''),
                    'EmailAddress': row.get('EmailAddress', '')
                }
                
                if employee_number:
                    # 有 Union ID，用 Union ID 作为键
                    existing_users[employee_number] = user_info
                else:
                    # 没有 Union ID，用 SamAccountName 作为键
                    users_without_union_id[sam] = user_info
        
        total_users = len(existing_users) + len(users_without_union_id)
        print(f"✓ 发现 {total_users} 个现有用户（{len(existing_users)} 个有 Union ID，{len(users_without_union_id)} 个无 Union ID）")
    except Exception as e:
        print(f"读取现有用户列表失败: {e}")
        # 如果失败，尝试其他编码
        for enc in ['utf-8-sig', 'gb2312', 'utf-16']:
            try:
                with open(get_output_path('ad_existing_users.csv'), 'r', encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        sam = row['SamAccountName']
                        employee_number = row.get('EmployeeNumber', '').strip()
                        user_info = {
                            'SamAccountName': sam,
                            'DisplayName': row.get('DisplayName', ''),
                            'EmailAddress': row.get('EmailAddress', '')
                        }
                        
                        if employee_number:
                            existing_users[employee_number] = user_info
                        else:
                            users_without_union_id[sam] = user_info
                
                total_users = len(existing_users) + len(users_without_union_id)
                print(f"✓ 使用 {enc} 编码读取成功，发现 {total_users} 个用户（{len(existing_users)} 个有 Union ID，{len(users_without_union_id)} 个无 Union ID）")
                break
            except:
                continue
        else:
            print(f"❌ 所有编码尝试均失败")
            sys.exit(1)
    
    # 验证数据完整性：从AD实时获取数量并对比
    print("正在验证用户数据完整性...")
    try:
        actual_count = get_ad_user_count()
        csv_count = len(existing_users) + len(users_without_union_id)
        print(f"  AD实际用户数: {actual_count}")
        print(f"  CSV文件用户数: {csv_count}")
        
        if actual_count != csv_count:
            print(f"❌ 用户数量不匹配！实际: {actual_count}, CSV: {csv_count}")
            print("数据可能不完整，同步失败！")
            sys.exit(1)
        print("✓ 用户数据验证通过")
    except Exception as e:
        print(f"❌ 验证用户数据失败: {e}")
        print("无法确认数据完整性，同步失败！")
        sys.exit(1)
    
    return existing_users, users_without_union_id

def split_users_for_sync(feishu_csv, existing_users, users_without_union_id):
    """将飞书用户分为新建和更新两类，处理拼音重名"""
    new_users = []
    update_users = []
    matched_ad_users = set()  # 记录匹配到的 AD 用户（Union ID）
    matched_ad_users_no_uid = set()  # 记录匹配到的 AD 用户（无 Union ID，用 SamAccountName）
    
    # 读取部门CSV，构建部门路径映射
    dept_path_map = {}  # {dept_id: "父部门\\子部门"}
    dept_csv = get_output_path('feishu_departments.csv')
    
    if os.path.exists(dept_csv):
        with open(dept_csv, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            dept_list = list(reader)
        
        # 递归构建部门路径
        def build_dept_path(dept_id, dept_list):
            if dept_id == "0":
                return ""
            
            dept = next((d for d in dept_list if d['dept_id'] == dept_id), None)
            if not dept:
                return ""
            
            parent_id = dept['parent_dept_id']
            dept_name = dept['dept_name']
            
            if parent_id == "0":
                return dept_name
            else:
                parent_path = build_dept_path(parent_id, dept_list)
                return f"{parent_path}\\{dept_name}" if parent_path else dept_name
        
        # 为所有部门构建路径
        for dept in dept_list:
            dept_id = dept['dept_id']
            dept_path_map[dept_id] = build_dept_path(dept_id, dept_list)
    
    # 读取所有飞书用户
    feishu_users_list = []
    with open(feishu_csv, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        feishu_users_list = list(reader)
    
    # 检测拼音重名
    pinyin_count = {}
    for user in feishu_users_list:
        pinyin = user['拼音']
        if pinyin not in pinyin_count:
            pinyin_count[pinyin] = []
        pinyin_count[pinyin].append(user)
    
    # 为每个用户确定 SamAccountName
    for row in feishu_users_list:
        pinyin = row['拼音']
        user_id = row['用户ID']  # 用户ID
        union_id = row.get('Union ID', '')  # Union ID
        user_uuid = row.get('UUID', '')  # UUID
        employee_no = row['工号']  # 工号
        display_name = row['姓名']
        email = row['企业邮箱']
        dept_id = row.get('dept_id', '')
        
        # 从部门ID获取完整部门路径
        dept_path = dept_path_map.get(dept_id, '') if dept_id else ''
        
        # 优先检查拼音例外映射
        if pinyin in PINYIN_EXCEPTIONS:
            sam_account = PINYIN_EXCEPTIONS[pinyin]
        # 确定 SamAccountName：拼音重名时，按工号排序
        elif len(pinyin_count[pinyin]) > 1:
            # 拼音重名，按工号排序，工号最小的用拼音，其他加序号
            sorted_users = sorted(pinyin_count[pinyin], key=lambda x: x['工号'])
            if employee_no == sorted_users[0]['工号']:
                sam_account = pinyin  # 工号最小，用拼音
            else:
                # 工号较大，在名字后、姓氏前加序号（从1开始）
                # 例如：jiayi.wang -> jiayi1.wang
                index = next(i for i, u in enumerate(sorted_users) if u['工号'] == employee_no)
                parts = pinyin.split('.')
                if len(parts) == 2:
                    sam_account = f"{parts[0]}{index}.{parts[1]}"
                else:
                    sam_account = f"{pinyin}{index}"
        else:
            sam_account = pinyin  # 拼音不重名，用拼音
        
        # 匹配逻辑：优先通过 Union ID，其次通过 SamAccountName
        matched_sam = None
        
        # 1. 通过 Union ID 查找现有用户
        if union_id and union_id in existing_users:
            ad_info = existing_users[union_id]
            matched_sam = ad_info['SamAccountName']
            matched_ad_users.add(union_id)  # 记录匹配的 AD 用户（使用 Union ID）
        # 2. 如果没有 Union ID 匹配，尝试通过 SamAccountName 匹配（针对旧用户）
        elif sam_account in users_without_union_id:
            ad_info = users_without_union_id[sam_account]
            matched_sam = sam_account
            matched_ad_users_no_uid.add(sam_account)  # 记录匹配的旧用户
        
        if matched_sam:
            # 用户已存在，更新用户（EmployeeID 使用工号，EmployeeNumber 使用 Union ID）
            update_users.append({
                'SamAccountName': matched_sam,
                'DisplayName': display_name,
                'EmailAddress': email,
                'EmployeeID': employee_no,
                'EmployeeNumber': union_id,
                'info': user_uuid,
                'DepartmentName': dept_path
            })
        else:
            # 新建用户（EmployeeID 使用工号，EmployeeNumber 使用 Union ID）
            new_users.append({
                'DisplayName': display_name,
                'SamAccountName': sam_account,
                'EmailAddress': email,
                'EmployeeID': employee_no,
                'EmployeeNumber': union_id,
                'info': user_uuid,
                'DepartmentName': dept_path
            })
    
    return new_users, update_users, matched_ad_users, matched_ad_users_no_uid

def create_csv_files(new_users, update_users):
    """生成新建和更新的CSV文件"""
    # 新建用户CSV
    if new_users:
        with open(get_output_path('ad_new_accounts.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=['DisplayName', 'SamAccountName', 'EmailAddress', 'EmployeeID', 'EmployeeNumber', 'info', 'DepartmentName'])
            writer.writeheader()
            writer.writerows(new_users)
        print(f"  - 待创建用户: {len(new_users)} 个")
    
    # 更新用户CSV
    if update_users:
        with open(get_output_path('ad_update_accounts.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=['SamAccountName', 'DisplayName', 'EmailAddress', 'EmployeeID', 'EmployeeNumber', 'info', 'DepartmentName'])
            writer.writeheader()
            writer.writerows(update_users)
        print(f"  - 待检查用户: {len(update_users)} 个（实际更新数量取决于字段差异）")
    
    return len(new_users), len(update_users)

def upload_dc_config():
    """上传配置到域控制器"""
    config_content = f"""# AD Domain Configuration
$DC_BASE_OU = "{DC_BASE_OU}"
$DC_DOMAIN = "{DC_DOMAIN}"
$DC_EXCLUDE_OU = "{DC_EXCLUDE_OU}"
$DC_RESIGNED_OU = "{DC_RESIGNED_OU}"
"""
    
    # 写入本地临时文件，使用 Windows 换行符
    config_file = get_output_path('dc_config.ps1')
    with open(config_file, 'w', encoding='utf-8-sig', newline='\r\n') as f:
        f.write(config_content)
    
    # 上传到域控制器（上传到远程用户主目录）
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {config_file} {DC_USER}@{DC_HOST}:~/dc_config.ps1"
    run_scp_with_retry(scp_cmd)
    
    print("✓ 已上传配置到域控制器")

def execute_on_dc(operation, csv_file, ps_script, use_local=False):
    """在域控制器上执行操作，返回实际处理的用户数"""
    if not os.path.exists(csv_file):
        return 0
    
    # 上传CSV到用户主目录
    remote_csv = f"~/{os.path.basename(csv_file)}"
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {csv_file} {DC_USER}@{DC_HOST}:{remote_csv}"
    
    # 重试上传
    for attempt in range(3):
        result = run_scp_with_retry(scp_cmd)
        if result.returncode == 0:
            break
        if attempt < 2:
            print(f"上传失败，{3-attempt} 秒后重试...")
            import time
            time.sleep(3)
        else:
            print(f"✗ 上传 CSV 失败: {result.stderr.decode('utf-8', errors='ignore')}")
            return 0
    
    # 上传PowerShell脚本
    if use_local:
        ps_local = ps_script
    else:
        ps_local = f"../domainusers/{ps_script}"
    
    actual_count = 0
    if os.path.exists(ps_local):
        remote_ps = f"~/{operation}.ps1"
        scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {ps_local} {DC_USER}@{DC_HOST}:{remote_ps}"
        
        # 重试上传
        for attempt in range(3):
            result = run_scp_with_retry(scp_cmd)
            if result.returncode == 0:
                break
            if attempt < 2:
                print(f"上传脚本失败，{3-attempt} 秒后重试...")
                import time
                time.sleep(3)
            else:
                print(f"✗ 上传脚本失败: {result.stderr.decode('utf-8', errors='ignore')}")
                return 0
        
        # 执行脚本，dry-run 模式传递 -DryRun 参数
        print(f"正在执行...")
        dry_run_flag = "-DryRun" if DRY_RUN else ""
        # PowerShell 脚本路径
        script_path = f"{DC_USER_HOME}/{operation}.ps1"
        # 使用 PowerShell 命令包装，重定向信息流
        ps_command = f"& '{script_path}' {dry_run_flag} *>&1"
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -Command \"{ps_command}\"'"
        result = run_ssh_with_retry(ssh_cmd)
        
        # 尝试解码输出（PowerShell 已设置 UTF-8 输出）
        try:
            output = result.stdout.decode('utf-8', errors='replace')
            if output.strip():
                print(output)
            else:
                print("（脚本无输出）")
            
            # 提取实际处理的用户数
            if DRY_RUN:
                match = re.search(r'将(?:更新|创建): (\d+) 个用户', output)
            else:
                match = re.search(r'成功: (\d+) 个用户', output)
            if match:
                actual_count = int(match.group(1))
        except Exception as e:
            print(f"脚本执行完成（输出解码失败: {e}）")
        
        if result.stderr:
            try:
                error = result.stderr.decode(encoding)
                print("错误:", error)
            except:
                pass
    
    return actual_count

def download_passwords():
    """下载生成的密码文件并发送邮件"""
    print("\n正在下载生成的密码...")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    password_file = get_output_path(f"ad_passwords_{timestamp}.csv")
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST}:~/GeneratedPasswords.csv {password_file}"
    result = run_scp_with_retry(scp_cmd)
    
    if result.returncode == 0:
        print(f"✓ 密码文件已保存到: {password_file}")
        
        # 自动发送密码邮件
        send_password_emails(password_file)
    else:
        print("未生成新密码（可能没有新建用户）")

def send_password_emails(password_file):
    """读取密码文件并发送邮件"""
    from send_password_email import send_password_email
    
    # 从环境变量获取公司名称
    company_name = os.getenv("FEISHU_COMPANY_NAME", "公司")
    
    print("\n正在发送密码邮件...")
    
    # 读取密码文件
    try:
        encoding = get_windows_encoding()
        with open(password_file, 'r', encoding=encoding) as f:
            reader = csv.DictReader(f)
            passwords = list(reader)
    except:
        # 尝试其他编码
        with open(password_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            passwords = list(reader)
    
    success_count = 0
    fail_count = 0
    
    for row in passwords:
        sam_account = row['SamAccountName']
        display_name = row['DisplayName']
        email = row['EmailAddress']
        password = row['Password']
        department = row.get('Department', '')
        
        # 如果部门为空，使用公司名称
        if not department or department.strip() == '':
            department = company_name
        
        if email and password != '[DRY-RUN]':
            try:
                success, message = send_password_email(
                    receiver_email=email,
                    new_password=password,
                    sam_account=sam_account,
                    display_name=display_name,
                    department=department
                )
                
                if success:
                    success_count += 1
                else:
                    print(f"✗ 发送失败: {sam_account} - {message}")
                    fail_count += 1
            except Exception as e:
                print(f"✗ 发送异常: {sam_account} - {e}")
                fail_count += 1
        else:
            if not email:
                print(f"⚠ 跳过: {sam_account} (无邮箱)")
            fail_count += 1
    
    print(f"\n邮件发送完成: 成功 {success_count} 个, 失败 {fail_count} 个")

def process_unmatched_users(unmatched_users):
    """处理未匹配的用户：禁用并移动到离职员工 OU"""
    print(f"\n正在处理 {len(unmatched_users)} 个未匹配用户...")
    
    # 生成用户列表 CSV
    with open(get_output_path('ad_resign_users.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['SamAccountName', 'DisplayName'])
        writer.writeheader()
        for user in unmatched_users:
            writer.writerow({
                'SamAccountName': user['SamAccountName'],
                'DisplayName': user['DisplayName']
            })
    
    # 调用 PowerShell 脚本处理
    return execute_on_dc('ResignUsers', get_output_path('ad_resign_users.csv'), get_ps_path('resign_users.ps1'), use_local=True)

def delete_extra_ous(extra_ous):
    """删除多余的 OU"""
    print(f"\n正在删除 {len(extra_ous)} 个多余的 OU...")
    
    # 生成 OU 列表 CSV
    with open(get_output_path('ad_delete_ous.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['OU名称'])
        for ou in extra_ous:
            writer.writerow([ou])
    
    # 调用 PowerShell 脚本删除
    execute_on_dc('DeleteOUs', get_output_path('ad_delete_ous.csv'), get_ps_path('delete_ou.ps1'), use_local=True)

def sync_departments():
    """同步部门OU结构到AD域"""
    dept_csv = get_output_path('feishu_departments.csv')
    if not os.path.exists(dept_csv):
        print(f"错误: 找不到 {dept_csv}，请先运行 fetch_feishu_data.py")
        return
    
    print("正在同步部门结构...")
    
    # 读取飞书部门列表
    feishu_depts = set()
    with open(dept_csv, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            feishu_depts.add(row['dept_name'])
    
    # 获取 AD 现有部门（返回 {DN: 名称}）
    ad_depts_dict = get_existing_ad_departments()
    ad_dept_names = set(ad_depts_dict.values())  # 提取所有 OU 名称
    
    # 找出 AD 中多余的 OU（AD 有但飞书没有）
    extra_ous = []
    for ou_name in ad_dept_names:
        if ou_name not in feishu_depts:
            extra_ous.append(ou_name)
    
    if extra_ous:
        print(f"\n⚠ 发现 {len(extra_ous)} 个 AD 中存在但飞书中不存在的 OU")
        # 导出多余的 OU 列表
        with open(get_output_path('ad_extra_ous.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['OU名称'])
            for ou in extra_ous:
                writer.writerow([ou])
        print(f"  - 多余 OU 列表已保存到: output/ad_extra_ous.csv")
        
        # 询问是否删除
        if not DRY_RUN and confirm(f"是否删除这 {len(extra_ous)} 个多余的 OU？", default=False):
            delete_extra_ous(extra_ous)
    
    # 上传部门CSV
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {dept_csv} {DC_USER}@{DC_HOST}:~/feishu_departments.csv"
    run_scp_with_retry(scp_cmd)
    
    # 上传PowerShell脚本
    ps_script = get_ps_path('create_ou.ps1')
    if not os.path.exists(ps_script):
        print(f"❌ 错误: 找不到 {ps_script}")
        print("无法同步部门结构，同步失败！")
        sys.exit(1)
    
    scp_cmd = f"sshpass -p '{DC_PASSWORD}' scp -o ControlPath={SSH_CONTROL_PATH} {ps_script} {DC_USER}@{DC_HOST}:~/CreateOU.ps1"
    run_scp_with_retry(scp_cmd)
    
    # 执行脚本，dry-run 模式传递 -DryRun 参数
    dry_run_flag = "-DryRun" if DRY_RUN else ""
    ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -ExecutionPolicy Bypass -File {DC_USER_HOME}/CreateOU.ps1 {dry_run_flag}'"
    result = run_ssh_with_retry(ssh_cmd)
    
    # 尝试解码输出（PowerShell 已设置 UTF-8 输出）
    try:
        output = result.stdout.decode('utf-8', errors='replace')
        print(output)
    except:
        print(f"（部分输出无法显示）")
    
    if result.stderr:
        try:
            error = result.stderr.decode('utf-8', errors='replace')
            print("错误:", error)
        except:
            print("有错误输出（解码失败）")
    
    print("✓ 部门CSV已上传到域控制器，可供后续脚本使用")

def cleanup_remote_files():
    """清理远程服务器上的临时文件"""
    remote_files = [
        'GetOUs.ps1',
        'ExportUsers.ps1',
        'CreateOU.ps1',
        'CreateAccounts.ps1',
        'UpdateAccounts.ps1',
        'dc_config.ps1',
        'ExistingOUs.csv',
        'ExportedUsers.csv',
        'feishu_departments.csv',
        'feishu_users_new.csv',
        'feishu_users_update.csv',
        'ad_new_accounts.csv',
        'ad_update_accounts.csv',
        'GeneratedPasswords.csv'
    ]
    
    try:
        # 批量删除所有文件（单次SSH调用）
        files_str = ','.join([f"~/{f}" for f in remote_files])
        ssh_cmd = f"sshpass -p '{DC_PASSWORD}' ssh -o StrictHostKeyChecking=no -o ControlPath={SSH_CONTROL_PATH} {DC_USER}@{DC_HOST} 'powershell -Command \"Remove-Item -Path {files_str} -ErrorAction SilentlyContinue\"'"
        run_ssh_with_retry(ssh_cmd, timeout=10)
        print("✓ 远程临时文件已清理")
    except:
        print("⚠ 清理远程文件时出现问题（可忽略）")

if __name__ == "__main__":
    # 清理旧的 SSH 控制套接字
    import glob
    for socket_file in glob.glob("/tmp/ssh-feishu-ad-sync-*"):
        try:
            os.remove(socket_file)
        except:
            pass
    
    # 检查命令行参数
    for arg in sys.argv[1:]:
        if arg == '--dry-run':
            DRY_RUN = True
        elif arg == '--yes' or arg == '-y':
            AUTO_YES = True
    
    if DRY_RUN:
        print("=" * 50)
        print("  DRY-RUN 模式 - 连接 AD 预览同步计划")
        print("=" * 50)
    else:
        print("=" * 50)
        print("  飞书用户增量同步到 AD 域")
        if AUTO_YES:
            print("  自动确认模式 - 跳过所有确认步骤")
        print("=" * 50)
    
    # 显示执行步骤概览
    print("\n执行步骤:")
    print("  1/6 - 同步飞书部门结构到 AD 域")
    print("  2/6 - 获取 AD 域现有用户")
    print("  3/6 - 分析飞书用户")
    print("  4/6 - 创建新用户、下载密码并发送邮件")
    print("  5/6 - 更新 AD 域现有用户")
    print("  6/6 - 完成同步")
    print("")
    
    # 0. 检查飞书数据文件，如果不存在或使用-y参数则自动获取
    feishu_users_csv = get_output_path('feishu_users.csv')
    feishu_depts_csv = get_output_path('feishu_departments.csv')
    fetch_script = os.path.join(SCRIPT_DIR, 'fetch_feishu_data.py')
    
    if AUTO_YES or not os.path.exists(feishu_users_csv) or not os.path.exists(feishu_depts_csv):
        print("\n【步骤 0/6】获取飞书数据")
        if AUTO_YES:
            print("自动确认模式：强制重新获取飞书数据...")
        else:
            print("未找到飞书数据文件，正在从飞书获取...")
        result = subprocess.run(['python3', fetch_script], cwd=SCRIPT_DIR)
        if result.returncode != 0:
            print("错误: 获取飞书数据失败")
            sys.exit(1)
        print("✓ 飞书数据获取完成")
    
    # 检查域控制器权限
    if not check_dc_permissions():
        print("\n错误: 域控制器权限检查失败")
        print("请确保:")
        print("  1. 域控制器地址、用户名、密码配置正确")
        print("  2. 用户具有 Active Directory 管理权限")
        print("  3. 域控制器已安装 Active Directory PowerShell 模块")
        sys.exit(1)
    
    # 上传配置到域控制器
    print("\n正在上传配置到域控制器...")
    upload_dc_config()
    
    # 1. 同步部门OU结构
    print("\n【步骤 1/6】同步飞书部门结构到 AD 域")
    sync_departments()
    
    # 2. 获取现有AD用户
    print("\n【步骤 2/6】获取 AD 域现有用户")
    if DRY_RUN:
        print("[DRY-RUN] 正在获取现有用户...")
    existing_users, users_without_union_id = get_existing_ad_users()
    
    # 3. 分类飞书用户
    print("\n【步骤 3/6】分析飞书用户")
    if DRY_RUN:
        print("[DRY-RUN] 分析飞书用户并生成同步计划")
    
    new_users, update_users, matched_ad_users, matched_ad_users_no_uid = split_users_for_sync(feishu_users_csv, existing_users, users_without_union_id)
    new_count, update_count = create_csv_files(new_users, update_users)
    
    # 显示未匹配的 AD 用户数量
    total_ad_users = len(existing_users) + len(users_without_union_id)
    total_matched = len(matched_ad_users) + len(matched_ad_users_no_uid)
    unmatched_ad_count = total_ad_users - total_matched
    
    if unmatched_ad_count > 0:
        print(f"  ⚠ {unmatched_ad_count} 个 AD 用户在飞书中未找到匹配")
        
        # 导出未匹配的 AD 用户（包括有 Union ID 和没有 Union ID 的）
        unmatched_users = []
        
        # 未匹配的有 Union ID 的用户
        for union_id, info in existing_users.items():
            if union_id not in matched_ad_users:
                unmatched_users.append({
                    'SamAccountName': info['SamAccountName'],
                    'DisplayName': info['DisplayName'],
                    'EmailAddress': info['EmailAddress']
                })
        
        # 未匹配的没有 Union ID 的用户
        for sam, info in users_without_union_id.items():
            if sam not in matched_ad_users_no_uid:
                unmatched_users.append({
                    'SamAccountName': info['SamAccountName'],
                    'DisplayName': info['DisplayName'],
                    'EmailAddress': info['EmailAddress']
                })
        
        with open(get_output_path('ad_unmatched_users.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=['SamAccountName', 'DisplayName', 'EmailAddress'])
            writer.writeheader()
            writer.writerows(unmatched_users)
        print(f"  - 未匹配用户列表已保存到: output/ad_unmatched_users.csv")
        
        # 询问是否处理未匹配用户
        if not DRY_RUN and DC_RESIGNED_OU:
            if confirm(f"是否将这 {unmatched_ad_count} 个用户禁用并移动到离职员工 OU？", default=False):
                actual_resign_count = process_unmatched_users(unmatched_users)
            else:
                actual_resign_count = 0
        else:
            actual_resign_count = 0
    else:
        actual_resign_count = 0
    
    if new_count == 0 and update_count == 0:
        print("\n没有需要同步的用户")
        sys.exit(0)
    
    # 4. 执行新建
    actual_new_count = 0
    if new_count > 0:
        print(f"\n【步骤 4/6】创建新用户、下载密码并发送邮件 ({new_count} 个)")
        actual_new_count = execute_on_dc('CreateAccounts', get_output_path('ad_new_accounts.csv'), get_ps_path('create_users.ps1'), use_local=True)
        # 创建成功后立即下载密码
        if actual_new_count > 0 and not DRY_RUN:
            download_passwords()
    
    # 5. 执行更新
    actual_update_count = 0
    if update_count > 0:
        print(f"\n【步骤 5/6】检查并更新 AD 域现有用户 ({update_count} 个)")
        actual_update_count = execute_on_dc('UpdateAccounts', get_output_path('ad_update_accounts.csv'), get_ps_path('update_users.ps1'), use_local=True)
    
    print("\n" + "=" * 50)
    if DRY_RUN:
        print("  DRY-RUN 完成 - 未执行实际操作")
        print("  以上为基于 AD 域真实数据的同步预览")
    else:
        print("  【步骤 6/6】同步完成")
    print("=" * 50)
    
    # 显示实际处理的数量
    if actual_new_count > 0 or actual_update_count > 0:
        print(f"新建: {actual_new_count} 个用户")
        print(f"更新: {actual_update_count} 个用户")
    else:
        print("所有用户信息已是最新，无需更新")
    
    # 清理远程临时文件
    print("\n清理远程临时文件...")
    cleanup_remote_files()
    
    # 清理SSH连接
    cleanup_ssh_control_master()
    
    # 发送青龙通知（如果在青龙环境中运行）
    try:
        if not DRY_RUN and (actual_new_count > 0 or actual_update_count > 0 or actual_resign_count > 0):
            notify_content = f"新建: {actual_new_count}\n更新: {actual_update_count}\n禁用: {actual_resign_count}"
            QLAPI.systemNotify({
                "title": "飞书用户同步AD域完成",
                "content": notify_content
            })
            print("✓ 青龙系统通知已发送")
    except (NameError, AttributeError):
        # 不在青龙环境中，跳过通知
        print("⚠ 未检测到青龙环境，跳过系统通知")
    except Exception as e:
        # 其他错误（如网络问题等）
        print(f"⚠ 系统通知发送失败: {e}")
