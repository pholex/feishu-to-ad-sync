#!/usr/bin/env python3
import requests
import json
import csv
import os
import time
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from dotenv import load_dotenv
from pypinyin import lazy_pinyin, Style

# 加载环境变量（使用脚本所在目录的.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(script_dir, '.env'))

APP_ID = os.getenv("FEISHU_APP_ID")
APP_SECRET = os.getenv("FEISHU_APP_SECRET")
FEISHU_COMPANY_NAME = os.getenv("FEISHU_COMPANY_NAME", "公司")  # 默认值为"公司"

# 配置重试
MAX_RETRIES = 3
RETRY_DELAY = 5  # 秒

# 姓氏多音字映射（作为姓氏时的正确读音）
SURNAME_PINYIN = {
    '曾': 'zeng',
    '卜': 'bu',
    '区': 'ou',
    '查': 'zha',
    '单': 'shan',
    '朴': 'piao',
    '仇': 'qiu',
    '黑': 'he',
    '盖': 'ge',
    '种': 'chong',
    '繁': 'po',
    '召': 'shao',
}

def name_to_pinyin(name):
    """将中文姓名转换为拼音（名.姓格式）"""
    if not name:
        return ""
    
    # 先替换多音字为姓氏读音
    for char, pinyin in SURNAME_PINYIN.items():
        name = name.replace(char, f'[{pinyin}]')
    
    # 转换为拼音
    pinyin_list = lazy_pinyin(name, style=Style.NORMAL)
    
    # 还原替换的拼音
    pinyin_list = [p.strip('[]') for p in pinyin_list]
    
    if len(pinyin_list) == 0:
        return ""
    elif len(pinyin_list) == 1:
        return pinyin_list[0].lower()
    else:
        # 姓是第一个字，名是后面的字
        surname = pinyin_list[0].lower()
        given_name = "".join(pinyin_list[1:]).lower()
        return f"{given_name}.{surname}"

def generate_uuid_from_email(email):
    """根据邮箱生成固定的 UUID"""
    if not email:
        return ""
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, email.lower()))

def request_with_retry(method, url, **kwargs):
    """带重试的请求"""
    for i in range(MAX_RETRIES):
        try:
            if method == "GET":
                response = requests.get(url, timeout=30, **kwargs)
            else:
                response = requests.post(url, timeout=30, **kwargs)
            return response
        except (requests.exceptions.ConnectionError, 
                requests.exceptions.Timeout) as e:
            if i < MAX_RETRIES - 1:
                print(f"请求失败，{RETRY_DELAY}秒后重试... ({i+1}/{MAX_RETRIES})")
                time.sleep(RETRY_DELAY)
            else:
                print(f"请求失败，已重试{MAX_RETRIES}次: {e}")
                raise

def get_tenant_access_token():
    """获取 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    data = {"app_id": APP_ID, "app_secret": APP_SECRET}
    response = request_with_retry("POST", url, headers=headers, json=data)
    return response.json()["tenant_access_token"]

def get_total_user_count(token):
    """获取企业总人数"""
    url = "https://open.feishu.cn/open-apis/contact/v3/departments/0"
    headers = {"Authorization": f"Bearer {token}"}
    params = {"department_id_type": "open_department_id"}
    
    try:
        response = request_with_retry("GET", url, headers=headers, params=params)
        data = response.json()
        if data["code"] == 0:
            return data.get("data", {}).get("department", {}).get("member_count", 0)
    except Exception as e:
        print(f"获取企业总人数异常: {e}")
    
    return None

def get_users_by_department(token, department_id, retry_counter=None):
    """获取指定部门的用户"""
    url = "https://open.feishu.cn/open-apis/contact/v3/users/find_by_department"
    headers = {"Authorization": f"Bearer {token}"}
    users = []
    page_token = None
    
    while True:
        params = {
            "department_id": department_id,
            "page_size": 50,
            "department_id_type": "open_department_id",
            "user_id_type": "user_id"
        }
        if page_token:
            params["page_token"] = page_token
        
        # 对限流错误进行重试
        for retry in range(3):
            response = request_with_retry("GET", url, headers=headers, params=params)
            data = response.json()
            
            if data["code"] == 0:
                break
            elif data["code"] in [99991400, 2200]:  # 频率限制或服务不可用
                if retry_counter is not None:
                    retry_counter['count'] += 1
                if retry < 2:
                    time.sleep(1)  # 等待1秒后重试
                    continue
                else:
                    # 重试3次仍失败
                    raise Exception(f"部门{department_id}用户获取失败: code={data['code']}, 重试3次均失败")
            else:
                # 其他错误
                raise Exception(f"部门{department_id}用户获取失败: code={data['code']}, msg={data.get('msg', 'unknown')}")
        
        if data["code"] != 0:
            raise Exception(f"部门{department_id}用户获取失败: code={data['code']}")
        
        items = data.get("data", {}).get("items", [])
        users.extend(items)
        
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")
    
    return users

def get_all_users_sequential(token, dept_ids_list):
    """获取所有用户信息（顺序，用于对比）"""
    dept_ids = ["0"] + dept_ids_list
    all_users = []
    seen_ids = set()
    
    start_time = time.time()
    for i, dept_id in enumerate(dept_ids, 1):
        users = get_users_by_department(token, dept_id)
        for user in users:
            user_id = user.get("user_id")
            if user_id and user_id not in seen_ids:
                all_users.append(user)
                seen_ids.add(user_id)
        print(f"\r  进度: {i}/{len(dept_ids)}, 已获取: {len(all_users)} 个用户", end='', flush=True)
    
    print()
    elapsed_time = time.time() - start_time
    print(f"共获取 {len(all_users)} 个用户，耗时 {elapsed_time:.2f} 秒")
    return all_users

def get_all_users(token, dept_ids_list):
    """获取所有用户信息（并发）"""
    # 包含根部门"0"，确保获取所有用户
    dept_ids = ["0"] + dept_ids_list
    
    start_time = time.time()
    
    # 第一阶段：并发获取所有部门的用户数据
    all_dept_users = []  # 存储每个部门的用户列表
    lock = Lock()
    processed_count = 0
    first_result = True
    
    print(f"\r正在从 {len(dept_ids_list)} 个部门获取用户（含根部门共 {len(dept_ids)} 个）... (启动中 0.0s)", end='', flush=True)
    stop_timer = threading.Event()
    retry_counter = {'count': 0}  # 限流重试计数
    
    def show_startup_time():
        while not stop_timer.is_set():
            elapsed = time.time() - start_time
            print(f"\r正在从 {len(dept_ids_list)} 个部门获取用户（含根部门共 {len(dept_ids)} 个）... (启动中 {elapsed:.1f}s)", end='', flush=True)
            time.sleep(0.1)
    
    timer_thread = threading.Thread(target=show_startup_time, daemon=True)
    timer_thread.start()
    
    with ThreadPoolExecutor(max_workers=20) as executor:
        future_to_dept = {executor.submit(get_users_by_department, token, dept_id, retry_counter): dept_id 
                         for dept_id in dept_ids}
        
        for future in as_completed(future_to_dept):
            if first_result:
                stop_timer.set()
                timer_thread.join(timeout=0.2)
                startup_time = time.time() - start_time
                print(f"\r正在从 {len(dept_ids_list)} 个部门获取用户（含根部门共 {len(dept_ids)} 个）... (启动完成 {startup_time:.2f}s)")
                first_result = False
            
            dept_id = future_to_dept[future]
            try:
                users = future.result()
                with lock:
                    all_dept_users.append(users)  # 收集每个部门的用户列表
                    processed_count += 1
                    print(f"\r  进度: {processed_count}/{len(dept_ids)}", end='', flush=True)
            except Exception as e:
                with lock:
                    processed_count += 1
                    print(f"\n⚠ 错误: 获取部门 {dept_id} 的用户失败: {e}")
                    print(f"\r  进度: {processed_count}/{len(dept_ids)} ⚠", end='', flush=True)
    
    print()
    
    # 第二阶段：去重处理（单线程，无竞争）
    all_users = []
    seen_ids = set()
    total_fetched = 0
    skipped_count = 0
    
    for users in all_dept_users:
        for user in users:
            total_fetched += 1
            user_id = user.get("user_id")
            if user_id:
                if user_id not in seen_ids:
                    all_users.append(user)
                    seen_ids.add(user_id)
            else:
                skipped_count += 1
                print(f"⚠ 警告: 发现没有user_id的用户: {user}")
    
    elapsed_time = time.time() - start_time
    duplicate_count = total_fetched - len(all_users) - skipped_count
    print(f"共获取 {len(all_users)} 个用户，耗时 {elapsed_time:.2f} 秒")
    if retry_counter['count'] > 0:
        print(f"  限流重试次数: {retry_counter['count']}")
    if duplicate_count > 0:
        print(f"  (其中 {duplicate_count} 个用户属于多个部门)")
    
    # 数据验证：统计所有部门的member_count总和
    return all_users, total_fetched

def export_to_csv(users, dept_map):
    """导出到 CSV 文件"""
    # 获取项目根目录（脚本所在目录）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'output')
    
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, 'feishu_users.csv')
    
    # 统计每个部门的用户数（用于验证）
    dept_user_count = {}
    
    with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["user_id", "open_id", "union_id", "uuid", "name", "pinyin", "enterprise_email", "mobile", "employee_no", "job_title", "status", "dept_id", "dept_name", "department_ids", "department_names"])
        
        for user in users:
            # 排除冻结和离职用户
            status_obj = user.get("status", {})
            if isinstance(status_obj, dict):
                is_frozen = status_obj.get("is_frozen", False)
                is_resigned = status_obj.get("is_resigned", False)
                if is_frozen or is_resigned:
                    continue
            
            user_id = user.get("user_id", "")
            open_id = user.get("open_id", "")
            union_id = user.get("union_id", "")
            name = user.get("name", "")
            pinyin = name_to_pinyin(name)
            enterprise_email = user.get("enterprise_email", "")
            user_uuid = generate_uuid_from_email(enterprise_email)
            mobile = user.get("mobile", "")
            employee_no = user.get("employee_no", "")
            job_title = user.get("job_title", "")
            
            # 获取用户状态信息
            is_activated = status_obj.get("is_activated", "")
            is_frozen = status_obj.get("is_frozen", "")
            is_resigned = status_obj.get("is_resigned", "")
            status = f"激活:{is_activated}|冻结:{is_frozen}|离职:{is_resigned}"
            
            # 获取部门信息（用户可能属于多个部门）
            department_ids = user.get("department_ids", [])
            dept_id = department_ids[0] if department_ids else ""
            dept_name = user.get("department_name", "")
            
            # 记录所有部门
            all_dept_ids = ";".join(department_ids)
            all_dept_names = ";".join([dept_map.get(did, "") for did in department_ids])
            
            # 统计每个部门的用户数
            for did in department_ids:
                dept_user_count[did] = dept_user_count.get(did, 0) + 1
            
            writer.writerow([user_id, open_id, union_id, user_uuid, name, pinyin, enterprise_email, mobile, employee_no, job_title, status, dept_id, dept_name, all_dept_ids, all_dept_names])
    
    print(f"已导出 {len(users)} 个用户到 {output_file}")
    
    # 返回统计结果用于验证
    return dept_user_count

def get_department_info(token):
    """获取所有部门信息，建立ID到名称的映射，并保存部门层级结构"""
    url = "https://open.feishu.cn/open-apis/contact/v3/departments"
    headers = {"Authorization": f"Bearer {token}"}
    dept_map = {}
    dept_list = []  # 保存完整的部门信息用于导出
    
    def fetch_dept_children(parent_id, level):
        """获取单个部门的子部门"""
        children = []
        page_token = None
        
        while True:
            params = {
                "parent_department_id": parent_id,
                "page_size": 50,
                "department_id_type": "open_department_id"
            }
            if page_token:
                params["page_token"] = page_token
            
            # 对限流错误进行重试
            for retry in range(3):
                response = request_with_retry("GET", url, headers=headers, params=params)
                data = response.json()
                
                if data["code"] == 0:
                    break
                elif data["code"] == 99991400:  # 频率限制
                    if retry < 2:
                        time.sleep(1)
                        continue
                    else:
                        # 重试3次仍失败
                        raise Exception(f"部门{parent_id}获取失败: 触发限流且重试3次均失败")
                else:
                    # 其他错误
                    raise Exception(f"部门{parent_id}获取失败: code={data['code']}, msg={data.get('msg', 'unknown')}")
            
            if data["code"] != 0:
                raise Exception(f"部门{parent_id}获取失败: code={data['code']}")
            
            items = data.get("data", {}).get("items", [])
            for dept in items:
                dept_id = dept.get("open_department_id")
                dept_name = dept.get("name")
                parent_dept_id = dept.get("parent_department_id")
                member_count = dept.get("member_count", 0)
                
                if dept_id:
                    children.append({
                        "dept_id": dept_id,
                        "dept_name": dept_name,
                        "parent_dept_id": parent_dept_id,
                        "level": level,
                        "member_count": member_count
                    })
            
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token")
        
        return children
    
    print("正在获取部门信息...")
    start_time = time.time()
    
    # 使用广度优先遍历，逐层并发获取
    current_level = [("0", 0)]  # (parent_id, level)
    
    while current_level:
        next_level = []
        level_depts = []
        
        # 并发获取当前层所有部门的子部门
        with ThreadPoolExecutor(max_workers=40) as executor:
            futures = {executor.submit(fetch_dept_children, parent_id, level): (parent_id, level) 
                      for parent_id, level in current_level}
            
            for future in as_completed(futures):
                try:
                    children = future.result()
                    for dept in children:
                        dept_map[dept["dept_id"]] = dept["dept_name"]
                        dept_list.append(dept)
                        level_depts.append(dept)
                        next_level.append((dept["dept_id"], dept["level"] + 1))
                except Exception as e:
                    print(f"  ⚠ 获取部门失败: {e}")
        
        # 实时显示当前层级统计
        if level_depts:
            level = level_depts[0]["level"]
            print(f"  层级{level + 1}: {len(level_depts)}个部门")
        
        current_level = next_level
    
    elapsed_time = time.time() - start_time
    print(f"共获取 {len(dept_map)} 个部门，耗时 {elapsed_time:.2f} 秒")
    
    # 导出部门结构到CSV
    export_departments_to_csv(dept_list, dept_map)
    
    return dept_map, dept_list

def export_departments_to_csv(dept_list, dept_map):
    """导出部门层级结构到CSV"""
    # 获取项目根目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'output')
    
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, 'feishu_departments.csv')
    
    with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["dept_id", "dept_name", "parent_dept_id", "parent_dept_name", "level"])
        writer.writeheader()
        
        for dept in dept_list:
            # 添加父部门名称
            parent_id = dept["parent_dept_id"]
            dept["parent_dept_name"] = dept_map.get(parent_id, "根部门" if parent_id == "0" else "")
            # 移除member_count字段
            dept_output = {k: v for k, v in dept.items() if k != "member_count"}
            writer.writerow(dept_output)
    
    print(f"已导出 {len(dept_list)} 个部门到 {output_file}")

if __name__ == "__main__":
    # 校验环境变量
    if not APP_ID or not APP_SECRET:
        print("错误: 请在 .env 文件中配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET")
        exit(1)
    
    # 获取项目根目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'output')
    
    try:
        token = get_tenant_access_token()
        
        # 获取企业总人数
        total_user_count = get_total_user_count(token)
        
        # 先获取企业名称
        tenant_name = None
        try:
            tenant_url = "https://open.feishu.cn/open-apis/tenant/v2/tenant/query"
            headers = {"Authorization": f"Bearer {token}"}
            tenant_response = request_with_retry("GET", tenant_url, headers=headers)
            if tenant_response.status_code == 200:
                tenant_data = tenant_response.json()
                if tenant_data.get("code") == 0:
                    tenant_name = tenant_data.get("data", {}).get("tenant", {}).get("name")
                    
                    # 更新 .env 文件中的企业名称
                    if tenant_name:
                        env_file = os.path.join(script_dir, '.env')
                        
                        if os.path.exists(env_file):
                            with open(env_file, 'r', encoding='utf-8') as f:
                                lines = f.readlines()
                            
                            # 查找并更新 FEISHU_COMPANY_NAME
                            found = False
                            for i, line in enumerate(lines):
                                if line.strip().startswith('FEISHU_COMPANY_NAME='):
                                    lines[i] = f'FEISHU_COMPANY_NAME={tenant_name}\n'
                                    found = True
                                    break
                            
                            # 如果没找到，添加到文件末尾
                            if not found:
                                if lines and not lines[-1].endswith('\n'):
                                    lines[-1] += '\n'
                                lines.append(f'FEISHU_COMPANY_NAME={tenant_name}\n')
                            
                            with open(env_file, 'w', encoding='utf-8') as f:
                                f.writelines(lines)
        except Exception as e:
            print(f"获取企业名称失败: {e}")
        
        # 显示开始信息
        if tenant_name:
            print(f"正在从 {tenant_name} 获取通讯录信息...")
        else:
            print("正在获取通讯录信息...")
        
        # 获取部门信息
        dept_map, dept_list = get_department_info(token)
        
        # 获取用户信息（传递部门ID列表）
        use_sequential = os.getenv("SEQUENTIAL_MODE", "false").lower() == "true"
        if use_sequential:
            print("使用顺序模式获取用户...")
            users = get_all_users_sequential(token, list(dept_map.keys()))
            total_fetched = len(users)
        else:
            users, total_fetched = get_all_users(token, list(dept_map.keys()))
        
        # 为每个用户添加部门名称
        for user in users:
            dept_ids = user.get("department_ids", [])
            if dept_ids:
                dept_id = dept_ids[0]
                # 如果部门ID为0，部门名称为空
                if dept_id == "0":
                    user["department_name"] = ""
                else:
                    user["department_name"] = dept_map.get(dept_id, "")
            else:
                # 没有部门时，部门名称为空
                user["department_name"] = ""
        
        # 导出用户信息并获取统计
        dept_user_count = export_to_csv(users, dept_map)
        
        # 数据验证
        print(f"\n数据统计:")
        print(f"  实际获取部门数: {len(dept_map)}")
        if total_user_count is not None:
            print(f"  企业总人数(飞书): {total_user_count}")
        print(f"  实际获取用户数: {len(users)}")
        print(f"  API返回总数(含重复): {total_fetched}")
        if total_fetched > len(users):
            print(f"  (其中 {total_fetched - len(users)} 个用户属于多个部门)")
        
        # 检查是否获取到数据
        if len(dept_map) == 0 or len(users) == 0:
            raise Exception(f"飞书数据获取失败: 部门数={len(dept_map)}, 用户数={len(users)}，请检查飞书应用配置和IP白名单")
        
        # 验证用户数据完整性 - 不一致视为失败
        if total_user_count is not None and len(users) != total_user_count:
            raise Exception(f"用户数据不完整: 实际获取{len(users)}个，企业总人数{total_user_count}个，差异{len(users) - total_user_count}个")
    
    except Exception as e:
        print(f"\n❌ 飞书数据获取失败: {e}")
        print("正在清理已生成的CSV文件...")
        
        # 删除可能已生成的CSV文件
        csv_files = [
            os.path.join(output_dir, 'feishu_users.csv'),
            os.path.join(output_dir, 'feishu_departments.csv')
        ]
        
        for csv_file in csv_files:
            if os.path.exists(csv_file):
                os.remove(csv_file)
                print(f"  已删除: {csv_file}")
        
        print("数据获取失败，请稍后重试")
        exit(1)
