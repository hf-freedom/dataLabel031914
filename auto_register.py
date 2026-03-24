# -*- coding: utf-8 -*-
import random
import string
import os
import time
import threading
from datetime import datetime
from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "register_data.xlsx")

FIRST_NAMES = ["张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗"]
LAST_NAMES = ["伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军", "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞"]

excel_lock = threading.Lock()

def generate_random_string(length=8):
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=length))

def generate_random_email():
    username = generate_random_string(10)
    return "{0}@163.com".format(username)

def generate_random_password():
    return generate_random_string(12) + random.choice(string.ascii_uppercase) + random.choice(string.digits)

def generate_random_name():
    first_name = random.choice(FIRST_NAMES)
    last_name = random.choice(LAST_NAMES)
    return first_name + last_name

def generate_random_age():
    return str(random.randint(18, 60))

def generate_random_phone():
    prefixes = ["130", "131", "132", "133", "134", "135", "136", "137", "138", "139",
                "150", "151", "152", "153", "155", "156", "157", "158", "159",
                "170", "176", "177", "178",
                "180", "181", "182", "183", "184", "185", "186", "187", "188", "189"]
    prefix = random.choice(prefixes)
    suffix = ''.join(random.choices(string.digits, k=8))
    return prefix + suffix



def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "注册数据"
        headers = ["序号", "用户名", "密码", "邮箱", "姓名", "年龄", "手机号", "注册开始时间", "注册结束时间", "注册耗时(秒)", "总耗时(秒)", "注册状态", "登录状态", "验证状态"]
        ws.append(headers)
        wb.save(EXCEL_FILE)
        print("创建Excel文件: {0}".format(EXCEL_FILE))
    return EXCEL_FILE

def save_to_excel(data):
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)

def save_profile_results(results):
    if not results:
        return
    
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        
        # 检查是否存在"个人信息验证"工作表
        sheet_name = "个人信息验证"
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            headers = ["序号", "要修改字段", "当前账号", "原字段值", "目标修改值", "接口返回结果"]
            ws.append(headers)
        else:
            ws = wb[sheet_name]
        
        # 获取当前最大行号，用于序号
        row_num = ws.max_row
        
        for result in results:
            row_num += 1
            excel_data = [
                row_num,
                result["field"],
                result["username"],
                result["old_value"],
                result["new_value"],
                result["message"]
            ]
            ws.append(excel_data)
        
        wb.save(EXCEL_FILE)

def find_input(page, selectors, field_name):
    for selector in selectors:
        try:
            element = page.query_selector(selector)
            if element:
                return element, selector
        except:
            continue
    return None, None

def perform_register(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    email = user_data["email"]
    name = user_data["name"]
    age = user_data["age"]
    phone = user_data["phone"]
    
    register_start_time = datetime.now()
    
    username_selectors = [
        'input[name="username"]',
        'input[name="user"]',
        'input[name="userName"]',
        'input[placeholder*="用户名"]',
        'input[placeholder*="账号"]',
        '#username',
        '#userName',
        'input[type="text"]:first-of-type'
    ]
    
    password_selectors = [
        'input[name="password"]',
        'input[name="pwd"]',
        'input[name="userPassword"]',
        'input[placeholder*="密码"]',
        '#password',
        '#pwd',
        'input[type="password"]'
    ]
    
    email_selectors = [
        'input[name="email"]',
        'input[name="mail"]',
        'input[placeholder*="邮箱"]',
        'input[placeholder*="Email"]',
        '#email',
        'input[type="email"]'
    ]
    
    name_selectors = [
        'input[name="name"]',
        'input[name="realName"]',
        'input[name="realname"]',
        'input[name="userName"]',
        'input[placeholder*="姓名"]',
        'input[placeholder*="真实姓名"]',
        '#name',
        '#realName',
        '#userName'
    ]
    
    age_selectors = [
        'input[name="age"]',
        'input[placeholder*="年龄"]',
        '#age',
        'input[type="number"]'
    ]
    
    phone_selectors = [
        'input[name="phone"]',
        'input[name="mobile"]',
        'input[name="tel"]',
        'input[name="phoneNumber"]',
        'input[placeholder*="手机"]',
        'input[placeholder*="电话"]',
        'input[placeholder*="手机号"]',
        '#phone',
        '#mobile',
        '#tel'
    ]
    
    print("[任务{0}] 查找注册表单字段...".format(task_id))
    
    username_input, _ = find_input(page, username_selectors, "用户名")
    password_input, _ = find_input(page, password_selectors, "密码")
    email_input, _ = find_input(page, email_selectors, "邮箱")
    name_input, _ = find_input(page, name_selectors, "姓名")
    age_input, _ = find_input(page, age_selectors, "年龄")
    phone_input, _ = find_input(page, phone_selectors, "手机号")
    
    print("[任务{0}] 填写注册信息...".format(task_id))
    
    if username_input:
        username_input.fill(username)
    if password_input:
        password_input.fill(password)
    if email_input:
        email_input.fill(email)
    if name_input:
        name_input.fill(name)
    if age_input:
        age_input.fill(age)
    if phone_input:
        phone_input.fill(phone)
    
    page.wait_for_timeout(500)
    
    submit_selectors = [
        'button:has-text("注册")',
        'button:has-text("提交")',
        'button:has-text("确定")',
        'input[type="submit"]',
        'input[value="注册"]',
        'input[value="提交"]',
        '.register-btn',
        '#register-btn',
        'button[type="submit"]'
    ]
    
    submit_btn = None
    for selector in submit_selectors:
        try:
            submit_btn = page.query_selector(selector)
            if submit_btn:
                break
        except:
            continue
    
    register_status = "失败"
    if submit_btn:
        print("[任务{0}] 点击注册按钮...".format(task_id))
        submit_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        register_status = "成功"
    
    register_end_time = datetime.now()
    register_duration = (register_end_time - register_start_time).total_seconds()
    
    return {
        "register_start_time": register_start_time,
        "register_end_time": register_end_time,
        "register_duration": register_duration,
        "register_status": register_status
    }

def perform_login(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    
    login_start_time = datetime.now()
    
    print("[任务{0}] 跳转到登录页面...".format(task_id))
    
    login_link_selectors = [
        'a:has-text("登录")',
        'text=登录',
        'a[href*="login"]',
        '.login-link',
        '#login-link'
    ]
    
    login_link = None
    for selector in login_link_selectors:
        try:
            login_link = page.query_selector(selector)
            if login_link:
                break
        except:
            continue
    
    if login_link:
        login_link.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
    else:
        page.goto("http://39.107.109.8:8082/", timeout=30000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
    
    username_selectors = [
        'input[name="username"]',
        'input[name="user"]',
        'input[name="userName"]',
        'input[placeholder*="用户名"]',
        'input[placeholder*="账号"]',
        '#username',
        '#userName',
        'input[type="text"]:first-of-type'
    ]
    
    password_selectors = [
        'input[name="password"]',
        'input[name="pwd"]',
        'input[name="userPassword"]',
        'input[placeholder*="密码"]',
        '#password',
        '#pwd',
        'input[type="password"]'
    ]
    
    username_input, _ = find_input(page, username_selectors, "用户名")
    password_input, _ = find_input(page, password_selectors, "密码")
    
    print("[任务{0}] 填写登录信息...".format(task_id))
    
    if username_input:
        username_input.fill(username)
    if password_input:
        password_input.fill(password)
    
    page.wait_for_timeout(500)
    
    login_btn_selectors = [
        'button:has-text("登录")',
        'button:has-text("Login")',
        'input[type="submit"]',
        'input[value="登录"]',
        '.login-btn',
        '#login-btn',
        'button[type="submit"]'
    ]
    
    login_btn = None
    for selector in login_btn_selectors:
        try:
            login_btn = page.query_selector(selector)
            if login_btn:
                break
        except:
            continue
    
    login_status = "失败"
    verify_status = "未验证"
    
    if login_btn:
        print("[任务{0}] 点击登录按钮...".format(task_id))
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        login_status = "成功"
        
        print("[任务{0}] 验证用户信息...".format(task_id))
        
        page_content = page.content()
        verify_keywords = [user_data["username"], user_data["name"], "个人信息", "欢迎", "我的", "用户中心"]
        
        for keyword in verify_keywords:
            if keyword in page_content:
                verify_status = "验证成功"
                break
        
        if verify_status != "验证成功":
            verify_status = "验证失败"
    
    login_end_time = datetime.now()
    login_duration = (login_end_time - login_start_time).total_seconds()
    
    return {
        "login_start_time": login_start_time,
        "login_end_time": login_end_time,
        "login_duration": login_duration,
        "login_status": login_status,
        "verify_status": verify_status
    }

def navigate_to_profile(page, task_id):
    print("[任务{0}] 导航到个人信息页面...".format(task_id))
    
    profile_selectors = [
        'a:has-text("个人信息")',
        'a:has-text("我的")',
        'a:has-text("用户中心")',
        'a[href*="profile"]',
        'a[href*="user"]',
        '.profile-link',
        '#profile-link',
        'text=个人信息',
        'text=我的',
        'text=用户中心'
    ]
    
    profile_link = None
    for selector in profile_selectors:
        try:
            profile_link = page.query_selector(selector)
            if profile_link:
                break
        except:
            continue
    
    if profile_link:
        profile_link.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        return True
    else:
        print("[任务{0}] 未找到个人信息页面链接".format(task_id))
        return False

def get_field_value(page, selectors):
    for selector in selectors:
        try:
            element = page.query_selector(selector)
            if element:
                return element.input_value()
        except:
            continue
    return ""

def update_profile_field(page, field_name, field_selectors, old_value, new_value, task_id):
    print("[任务{0}] 修改{1}: {2} -> {3}".format(task_id, field_name, old_value, new_value))
    
    field_input, _ = find_input(page, field_selectors, field_name)
    if not field_input:
        print("[任务{0}] 未找到{1}输入框".format(task_id, field_name))
        return False, "未找到输入框"
    
    field_input.fill(new_value)
    page.wait_for_timeout(500)
    
    submit_selectors = [
        'button:has-text("保存")',
        'button:has-text("提交")',
        'button:has-text("确定")',
        'input[type="submit"]',
        'input[value="保存"]',
        'input[value="提交"]',
        '.save-btn',
        '#save-btn',
        'button[type="submit"]'
    ]
    
    submit_btn = None
    for selector in submit_selectors:
        try:
            submit_btn = page.query_selector(selector)
            if submit_btn:
                break
        except:
            continue
    
    if not submit_btn:
        print("[任务{0}] 未找到保存按钮".format(task_id))
        return False, "未找到保存按钮"
    
    submit_btn.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    
    page_content = page.content()
    
    success_keywords = ["成功", "保存成功", "修改成功", "更新成功"]
    error_keywords = ["错误", "失败", "请输入", "格式错误", "长度错误"]
    
    for keyword in success_keywords:
        if keyword in page_content:
            return True, "修改成功"
    
    for keyword in error_keywords:
        if keyword in page_content:
            return False, "验证失败: " + keyword
    
    return False, "未知结果"

def perform_profile_update(page, user_data, task_id):
    print("[任务{0}] 开始修改个人信息...".format(task_id))
    
    if not navigate_to_profile(page, task_id):
        return []
    
    name_selectors = [
        'input[name="name"]',
        'input[name="realName"]',
        'input[name="realname"]',
        'input[name="userName"]',
        'input[placeholder*="姓名"]',
        'input[placeholder*="真实姓名"]',
        '#name',
        '#realName',
        '#userName'
    ]
    
    age_selectors = [
        'input[name="age"]',
        'input[placeholder*="年龄"]',
        '#age',
        'input[type="number"]'
    ]
    
    phone_selectors = [
        'input[name="phone"]',
        'input[name="mobile"]',
        'input[name="tel"]',
        'input[name="phoneNumber"]',
        'input[placeholder*="手机"]',
        'input[placeholder*="电话"]',
        'input[placeholder*="手机号"]',
        '#phone',
        '#mobile',
        '#tel'
    ]
    
    email_selectors = [
        'input[name="email"]',
        'input[name="mail"]',
        'input[placeholder*="邮箱"]',
        'input[placeholder*="Email"]',
        '#email',
        'input[type="email"]'
    ]
    
    old_name = get_field_value(page, name_selectors)
    old_age = get_field_value(page, age_selectors)
    old_phone = get_field_value(page, phone_selectors)
    old_email = get_field_value(page, email_selectors)
    
    test_cases = [
        ("姓名", name_selectors, old_name, ""),  # 空字段
        ("姓名", name_selectors, old_name, "a" * 100),  # 100个字符
        ("姓名", name_selectors, old_name, "123456"),  # 纯数字
        ("姓名", name_selectors, old_name, "!@#$%^"),  # 特殊符号
        ("年龄", age_selectors, old_age, ""),  # 空字段
        ("年龄", age_selectors, old_age, "abc"),  # 非数字
        ("年龄", age_selectors, old_age, "1000"),  # 过大的数字
        ("手机号", phone_selectors, old_phone, ""),  # 空字段
        ("手机号", phone_selectors, old_phone, "123"),  # 过短
        ("手机号", phone_selectors, old_phone, "1234567890123"),  # 过长
        ("手机号", phone_selectors, old_phone, "abc12345678"),  # 包含字母
        ("邮箱", email_selectors, old_email, ""),  # 空字段
        ("邮箱", email_selectors, old_email, "invalid-email"),  # 无效邮箱
        ("邮箱", email_selectors, old_email, "test@"),  # 不完整邮箱
    ]
    
    results = []
    for field_name, selectors, old_value, new_value in test_cases:
        success, message = update_profile_field(page, field_name, selectors, old_value, new_value, task_id)
        results.append({
            "field": field_name,
            "username": user_data["username"],
            "old_value": old_value,
            "new_value": new_value,
            "success": success,
            "message": message
        })
        
        # 每次测试后刷新页面，恢复原始值
        page.reload()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        navigate_to_profile(page, task_id)
    
    return results

def single_task(task_id, user_data):
    print("\n[任务{0}] 开始执行...".format(task_id))
    print("[任务{0}] 用户名: {1}".format(task_id, user_data['username']))
    
    task_start_time = datetime.now()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print("[任务{0}] 正在打开网站...".format(task_id))
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            print("[任务{0}] 查找注册链接...".format(task_id))
            register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
            
            if register_link:
                print("[任务{0}] 点击注册链接...".format(task_id))
                register_link.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
            
            register_result = perform_register(page, user_data, task_id)
            
            login_result = perform_login(page, user_data, task_id)
            
            profile_results = []
            if login_result["login_status"] == "成功":
                profile_results = perform_profile_update(page, user_data, task_id)
            
            task_end_time = datetime.now()
            total_duration = (task_end_time - task_start_time).total_seconds()
            
            excel_data = [
                task_id,
                user_data["username"],
                user_data["password"],
                user_data["email"],
                user_data["name"],
                user_data["age"],
                user_data["phone"],
                register_result["register_start_time"].strftime("%Y-%m-%d %H:%M:%S"),
                register_result["register_end_time"].strftime("%Y-%m-%d %H:%M:%S"),
                round(register_result["register_duration"], 2),
                round(total_duration, 2),
                register_result["register_status"],
                login_result["login_status"],
                login_result["verify_status"]
            ]
            
            save_to_excel(excel_data)
            
            # 保存个人信息修改验证结果
            save_profile_results(profile_results)
            
            print("\n[任务{0}] ========== 执行完成 ==========".format(task_id))
            print("[任务{0}] 注册状态: {1}".format(task_id, register_result['register_status']))
            print("[任务{0}] 注册耗时: {1:.2f}秒".format(task_id, register_result['register_duration']))
            print("[任务{0}] 登录状态: {1}".format(task_id, login_result['login_status']))
            print("[任务{0}] 验证状态: {1}".format(task_id, login_result['verify_status']))
            print("[任务{0}] 总耗时: {1:.2f}秒".format(task_id, total_duration))
            print("[任务{0}] ==============================\n".format(task_id))
            
            return {
                "task_id": task_id,
                "status": "成功",
                "total_duration": total_duration,
                "register_status": register_result["register_status"],
                "login_status": login_result["login_status"],
                "verify_status": login_result["verify_status"]
            }
            
        except Exception as e:
            print("[任务{0}] 发生错误: {1}".format(task_id, e))
            
            return {
                "task_id": task_id,
                "status": "失败",
                "error": str(e)
            }
        finally:
            browser.close()

def generate_user_data():
    return {
        "username": "user_" + generate_random_string(6),
        "password": generate_random_password(),
        "email": generate_random_email(),
        "name": generate_random_name(),
        "age": generate_random_age(),
        "phone": generate_random_phone()
    }

def run_parallel_register(num_tasks=5):
    init_excel()
    
    print("=" * 60)
    print("开始并行执行 {0} 个注册任务".format(num_tasks))
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print("  任务{0}: {1}".format(i, user['username']))
    
    results = []
    threads = []
    
    def thread_task(task_id, user_data, result_list):
        try:
            result = single_task(task_id, user_data)
            result_list.append(result)
        except Exception as e:
            print("任务{0}执行异常: {1}".format(task_id, e))
            result_list.append({"task_id": task_id, "status": "异常", "error": str(e)})
    
    for i, user in enumerate(users_data, 1):
        thread = threading.Thread(target=thread_task, args=(i, user, results))
        threads.append(thread)
        thread.start()
    
    for thread in threads:
        thread.join()
    
    overall_end_time = datetime.now()
    overall_duration = (overall_end_time - overall_start_time).total_seconds()
    
    print("\n" + "=" * 60)
    print("所有任务执行完成!")
    print("=" * 60)
    
    success_count = sum(1 for r in results if r.get("status") == "成功")
    fail_count = num_tasks - success_count
    
    print("\n执行统计:")
    print("  总任务数: {0}".format(num_tasks))
    print("  成功: {0}".format(success_count))
    print("  失败: {0}".format(fail_count))
    print("  总耗时: {0:.2f}秒".format(overall_duration))
    print("  平均耗时: {0:.2f}秒/任务".format(overall_duration/num_tasks))
    
    print("\n数据已保存到: {0}".format(EXCEL_FILE))
    
    return results

if __name__ == "__main__":
    results = run_parallel_register(5)
