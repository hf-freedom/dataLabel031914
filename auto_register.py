# -*- coding: utf-8 -*-
import random
import string
import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "validation_results.xlsx")

FIRST_NAMES = ["张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗"]
LAST_NAMES = ["伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军", "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞"]

excel_lock = threading.Lock()


def generate_random_string(length=8):
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=length))


def generate_random_email():
    username = generate_random_string(10)
    return f"{username}@163.com"


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
        ws.title = "字段校验结果"
        headers = ["序号", "测试时间", "账号", "字段名称", "原字段值", "目标修改值", "测试类型", "接口返回结果", "校验状态"]
        ws.append(headers)
        wb.save(EXCEL_FILE)
        print(f"创建Excel文件: {EXCEL_FILE}")
    return EXCEL_FILE


def save_validation_result(data):
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
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
    
    print(f"[任务{task_id}] 查找注册表单字段...")
    
    username_input, _ = find_input(page, username_selectors, "用户名")
    password_input, _ = find_input(page, password_selectors, "密码")
    email_input, _ = find_input(page, email_selectors, "邮箱")
    name_input, _ = find_input(page, name_selectors, "姓名")
    age_input, _ = find_input(page, age_selectors, "年龄")
    phone_input, _ = find_input(page, phone_selectors, "手机号")
    
    print(f"[任务{task_id}] 填写注册信息...")
    
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
        print(f"[任务{task_id}] 点击注册按钮...")
        submit_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        register_status = "成功"
    
    return register_status == "成功"


def perform_login(page, user_data, task_id):
    username = user_data["username"]
    password = user_data["password"]
    
    print(f"[任务{task_id}] 跳转到登录页面...")
    
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
    
    print(f"[任务{task_id}] 填写登录信息...")
    
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
    
    login_success = False
    if login_btn:
        print(f"[任务{task_id}] 点击登录按钮...")
        login_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        login_success = True
    
    return login_success


def navigate_to_profile(page, task_id):
    print(f"[任务{task_id}] 查找编辑按钮...")
    
    # 首先打印页面上所有按钮，方便调试
    all_buttons = page.query_selector_all("button, a[role='button'], input[type='button']")
    print(f"[任务{task_id}] 页面上有 {len(all_buttons)} 个按钮/链接")
    for i, btn in enumerate(all_buttons[:10]):  # 只显示前10个
        try:
            text = btn.inner_text().strip() if btn else ""
            class_attr = btn.get_attribute("class") or ""
            print(f"  按钮{i+1}: text='{text[:20]}', class='{class_attr[:30]}'")
        except:
            pass
    
    edit_btn_selectors = [
        'button:has-text("编辑")',
        'a:has-text("编辑")',
        'button:has-text("修改")',
        'a:has-text("修改")',
        '.edit-btn',
        '#edit-btn',
        'button[class*="edit"]',
        'a[class*="edit"]',
        'button:has-text("Edit")',
        'input[value="编辑"]',
        'input[value="修改"]',
        'span:has-text("编辑")',
        'span:has-text("修改")',
        'i:has-text("编辑")',
        'i:has-text("修改")',
        '[title="编辑"]',
        '[title="修改"]',
        '.btn-edit',
        '#btn-edit'
    ]
    
    edit_btn = None
    found_selector = None
    for selector in edit_btn_selectors:
        try:
            edit_btn = page.query_selector(selector)
            if edit_btn and edit_btn.is_visible():
                found_selector = selector
                print(f"[任务{task_id}] 找到编辑按钮: {selector}")
                break
        except Exception as e:
            continue
    
    if edit_btn:
        print(f"[任务{task_id}] 点击编辑按钮进入修改页面...")
        try:
            edit_btn.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            print(f"[任务{task_id}] 成功进入编辑页面")
            return True
        except Exception as e:
            print(f"[任务{task_id}] 点击编辑按钮失败: {e}")
    
    print(f"[任务{task_id}] 未找到编辑按钮，尝试查找个人信息入口...")
    
    profile_selectors = [
        'a:has-text("个人信息")',
        'a:has-text("个人中心")',
        'a:has-text("我的")',
        'a:has-text("用户中心")',
        'text=个人信息',
        'text=个人中心',
        '.profile-link',
        '#profile-link',
        'a[href*="profile"]',
        'a[href*="user"]',
        'a[href*="personal"]'
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
        print(f"[任务{task_id}] 进入个人信息页面...")
        profile_link.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        
        print(f"[任务{task_id}] 在个人信息页面查找编辑按钮...")
        for selector in edit_btn_selectors:
            try:
                edit_btn = page.query_selector(selector)
                if edit_btn:
                    print(f"[任务{task_id}] 找到编辑按钮并点击...")
                    edit_btn.click()
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(1500)
                    return True
            except:
                continue
    
    print(f"[任务{task_id}] 未找到个人信息入口，尝试直接访问...")
    try:
        page.goto("http://39.107.109.8:8082/profile", timeout=10000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        return True
    except:
        pass
    
    try:
        page.goto("http://39.107.109.8:8082/user/profile", timeout=10000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        return True
    except:
        pass
    
    return False


def get_profile_fields(page):
    fields = {}
    
    # 首先尝试获取页面上所有可见的输入框
    all_inputs = page.query_selector_all('input:not([type="hidden"]):not([type="submit"]):not([type="button"])')
    print(f"  找到 {len(all_inputs)} 个输入框")
    
    # 尝试通过placeholder或label文本识别字段
    field_configs = {
        "username": [
            'input[name="username"]', 'input[name="userName"]', 'input[name="loginname"]',
            '#username', '#userName', '#loginName',
            'input[placeholder*="用户名"]', 'input[placeholder*="账号"]',
            'input[id*="user"]', 'input[id*="name"]'
        ],
        "name": [
            'input[name="name"]', 'input[name="realName"]', 'input[name="realname"]', 
            'input[name="nickname"]', 'input[name="nickName"]',
            '#name', '#realName', '#nickname',
            'input[placeholder*="姓名"]', 'input[placeholder*="真实姓名"]', 'input[placeholder*="昵称"]'
        ],
        "email": [
            'input[name="email"]', 'input[name="mail"]', 
            '#email', '#mail', 
            'input[type="email"]',
            'input[placeholder*="邮箱"]', 'input[placeholder*="Email"]',
            'input[id*="email"]', 'input[id*="mail"]'
        ],
        "phone": [
            'input[name="phone"]', 'input[name="mobile"]', 'input[name="tel"]', 
            'input[name="phoneNumber"]', 'input[name="telephone"]',
            '#phone', '#mobile', '#tel', '#telephone',
            'input[placeholder*="手机"]', 'input[placeholder*="电话"]', 'input[placeholder*="手机号"]',
            'input[type="tel"]',
            'input[id*="phone"]', 'input[id*="mobile"]'
        ],
        "age": [
            'input[name="age"]', '#age', 
            'input[type="number"]', 
            'input[placeholder*="年龄"]',
            'input[id*="age"]'
        ],
        "address": [
            'input[name="address"]', '#address',
            'input[placeholder*="地址"]', 'input[placeholder*="住址"]',
            'input[id*="address"]'
        ],
        "gender": [
            'select[name="gender"]', 'select[name="sex"]',
            '#gender', '#sex',
            'input[name="gender"]', 'input[name="sex"]'
        ]
    }
    
    for field_name, selectors in field_configs.items():
        element, selector = find_input(page, selectors, field_name)
        if element:
            try:
                # 检查元素是否可见和可编辑
                is_visible = element.is_visible()
                is_enabled = element.is_enabled()
                if is_visible and is_enabled:
                    original_value = element.input_value() or ""
                    fields[field_name] = {
                        "element": element,
                        "selector": selector,
                        "original_value": original_value
                    }
                    print(f"  找到字段 '{field_name}': 值='{original_value}', 选择器='{selector}'")
            except Exception as e:
                print(f"  字段 '{field_name}' 检查失败: {e}")
    
    # 如果没有找到任何字段，尝试获取所有文本输入框
    if not fields:
        print("  未通过配置找到字段，尝试获取所有文本输入框...")
        text_inputs = page.query_selector_all('input[type="text"], input:not([type])')
        for i, inp in enumerate(text_inputs):
            try:
                if inp.is_visible() and inp.is_enabled():
                    name = inp.get_attribute("name") or f"field_{i}"
                    original_value = inp.input_value() or ""
                    fields[name] = {
                        "element": inp,
                        "selector": f'input[name="{name}"]' if name != f"field_{i}" else f'input[type="text"]:nth-of-type({i+1})',
                        "original_value": original_value
                    }
                    print(f"  找到字段 '{name}': 值='{original_value}'")
            except:
                pass
    
    return fields


def get_save_button(page):
    save_selectors = [
        'button:has-text("保存")',
        'button:has-text("提交")',
        'button:has-text("更新")',
        'button:has-text("确认")',
        'input[value="保存"]',
        'input[value="提交"]',
        'input[value="更新"]',
        '.save-btn',
        '#save-btn',
        'button[type="submit"]'
    ]
    
    for selector in save_selectors:
        try:
            btn = page.query_selector(selector)
            if btn:
                return btn
        except:
            continue
    return None


def test_field_validation(page, user_data, task_id, validation_results):
    fields = get_profile_fields(page)
    
    if not fields:
        print(f"[任务{task_id}] 未找到可编辑的个人信息字段")
        return validation_results
    
    print(f"[任务{task_id}] 找到以下可编辑字段: {list(fields.keys())}")
    
    test_cases = generate_test_cases()
    
    for field_name, field_info in fields.items():
        original_value = field_info["original_value"]
        element = field_info["element"]
        
        print(f"[任务{task_id}] 开始测试字段 '{field_name}'，原值: '{original_value}'")
        
        for test_case in test_cases:
            test_value = test_case["value"]
            test_type = test_case["type"]
            
            try:
                element.fill(test_value)
                page.wait_for_timeout(300)
                
                save_btn = get_save_button(page)
                if save_btn:
                    save_btn.click()
                    page.wait_for_timeout(1500)
                
                response_result = check_response_result(page)
                
                validation_status = "通过" if response_result.get("success", False) else "失败"
                
                result_data = [
                    len(validation_results) + 1,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    user_data["username"],
                    field_name,
                    original_value,
                    test_value if len(test_value) <= 50 else test_value[:50] + "...",
                    test_type,
                    response_result.get("message", "未知"),
                    validation_status
                ]
                
                save_validation_result(result_data)
                validation_results.append(result_data)
                
                print(f"[任务{task_id}] 字段 '{field_name}' - 测试类型: {test_type} - 结果: {validation_status}")
                
                element.fill(original_value)
                page.wait_for_timeout(300)
                
            except Exception as e:
                print(f"[任务{task_id}] 测试字段 '{field_name}' 时出错: {e}")
                
                result_data = [
                    len(validation_results) + 1,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    user_data["username"],
                    field_name,
                    original_value,
                    test_value if len(test_value) <= 50 else test_value[:50] + "...",
                    test_type,
                    f"异常: {str(e)}",
                    "错误"
                ]
                
                save_validation_result(result_data)
                validation_results.append(result_data)
    
    return validation_results


def generate_test_cases():
    test_cases = [
        {"value": "", "type": "空字段"},
        {"value": "a" * 100, "type": "100字符超长字段"},
        {"value": "12345678901234567890", "type": "纯数字(20位)"},
        {"value": "!@#$%^&*()_+-=[]{}|;':\",./<>?", "type": "特殊符号"},
        {"value": "<script>alert('xss')</script>", "type": "XSS攻击脚本"},
        {"value": "' OR '1'='1", "type": "SQL注入尝试"},
        {"value": "测试中文特殊字符【】、；'\"，。/", "type": "中文特殊字符"},
        {"value": "  前后空格  ", "type": "前后空格"},
        {"value": "\n\t\r", "type": "换行制表符"},
        {"value": "null", "type": "字符串null"},
        {"value": "undefined", "type": "字符串undefined"},
    ]
    return test_cases


def check_response_result(page):
    result = {"success": False, "message": "未检测到响应"}
    
    try:
        page_content = page.content()
        
        success_keywords = ["成功", "保存成功", "更新成功", "修改成功", "success", "ok"]
        error_keywords = ["失败", "错误", "非法", "无效", "不能为空", "格式错误", "error", "fail", "invalid"]
        
        page_text = page_content.lower()
        
        for keyword in success_keywords:
            if keyword in page_text:
                result["success"] = True
                result["message"] = f"检测到成功提示: {keyword}"
                return result
        
        for keyword in error_keywords:
            if keyword in page_text:
                result["success"] = False
                result["message"] = f"检测到错误提示: {keyword}"
                return result
        
        result["message"] = "未检测到明确的成功或失败提示"
        
    except Exception as e:
        result["message"] = f"检测异常: {str(e)}"
    
    return result


def single_task(task_id, user_data):
    print(f"\n[任务{task_id}] 开始执行...")
    print(f"[任务{task_id}] 用户名: {user_data['username']}")
    
    validation_results = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print(f"[任务{task_id}] 正在打开网站...")
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            print(f"[任务{task_id}] 查找注册链接...")
            register_link = page.query_selector('text=注册') or page.query_selector('text=立即注册') or page.query_selector('a:has-text("注册")')
            
            if register_link:
                print(f"[任务{task_id}] 点击注册链接...")
                register_link.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(1000)
            
            register_success = perform_register(page, user_data, task_id)
            
            if not register_success:
                print(f"[任务{task_id}] 注册失败，跳过后续操作")
                return {
                    "task_id": task_id,
                    "status": "注册失败",
                    "validation_count": 0
                }
            
            print(f"[任务{task_id}] 注册成功，开始登录...")
            login_success = perform_login(page, user_data, task_id)
            
            if not login_success:
                print(f"[任务{task_id}] 登录失败，跳过后续操作")
                return {
                    "task_id": task_id,
                    "status": "登录失败",
                    "validation_count": 0
                }
            
            print(f"[任务{task_id}] 登录成功，进入个人信息页面...")
            profile_success = navigate_to_profile(page, task_id)
            
            if not profile_success:
                print(f"[任务{task_id}] 无法进入个人信息页面")
                return {
                    "task_id": task_id,
                    "status": "进入个人信息失败",
                    "validation_count": 0
                }
            
            print(f"[任务{task_id}] 开始进行字段校验测试...")
            validation_results = test_field_validation(page, user_data, task_id, validation_results)
            
            print(f"\n[任务{task_id}] ========== 执行完成 ==========")
            print(f"[任务{task_id}] 完成字段校验测试: {len(validation_results)} 项")
            print(f"[任务{task_id}] ==============================\n")
            
            return {
                "task_id": task_id,
                "status": "成功",
                "validation_count": len(validation_results)
            }
            
        except Exception as e:
            print(f"[任务{task_id}] 发生错误: {e}")
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


def run_parallel_validation(num_tasks=3):
    init_excel()
    
    print("=" * 60)
    print(f"开始并行执行 {num_tasks} 个字段校验任务")
    print("=" * 60)
    
    overall_start_time = datetime.now()
    
    users_data = [generate_user_data() for _ in range(num_tasks)]
    
    print("\n生成的用户信息:")
    for i, user in enumerate(users_data, 1):
        print(f"  任务{i}: {user['username']}")
    
    results = []
    
    with ThreadPoolExecutor(max_workers=num_tasks) as executor:
        futures = {executor.submit(single_task, i+1, user): i+1 for i, user in enumerate(users_data)}
        
        for future in as_completed(futures):
            task_id = futures[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                print(f"任务{task_id}执行异常: {e}")
                results.append({"task_id": task_id, "status": "异常", "error": str(e)})
    
    overall_end_time = datetime.now()
    overall_duration = (overall_end_time - overall_start_time).total_seconds()
    
    print("\n" + "=" * 60)
    print("所有任务执行完成!")
    print("=" * 60)
    
    success_count = sum(1 for r in results if r.get("status") == "成功")
    fail_count = num_tasks - success_count
    total_validations = sum(r.get("validation_count", 0) for r in results)
    
    print(f"\n执行统计:")
    print(f"  总任务数: {num_tasks}")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    print(f"  总校验项: {total_validations}")
    print(f"  总耗时: {overall_duration:.2f}秒")
    
    print(f"\n数据已保存到: {EXCEL_FILE}")
    
    return results


if __name__ == "__main__":
    results = run_parallel_validation(5)
