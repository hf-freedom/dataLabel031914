# -*- coding: utf-8 -*-
import random
import string
from playwright.sync_api import sync_playwright
import time

def generate_random_string(length=8):
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=length))

def generate_random_password():
    return generate_random_string(12) + random.choice(string.ascii_uppercase) + random.choice(string.digits)

def generate_random_name():
    first_names = ["张", "王", "李", "赵", "刘"]
    last_names = ["伟", "芳", "娜", "秀英", "敏"]
    return random.choice(first_names) + random.choice(last_names)

def debug_site():
    print("启动调试模式...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=300)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print("1. 打开网站...")
            page.goto("http://39.107.109.8:8082/", timeout=30000)
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            
            print("\n" + "=" * 60)
            print("2. 点击注册链接...")
            register_link = page.query_selector('a:has-text("注册")')
            
            if register_link:
                register_link.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                
                username = "user_" + generate_random_string(6)
                password = generate_random_password()
                name = generate_random_name()
                age = str(random.randint(18, 60))
                phone = "138" + ''.join(random.choices(string.digits, k=8))
                email = f"{generate_random_string(10)}@163.com"
                
                print(f"注册用户: {username}")
                
                all_inputs = page.query_selector_all('input')
                for inp in all_inputs:
                    try:
                        placeholder = inp.get_attribute('placeholder') or ""
                        input_type = inp.get_attribute('type') or ""
                        
                        if '账号' in placeholder:
                            inp.fill(username)
                        elif '姓名' in placeholder:
                            inp.fill(name)
                        elif '年龄' in placeholder:
                            inp.fill(age)
                        elif '密码' in placeholder and input_type == 'password':
                            inp.fill(password)
                        elif '手机' in placeholder:
                            inp.fill(phone)
                        elif '邮箱' in placeholder:
                            inp.fill(email)
                    except:
                        pass
                
                register_btn = page.query_selector('button:has-text("注册")')
                if register_btn:
                    register_btn.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(3)
            
            print("\n" + "=" * 60)
            print("3. 登录...")
            
            page.goto("http://39.107.109.8:8082/login", timeout=30000)
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            
            all_inputs = page.query_selector_all('input')
            for inp in all_inputs:
                try:
                    placeholder = inp.get_attribute('placeholder') or ""
                    input_type = inp.get_attribute('type') or ""
                    
                    if '账号' in placeholder:
                        inp.fill(username)
                    elif '密码' in placeholder and input_type == 'password':
                        inp.fill(password)
                except:
                    pass
            
            login_btn = page.query_selector('button:has-text("登录")')
            if login_btn:
                login_btn.click()
                page.wait_for_load_state("networkidle")
                time.sleep(3)
            
            print(f"登录后URL: {page.url}")
            
            print("\n" + "=" * 60)
            print("4. 点击编辑按钮...")
            
            edit_btn = page.query_selector('button:has-text("编辑")')
            if edit_btn:
                edit_btn.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                
                print(f"编辑页面URL: {page.url}")
                
                print("\n查找所有输入框...")
                all_inputs = page.query_selector_all('input')
                print(f"输入框数量: {len(all_inputs)}")
                for i, inp in enumerate(all_inputs):
                    try:
                        placeholder = inp.get_attribute('placeholder')
                        input_type = inp.get_attribute('type')
                        name_attr = inp.get_attribute('name')
                        value = inp.input_value()
                        print(f"  输入框{i+1}: name='{name_attr}', placeholder='{placeholder}', type='{input_type}', value='{value}'")
                    except:
                        pass
                
                print("\n查找所有按钮...")
                all_buttons = page.query_selector_all('button')
                print(f"按钮数量: {len(all_buttons)}")
                for i, btn in enumerate(all_buttons):
                    try:
                        text = btn.inner_text()
                        btn_class = btn.get_attribute('class')
                        print(f"  按钮{i+1}: '{text}', class='{btn_class}'")
                    except:
                        pass
                
                print("\n页面HTML片段...")
                body_html = page.evaluate("document.body.innerHTML.substring(0, 4000)")
                print(body_html)
                
                print("\n" + "=" * 60)
                print("5. 测试修改字段...")
                
                name_input = None
                for inp in all_inputs:
                    try:
                        placeholder = inp.get_attribute('placeholder') or ""
                        if '姓名' in placeholder:
                            name_input = inp
                            break
                    except:
                        pass
                
                if name_input:
                    print("找到姓名输入框，尝试修改...")
                    original_value = name_input.input_value()
                    print(f"原值: {original_value}")
                    
                    name_input.fill("")
                    name_input.fill("测试XSS<script>alert(1)</script>")
                    print("已填写测试值: 测试XSS<script>alert(1)</script>")
                    
                    save_btn = page.query_selector('button:has-text("保存")') or page.query_selector('button:has-text("提交")')
                    if save_btn:
                        print("点击保存按钮...")
                        save_btn.click()
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)
                        
                        print(f"保存后URL: {page.url}")
                        
                        error_el = page.query_selector('.error-message')
                        if error_el:
                            print(f"错误信息: {error_el.inner_text()}")
                        
                        success_el = page.query_selector('.success-message')
                        if success_el:
                            print(f"成功信息: {success_el.inner_text()}")
                        
                        print("\n保存后页面HTML...")
                        body_html = page.evaluate("document.body.innerHTML.substring(0, 2000)")
                        print(body_html)
                    else:
                        print("未找到保存按钮")
                else:
                    print("未找到姓名输入框")
            
            print("\n" + "=" * 60)
            print("调试完成，浏览器将在20秒后关闭...")
            time.sleep(20)
            
        except Exception as e:
            print(f"发生错误: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(30)
        finally:
            browser.close()

if __name__ == "__main__":
    debug_site()
