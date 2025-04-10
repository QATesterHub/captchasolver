import os, re, time, random
from flask import Flask, request, render_template, send_file, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

progress = {'total': 0, 'current': 0, 'done': False}

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))

def run_bot(excel_path):
    global progress
    wb = load_workbook(excel_path)
    sh = wb.active
    driver = webdriver.Chrome()
    driver.maximize_window()

    total_rows = sh.max_row - 1
    progress.update({'total': total_rows, 'current': 0, 'done': False})

    for i in range(2, sh.max_row + 1):
        try:
            url, title, desc, name, email = [sh[f'{col}{i}'].value for col in 'ABCDE']
            if not all([url, title, desc, name, email]):
                sh[f'F{i}'] = '❌ Missing data'
                continue

            driver.get(url)
            time.sleep(random.uniform(2, 3))

            human_type(driver.find_element(By.ID, 'TITLE'), title)
            human_type(driver.find_element(By.ID, 'URL'), url)
            human_type(driver.find_element(By.ID, 'DESCRIPTION'), desc)
            human_type(driver.find_element(By.ID, 'OWNER_NAME'), name)
            human_type(driver.find_element(By.ID, 'OWNER_EMAIL'), email)

            for font in driver.find_elements(By.TAG_NAME, 'font'):
                if '=' in font.text:
                    expr = font.text.replace('x', '*').replace('=', '')
                    result = str(eval(expr))
                    human_type(driver.find_element(By.ID, 'DO_MATH'), result)
                    break

            driver.find_element(By.XPATH, "//input[@id='AGREERULES']").click()
            time.sleep(1)
            driver.find_element(By.NAME, 'continue').click()
            time.sleep(3)

            body = driver.find_element(By.TAG_NAME, 'body').text
            link = re.search(r'https?://\S+\.html', body)
            sh[f'F{i}'] = '✅ Posted'
            sh[f'G{i}'] = link.group() if link else 'No link'

        except Exception as e:
            sh[f'F{i}'] = '❌ Failed'

        progress['current'] += 1

    driver.quit()
    progress['done'] = True

    result_path = os.path.join(RESULT_FOLDER, f"result_{int(time.time())}.xlsx")
    wb.save(result_path)
    return result_path

@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file:
        return "No file uploaded."

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    result = run_bot(filepath)
    return send_file(result, as_attachment=True)

@app.route('/progress')
def get_progress():
    return jsonify(progress)

# ✅ PORT-compatible main block
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
