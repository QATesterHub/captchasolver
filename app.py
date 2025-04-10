import os, re, time
from flask import Flask, request, render_template, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return "No file uploaded."

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        run_bot(path)
        return send_file(path, as_attachment=True)

    return render_template('index.html')

def slow_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(0.05)

def run_bot(excel_path):
    wb = load_workbook(excel_path)
    sh = wb.active
    driver = webdriver.Chrome()
    driver.maximize_window()

    for i in range(2, sh.max_row + 1):
        try:
            url, title, desc, name, email = [sh[f'{col}{i}'].value for col in 'ABCDE']
            if not all([url, title, desc, name, email]):
                sh[f'F{i}'] = '❌ Missing data'
                continue

            # Use hardcoded URL always
            driver.get("https://ebay-dir.com/submit?c=51&LINK_TYPE=1")
            time.sleep(2)

            slow_type(driver.find_element(By.ID, 'TITLE'), title)
            slow_type(driver.find_element(By.ID, 'URL'), url)
            slow_type(driver.find_element(By.ID, 'DESCRIPTION'), desc)
            slow_type(driver.find_element(By.ID, 'OWNER_NAME'), name)
            slow_type(driver.find_element(By.ID, 'OWNER_EMAIL'), email)

            for font in driver.find_elements(By.TAG_NAME, 'font'):
                if '=' in font.text:
                    expr = font.text.replace('x', '*').replace('=', '')
                    result = str(eval(expr))
                    driver.find_element(By.ID, "DO_MATH").send_keys(result)
                    break

            # ✅ Always include this line
            driver.find_element(By.XPATH, "//input[@id='AGREERULES']").click()

            driver.find_element(By.NAME, 'continue').click()
            time.sleep(2)

            body = driver.find_element(By.TAG_NAME, 'body').text
            link = re.search(r'https?://\S+\.html', body)
            sh[f'F{i}'] = '✅ Posted'
            sh[f'G{i}'] = link.group() if link else 'No link'

        except Exception as e:
            sh[f'F{i}'] = '❌ Failed'

    wb.save(excel_path)
    driver.quit()

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
