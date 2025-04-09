import os, re, time
from flask import Flask, request, render_template, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return "No file uploaded."
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        result_file = run_bot(path)
        return send_file(result_file, as_attachment=True)

    return render_template('index.html')

def run_bot(excel_path):
    wb = load_workbook(excel_path)
    sh = wb.active
    driver = webdriver.Chrome()
    driver.maximize_window()

    count = 0
    for i in range(2, sh.max_row + 1):
        if count == 7:
            break

        try:
            data = [sh[f'{col}{i}'].value for col in 'ABCDE']
            if not all(data):
                sh[f'F{i}'] = '❌ Missing data'
                continue

            url, title, desc, name, email = data
            driver.get("https://ebay-dir.com/submit?c=51&LINK_TYPE=1")

            driver.find_element(By.ID, 'TITLE').send_keys(title)
            driver.find_element(By.ID, 'URL').send_keys(url)
            driver.find_element(By.ID, 'DESCRIPTION').send_keys(desc)
            driver.find_element(By.ID, 'OWNER_NAME').send_keys(name)
            driver.find_element(By.ID, 'OWNER_EMAIL').send_keys(email)

            for f in driver.find_elements(By.TAG_NAME, 'font'):
                if '=' in f.text:
                    expr = f.text.replace('x', '*').replace('=', '')
                    driver.find_element(By.ID, "DO_MATH").send_keys(str(eval(expr)))
                    break

            driver.find_element(By.XPATH, "//input[@id='AGREERULES']").click()
            driver.find_element(By.NAME, 'continue').click()
            time.sleep(2)

            body = driver.find_element(By.TAG_NAME, 'body').text
            link = re.search(r'https?://\S+\.html', body)
            sh[f'F{i}'] = '✅ Posted'
            sh[f'G{i}'] = link.group() if link else 'No link'

            count += 1
        except Exception as e:
            sh[f'F{i}'] = '❌ Failed'

    driver.quit()
    result_path = os.path.join(RESULT_FOLDER, f"result_{int(time.time())}.xlsx")
    wb.save(result_path)
    return result_path

if __name__ == '__main__':
    app.run(debug=True)
