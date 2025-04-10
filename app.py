import os, re, time, random
from flask import Flask, request, render_template, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))  # Simulate human typing

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

            driver.find_element(By.ID, 'AGREERULES').click()
            time.sleep(1)
            driver.find_element(By.NAME, 'continue').click()
            time.sleep(3)

            body = driver.find_element(By.TAG_NAME, 'body').text
            link = re.search(r'https?://\S+\.html', body)
            sh[f'F{i}'] = '✅ Posted'
            sh[f'G{i}'] = link.group() if link else 'No link'
        except Exception as e:
            sh[f'F{i}'] = '❌ Failed'

    driver.quit()
    result_path = os.path.join(RESULT_FOLDER, f"result_{int(time.time())}.xlsx")
    wb.save(result_path)
    return result_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            return "No file uploaded."

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        result = run_bot(filepath)
        return send_file(result, as_attachment=True)

    return '''
        <h2>Upload Excel File for Ad Posting</h2>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file" required>
            <input type="submit" value="Start Posting">
        </form>
    '''

# ✅ PORT-compatible main block
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Use Render’s PORT or default to 5000
    app.run(host='0.0.0.0', port=port)
