import os, re, time, threading
from flask import Flask, request, render_template, send_file, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook

app = Flask(__name__)
os.makedirs('uploads', exist_ok=True)
progress = {"total": 0, "current": 0, "status": ""}

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    path = os.path.join('uploads', file.filename)
    file.save(path)
    threading.Thread(target=run_bot, args=(path,)).start()
    return jsonify({"message": "Processing started"})

@app.route('/progress')
def check_progress():
    return jsonify(progress)

@app.route('/download/<name>')
def download(name):
    return send_file(os.path.join('uploads', name), as_attachment=True)

def run_bot(path):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=chrome_options)

    wb = load_workbook(path)
    sh = wb.active
    progress.update(total=sh.max_row - 1, current=0, status="Running")

    for i in range(2, sh.max_row + 1):
        data = [sh[f'{c}{i}'].value for c in 'ABCDE']
        if not all(data): break
        try:
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

            driver.find_element(By.ID, 'AGREERULES').click()
            driver.find_element(By.NAME, 'continue').click()
            time.sleep(1)

            body = driver.find_element(By.TAG_NAME, 'body').text
            link = re.search(r'https?://\S+\.html', body)
            sh[f'F{i}'] = '✅ Posted'
            sh[f'G{i}'] = link.group() if link else 'No link'
        except:
            sh[f'F{i}'] = '❌ Failed'
        progress['current'] += 1

    driver.quit()
    wb.save(path)
    progress['status'] = 'Done'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050)
