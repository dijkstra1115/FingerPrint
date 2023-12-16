from unicodedata import name
from flask import Flask, request, render_template, send_from_directory, redirect, url_for
from modules.report_generator import generate_report
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/submit-personality-test', methods=['POST'])
def submit_personality_test():
    if request.method == 'POST':
        # 獲取表單數據
        data = {}
        fingers = ['L1', 'L2', 'L3', 'L4', 'L5', 'R1', 'R2', 'R3', 'R4', 'R5']

        for finger in fingers:
            data[f'{finger}_code'] = request.form.get(f'{finger}_code')
            data[f'{finger}_left_value'] = request.form.get(f'{finger}_left_value')
            data[f'{finger}_right_value'] = request.form.get(f'{finger}_right_value')

        user_name = request.form.get('user_name')
        pricing_plan = request.form.get('pricing_plan')

        # 在這裡處理這些數據，例如存儲到數據庫或進行計算

        generate_report(user_name, pricing_plan, data)

        file_path = os.path.join("download", f"{user_name}_{pricing_plan}.docx")

        return render_template('download.html', file_path = file_path)

        # 處理完後重定向到某個頁面，或返回結果
        # return redirect(url_for('index'))  # 重定向回主頁面

@app.route('/download/<path:filename>', methods=['GET'])
def download_file(filename):
    dirname = os.path.dirname(os.path.realpath(__file__))
    output_dir = os.path.join(dirname, "output")
    return send_from_directory(output_dir, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
