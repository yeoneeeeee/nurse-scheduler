from flask import Flask, render_template, request, send_file
import os
import subprocess
import sys

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    success = False
    last_error = ""
     
    for i in range(100):
        result = subprocess.run(
            [sys.executable, "scheduler.py"],
            capture_output = True
            text = True
        )

        if result.returncode == 0:
            sucess = True
            break

        last_error = result.stdout + "\n" + result.stderr
    
    if not success:
        return f"""
        <h2>조건 만족 근무표 생성 실패</h2>
        <p>100번 시도했지만 조건을 만족하는 표를 만들지 못했습니다.</p>
        <pre>{last_error}</pre>
        """, 500
    

    if result.returncode != 0:
        return f"""
        <h2>근무표 생성 중 오류 발생</h2>
        <pre>{result.stderr}</pre>
        """, 500

    if not os.path.exists("schedule_output.xlsx"):
        return "schedule_output.xlsx 파일이 생성되지 않았습니다.", 500

    return send_file(
        "schedule_output.xlsx",
        as_attachment=True,
        download_name="schedule_output.xlsx"
    )

if __name__ == "__main__":
    app.run(debug=True)
