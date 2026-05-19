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
    if "file" not in request.files:
        return "파일이 업로드되지 않았습니다.", 400

    file = request.files["file"]

    # scheduler.py가 읽는 이름으로 저장
    file.save("template.xlsx")

    # scheduler.py 실행
    result = subprocess.run(
        [sys.executable, "scheduler.py"],
        capture_output=True,
        text=True
    )

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
