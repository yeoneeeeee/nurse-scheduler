from flask import Flask, render_template, request, send_file
import subprocess
import sys
import os

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():

    if "file" not in request.files:
        return "파일 없음", 400

    file = request.files["file"]

    # 업로드 저장
    file.save("template.xlsx")

    success = False
    last_error = ""

    # 조건 만족할 때까지 최대 100번 시도
    for i in range(100):

        print(f"\n {i+1}번째 스케줄 생성 시도 중...")

        result = subprocess.run(
            [sys.executable, "scheduler.py"],
            capture_output=True,
            text=True
        )

        if result.returncode == 0:
            print("모든 조건 만족 스케줄 생성 성공")
            print(result.stdout)
            success = True
            break

        else:
            print("조건 불만족-> 자동 재생성")
            print(result.stdout)
            print(result.stderr)

        last_error = result.stdout + "\n" + result.stderr

    if not success:
        return f"""
        <h2>조건 만족 스케줄 생성 실패</h2>
        <pre>{last_error}</pre>
        """, 500

    if not os.path.exists("schedule_output.xlsx"):
        return "엑셀 파일 생성 실패", 500

    return send_file(
        "schedule_output.xlsx",
        as_attachment=True
    )

if __name__ == "__main__":
    app.run(debug=True)
