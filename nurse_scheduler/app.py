from flask import Flask, render_template, request, send_file
import os
import shutil
import subprocess

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():

    file = request.files["file"]

    upload_path = os.path.join(UPLOAD_FOLDER, "template.xlsx")
    file.save(upload_path)

    # scheduler.py 실행
    subprocess.run(["python", "scheduler.py"])

    output_file = "schedule_output.xlsx"

    return send_file(
        output_file,
        as_attachment=True
    )

if __name__ == "__main__":
    app.run(debug=True)