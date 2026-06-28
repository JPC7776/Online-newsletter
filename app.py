#!/usr/bin/env python3
import os
import base64
from flask import Flask, render_template, request, jsonify, send_file, Response
from generate_newsletter import generate_newsletter

app = Flask(__name__, static_folder='static')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_FILE = os.path.join(BASE_DIR, "Newsletter_Ready.html")
TEMPLATE_FILE = os.path.join(BASE_DIR, "Untitled Template.html")
LOGO_FILE = os.path.join(BASE_DIR, "imgi_3_logo-LG-WEB-2024.png")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(os.path.join(BASE_DIR, "static"), exist_ok=True)

@app.route("/")
def index():
    # Encode logo as base64 for embedding in template
    logo_b64 = ""
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode("utf-8")
    return render_template("index.html", logo_b64=logo_b64)

@app.route("/generate", methods=["POST"])
def generate():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    template_id = request.form.get("template", "executive-dark")
    greeting_style = request.form.get("greeting", "none")

    docx_path = os.path.join(UPLOAD_DIR, file.filename)
    file.save(docx_path)

    try:
        generate_newsletter(TEMPLATE_FILE, docx_path, OUTPUT_FILE, template_id=template_id, greeting_style=greeting_style)
        return jsonify({"status": "success"})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/preview")
def preview():
    if os.path.exists(OUTPUT_FILE):
        return send_file(OUTPUT_FILE)
    return "No newsletter compiled yet.", 404

@app.route("/raw_html")
def raw_html():
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
            return Response(f.read(), mimetype="text/plain")
    return "No html", 404

@app.route("/download")
def download():
    if os.path.exists(OUTPUT_FILE):
        return send_file(OUTPUT_FILE, as_attachment=True, download_name="Mailchimp_Newsletter_Ready.html")
    return "No file", 404

if __name__ == "__main__":
    print("")
    print("  ╔══════════════════════════════════════════════════╗")
    print("  ║  ✦  LodgeGate Newsletter Studio                 ║")
    print("  ║     Running at: http://127.0.0.1:5050            ║")
    print("  ╚══════════════════════════════════════════════════╝")
    print("")
    app.run(host="127.0.0.1", port=5050, debug=True)
