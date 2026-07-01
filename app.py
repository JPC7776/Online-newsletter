#!/usr/bin/env python3
import os
import re
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
    hero_image_url = request.form.get("hero_image_url", "").strip()
    article_images_raw = request.form.get("article_images", "").strip()
    custom_images = [u.strip() for u in article_images_raw.splitlines() if u.strip()] if article_images_raw else None

    docx_path = os.path.join(UPLOAD_DIR, file.filename)
    file.save(docx_path)

    try:
        generate_newsletter(TEMPLATE_FILE, docx_path, OUTPUT_FILE, template_id=template_id, greeting_style=greeting_style, custom_images=custom_images, hero_image_url=hero_image_url)
        return jsonify({"status": "success"})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/get_images")
def get_images():
    if not os.path.exists(OUTPUT_FILE):
        return jsonify({"error": "No newsletter compiled yet"}), 404
    
    with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    images = []
    hero_match = re.search(r"background-image:url\(['\"]?([^'\")]+)['\"]?\)", html)
    if hero_match:
        images.append({
            "id": "hero_bg",
            "label": "🖼️ Hero Banner Achtergrond",
            "url": hero_match.group(1),
            "type": "hero"
        })

    for img_tag in re.finditer(r'<img\b[^>]+>', html, re.IGNORECASE):
        tag = img_tag.group(0)
        mc_match = re.search(r'mc:edit="([^"]+)"', tag)
        src_match = re.search(r'src="([^"]+)"', tag)
        alt_match = re.search(r'alt="([^"]*)"', tag)
        if mc_match and src_match:
            edit_key = mc_match.group(1)
            if edit_key.startswith("header_logo"):
                continue
            url = src_match.group(1)
            alt = alt_match.group(1) if alt_match else ""
            lang_flag = "🇳🇱 NL" if "_nl_" in edit_key else ("🇬🇧 EN" if "_en_" in edit_key else "🖼️")
            label = f"{lang_flag} Artikel: {alt}" if alt else f"{lang_flag} Foto ({edit_key})"
            images.append({
                "id": edit_key,
                "label": label,
                "url": url,
                "type": "article"
            })

    return jsonify({"status": "success", "images": images})

@app.route("/update_images", methods=["POST"])
def update_images():
    if not os.path.exists(OUTPUT_FILE):
        return jsonify({"error": "No newsletter compiled yet"}), 404
    
    data = request.get_json()
    if not data or "updates" not in data:
        return jsonify({"error": "Invalid updates data"}), 400

    updates = data["updates"]
    with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    if "hero_bg" in updates and updates["hero_bg"].strip():
        new_url = updates["hero_bg"].strip()
        html = re.sub(
            r"(background-image:url\(['\"]?)[^'\")]+(['\"]?\))",
            lambda m: f"{m.group(1)}{new_url}{m.group(2)}",
            html
        )
        html = re.sub(
            r'(<v:fill\b[^>]*?\bsrc=["\'])[^"\']+(["\'])',
            lambda m: f"{m.group(1)}{new_url}{m.group(2)}",
            html,
            flags=re.IGNORECASE
        )

    def replace_img_tag(match):
        tag = match.group(0)
        mc_match = re.search(r'mc:edit=["\']([^"\']+)["\']', tag)
        if mc_match:
            key = mc_match.group(1)
            if key in updates and updates[key].strip():
                new_url = updates[key].strip()
                return re.sub(r'src=["\'][^"\']+["\']', lambda m: f'src="{new_url}"', tag)
        return tag

    html = re.sub(r'<img\b[^>]+>', replace_img_tag, html, flags=re.IGNORECASE)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    return jsonify({"status": "success"})

@app.route("/preview")
def preview():
    if os.path.exists(OUTPUT_FILE):
        response = send_file(OUTPUT_FILE)
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response
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
