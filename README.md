# Mailchimp Automated Bilingual Newsletter Studio

An "Uncle-Proof" automated tool to convert LodgeGate & RezXS Microsoft Word drafts (`.docx`) into responsive, bilingual Mailchimp HTML templates.

## Quickstart (Monthly Use)

1. Double click **`Start_Newsletter_Studio.command`** inside this folder.
2. Your browser will automatically open the Studio at `http://127.0.0.1:5050`.
3. Drag your monthly `.docx` file onto the screen and click **✦ Generate Mailchimp Template**.
4. Click **📋 Copy Code for Mailchimp** and paste it into Mailchimp under *Code your own -> Paste in code*.

## Files in this Directory
- `Start_Newsletter_Studio.command` — Double click Mac app launcher
- `app.py` — Local web studio server
- `generate_newsletter.py` — Standalone docx parser & HTML generator engine
- `templates/index.html` — Studio UI aesthetic layout
- `Untitled Template.html` — Custom Mailchimp base design
- `Newsletter LodgeGate 2026.docx` — Sample monthly input draft
- `Newsletter_LodgeGate_2026_Ready.html` — Final compiled output template
