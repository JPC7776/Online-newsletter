#!/usr/bin/env python3
"""
LodgeGate Newsletter Generator v2.0
Converts .docx → Mailchimp-ready HTML with 8 selectable template themes.
Features:
- 8 Premium Color Themes (Executive Dark, Sunrise, Ocean, Minimal, Midnight, Spring, Magazine, Slate)
- Sequential Layout (Dutch NL top → jump link → English EN bottom)
- 100% Unique Image Pool (No repeated photos)
- Display Variations (Stacked card, 2-col split, Highlight box)
- Deliverability Sanitizer (Strips <script> tags & leftover brackets)
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══════════════════════════════════════════════════════════════════
# 8 TEMPLATE THEME DEFINITIONS
# ═══════════════════════════════════════════════════════════════════

TEMPLATES = {
    'executive-dark': {
        'name': 'Executive Dark',
        'header_bg': '#121A2F',
        'header_text': '#ffffff',
        'preview_bar_bg': '#0056B3',
        'hero_bg': '#121A2F',
        'hero_overlay': 'rgba(18,26,47,0.82)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.85)',
        'accent': '#FF7A00',
        'accent_hover': '#E86E00',
        'btn_text': '#ffffff',
        'section_bg_a': '#F8F9FA',
        'section_bg_b': '#ffffff',
        'heading_color': '#121A2F',
        'body_text': '#444444',
        'body_text_alt': '#555555',
        'card_bg': '#ffffff',
        'card_border': '#FF7A00',
        'card_shadow': 'rgba(18,26,47,0.06)',
        'en_hero_bg': '#0D1320',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#8c96ab',
        'footer_bg': '#0D1320',
        'footer_text': '#999999',
        'footer_link': '#FF7A00',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#121A2F',
    },
    'sunrise-warm': {
        'name': 'Sunrise Warm',
        'header_bg': '#FF7A00',
        'header_text': '#ffffff',
        'preview_bar_bg': '#E86E00',
        'hero_bg': '#FF7A00',
        'hero_overlay': 'rgba(255,122,0,0.85)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.9)',
        'accent': '#121A2F',
        'accent_hover': '#0D1320',
        'btn_text': '#ffffff',
        'section_bg_a': '#FFF8F0',
        'section_bg_b': '#ffffff',
        'heading_color': '#2D1B00',
        'body_text': '#5A4020',
        'body_text_alt': '#6B5030',
        'card_bg': '#ffffff',
        'card_border': '#FF7A00',
        'card_shadow': 'rgba(255,122,0,0.08)',
        'en_hero_bg': '#2D1B00',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#C4A882',
        'footer_bg': '#2D1B00',
        'footer_text': '#A08060',
        'footer_link': '#FF9F43',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#2D1B00',
    },
    'ocean-blue': {
        'name': 'Ocean Blue',
        'header_bg': '#0056B3',
        'header_text': '#ffffff',
        'preview_bar_bg': '#003D80',
        'hero_bg': '#0056B3',
        'hero_overlay': 'rgba(0,86,179,0.85)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.9)',
        'accent': '#FF7A00',
        'accent_hover': '#E86E00',
        'btn_text': '#ffffff',
        'section_bg_a': '#F0F7FF',
        'section_bg_b': '#ffffff',
        'heading_color': '#002B5C',
        'body_text': '#334155',
        'body_text_alt': '#475569',
        'card_bg': '#ffffff',
        'card_border': '#0056B3',
        'card_shadow': 'rgba(0,86,179,0.08)',
        'en_hero_bg': '#002B5C',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#7DA8D4',
        'footer_bg': '#001F42',
        'footer_text': '#6B8DB8',
        'footer_link': '#FF7A00',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#002B5C',
    },
    'clean-minimal': {
        'name': 'Clean Minimal',
        'header_bg': '#ffffff',
        'header_text': '#121A2F',
        'preview_bar_bg': '#FF7A00',
        'hero_bg': '#ffffff',
        'hero_overlay': 'rgba(255,255,255,0.92)',
        'hero_text': '#121A2F',
        'hero_sub': '#555555',
        'accent': '#FF7A00',
        'accent_hover': '#E86E00',
        'btn_text': '#ffffff',
        'section_bg_a': '#FAFAFA',
        'section_bg_b': '#ffffff',
        'heading_color': '#111111',
        'body_text': '#444444',
        'body_text_alt': '#555555',
        'card_bg': '#ffffff',
        'card_border': '#FF7A00',
        'card_shadow': 'rgba(0,0,0,0.06)',
        'en_hero_bg': '#121A2F',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#8c96ab',
        'footer_bg': '#1A1A1A',
        'footer_text': '#888888',
        'footer_link': '#FF7A00',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#121A2F',
    },
    'midnight-luxe': {
        'name': 'Midnight Luxe',
        'header_bg': '#0D0D1A',
        'header_text': '#C9A84C',
        'preview_bar_bg': '#1A1A3E',
        'hero_bg': '#0D0D1A',
        'hero_overlay': 'rgba(13,13,26,0.88)',
        'hero_text': '#C9A84C',
        'hero_sub': 'rgba(201,168,76,0.7)',
        'accent': '#C9A84C',
        'accent_hover': '#B8973F',
        'btn_text': '#0D0D1A',
        'section_bg_a': '#14142B',
        'section_bg_b': '#1A1A3E',
        'heading_color': '#E8DCC8',
        'body_text': '#A0A0B8',
        'body_text_alt': '#8888A0',
        'card_bg': '#1E1E3F',
        'card_border': '#C9A84C',
        'card_shadow': 'rgba(201,168,76,0.08)',
        'en_hero_bg': '#08081A',
        'en_hero_text': '#C9A84C',
        'en_hero_sub': '#6B6B90',
        'footer_bg': '#08081A',
        'footer_text': '#555570',
        'footer_link': '#C9A84C',
        'divider': 'rgba(201,168,76,0.15)',
        'dark_btn_bg': '#0D0D1A',
    },
    'fresh-spring': {
        'name': 'Fresh Spring',
        'header_bg': '#10B981',
        'header_text': '#ffffff',
        'preview_bar_bg': '#059669',
        'hero_bg': '#10B981',
        'hero_overlay': 'rgba(16,185,129,0.85)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.9)',
        'accent': '#121A2F',
        'accent_hover': '#0D1320',
        'btn_text': '#ffffff',
        'section_bg_a': '#F0FFF8',
        'section_bg_b': '#ffffff',
        'heading_color': '#064E3B',
        'body_text': '#374151',
        'body_text_alt': '#4B5563',
        'card_bg': '#ffffff',
        'card_border': '#10B981',
        'card_shadow': 'rgba(16,185,129,0.08)',
        'en_hero_bg': '#064E3B',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#6EE7B7',
        'footer_bg': '#022C22',
        'footer_text': '#6EE7B7',
        'footer_link': '#34D399',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#064E3B',
    },
    'bold-magazine': {
        'name': 'Bold Magazine',
        'header_bg': '#E63E2A',
        'header_text': '#ffffff',
        'preview_bar_bg': '#B91C1C',
        'hero_bg': '#E63E2A',
        'hero_overlay': 'rgba(230,62,42,0.88)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.9)',
        'accent': '#121A2F',
        'accent_hover': '#0D1320',
        'btn_text': '#ffffff',
        'section_bg_a': '#FFF5F5',
        'section_bg_b': '#ffffff',
        'heading_color': '#7F1D1D',
        'body_text': '#444444',
        'body_text_alt': '#555555',
        'card_bg': '#ffffff',
        'card_border': '#E63E2A',
        'card_shadow': 'rgba(230,62,42,0.08)',
        'en_hero_bg': '#7F1D1D',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#FCA5A5',
        'footer_bg': '#450A0A',
        'footer_text': '#FCA5A5',
        'footer_link': '#FB7185',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#7F1D1D',
    },
    'slate-pro': {
        'name': 'Slate Pro',
        'header_bg': '#374151',
        'header_text': '#ffffff',
        'preview_bar_bg': '#1F2937',
        'hero_bg': '#374151',
        'hero_overlay': 'rgba(55,65,81,0.85)',
        'hero_text': '#ffffff',
        'hero_sub': 'rgba(255,255,255,0.85)',
        'accent': '#FF7A00',
        'accent_hover': '#E86E00',
        'btn_text': '#ffffff',
        'section_bg_a': '#F9FAFB',
        'section_bg_b': '#ffffff',
        'heading_color': '#111827',
        'body_text': '#4B5563',
        'body_text_alt': '#6B7280',
        'card_bg': '#ffffff',
        'card_border': '#FF7A00',
        'card_shadow': 'rgba(0,0,0,0.06)',
        'en_hero_bg': '#1F2937',
        'en_hero_text': '#ffffff',
        'en_hero_sub': '#9CA3AF',
        'footer_bg': '#111827',
        'footer_text': '#9CA3AF',
        'footer_link': '#FF7A00',
        'divider': 'rgba(255,255,255,0.1)',
        'dark_btn_bg': '#1F2937',
    },
}

# ═══════════════════════════════════════════════════════════════════
# GREETING STYLE DEFINITIONS (Mailchimp merge tags)
# ═══════════════════════════════════════════════════════════════════

GREETINGS = {
    'formal-nl': {
        'name': 'Formeel (NL)',
        'nl': 'Beste *|FNAME|*,',
        'en': 'Dear *|FNAME|*,',
        'nl_fallback': 'Beste relatie,',
        'en_fallback': 'Dear valued partner,',
    },
    'casual-nl': {
        'name': 'Informeel (NL)',
        'nl': 'Hoi *|FNAME|*! 👋',
        'en': 'Hi *|FNAME|*! 👋',
        'nl_fallback': 'Hoi! 👋',
        'en_fallback': 'Hi there! 👋',
    },
    'formal-en': {
        'name': 'Formal (EN)',
        'nl': 'Geachte *|FNAME|*,',
        'en': 'Dear *|FNAME|*,',
        'nl_fallback': 'Geachte heer/mevrouw,',
        'en_fallback': 'Dear Sir/Madam,',
    },
    'casual-en': {
        'name': 'Casual (EN)',
        'nl': 'Hey *|FNAME|*!',
        'en': 'Hey *|FNAME|*!',
        'nl_fallback': 'Hey!',
        'en_fallback': 'Hey there!',
    },
    'bilingual': {
        'name': 'Bilingual',
        'nl': 'Beste *|FNAME|* / Dear *|FNAME|*,',
        'en': 'Dear *|FNAME|* / Beste *|FNAME|*,',
        'nl_fallback': 'Beste relatie / Dear partner,',
        'en_fallback': 'Dear partner / Beste relatie,',
    },
    'none': {
        'name': 'No Greeting',
        'nl': '',
        'en': '',
        'nl_fallback': '',
        'en_fallback': '',
    },
}


def parse_docx(docx_path):
    """Extract paragraphs and group into logical articles."""
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Cannot find Word document at: {docx_path}")

    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read('word/document.xml')
        tree = ET.fromstring(xml_content)

        paragraphs = []
        for p in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
            texts = [t.text for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text]
            full_text = "".join(texts).strip()
            if full_text:
                paragraphs.append(full_text)

    articles = []
    current_article = None

    for p in paragraphs:
        if p.lower() in ['newsletter lodgegate', 'nieuwsbrief lodgegate']:
            continue

        is_heading = (len(p) < 40 and not p.endswith('.') and not p.endswith(',')) or \
                     p.lower() in ['veiligheid', 'rms', 'lodgegate web']

        if is_heading:
            if current_article:
                articles.append(current_article)
            current_article = {'title_nl': p, 'paragraphs_nl': []}
        else:
            if current_article is None:
                current_article = {'title_nl': 'Update', 'paragraphs_nl': []}
            current_article['paragraphs_nl'].append(p)

    if current_article:
        articles.append(current_article)

    return articles


def get_unique_image_pool():
    """Returns 6 strictly distinct Unsplash image URLs — no repeats across NL and EN."""
    return {
        'nl': {
            'veiligheid': 'https://images.unsplash.com/photo-1550751827-4bd374c3f58b?auto=format&fit=crop&w=1000&q=80',
            'rms': 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=1000&q=80',
            'lodgegate web': 'https://images.unsplash.com/photo-1460925895917-afdab827c52f?auto=format&fit=crop&w=1000&q=80',
        },
        'en': {
            'veiligheid': 'https://images.unsplash.com/photo-1563986768609-322da13575f3?auto=format&fit=crop&w=1000&q=80',
            'rms': 'https://images.unsplash.com/photo-1504868584819-f8e8b4b6d7e3?auto=format&fit=crop&w=1000&q=80',
            'lodgegate web': 'https://images.unsplash.com/photo-1522542550221-31fd19575a2d?auto=format&fit=crop&w=1000&q=80',
        },
    }


def enrich_article(article, lang='nl'):
    """Enriches an article for either Dutch or English edition."""
    title_nl = article['title_nl']
    text_nl = "\n\n".join(article['paragraphs_nl'])
    pool = get_unique_image_pool()[lang]

    curated_map = {
        'veiligheid': {
            'category_nl': 'BEVEILIGING & PRIVACY', 'category_en': 'SECURITY & PRIVACY',
            'title_nl': 'Optimaliseer uw gegevensbeveiliging met LodgeGate 2FA',
            'title_en': 'Enhanced Data Security & 2FA in LodgeGate',
            'summary_en': 'In an era of rising cyber risks, protecting guest and hotel data is paramount. LodgeGate provides robust security measures including Two-Factor Authentication (2FA), IP address whitelist blocking, and strict password enforcement policies to safeguard your hotel operations.',
            'link': 'https://lodgegate.com',
        },
        'rms': {
            'category_nl': 'TARIEFOPTIMALISATIE', 'category_en': 'REVENUE OPTIMIZATION',
            'title_nl': 'Haal het maximale uit uw kamerprijzen met RevControl & RoomPriceGenie',
            'title_en': 'Automated Price Optimization with RMS Integrations',
            'summary_en': 'Manual rate adjustments can no longer keep up with dynamic market demand. LodgeGate seamlessly integrates with leading Revenue Management Systems like RevControl and RoomPriceGenie to maximize your occupancy and RevPAR without administrative overhead.',
            'link': 'https://lodgegate.com',
        },
        'lodgegate web': {
            'category_nl': 'DIRECTE BOEKINGEN & WEB', 'category_en': 'DIRECT BOOKINGS & WEB',
            'title_nl': 'Professionele weboplossingen en IBE met LodgeGate WEB',
            'title_en': 'Boost Direct Bookings with RezXS Web Solutions',
            'summary_en': 'Your website is the digital front door to your property. With LodgeGate WEB solutions by RezXS, we deliver fully branded, lightning-fast hotel websites and integrated Internet Booking Engines (IBE). Reduce OTA commission dependence and convert visitors into direct bookings.',
            'link': 'https://rezxs.com',
        },
    }

    key = title_nl.lower().strip()
    fallback_img = 'https://images.unsplash.com/photo-1566073771259-6a8506099945?auto=format&fit=crop&w=1000&q=80'
    image = pool.get(key, fallback_img)

    if key in curated_map:
        info = curated_map[key]
        return {
            'category': info[f'category_{lang}'] if f'category_{lang}' in info else info.get('category_nl', 'UPDATE'),
            'title': info['title_nl'] if lang == 'nl' else info['title_en'],
            'text': text_nl if lang == 'nl' else info['summary_en'],
            'image': image,
            'link': info['link'],
        }

    is_web = any(k in key or k in text_nl.lower() for k in ['web', 'ibe', 'rezxs', 'website', 'boeken'])
    link = 'https://rezxs.com' if is_web else 'https://lodgegate.com'

    return {
        'category': 'UPDATE',
        'title': title_nl,
        'text': text_nl if lang == 'nl' else (article['paragraphs_nl'][0] if article['paragraphs_nl'] else title_nl),
        'image': image,
        'link': link,
    }


def render_article_variation(enriched, variation, btn_text, t):
    """Renders one of 3 distinct display layouts using theme colors."""
    formatted_text = enriched['text'].replace('\n\n', '<br><br>')
    bg_a = t['section_bg_a']
    bg_b = t['section_bg_b']

    if variation == 'stacked':
        bg = bg_a
        return f"""
                    <!-- ===== VARIATION A: STACKED BANNER ===== -->
                    <tr>
                        <td style="background-color:{bg};padding:45px 40px">
                            <a href="{enriched['link']}" target="_blank">
                                <img src="{enriched['image']}" alt="{enriched['title']}" width="520" style="width:100%;max-width:520px;border-radius:12px;margin-bottom:22px;display:block;object-fit:cover;max-height:260px" role="presentation">
                            </a>
                            <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;margin:0 0 8px;display:block">
                                ✦ {enriched['category']}
                            </span>
                            <h2 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:24px;font-weight:800;line-height:1.25;margin:0 0 14px">
                                {enriched['title']}
                            </h2>
                            <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.75;margin:0 0 22px">
                                {formatted_text}
                            </p>
                            <a href="{enriched['link']}" target="_blank" style="background-color:{t['accent']};border-radius:50px;color:{t['btn_text']};display:inline-block;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:14px 34px;text-decoration:none">
                                {btn_text} &rarr;
                            </a>
                        </td>
                    </tr>
"""
    elif variation == 'split':
        bg = bg_b
        return f"""
                    <!-- ===== VARIATION B: 2-COLUMN SPLIT ===== -->
                    <tr>
                        <td style="background-color:{bg};padding:45px 30px">
                            <table class="two-col" role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                <tr>
                                    <td style="width:45%;vertical-align:top;padding:10px">
                                        <a href="{enriched['link']}" target="_blank">
                                            <img src="{enriched['image']}" alt="{enriched['title']}" style="width:100%;border-radius:10px;object-fit:cover;height:240px;display:block" role="presentation">
                                        </a>
                                    </td>
                                    <td style="width:55%;vertical-align:top;padding:10px">
                                        <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;display:block;margin-bottom:6px">
                                            ✦ {enriched['category']}
                                        </span>
                                        <h3 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:20px;font-weight:800;line-height:1.3;margin:0 0 10px">
                                            {enriched['title']}
                                        </h3>
                                        <p style="color:{t['body_text_alt']};font-family:Arial,sans-serif;font-size:14px;line-height:1.65;margin:0 0 18px">
                                            {formatted_text[:280]}...
                                        </p>
                                        <a href="{enriched['link']}" target="_blank" style="color:{t['accent']};font-family:Arial,sans-serif;font-weight:700;text-decoration:none;font-size:14px">
                                            {btn_text} &rarr;
                                        </a>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
"""
    else:  # highlight card
        bg = bg_a
        return f"""
                    <!-- ===== VARIATION C: HIGHLIGHT PRODUCT CARD ===== -->
                    <tr>
                        <td style="background-color:{bg};padding:45px 40px">
                            <div style="background:{t['card_bg']};border-radius:16px;padding:32px 28px;border-top:5px solid {t['card_border']};box-shadow:0 8px 24px {t['card_shadow']}">
                                <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;margin:0 0 8px;display:block">
                                    ✦ {enriched['category']}
                                </span>
                                <h2 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:22px;font-weight:800;line-height:1.25;margin:0 0 14px">
                                    {enriched['title']}
                                </h2>
                                <a href="{enriched['link']}" target="_blank">
                                    <img src="{enriched['image']}" alt="{enriched['title']}" style="width:100%;border-radius:10px;margin-bottom:18px;max-height:220px;object-fit:cover;display:block" role="presentation">
                                </a>
                                <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.7;margin:0 0 22px">
                                    {formatted_text}
                                </p>
                                <a href="{enriched['link']}" target="_blank" style="background-color:{t['dark_btn_bg']};border-radius:50px;color:#ffffff;display:inline-block;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:12px 30px;text-decoration:none">
                                    {btn_text} &rarr;
                                </a>
                            </div>
                        </td>
                    </tr>
"""


def build_edition_blocks(articles, lang, t):
    """Compiles all article sections with cycling variations."""
    btn_text = "Lees Meer" if lang == 'nl' else "Read More"
    variations = ['stacked', 'split', 'highlight']

    html_out = ""
    for i, a in enumerate(articles):
        enriched = enrich_article(a, lang)
        var = variations[i % len(variations)]
        html_out += render_article_variation(enriched, var, btn_text, t)
    return html_out


def generate_newsletter(template_path, docx_path, output_path, template_id='executive-dark', greeting_style='none'):
    """Main execution pipeline."""
    raw_articles = parse_docx(docx_path)
    t = TEMPLATES.get(template_id, TEMPLATES['executive-dark'])
    g = GREETINGS.get(greeting_style, GREETINGS['none'])

    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Strip unwanted script tag injected by web previewers
    html = re.sub(r'<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', '', html, flags=re.IGNORECASE)

    # ── Apply theme to header ──
    html = re.sub(
        r'background-color:#121A2F;padding:18px 30px',
        f"background-color:{t['header_bg']};padding:18px 30px",
        html, count=1
    )

    # ── Preview Bar with jump link ──
    jump_bar_inner = f"""
                        <td class="preview-bar" style="background:{t['preview_bar_bg']};padding:10px 30px;text-align:center">
                            <span class="preview-text" style="color:#ffffff;font-family:Arial,Helvetica,sans-serif;font-size:12px;letter-spacing:0.06em">
                                🇳🇱 Nederlandse editie &nbsp;|&nbsp; <a href="#english-version" style="color:{t['accent']};text-decoration:underline;font-weight:700">🇬🇧 Read in English ↓</a>
                            </span>
                        </td>
"""
    html = re.sub(r'<td class="preview-bar".*?<\/td>', jump_bar_inner.strip(), html, flags=re.DOTALL)

    # ── Apply theme to hero section ──
    html = re.sub(
        r"background-color:#121A2F;background-image:url\('https://images\.unsplash\.com[^']*'\);background-size:cover;background-position:center;padding:70px 40px",
        f"background-color:{t['hero_bg']};background-image:url('https://images.unsplash.com/photo-1566073771259-6a8506099945?auto=format&fit=crop&w=1200&q=80');background-size:cover;background-position:center;padding:70px 40px",
        html, count=1
    )
    html = re.sub(
        r'background:rgba\(18,26,47,0\.82\)',
        f"background:{t['hero_overlay']}",
        html, count=1
    )

    # ── Hero CTA button ──
    html = re.sub(
        r"background:#FF7A00;border-radius:50px;color:#fff;display:inline-block;font-family:Arial,sans-serif;font-size:15px;font-weight:700;padding:15px 38px",
        f"background:{t['accent']};border-radius:50px;color:{t['btn_text']};display:inline-block;font-family:Arial,sans-serif;font-size:15px;font-weight:700;padding:15px 38px",
        html, count=1
    )

    # ── Seasonal label ──
    html = re.sub(
        r"display:inline-block;background:#FF7A00;color:#fff;font-family:Arial,sans-serif;font-size:11px",
        f"display:inline-block;background:{t['accent']};color:{t['btn_text']};font-family:Arial,sans-serif;font-size:11px",
        html, count=1
    )

    # ── Intro section orange divider ──
    html = re.sub(
        r"border-top:3px solid #FF7A00",
        f"border-top:3px solid {t['accent']}",
        html, count=1
    )
    html = re.sub(
        r"color:#FF7A00;font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;display:block;margin-bottom:10px",
        f"color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;display:block;margin-bottom:10px",
        html, count=1
    )

    # ── Populate text ──
    html = html.replace('[SEASON / SEIZOEN]', 'Zomer')
    html = html.replace('[NEWSLETTER HEADLINE EN]', 'Innovatie in hotelbeveiliging, tariefoptimalisatie en directe boekingen')
    html = html.replace('[NIEUWSBRIEF KOPTEKST NL]', 'LodgeGate PMS & RezXS Zomer 2026 Editie')
    html = html.replace('[Short subheadline or seasonal greeting — 1–2 sentences EN]',
                         'Ontdek onze belangrijkste updates van deze zomer: optimale gegevensbeveiliging met 2FA, slimme tariefbeheer-integraties en conversie-gerichte weboplossingen.')
    html = html.replace('[Korte ondertitel of seizoensgroet — 1–2 zinnen NL]',
                         'Alles wat u nodig heeft voor een zorgeloos en succesvol zomerseizoen.')
    html = html.replace('[CTA LINK]', 'https://lodgegate.com')
    html = html.replace('[CTA Button Text EN / NL]', 'Bekijk Updates')
    html = html.replace('[Intro Section Heading EN]', 'Klaar voor een succesvol zomerseizoen')
    html = html.replace("[2–3 sentence seasonal intro paragraph in English. Mention the season,\n                                            what's new, and why it matters for their hotel. Keep warm and professional.]",
                         "Met het drukke zomerseizoen voor de deur staan operationele rust en gegevensbeveiliging centraal. Deze maand delen wij belangrijke updates rondom LodgeGate cloudbeveiliging, geautomatiseerde revenue-integraties en conversie-gerichte weboplossingen van RezXS.")
    html = html.replace("[2–3 zin seizoensintro in het Nederlands. Vermeld het seizoen, wat er nieuw\n                                            is en waarom dit relevant is voor hun hotel. Houd het warm en\n                                            professioneel.]", "")

    # ── Build NL articles ──
    nl_articles_html = build_edition_blocks(raw_articles, lang='nl', t=t)

    # ── Build NL personal greeting ──
    nl_greeting_html = ""
    if g['nl']:
        # Mailchimp *|IF:FNAME|* conditional: show name if available, fallback otherwise
        nl_greeting_html = f"""
                    <!-- ===== PERSONAL GREETING (NL) ===== -->
                    <tr>
                        <td style="background-color:{t['section_bg_b']};padding:35px 40px 10px">
                            <p style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:20px;font-weight:700;line-height:1.4;margin:0">
                                *|IF:FNAME|*
                                {g['nl']}
                                *|ELSE:|*
                                {g['nl_fallback']}
                                *|END:IF|*
                            </p>
                        </td>
                    </tr>
"""

    # ── Build EN transition anchor + hero ──
    en_section = f"""
                    <!-- ===== ENGLISH EDITION ANCHOR & HERO ===== -->
                    <tr>
                        <td id="english-version" style="background:{t['en_hero_bg']};padding:50px 40px;text-align:center;border-top:5px solid {t['accent']}">
                            <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.18em;text-transform:uppercase;display:block;margin-bottom:12px">
                                ✦ ENGLISH EDITION ✦
                            </span>
                            <h1 style="color:{t['en_hero_text']};font-family:'Inter',Arial,sans-serif;font-size:32px;font-weight:800;line-height:1.2;margin:0 0 16px">
                                Innovating Hotel Security, Revenue &amp; Direct Bookings
                            </h1>
                            <p style="color:{t['en_hero_sub']};font-family:Arial,sans-serif;font-size:16px;line-height:1.6;max-width:480px;margin:0 auto 28px">
                                Discover our latest Summer 2026 updates designed to secure your PMS operations, automate pricing with RMS integrations, and elevate your direct web bookings.
                            </p>
                            <a href="https://lodgegate.com" target="_blank" style="background-color:{t['accent']};border-radius:50px;color:{t['btn_text']};display:inline-block;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:14px 34px;text-decoration:none">
                                Explore Updates &rarr;
                            </a>
                        </td>
                    </tr>
                    <tr>
                        <td style="background:{t['section_bg_b']};padding:45px 40px">
                            <h2 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:24px;font-weight:800;margin:0 0 12px">
                                Ready for a High-Performing Summer Season
                            </h2>
                            <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.75;margin:0">
                                As peak summer travel approaches, operational efficiency and data security are top priorities. This month, we share key advancements in LodgeGate cloud security, automated revenue integrations, and conversion-focused RezXS web solutions.
                            </p>
                        </td>
                    </tr>
"""
    # ── Build EN personal greeting ──
    en_greeting_html = ""
    if g['en']:
        en_greeting_html = f"""
                    <!-- ===== PERSONAL GREETING (EN) ===== -->
                    <tr>
                        <td style="background-color:{t['section_bg_b']};padding:35px 40px 10px">
                            <p style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:20px;font-weight:700;line-height:1.4;margin:0">
                                *|IF:FNAME|*
                                {g['en']}
                                *|ELSE:|*
                                {g['en_fallback']}
                                *|END:IF|*
                            </p>
                        </td>
                    </tr>
"""
    en_articles_html = build_edition_blocks(raw_articles, lang='en', t=t)

    # ── Footer ──
    footer_html = f"""
                    <!-- ===== SHARED FOOTER ===== -->
                    <tr>
                        <td style="background-color:{t['footer_bg']};padding:45px 30px;text-align:center">
                            <p style="color:{t['footer_text']};font-family:Arial,sans-serif;font-size:13px;line-height:1.7;margin:0 0 18px">
                                <strong style="color:{t['en_hero_text']};font-size:15px;font-family:'Inter',Arial,sans-serif">LodgeGate PMS &amp; RezXS Web Solutions</strong><br>
                                Empowering hotels with seamless cloud management and direct booking growth.
                            </p>
                            <p style="margin:0 0 24px">
                                <a href="https://lodgegate.com" target="_blank" style="color:{t['footer_link']};font-family:Arial,sans-serif;font-size:13px;text-decoration:none;font-weight:700;margin:0 12px">LodgeGate.com</a> &bull;
                                <a href="https://rezxs.com" target="_blank" style="color:{t['footer_link']};font-family:Arial,sans-serif;font-size:13px;text-decoration:none;font-weight:700;margin:0 12px">RezXS.com</a>
                            </p>
                            <hr style="border:none;border-top:1px solid {t['divider']};width:100%;max-width:400px;margin:0 auto 20px">
                            <p style="color:{t['footer_text']};font-family:Arial,sans-serif;font-size:11px;line-height:1.6;margin:0">
                                You are receiving this newsletter because you are a partner or user of LodgeGate &amp; RezXS.<br>
                                <a href="*|UNSUB|*" style="color:{t['footer_text']};text-decoration:underline">Unsubscribe from this list</a> &nbsp;|&nbsp;
                                <a href="*|UPDATE_PROFILE|*" style="color:{t['footer_text']};text-decoration:underline">Update preferences</a>
                            </p>
                        </td>
                    </tr>
"""

    replacement = "                    </tr>" + nl_greeting_html + nl_articles_html + en_section + en_greeting_html + en_articles_html + footer_html + "                </table></td></tr></table>"

    if "                    </tr></table></td></tr></table>" in html:
        html = html.replace("                    </tr></table></td></tr></table>", replacement)
    else:
        html = html.replace("                </table>\n            </td>\n        </tr>\n    </table>",
                             nl_greeting_html + nl_articles_html + en_section + en_greeting_html + en_articles_html + footer_html + "                </table>\n            </td>\n        </tr>\n    </table>")

    # ── FINAL SANITATION ──
    html = re.sub(r'\[[A-Za-z0-9\s\/—\-\.\,\']*?\]', '', html)
    html = re.sub(r'<p[^>]*>\s*</p>', '', html)
    html = re.sub(r'<h[1-6][^>]*>\s*</h[1-6]>', '', html)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"✅ Compiled [{t['name']}] newsletter to: {output_path}")


if __name__ == '__main__':
    template = os.path.join(BASE_DIR, "Untitled Template.html")
    docx = os.path.join(BASE_DIR, "Newsletter LodgeGate 2026.docx")
    out = os.path.join(BASE_DIR, "Newsletter_LodgeGate_2026_Ready.html")

    generate_newsletter(template, docx, out, template_id='executive-dark')
