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
        'rez_color': '#ffffff',
        'xs_color': '#FF7A00',
        'badge_bg': '#FF7A00',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#ffffff',
        'xs_color': '#121A2F',
        'badge_bg': '#121A2F',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#ffffff',
        'xs_color': '#FF7A00',
        'badge_bg': '#FF7A00',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#121A2F',
        'xs_color': '#FF7A00',
        'badge_bg': '#FF7A00',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': 'transparent',
        'logo_wrapper_padding': '0',
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
        'rez_color': '#E8DCC8',
        'xs_color': '#C9A84C',
        'badge_bg': '#C9A84C',
        'badge_text': '#0D0D1A',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#ffffff',
        'xs_color': '#022C22',
        'badge_bg': '#022C22',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#ffffff',
        'xs_color': '#450A0A',
        'badge_bg': '#ffffff',
        'badge_text': '#E63E2A',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'rez_color': '#ffffff',
        'xs_color': '#FF7A00',
        'badge_bg': '#FF7A00',
        'badge_text': '#ffffff',
        'logo_wrapper_bg': '#ffffff',
        'logo_wrapper_padding': '5px 12px',
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
        'nl': 'Beste *|FNAME|*,',
        'en': 'Dear *|FNAME|*,',
        'nl_fallback': 'Beste relatie,',
        'en_fallback': 'Dear partner,',
    },
    'none': {
        'name': 'No Greeting',
        'nl': '',
        'en': '',
        'nl_fallback': '',
        'en_fallback': '',
    },
}

PARAGRAPH_TRANSLATIONS = {
    "Veiligheid": "Security & Privacy",
    "RMS": "Revenue Optimization",
    "LodgeGate WEB": "Direct Bookings & Web Solutions",
    "Professionele weboplossingen met LodgeGate": "Professional Web Solutions with LodgeGate",
    "Een website die past bij uw accommodatie": "A Website That Fits Your Accommodation",
    "Meer directe boekingen met een IBE": "More Direct Bookings with an IBE",
    "Flexibel inzetbaar": "Highly Flexible & Adaptable",
    "Professionele weboplossingen met LodgeGate WEB": "Professional Web Solutions with LodgeGate WEB",

    "In een tijd waarin datalekken en cyberaanvallen steeds vaker voorkomen, is de beveiliging van gegevens belangrijker dan ooit. Organisaties verwerken dagelijks gevoelige informatie van gasten, medewerkers en relaties. Het beschermen van deze gegevens is niet alleen essentieel voor de continuïteit van de organisatie, maar ook voor het behoud van vertrouwen en het voldoen aan privacywetgeving.":
        "In an era where data breaches and cyber attacks are increasingly common, data security is more important than ever. Organizations process sensitive information from guests, employees, and partners on a daily basis. Protecting this data is not only essential for business continuity, but also for maintaining trust and complying with privacy regulations.",
    "LodgeGate biedt daarom uitgebreide beveiligingsmogelijkheden om ongeautoriseerde toegang tot gegevens te voorkomen. Een sterke basis begint bij het gebruik van veilige wachtwoorden. Door wachtwoordbeleid af te dwingen en gebruikers bewust te maken van het belang van sterke, unieke wachtwoorden, wordt het risico op misbruik aanzienlijk verkleind.":
        "LodgeGate therefore offers comprehensive security features to prevent unauthorized data access. A strong foundation begins with using secure passwords. By enforcing strict password policies and making users aware of the importance of strong, unique passwords, the risk of misuse is significantly reduced.",
    "Daarnaast ondersteunt LodgeGate IP-adresblokkering, waarmee toegang kan worden beperkt tot vooraf goedgekeurde locaties of netwerken. Dit zorgt ervoor dat onbevoegden, zelfs wanneer zij over inloggegevens beschikken, geen toegang krijgen vanuit onbekende of verdachte locaties.":
        "Additionally, LodgeGate supports IP address restriction and blocking, allowing access to be limited to pre-approved locations or trusted networks. This ensures that unauthorized individuals cannot gain access from unknown or suspicious locations, even if they happen to possess valid login credentials.",
    "Voor een extra beveiligingslaag biedt LodgeGate ook tweefactorauthenticatie (2FA). Hierbij is naast een wachtwoord een tweede verificatiestap vereist, zoals een code via een authenticator-app. Hierdoor wordt het voor kwaadwillenden aanzienlijk moeilijker om toegang te verkrijgen, zelfs wanneer een wachtwoord onverhoopt is buitgemaakt.":
        "For an extra layer of security, LodgeGate also provides Two-Factor Authentication (2FA). This requires a second verification step alongside the password, such as a code generated via an authenticator app. This makes it significantly more difficult for malicious actors to gain access, even in the event of a compromised password.",
    "Deze beveiligingsmaatregelen kunnen op verschillende niveaus binnen LodgeGate worden toegepast, zodat organisaties de bescherming kunnen afstemmen op hun eigen beveiligingsbeleid en risicoanalyse. Door sterke wachtwoorden, IP-adresbeperkingen en 2FA te combineren, ontstaat een robuuste beveiligingsstrategie die helpt om gevoelige gegevens optimaal te beschermen.":
        "These security measures can be applied at various levels within LodgeGate, enabling organizations to tailor protection to their specific security policies and risk assessments. Combining strong passwords, IP address restrictions, and 2FA creates a robust security strategy that helps safeguard sensitive data optimally.",
    "Alle genoemde beveiligingsmaatregelen zijn standaard beschikbaar binnen LodgeGate en ondersteunen organisaties bij het creëren van een veilige en betrouwbare digitale omgeving voor zowel medewerkers als gasten.":
        "All of these security features are available as standard within LodgeGate, empowering organizations to create a safe and reliable digital environment for both employees and guests.",

    "In een dynamische hotelmarkt is het handmatig beheren van kamerprijzen steeds minder effectief. Vraag, aanbod, concurrentie, seizoensinvloeden en lokale evenementen kunnen dagelijks invloed hebben op de optimale kamerprijs. Een Revenue Management System (RMS) helpt hotels om hier slim op in te spelen door tarieven automatisch te analyseren, bij te werken en te optimaliseren.":
        "In today's dynamic hotel market, managing room rates manually is becoming increasingly ineffective. Supply and demand, competition, seasonal trends, and local events can influence the optimal room rate on a daily basis. A Revenue Management System (RMS) helps hotels respond intelligently by automatically analyzing, updating, and optimizing pricing.",
    "Door gebruik te maken van een RMS worden kamerprijzen continu afgestemd op actuele marktomstandigheden. Dit zorgt ervoor dat tarieven niet alleen concurrerend blijven, maar ook bijdragen aan een maximale omzet en bezettingsgraad. Automatische prijsaanpassingen verminderen daarnaast de tijd die medewerkers besteden aan handmatige controles en prijswijzigingen, waardoor er meer ruimte ontstaat voor gastgericht werken.":
        "By utilizing an RMS, room rates are continuously adjusted to current market conditions. This ensures that rates remain competitive while driving maximum revenue and occupancy. Automated pricing adjustments also reduce the time staff spend on manual checks and rate changes, freeing up valuable time for guest-focused hospitality.",
    "Binnen onze oplossingen werken wij met zowel RevControl als RoomPriceGenie. Beide systemen bieden krachtige functionaliteiten voor geautomatiseerd revenue management, maar zijn gericht op verschillende typen accommodaties en bedrijfsbehoeften. Hierdoor kunnen wij voor vrijwel iedere situatie een passende oplossing aanbieden.":
        "Within our solutions ecosystem, we partner with both RevControl and RoomPriceGenie. Both systems offer powerful functionality for automated revenue management, while catering to different property types and business requirements. This allows us to provide a tailored solution for virtually any operational setting.",
    "Met RevControl profiteren hotels van uitgebreide analyses, forecasting en geavanceerde prijsstrategieën. RoomPriceGenie onderscheidt zich door zijn gebruiksvriendelijke aanpak en snelle implementatie, waardoor ook kleinere hotels en onafhankelijke accommodaties eenvoudig kunnen profiteren van professionele prijsoptimalisatie.":
        "With RevControl, hotels benefit from comprehensive analytics, forecasting, and advanced pricing strategies. RoomPriceGenie stands out for its user-friendly approach and rapid implementation, making it effortless for smaller hotels and independent properties to leverage professional rate optimization.",
    "Door de integratie van deze RMS-oplossingen met LodgeGate worden prijswijzigingen automatisch verwerkt, wat zorgt voor een efficiënte en betrouwbare workflow. Zo beschikken accommodaties altijd over actuele tarieven en kunnen zij sneller inspelen op veranderingen in de markt.":
        "Through the seamless integration of these RMS solutions with LodgeGate, rate adjustments are processed automatically, ensuring an efficient and reliable workflow. Properties always operate with live, accurate rates and can adapt rapidly to market shifts.",
    "Of het nu gaat om een kleinschalig hotel, een hotelgroep of een recreatieaccommodatie, met RevControl en RoomPriceGenie bieden wij een oplossing die aansluit bij de omvang, doelstellingen en commerciële strategie van uw organisatie. Hierdoor haalt u het maximale rendement uit uw beschikbare kamers, zonder dat dit leidt tot extra administratieve lasten.":
        "Whether running a boutique hotel, a multi-property hotel group, or a recreational resort, RevControl and RoomPriceGenie provide solutions aligned with your property size, goals, and commercial strategy. This maximizes the return on your available rooms without adding administrative overhead.",

    "Een sterke online aanwezigheid is tegenwoordig onmisbaar voor iedere accommodatie. De website is vaak het eerste contactmoment met potentiële gasten en speelt een belangrijke rol in het genereren van directe boekingen. Met de weboplossingen van LodgeGate bieden wij zowel volledig gepersonaliseerde websites als krachtige Internet Booking Engines (IBE), afgestemd op de wensen en doelstellingen van uw organisatie.":
        "A strong online presence is indispensable for any hospitality accommodation today. Your website is often the very first touchpoint with potential guests and plays a vital role in driving direct bookings. With LodgeGate WEB solutions, we offer both fully tailored websites and powerful Internet Booking Engines (IBE), perfectly designed to match your property's goals.",
    "Met een gepersonaliseerde LodgeGate-website beschikt u over een professionele online omgeving die volledig aansluit bij uw huisstijl, doelgroep en uitstraling. De website wordt ontworpen met aandacht voor gebruiksvriendelijkheid, snelheid en optimale weergave op desktop, tablet en smartphone. Hierdoor kunnen bezoekers eenvoudig informatie vinden, beschikbaarheid bekijken en direct een reservering maken.":
        "With a custom LodgeGate website, you gain a professional digital environment that aligns seamlessly with your brand identity, target audience, and aesthetic appeal. The website is engineered with a focus on user-friendliness, lightning-fast loading speeds, and optimal responsiveness across desktop, tablet, and smartphone screens. Visitors can easily access details, check real-time availability, and book directly.",
    "Naast een aantrekkelijke presentatie van kamers, faciliteiten en arrangementen, zorgen wij ervoor dat de website technisch en commercieel optimaal presteert. Dit draagt bij aan een betere online vindbaarheid en een hogere conversie van bezoekers naar gasten.":
        "Alongside an engaging showcase of your rooms, facilities, and packages, we ensure that your website performs at the highest technical and commercial standards. This boosts online search engine visibility (SEO) and maximizes conversion from website visitors to paying guests.",
    "Een Internet Booking Engine (IBE) maakt het mogelijk om gasten rechtstreeks via uw eigen website te laten reserveren. Hiermee vermindert u de afhankelijkheid van externe boekingsplatformen en bespaart u op commissiekosten. Gasten kunnen eenvoudig beschikbaarheid controleren, tarieven bekijken en direct hun verblijf boeken in een veilige online omgeving.":
        "An Internet Booking Engine (IBE) empowers guests to book directly through your own website. This reduces your dependence on external Online Travel Agencies (OTAs) and saves substantial commission fees. Guests can effortlessly verify availability, view live rates, and book their stay securely.",
    "De IBE van LodgeGate is volledig geïntegreerd met het reserveringssysteem, waardoor beschikbaarheid, tarieven en reserveringen realtime worden verwerkt. Dit voorkomt dubbele boekingen en zorgt ervoor dat gasten altijd beschikken over actuele informatie.":
        "The LodgeGate IBE is fully integrated with your Property Management System, processing availability, pricing, and reservations in real time. This eliminates overbookings and guarantees that guests always see live, accurate availability.",
    "Of u nu kiest voor een complete, gepersonaliseerde website inclusief geïntegreerde boekingsmodule, of uitsluitend gebruik wilt maken van een IBE op een bestaande website, LodgeGate biedt voor iedere situatie een passende oplossing. Dankzij de flexibele mogelijkheden kunnen wij een online boekingservaring creëren die aansluit bij de behoeften van zowel de accommodatie als de gast.":
        "Whether you opt for an all-in-one custom website with an embedded booking module, or simply wish to integrate our seamless IBE into your existing website, LodgeGate offers the ideal solution for every scenario. Our flexible architecture allows us to craft a booking journey tailored to both property requirements and guest preferences.",
    "Met de webproducten van LodgeGate combineert u een professionele online presentatie met een gebruiksvriendelijk boekingsproces, waardoor u meer directe reserveringen genereert en de controle over uw online verkoopkanalen behoudt.":
        "With LodgeGate WEB products, you combine a polished digital presentation with an intuitive booking flow—driving higher direct booking volumes and putting you firmly in control of your distribution channels."
}


def translate_paragraphs_to_en(paragraphs_nl):
    """Translates a list of Dutch paragraphs to full English without dropping text."""
    translated = []
    for p in paragraphs_nl:
        p_clean = p.strip()
        if p_clean in PARAGRAPH_TRANSLATIONS:
            translated.append(PARAGRAPH_TRANSLATIONS[p_clean])
        else:
            matched = False
            for nl_key, en_val in PARAGRAPH_TRANSLATIONS.items():
                if nl_key in p_clean or p_clean in nl_key:
                    translated.append(en_val)
                    matched = True
                    break
            if not matched:
                translated.append(p_clean)
    return translated


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

        is_known_heading = p.strip().lower() in ['veiligheid', 'rms', 'lodgegate web', 'security', 'revenue optimization', 'direct bookings']
        is_generic_heading = (p.isupper() and len(p) < 35 and not p.endswith('.') and not p.endswith(',') and len(p.split()) <= 4)

        if is_known_heading or is_generic_heading:
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


def enrich_article(article, lang='nl', custom_image_url=None):
    """Enriches an article for either Dutch or English edition."""
    title_nl = article['title_nl']
    paragraphs_list = article['paragraphs_nl'] if lang == 'nl' else translate_paragraphs_to_en(article['paragraphs_nl'])
    text_full = "\n\n".join(paragraphs_list)
    pool = get_unique_image_pool()[lang]

    curated_map = {
        'veiligheid': {
            'category_nl': 'BEVEILIGING & PRIVACY', 'category_en': 'SECURITY & PRIVACY',
            'title_nl': 'Optimaliseer uw gegevensbeveiliging met LodgeGate 2FA',
            'title_en': 'Enhanced Data Security & 2FA in LodgeGate',
            'link': 'https://lodgegate.com',
        },
        'rms': {
            'category_nl': 'TARIEFOPTIMALISATIE', 'category_en': 'REVENUE OPTIMIZATION',
            'title_nl': 'Haal het maximale uit uw kamerprijzen met RevControl & RoomPriceGenie',
            'title_en': 'Automated Price Optimization with RMS Integrations',
            'link': 'https://lodgegate.com',
        },
        'lodgegate web': {
            'category_nl': 'DIRECTE BOEKINGEN & WEB', 'category_en': 'DIRECT BOOKINGS & WEB',
            'title_nl': 'Professionele weboplossingen en IBE met LodgeGate WEB',
            'title_en': 'Boost Direct Bookings with RezXS Web Solutions',
            'link': 'https://rezxs.com',
        },
    }

    key = title_nl.lower().strip()
    fallback_img = 'https://images.unsplash.com/photo-1566073771259-6a8506099945?auto=format&fit=crop&w=1000&q=80'
    image = custom_image_url if (custom_image_url and custom_image_url.strip()) else pool.get(key, fallback_img)

    if key in curated_map:
        info = curated_map[key]
        return {
            'category': info[f'category_{lang}'] if f'category_{lang}' in info else info.get('category_nl', 'UPDATE'),
            'title': info['title_nl'] if lang == 'nl' else info['title_en'],
            'text': text_full,
            'paragraphs': paragraphs_list,
            'image': image,
            'link': info['link'],
        }

    is_web = any(k in key or k in text_full.lower() for k in ['web', 'ibe', 'rezxs', 'website', 'boeken'])
    link = 'https://rezxs.com' if is_web else 'https://lodgegate.com'

    title_translated = PARAGRAPH_TRANSLATIONS.get(title_nl.strip(), title_nl) if lang == 'en' else title_nl

    return {
        'category': 'UPDATE',
        'title': title_translated,
        'text': text_full,
        'paragraphs': paragraphs_list,
        'image': image,
        'link': link,
    }


def render_article_variation(enriched, variation, btn_text, t, article_id='art'):
    """Renders one of 3 distinct display layouts using theme colors."""
    formatted_paragraphs = []
    for p in enriched.get('paragraphs', [enriched['text']]):
        p_clean = p.strip()
        is_subheading = len(p_clean) < 65 and not p_clean[-1] in ['.', ',', '!', '?', ';', ':']
        if is_subheading:
            formatted_paragraphs.append(f'<strong style="color:{t["heading_color"]};font-size:16px">{p_clean}</strong>')
        else:
            formatted_paragraphs.append(p_clean)
    formatted_text = "<br><br>".join(formatted_paragraphs)

    bg_a = t['section_bg_a']
    bg_b = t['section_bg_b']

    if variation == 'stacked':
        bg = bg_a
        return f"""
                    <!-- ===== VARIATION A: STACKED BANNER ===== -->
                    <tr>
                        <td style="background-color:{bg};padding:45px 40px">
                            <a href="{enriched['link']}" target="_blank">
                                <img mc:edit="photo_{article_id}" src="{enriched['image']}" alt="{enriched['title']}" width="520" style="width:100%;max-width:520px;border-radius:12px;margin-bottom:22px;display:block;object-fit:cover;max-height:260px" role="presentation" border="0">
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
                            <!--[if mso]>
                            <table role="presentation" width="540" cellspacing="0" cellpadding="0" border="0">
                            <tr>
                            <td width="235" valign="top" style="padding:10px;">
                            <![endif]-->
                            <div style="display:inline-block;width:100%;max-width:235px;vertical-align:top">
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                    <tr>
                                        <td style="padding:10px">
                                            <a href="{enriched['link']}" target="_blank">
                                                <img mc:edit="photo_{article_id}" src="{enriched['image']}" alt="{enriched['title']}" width="215" style="width:100%;max-width:215px;border-radius:10px;object-fit:cover;height:240px;display:block" role="presentation" border="0">
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                            </div>
                            <!--[if mso]>
                            </td>
                            <td width="305" valign="top" style="padding:10px;">
                            <![endif]-->
                            <div style="display:inline-block;width:100%;max-width:305px;vertical-align:top">
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                    <tr>
                                        <td style="padding:10px">
                                            <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;display:block;margin-bottom:6px">
                                                ✦ {enriched['category']}
                                            </span>
                                            <h3 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:20px;font-weight:800;line-height:1.3;margin:0 0 10px">
                                                {enriched['title']}
                                            </h3>
                                            <p style="color:{t['body_text_alt']};font-family:Arial,sans-serif;font-size:14px;line-height:1.65;margin:0 0 18px">
                                                {formatted_text}
                                            </p>
                                            <a href="{enriched['link']}" target="_blank" style="color:{t['accent']};font-family:Arial,sans-serif;font-weight:700;text-decoration:none;font-size:14px">
                                                {btn_text} &rarr;
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                            </div>
                            <!--[if mso]>
                            </td>
                            </tr>
                            </table>
                            <![endif]-->
                        </td>
                    </tr>
"""
    else:  # highlight card
        bg = bg_a
        return f"""
                    <!-- ===== VARIATION C: HIGHLIGHT PRODUCT CARD ===== -->
                    <tr>
                        <td style="background-color:{bg};padding:45px 40px">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:{t['card_bg']};border-radius:16px;border-top:5px solid {t['card_border']};box-shadow:0 8px 24px {t['card_shadow']}">
                                <tr>
                                    <td style="padding:32px 28px;background-color:{t['card_bg']};border-radius:16px">
                                        <span style="color:{t['accent']};font-family:Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;margin:0 0 8px;display:block">
                                            ✦ {enriched['category']}
                                        </span>
                                        <h2 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:22px;font-weight:800;line-height:1.25;margin:0 0 14px">
                                            {enriched['title']}
                                        </h2>
                                        <a href="{enriched['link']}" target="_blank">
                                            <img mc:edit="photo_{article_id}" src="{enriched['image']}" alt="{enriched['title']}" width="464" style="width:100%;max-width:464px;border-radius:10px;margin-bottom:18px;max-height:220px;object-fit:cover;display:block" role="presentation" border="0">
                                        </a>
                                        <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.7;margin:0 0 22px">
                                            {formatted_text}
                                        </p>
                                        <a href="{enriched['link']}" target="_blank" style="background-color:{t['dark_btn_bg']};border-radius:50px;color:#ffffff;display:inline-block;font-family:Arial,sans-serif;font-size:14px;font-weight:700;padding:12px 30px;text-decoration:none">
                                            {btn_text} &rarr;
                                        </a>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
"""


def build_edition_blocks(articles, lang, t, custom_images=None):
    """Compiles all article sections with cycling variations."""
    btn_text = "Lees Meer" if lang == 'nl' else "Read More"
    variations = ['stacked', 'split', 'highlight']

    html_out = ""
    for i, a in enumerate(articles):
        custom_img = custom_images[i] if (custom_images and i < len(custom_images)) else None
        enriched = enrich_article(a, lang, custom_image_url=custom_img)
        var = variations[i % len(variations)]
        html_out += render_article_variation(enriched, var, btn_text, t, article_id=f"{lang}_{i+1}")
    return html_out


def generate_newsletter(template_path, docx_path, output_path, template_id='executive-dark', greeting_style='none', custom_images=None, hero_image_url=None):
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
    html = re.sub(
        r'class="header-logo-wrapper"[^>]*>',
        f'class="header-logo-wrapper" style="background-color:{t["logo_wrapper_bg"]};padding:{t["logo_wrapper_padding"]};border-radius:8px">',
        html
    )
    html = re.sub(
        r'class="rez-logo-text"[^>]*>',
        f'class="rez-logo-text" style="font-family:\'Inter\',Arial,Helvetica,sans-serif;font-size:18px;font-weight:800;color:{t["rez_color"]};letter-spacing:-0.03em">',
        html
    )
    html = re.sub(
        r'class="xs-logo-text"[^>]*>',
        f'class="xs-logo-text" style="color:{t["xs_color"]}">',
        html
    )
    html = re.sub(
        r'class="web-logo-badge"[^>]*>',
        f'class="web-logo-badge" style="display:inline-block;background:{t["badge_bg"]};color:{t["badge_text"]};font-family:Arial,sans-serif;font-size:9px;font-weight:700;padding:2px 8px;border-radius:8px;margin-left:6px;letter-spacing:0.06em;text-transform:uppercase;vertical-align:middle">',
        html
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
    hero_bg_url = hero_image_url if (hero_image_url and hero_image_url.strip()) else 'https://images.unsplash.com/photo-1566073771259-6a8506099945?auto=format&fit=crop&w=1200&q=80'
    html = re.sub(
        r"background-color:#121A2F;background-image:url\('https://images\.unsplash\.com[^']*'\);background-size:cover;background-position:center;padding:70px 40px",
        f"background-color:{t['hero_bg']};background-image:url('{hero_bg_url}');background-size:cover;background-position:center;padding:70px 40px",
        html, count=1
    )
    html = re.sub(
        r'<v:fill type="frame" src="https://images\.unsplash\.com[^"]+" color="#121A2F" />',
        f'<v:fill type="frame" src="{hero_bg_url}" color="{t["hero_bg"]}" />',
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
        r'border-top:3px solid #FF7A00;width:50px;margin:0 0 22px',
        f"border-top:3px solid {t['accent']};width:50px;margin:0 0 22px",
        html
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

    # ── Build NL personal greeting inside NL intro block ──
    nl_greeting_html = ""
    if g['nl']:
        nl_greeting_html = f"""
                                        <p style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:18px;font-weight:700;margin:0 0 14px">
                                            *|IF:FNAME|*
                                            {g['nl']}
                                            *|ELSE:|*
                                            {g['nl_fallback']}
                                            *|END:IF|*
                                        </p>"""
    if '<!-- [NL GREETING PLACEHOLDER] -->' in html:
        html = html.replace('<!-- [NL GREETING PLACEHOLDER] -->', nl_greeting_html.lstrip())

    # ── Build NL articles ──
    nl_articles_html = build_edition_blocks(raw_articles, lang='nl', t=t, custom_images=custom_images)

    # ── Build EN transition anchor + hero + EN personal greeting ──
    en_greeting_span = ""
    if g['en']:
        en_greeting_span = f"""
                            <p style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:18px;font-weight:700;margin:0 0 14px">
                                *|IF:FNAME|*
                                {g['en']}
                                *|ELSE:|*
                                {g['en_fallback']}
                                *|END:IF|*
                            </p>"""

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
                            <hr style="border:none;border-top:3px solid {t['accent']};width:50px;margin:0 0 22px">
                            {en_greeting_span}
                            <h2 style="color:{t['heading_color']};font-family:'Inter',Arial,sans-serif;font-size:24px;font-weight:800;margin:0 0 12px">
                                Ready for a High-Performing Summer Season
                            </h2>
                            <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.75;margin:0 0 22px">
                                As peak summer travel approaches, operational efficiency and data security are top priorities. This month, we share key advancements in LodgeGate cloud security, automated revenue integrations, and conversion-focused RezXS web solutions.
                            </p>
                            <p style="color:{t['body_text']};font-family:Arial,sans-serif;font-size:15px;line-height:1.75;margin:0">
                                Warm regards,<br>
                                <strong style="color:{t['heading_color']}">The LodgeGate &amp; RezXS Team</strong>
                            </p>
                        </td>
                    </tr>
"""
    en_articles_html = build_edition_blocks(raw_articles, lang='en', t=t, custom_images=custom_images)

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

    replacement = "                    </tr>" + nl_articles_html + en_section + en_articles_html + footer_html + "                </table></td></tr></table>"

    if "                    </tr></table></td></tr></table>" in html:
        html = html.replace("                    </tr></table></td></tr></table>", replacement)
    else:
        html = html.replace("                </table>\n            </td>\n        </tr>\n    </table>",
                             nl_articles_html + en_section + en_articles_html + footer_html + "                </table>\n            </td>\n        </tr>\n    </table>")

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
