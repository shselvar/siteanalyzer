#!/usr/bin/env python3
"""Generate comprehensive site analysis Word document for myaerotel.com"""

import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)

def set_cell_shading(cell, color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_header_row(table, row_idx, texts, color="4B2E20"):
    row = table.rows[row_idx]
    for i, text in enumerate(texts):
        cell = row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)
        set_cell_shading(cell, color)

def add_data_row(table, row_idx, texts):
    row = table.rows[row_idx]
    for i, text in enumerate(texts):
        cell = row.cells[i]
        cell.text = str(text)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.size = Pt(10)

def add_screenshot(doc, path, caption, width=5.5):
    full_path = os.path.join("/workspace/screenshots", path)
    if os.path.exists(full_path):
        try:
            doc.add_picture(full_path, width=Inches(width))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = doc.add_paragraph(caption)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cap.runs:
                cap.runs[0].italic = True
                cap.runs[0].font.size = Pt(9)
            return True
        except Exception as e:
            doc.add_paragraph(f"[Screenshot: {path} - {str(e)}]")
            return False
    else:
        doc.add_paragraph(f"[Screenshot not available: {path}]")
        return False

# ═══════════════════════════════════════════
# COVER PAGE
# ═══════════════════════════════════════════
doc.add_paragraph("")
doc.add_paragraph("")
doc.add_paragraph("")
title = doc.add_heading("Website Analysis Report", level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle = doc.add_heading("www.myaerotel.com", level=1)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph("")
cover_info = doc.add_paragraph("AEM Edge Delivery Services Migration Assessment")
cover_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph("")
date_p = doc.add_paragraph("Date: February 2025")
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph("")
doc.add_paragraph("")
platform_p = doc.add_paragraph()
platform_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = platform_p.add_run("Current Platform: Kentico CMS (ASP.NET WebForms)")
run.bold = True
run.font.size = Pt(14)
doc.add_page_break()

# ═══════════════════════════════════════════
# TABLE OF CONTENTS
# ═══════════════════════════════════════════
doc.add_heading("Table of Contents", level=1)
toc_items = [
    "1. Executive Summary",
    "2. Templates Inventory",
    "3. Blocks / Components Catalog",
    "4. Page Counts by Template",
    "5. Integrations Analysis",
    "6. Complex Use Cases & Observations",
    "7. Migration Estimates",
    "Appendix: Screenshots"
]
for item in toc_items:
    doc.add_paragraph(item, style='List Number')
doc.add_page_break()

# ═══════════════════════════════════════════
# 1. EXECUTIVE SUMMARY
# ═══════════════════════════════════════════
doc.add_heading("1. Executive Summary", level=1)
doc.add_paragraph(
    "This report provides a comprehensive analysis of www.myaerotel.com for migration "
    "to Adobe Experience Manager (AEM) Edge Delivery Services. The site is currently built "
    "on Kentico CMS running ASP.NET WebForms and serves as the online presence for Aerotel, "
    "an airport transit hotel brand under the Plaza Premium Group."
)
doc.add_heading("Key Findings", level=2)
findings = [
    "Platform: Kentico CMS (ASP.NET WebForms) with Portal Engine architecture",
    "Total Pages in Sitemap: 25 URLs across 9 distinct template types",
    "Hotel Properties: 12 active hotel locations with individual detail pages (plus sub-pages for rooms, offers, dining)",
    "Estimated Total Pages: ~97-102 including hotel sub-pages",
    "Reusable Components: 21 distinct block/component types identified",
    "Third-Party Integrations: 18 external services including analytics, marketing automation, chat, and consent management",
    "Forms: 4 distinct form types (Booking Widget, Contact/Inquiry, Newsletter Signup, Login)",
    "Complexity: Medium-High due to booking widget, ASP.NET postback forms, geo-conditional tracking, and loyalty program integration",
]
for f in findings:
    doc.add_paragraph(f, style='List Bullet')
doc.add_page_break()

# ═══════════════════════════════════════════
# 2. TEMPLATES INVENTORY
# ═══════════════════════════════════════════
doc.add_heading("2. Templates Inventory", level=1)
doc.add_paragraph(
    "The following table lists all unique page templates identified across the site. "
    "Each template represents a distinct layout pattern used for one or more pages."
)

templates_data = [
    ["T1", "Homepage", "High",
     "Complex hero image carousel with rotating hotel highlights, inline booking widget with city autocomplete/date-time pickers/guest selectors, value proposition section with icon cards, latest offers carousel, brand CTA section with dual buttons, video embed link. Most interactive page on the site.",
     "https://www.myaerotel.com/en-uk"],
    ["T2", "Hotel Detail Page", "High",
     "Hero image carousel, inline booking widget, sub-navigation tabs (Hotel | Offers | Rooms | Dining | What's Around), important notice panel, feature badges, contact details, multiple content carousels (promotions, rooms, dining), downloadable PDF factsheet, value proposition section.",
     "https://www.myaerotel.com/en-uk/find/china-regions/mainland-china/shanghai/aerotel-shanghai"],
    ["T3", "Card Listing Page", "Medium",
     "Page title, optional dropdown filter, grid of cards with image/title/description/CTA. Used for offers and news. Cards may include image sliders. Some listing pages include additional content sections below the grid.",
     "https://www.myaerotel.com/en-uk/latest-offers\nhttps://www.myaerotel.com/en-uk/about-aerotel/news-room"],
    ["T4", "Content Detail Page", "Low",
     "Back navigation link, H1 title, rich text body with paragraphs/lists/bold, CTA buttons, horizontal rule separators. Used for individual offer details and informational pages.",
     "https://www.myaerotel.com/en-uk/latest-offers/ramadan-at-aerotel\nhttps://www.myaerotel.com/en-uk/about-aerotel/work-with-us"],
    ["T5", "Form Page", "High",
     "Page title, introductory paragraph, multi-field form with dropdowns, text inputs, international phone selector, date/time pickers, textarea, CAPTCHA image, checkbox, submit button. Cascading dropdowns for location selection. Uses ASP.NET postback.",
     "https://www.myaerotel.com/en-uk/about-aerotel/get-in-touch\nhttps://www.myaerotel.com/en-uk/about-aerotel/group-bookings\nhttps://www.myaerotel.com/en-uk/about-aerotel/chatbot-bookings"],
    ["T6", "FAQ / Accordion Page", "Medium",
     "Left sidebar navigation with category links, main content area with accordion/expandable Q&A items. Expand/collapse buttons with +/- icons. Used for FAQs and country-specific contact information.",
     "https://www.myaerotel.com/en-uk/about-aerotel/faq/general-faqs\nhttps://www.myaerotel.com/en-uk/about-aerotel/contact-us"],
    ["T7", "Long-form Legal Page", "Low",
     "Page title with last-updated date, extensive continuous text with numbered sections, bold headings, bulleted lists, inline links. No interactive elements. Very long pages.",
     "https://www.myaerotel.com/en-uk/terms-and-conditions\nhttps://www.myaerotel.com/en-uk/data-privacy-statement-security-policy"],
    ["T8", "Marketing Landing Page", "Medium",
     "Hero title with image, subtitle, descriptive paragraph, embedded YouTube video, feature cards in grid layout (2x2), horizontal rule separators, section headings, clickable banner images linking to external portals.",
     "https://www.myaerotel.com/en-uk/about-aerotel/smart-traveller"],
    ["T9", "Content Page (About)", "Medium",
     "Page title, alternating text and image sections, Why Choose Us heading, icon-based value proposition cards, YouTube video link.",
     "https://www.myaerotel.com/en-uk/about-aerotel/about-aerotel"],
]

table = doc.add_table(rows=1 + len(templates_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["ID", "Template Name", "Complexity", "Description & Reasoning", "Reference URL(s)"])
for i, row_data in enumerate(templates_data):
    add_data_row(table, i + 1, row_data)
for row in table.rows:
    row.cells[0].width = Cm(1.2)
    row.cells[1].width = Cm(3)
    row.cells[2].width = Cm(2)
    row.cells[3].width = Cm(7)
    row.cells[4].width = Cm(4.5)

doc.add_paragraph("")
doc.add_paragraph("Screenshots of each template type are provided in the Appendix section.")
doc.add_page_break()

# ═══════════════════════════════════════════
# 3. BLOCKS / COMPONENTS CATALOG
# ═══════════════════════════════════════════
doc.add_heading("3. Blocks / Components Catalog", level=1)
doc.add_paragraph(
    "The following catalog identifies all reusable blocks and components across the site. "
    "Where the same content model has different visual layouts, design variations are noted "
    "rather than separate blocks."
)

blocks_data = [
    ["B01", "Header / Navigation", "Medium",
     "Global sticky header with hamburger menu, Aerotel logo (links to /Home), LOGIN button (opens modal), BOOK button. Responsive mobile navigation (WebSlideMenu). Consistent across all pages.",
     "All pages"],
    ["B02", "Footer", "Medium",
     "Multi-section global footer: (1) Smart Traveller promotional banner/carousel, (2) Payment options bar (Visa, Mastercard, Alipay, WeChat Pay, Apple Pay), (3) Newsletter signup form, (4) Privacy consent text, (5) Quick links (Group Booking, Work with Us), (6) Social media links (Facebook, Instagram, WeChat), (7) Plaza Premium Group brand family grid, (8) Copyright bar.",
     "All pages"],
    ["B03", "Hero Carousel", "High",
     "Full-width image carousel using MasterSlider. On homepage: rotates through hotel locations with hotel name, location tag (Airside/Landside), description, and 'Discover More' CTA. On hotel detail pages: gallery of hotel images with navigation arrows. Supports auto-play and manual navigation.\nDesign Variation 1 (Homepage): Hotel info overlay with CTA.\nDesign Variation 2 (Hotel Detail): Pure image gallery.",
     "Homepage, Hotel Detail pages"],
    ["B04", "Booking Widget", "High",
     "Complex inline booking form: city search with autocomplete, booking type dropdown (Hourly/Overnight), check-in/check-out date-time pickers, room count selector (1-4), adult/children guest selector with dynamic age sub-selectors, promo code input, Book Now button. Uses ASP.NET postback. Fires analytics events.",
     "Homepage, Hotel Detail pages (~13 pages)"],
    ["B05", "Value Proposition Cards", "Low",
     "Three icon-based feature cards: 'Right at the Airport', 'Flexible Hourly Booking', 'Feels Like Home'. Each: icon, H4 heading, description paragraph.\nDesign Variation 1 (Homepage): Larger text, inline layout.\nDesign Variation 2 (About): Bordered card layout.\nDesign Variation 3 (Hotel Detail): Compact layout.",
     "Homepage, About, Hotel Detail"],
    ["B06", "Offers Card Grid", "Medium",
     "Grid of promotional offer cards with image, title, description, CTA.\nDesign Variation 1 (Homepage): 2 featured cards, large images.\nDesign Variation 2 (Listing): Full grid with dropdown filter.\nDesign Variation 3 (Hotel Detail): Carousel with dot pagination.",
     "Homepage, Latest Offers, Hotel Detail"],
    ["B07", "Brand CTA Section", "Low",
     "Content section with heading, paragraph, dual CTA buttons ('Discover More', 'View Brochure'), large illustrative image.",
     "Homepage"],
    ["B08", "Video Embed / Link", "Low",
     "Video content component.\nDesign Variation 1: External link to MP4 with thumbnail and play icon.\nDesign Variation 2: Embedded YouTube iframe with Plyr controls.",
     "Homepage, Smart Traveller, About"],
    ["B09", "Content Section (Text + Image)", "Low",
     "Alternating text and image layout. Heading, paragraph text, optional CTA. Images left or right.",
     "About, Work with Us"],
    ["B10", "Contact Form", "High",
     "Multi-field inquiry form: hotel selector (cascading), inquiry type, name, country (200+ options), intl phone input with flag selector, email, booking ref, message textarea, newsletter checkbox, image CAPTCHA, submit.\nDesign Variation (Chatbot Bookings): Additional date/time pickers, room type, adults/children, T&C checkbox.",
     "Get in Touch, Group Bookings, Chatbot Bookings"],
    ["B11", "Accordion / FAQ", "Medium",
     "Expandable/collapsible sections with +/- icon toggle.\nDesign Variation 1 (FAQ): Left sidebar nav + accordion Q&A.\nDesign Variation 2 (Contact Us): Country-specific phone numbers accordion.",
     "FAQ pages, Contact Us"],
    ["B12", "Newsletter Signup", "Low",
     "Footer form: Title dropdown, First Name, Last Name, Email, 'Sign Me Up!' button. Submits to Netcore Smartech.",
     "All pages (footer)"],
    ["B13", "Smart Traveller Banner", "Low",
     "Full-width promotional banner carousel in pre-footer. Links to loyalty program.",
     "All pages (pre-footer)"],
    ["B14", "Cookie Consent Banner", "Low",
     "Bottom overlay with cookie notice text, 'I Understand' dismiss button, privacy policy link. Managed by consenTag.eu.",
     "All pages"],
    ["B15", "Login Modal", "Medium",
     "Modal: email input, password with show/hide toggle, 'Remember Me' checkbox, login button. Tracks state via hidden field.",
     "All pages (header)"],
    ["B16", "Hotel Sub-Nav Tabs", "Medium",
     "Horizontal tab bar: Hotel | Special Offers | Rooms | Dining | What's Around. Active tab highlighted.",
     "Hotel Detail pages"],
    ["B17", "Important Notice Panel", "Low",
     "Bordered info panel with hotel access directions in table format. Transit vs. departure passenger instructions.",
     "Hotel Detail pages"],
    ["B18", "Feature Card Grid (Loyalty)", "Medium",
     "Grid of benefit cards for Smart Traveller loyalty program. Image, heading, description, CTA. 2x2 layout.",
     "Smart Traveller page"],
    ["B19", "Chat Widget", "Medium",
     "Persistent floating chat button (bottom-right) powered by Pantheon Lab AI chatbot. Iframe-based. All pages.",
     "All pages"],
    ["B20", "Page Title / Breadcrumb", "Low",
     "Page header with H1 title. Some pages include 'BACK' navigation link.",
     "All interior pages"],
    ["B21", "Payment Options Bar", "Low",
     "Horizontal bar with payment method icons: Visa, Mastercard, Alipay, WeChat Pay, Apple Pay.",
     "All pages (footer)"],
]

table = doc.add_table(rows=1 + len(blocks_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["ID", "Block Name", "Complexity", "Description & Behaviour", "Reference URL(s)"])
for i, row_data in enumerate(blocks_data):
    add_data_row(table, i + 1, row_data)
for row in table.rows:
    row.cells[0].width = Cm(1.2)
    row.cells[1].width = Cm(3)
    row.cells[2].width = Cm(1.8)
    row.cells[3].width = Cm(7.5)
    row.cells[4].width = Cm(4)

doc.add_paragraph("")
doc.add_paragraph(
    "Note: Design variations of the same block are documented as variations rather than separate blocks, "
    "as they share the same content model but differ in visual layout."
)
doc.add_page_break()

# ═══════════════════════════════════════════
# 3.1 BLOCK SCREENSHOTS
# ═══════════════════════════════════════════
doc.add_heading("3.1 Block Screenshots by Template", level=2)

screenshot_map = [
    ("01-homepage.jpeg", "T1 - Homepage: Hero Carousel, Booking Widget, Value Props, Offers, Brand CTA"),
    ("16-hotel-detail.jpeg", "T2 - Hotel Detail: Hero, Booking Widget, Sub-Nav Tabs, Content Carousels"),
    ("02-latest-offers.jpeg", "T3 - Card Listing: Offers Grid with Filter"),
    ("03-offer-detail.jpeg", "T4 - Content Detail: Back Nav, Rich Text, CTA Buttons"),
    ("06-contact-form.jpeg", "T5 - Form Page: Contact Form with CAPTCHA"),
    ("05-faq.jpeg", "T6 - FAQ / Accordion: Sidebar Nav + Expandable Q&A"),
    ("13-terms.jpeg", "T7 - Legal Page: Long-form Terms & Conditions"),
    ("10-smart-traveller.jpeg", "T8 - Marketing Landing: Video Embed, Feature Cards"),
    ("04-about.jpeg", "T9 - Content Page: Text+Image Sections, Value Props"),
]

for filename, caption in screenshot_map:
    add_screenshot(doc, filename, caption, width=5.0)
    doc.add_paragraph("")

doc.add_page_break()

# ═══════════════════════════════════════════
# 4. PAGE COUNTS BY TEMPLATE
# ═══════════════════════════════════════════
doc.add_heading("4. Page Counts by Template", level=1)
doc.add_paragraph(
    "The following table provides page counts by template type, based on the sitemap (25 URLs) "
    "plus estimated hotel detail pages and sub-pages discovered during analysis. "
    "The site has 12 active hotel locations, each with ~5 sub-pages."
)

page_counts_data = [
    ["T1", "Homepage", "1", "Automatic", "Standard content migration. Booking widget requires custom re-implementation."],
    ["T2", "Hotel Detail Page", "~60", "Manual", "12 hotels x ~5 sub-pages each. Complex due to booking widget, multiple carousels, sub-navigation, dynamic content."],
    ["T3", "Card Listing Page", "3", "Automatic", "Latest Offers listing, News Room, About section listing. Standardized card grid layout."],
    ["T4", "Content Detail Page", "15-20", "Automatic", "Individual offer detail pages (~12+), Work with Us, and similar content pages."],
    ["T5", "Form Page", "3", "Manual", "Get in Touch, Group Bookings, Chatbot Bookings. Complex forms require custom re-implementation."],
    ["T6", "FAQ / Accordion Page", "9", "Automatic", "7 FAQ category pages + Contact Us + FAQ parent."],
    ["T7", "Legal / Text Page", "3", "Automatic", "Terms & Conditions, Privacy Policy, Terms-Conditions (duplicate)."],
    ["T8", "Marketing Landing Page", "1", "Semi-Auto", "Smart Traveller page. YouTube embed and external links need manual verification."],
    ["T9", "Content Page (About)", "2", "Automatic", "About Aerotel, Photo Gallery (currently empty/stub)."],
    ["", "TOTAL", "~97-102", "", ""],
]

table = doc.add_table(rows=1 + len(page_counts_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["ID", "Template", "Page Count", "Migration Type", "Notes"])
for i, row_data in enumerate(page_counts_data):
    add_data_row(table, i + 1, row_data)
    if row_data[0] == "":
        for j in range(5):
            cell = table.rows[i+1].cells[j]
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True
for row in table.rows:
    row.cells[0].width = Cm(1.2)
    row.cells[1].width = Cm(3.5)
    row.cells[2].width = Cm(2)
    row.cells[3].width = Cm(2.5)
    row.cells[4].width = Cm(8.5)

doc.add_paragraph("")

doc.add_heading("Migration Approach Summary", level=2)
summary_data = [
    ["Automatic Migration", "~33-38 pages", "Content detail, FAQ, legal, listing, about pages. Standardized layouts."],
    ["Semi-Automatic", "~1 page", "Pages with embedded media needing manual verification."],
    ["Manual Migration", "~63 pages", "Hotel detail pages (complex carousels, booking widget) and form pages (CAPTCHA, cascading dropdowns)."],
]
table = doc.add_table(rows=1 + len(summary_data), cols=3)
table.style = 'Table Grid'
add_header_row(table, 0, ["Approach", "Page Count", "Description"])
for i, row_data in enumerate(summary_data):
    add_data_row(table, i + 1, row_data)
doc.add_page_break()

# ═══════════════════════════════════════════
# 5. INTEGRATIONS ANALYSIS
# ═══════════════════════════════════════════
doc.add_heading("5. Integrations Analysis", level=1)
doc.add_paragraph(
    "The following table catalogs all third-party integrations and embedded services detected across the site."
)

integrations_data = [
    ["Google Analytics 4", "Embed / Script", "Low", "GA4 property G-PCMKG7JE41. Standard page tracking.", "All pages"],
    ["Google Analytics (Universal)", "Embed / Script", "Low", "UA-109911887-5. Legacy Universal Analytics.", "All pages"],
    ["Google Tag Manager", "Embed / Script", "Medium", "GTM-NTQ54X2. Central tag container managing analytics/marketing tags.", "All pages"],
    ["Facebook Pixel", "Embed / Script", "Low", "Pixel ID 2147531685521875. Conversion and remarketing.", "All pages"],
    ["Bing UET", "Embed / Script", "Low", "Tag ID 343150333. Microsoft Ads conversion tracking.", "All pages"],
    ["DoubleClick Floodlight", "Embed (iframe)", "Low", "Source 9840322. Google DV360 campaign attribution.", "All pages"],
    ["Netcore Smartech", "API / Script", "High", "Full marketing automation suite incl. web push, activity tracking, in-page creatives, newsletter integration, Hansel A/B testing, Boxx.ai personalization.", "All pages"],
    ["Pantheon Lab AI Chatbot", "Embed (iframe)", "High", "Custom AI chatbot. Service: 'Aerotel'. Floating chat button + full interface. Auth token based.", "All pages"],
    ["consenTag.eu", "Embed / API", "Medium", "Cookie consent management v3.0.1. Cross-domain consent storage via iframes.", "All pages"],
    ["Optimix Asia", "Embed / Script", "Low", "j02.optimix.asia - optimization platform. Currently BROKEN (ERR_TUNNEL_CONNECTION_FAILED).", "All pages (broken)"],
    ["Affilired", "Embed / Script", "Low", "Affiliate tracking. Merchant ID 4749.", "All pages"],
    ["Denomatic", "Embed / Script", "Low", "Advertising/retargeting platform.", "All pages"],
    ["Connatix (ctnsnet.com)", "Embed / Script", "Low", "Scraper/pixel tracking scripts.", "All pages"],
    ["YouTube", "Embed (iframe)", "Low", "Embedded video player on Smart Traveller. External video links on Homepage/About.", "Smart Traveller, Homepage, About"],
    ["ipinfo.io", "API", "Low", "IP geolocation for auto-detecting user country on phone code selector.", "Chatbot Bookings"],
    ["GeoLookup (.ashx)", "Custom API", "Medium", "Server-side endpoint returning visitor country. Blocks ALL tracking for China-based visitors.", "All pages"],
    ["Smart Traveller Portal", "External Link", "Low", "External loyalty program at mysmarttraveller.com.", "Smart Traveller, Footer"],
    ["Plaza Premium Careers", "External Link", "Low", "External careers portal at plazapremiumgroup.com/careers.", "Work with Us"],
]

table = doc.add_table(rows=1 + len(integrations_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["Integration", "Type", "Complexity", "Description", "Reference Page(s)"])
for i, row_data in enumerate(integrations_data):
    add_data_row(table, i + 1, row_data)
for row in table.rows:
    row.cells[0].width = Cm(3)
    row.cells[1].width = Cm(2)
    row.cells[2].width = Cm(1.8)
    row.cells[3].width = Cm(7)
    row.cells[4].width = Cm(3.5)
doc.add_page_break()

# ═══════════════════════════════════════════
# 6. COMPLEX USE CASES
# ═══════════════════════════════════════════
doc.add_heading("6. Complex Use Cases & Observations", level=1)
doc.add_paragraph(
    "The following identifies complex behaviours and edge cases requiring special attention during migration."
)

complex_data = [
    ["CU01", "Booking Widget with Hourly Bookings", "~13 pages", "Homepage + 12 hotel detail pages",
     "Supports both hourly and overnight bookings with city autocomplete, date/time pickers, guest selectors. Hourly booking model is unique to airport hotels. Requires custom booking API integration or replacement."],
    ["CU02", "ASP.NET WebForms Single-Form Architecture", "All pages", "Entire site",
     "Entire site uses a single <form> element with __VIEWSTATE, __EVENTTARGET, __EVENTARGUMENT, __CMSCsrfToken. All form submissions use WebForm_DoPostBackWithOptions. Must be completely replaced with client-side JavaScript API calls."],
    ["CU03", "Geo-Conditional Script Loading", "All pages", "All pages (JavaScript)",
     "Calls /GeoLookup.ashx to detect visitor country. If in China, ALL analytics/marketing scripts are blocked. Must be replicated in EDS via edge-side logic."],
    ["CU04", "Custom CAPTCHA System", "3 pages", "Get in Touch, Group Bookings, Chatbot Bookings",
     "Server-generated image CAPTCHA (not reCAPTCHA). Must be replaced with modern solution (reCAPTCHA, Cloudflare Turnstile)."],
    ["CU05", "International Phone Input + Geo-Detection", "3 pages", "Contact forms",
     "Uses intl-tel-input library with country flag/code selector. Chatbot Bookings auto-detects country via ipinfo.io. Requires JavaScript library integration."],
    ["CU06", "Cascading Location Dropdowns", "3 pages", "Contact forms",
     "Hotel selection triggers dynamic population of inquiry type dropdown. Country triggers city selection. Server-side cascading must be reimplemented client-side."],
    ["CU07", "Login / Authentication System", "All pages", "Header (all pages)",
     "Login modal with email/password, remember-me. Server-side Kentico auth. Requires new authentication backend if user accounts need migration."],
    ["CU08", "Multi-Language / Locale Support", "All pages", "URL structure (/en-uk/...)",
     "URL structure includes locale prefix. Only English (UK) observed. If additional languages exist, migration must account for locale-based routing."],
    ["CU09", "Hotel Location Hierarchy", "~60 pages", "Hotel detail pages under /find/",
     "Deep URL hierarchy: /find/{region}/{sub-region}/{city}/{hotel-name}. Each hotel has ~5 sub-pages. Complex IA must be preserved or restructured."],
    ["CU10", "Netcore Smartech Deep Integration", "All pages", "All pages + newsletter form",
     "Deeply integrated: activity tracking, push tokens, in-page creatives, newsletter handler, booking events, A/B testing, AI personalization. Requires reconnecting all touchpoints."],
    ["CU11", "PDF Factsheet Downloads", "~12 pages", "Hotel detail pages",
     "Each hotel has downloadable PDF factsheet with map. Assets need migration to EDS media storage."],
    ["CU12", "Broken Integration (Optimix Asia)", "All pages", "All pages",
     "Consistently fails with ERR_TUNNEL_CONNECTION_FAILED. Non-functional; should be removed during migration."],
]

table = doc.add_table(rows=1 + len(complex_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["ID", "Use Case", "Instances", "Where Found", "Why It's Complex"])
for i, row_data in enumerate(complex_data):
    add_data_row(table, i + 1, row_data)
for row in table.rows:
    row.cells[0].width = Cm(1.2)
    row.cells[1].width = Cm(3)
    row.cells[2].width = Cm(1.8)
    row.cells[3].width = Cm(3)
    row.cells[4].width = Cm(8.5)
doc.add_page_break()

# ═══════════════════════════════════════════
# 7. MIGRATION ESTIMATES
# ═══════════════════════════════════════════
doc.add_heading("7. Migration Estimates", level=1)
doc.add_paragraph(
    "The following estimates assume an experienced AEM EDS development team "
    "of 2-3 developers with 1 QA resource."
)

doc.add_heading("7.1 Effort Breakdown by Work Stream", level=2)

effort_data = [
    ["WS1", "Project Setup & Architecture", "EDS project structure, environments, CI/CD, content models, coding standards.", "3-4 days", "-"],
    ["WS2", "Design System Migration", "Extract design tokens (colors, typography, spacing). Implement CSS custom properties. Map fonts (Nunito, PT Sans).", "3-4 days", "-"],
    ["WS3", "Global Components", "Header/Nav (B01), Footer (B02), Cookie Consent (B14), Newsletter (B12), Payment Bar (B21), Smart Traveller Banner (B13), Chat Widget (B19).", "5-7 days", "-"],
    ["WS4", "Blocks - Low Complexity", "Value Prop Cards (B05), Brand CTA (B07), Video Embed (B08), Content Section (B09), Page Title (B20), Notice Panel (B17).", "4-5 days", "-"],
    ["WS5", "Blocks - Medium Complexity", "Hero Carousel (B03), Offers Card Grid (B06), Accordion/FAQ (B11), Sub-Nav Tabs (B16), Login Modal (B15), Feature Cards (B18).", "8-10 days", "-"],
    ["WS6", "Blocks - High Complexity", "Booking Widget (B04) - city autocomplete, hourly/overnight, date-time, guests, API. Contact Form (B10) - cascading dropdowns, intl phone, CAPTCHA.", "10-14 days", "-"],
    ["WS7", "Template Development", "Create 9 page templates (T1-T9) with proper block composition and responsive layouts.", "5-7 days", "-"],
    ["WS8", "Content Migration - Automated", "Scripted migration of ~33-38 pages: content detail, FAQ, legal, listing pages.", "-", "4-5 days"],
    ["WS9", "Content Migration - Manual", "Manual migration of ~63+ hotel detail pages and form pages. Includes image/PDF asset migration.", "-", "8-10 days"],
    ["WS10", "Integration Setup", "GTM, GA4, Facebook Pixel, Bing UET, Netcore Smartech, Pantheon Lab chatbot, consenTag.eu, geo-conditional loading.", "5-7 days", "-"],
    ["WS11", "Authentication System", "Login/registration flow - AEM-native or third-party auth. User account migration if required.", "3-5 days", "-"],
    ["WS12", "Booking API Integration", "Connect booking widget to backend API. City search, availability, booking submission. Hourly + overnight modes.", "5-8 days", "-"],
    ["WS13", "QA & Testing", "Cross-browser, responsive, accessibility (WCAG 2.1 AA), performance (Lighthouse 100), visual regression, form testing, analytics verification.", "-", "8-10 days"],
    ["WS14", "UAT & Bug Fixes", "User acceptance testing, bug fixing, content review, stakeholder sign-off.", "-", "5-7 days"],
]

table = doc.add_table(rows=1 + len(effort_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["ID", "Work Stream", "Description", "Dev Effort", "Migration/QA"])
for i, row_data in enumerate(effort_data):
    add_data_row(table, i + 1, row_data)
for row in table.rows:
    row.cells[0].width = Cm(1)
    row.cells[1].width = Cm(3)
    row.cells[2].width = Cm(7)
    row.cells[3].width = Cm(2.5)
    row.cells[4].width = Cm(3)

doc.add_paragraph("")

doc.add_heading("7.2 Overall Estimates Summary", level=2)

summary_effort = [
    ["Development (Block & Template)", "47-67 days", "WS1-WS7, WS10-WS12"],
    ["Automated Content Migration", "4-5 days", "WS8"],
    ["Manual Content Migration", "8-10 days", "WS9"],
    ["QA & Testing", "8-10 days", "WS13"],
    ["UAT & Bug Fixes", "5-7 days", "WS14"],
    ["TOTAL (Sequential)", "72-99 days", "~15-20 weeks with 2-3 devs"],
    ["TOTAL (Parallel Execution)", "50-65 days", "~10-13 weeks with 2-3 devs + 1 QA"],
]

table = doc.add_table(rows=1 + len(summary_effort), cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(table, 0, ["Category", "Effort Estimate", "Work Streams"])
for i, row_data in enumerate(summary_effort):
    add_data_row(table, i + 1, row_data)
    if "TOTAL" in row_data[0]:
        for j in range(3):
            cell = table.rows[i+1].cells[j]
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True

doc.add_paragraph("")

doc.add_heading("7.3 Key Assumptions & Risks", level=2)
assumptions = [
    "Booking backend API exists and is accessible. If the booking system needs to be rebuilt, add 15-20 additional days.",
    "User authentication requirements are limited to basic login. If SSO, OAuth, or complex role-based access is needed, add 5-10 days.",
    "Content is primarily in English (UK). If multi-language migration is required, multiply content migration effort by number of languages.",
    "Hotel detail pages share a consistent template. If each hotel has significant custom content, manual migration effort increases.",
    "Netcore Smartech integration can be reconnected via JavaScript SDK. If deep server-side integration is required, add 3-5 days.",
    "The broken Optimix Asia integration will be removed, not replaced.",
    "PDF factsheets and image assets can be batch-migrated to EDS media storage.",
    "Performance target is Lighthouse score of 100.",
]
for a in assumptions:
    doc.add_paragraph(a, style='List Bullet')

doc.add_heading("7.4 Recommended Migration Phases", level=2)
phases = [
    ["Phase 1: Foundation", "Weeks 1-3", "Project setup, design system extraction, global components (header, footer, cookie consent)."],
    ["Phase 2: Core Blocks", "Weeks 3-7", "All block development. Parallel low/medium blocks while high-complexity blocks are architected."],
    ["Phase 3: Templates & Integration", "Weeks 6-9", "Template assembly, integration reimplementation, authentication system."],
    ["Phase 4: Content Migration", "Weeks 8-11", "Automated migration scripting, followed by manual hotel/form page migration."],
    ["Phase 5: QA & Launch", "Weeks 10-13", "Testing, accessibility audit, performance optimization, UAT, staged rollout."],
]
table = doc.add_table(rows=1 + len(phases), cols=3)
table.style = 'Table Grid'
add_header_row(table, 0, ["Phase", "Timeline", "Description"])
for i, row_data in enumerate(phases):
    add_data_row(table, i + 1, row_data)
doc.add_page_break()

# ═══════════════════════════════════════════
# APPENDIX: ADDITIONAL SCREENSHOTS
# ═══════════════════════════════════════════
doc.add_heading("Appendix: Additional Screenshots", level=1)

additional_screenshots = [
    ("07-newsroom.jpeg", "News Room - Card Listing Template (T3 variation)"),
    ("09-work-with-us.jpeg", "Work with Us - Content Detail Page (T4 variation)"),
    ("08-group-bookings.jpeg", "Group Bookings - Form Page (T5 variation)"),
    ("15-chatbot-bookings.jpeg", "Chatbot Bookings / Hotel Booking Request - Form Page (T5 variation)"),
    ("12-contact-us.jpeg", "Contact Us - Accordion for Phone Numbers (T6 variation)"),
    ("14-privacy.jpeg", "Privacy Policy - Legal Page (T7)"),
    ("11-photo-gallery.jpeg", "Photo Gallery - Empty/Stub Page"),
]

for filename, caption in additional_screenshots:
    add_screenshot(doc, filename, caption, width=4.5)
    doc.add_paragraph("")

# ── Save ──
output_path = "/workspace/Aerotel_Website_Analysis_Report.docx"
doc.save(output_path)
print(f"Report saved to: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024 / 1024:.1f} MB")
