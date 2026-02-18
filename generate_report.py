#!/usr/bin/env python3
"""Generate ALLWAYS VIP Site Analysis Word Document."""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

SCREENSHOTS_DIR = "/workspace/screenshots"

def set_cell_shading(cell, color):
    """Set cell background color."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_styled_table(doc, headers, rows, col_widths=None):
    """Add a styled table with purple header row."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = 'Table Grid'

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_shading(cell, "3C1053")

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(cell_text)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(8.5)
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F3F0F5")

    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)

    return table

def add_screenshot(doc, filename, caption, width=5.5):
    """Add a screenshot image with caption."""
    filepath = os.path.join(SCREENSHOTS_DIR, filename)
    if os.path.exists(filepath):
        doc.add_picture(filepath, width=Inches(width))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.style = 'Caption' if 'Caption' in [s.name for s in doc.styles] else 'Normal'
        for run in cap.runs:
            run.italic = True
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    else:
        doc.add_paragraph(f"[Screenshot not found: {filename}]")

def main():
    doc = Document()

    # -- Page setup --
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # -- Styles --
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10)

    # ========== COVER PAGE ==========
    for _ in range(6):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("ALLWAYS VIP")
    run.bold = True
    run.font.size = Pt(36)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Site Analysis & Migration Assessment")
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x00, 0xB4, 0xB4)

    doc.add_paragraph()

    url_p = doc.add_paragraph()
    url_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = url_p.add_run("https://www.allwaysvip.com")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run("February 18, 2026")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

    # ========== TABLE OF CONTENTS ==========
    h = doc.add_heading('Table of Contents', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    toc_items = [
        "1. Templates Inventory",
        "2. Blocks / Components Catalog",
        "3. Page Counts by Template",
        "4. Integrations Analysis",
        "5. Complex Use Cases & Observations",
        "6. Migration Estimates",
    ]
    for item in toc_items:
        p = doc.add_paragraph(item)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
            run.font.size = Pt(11)

    doc.add_page_break()

    # ========== EXECUTIVE SUMMARY ==========
    h = doc.add_heading('Executive Summary', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    doc.add_paragraph(
        "ALLWAYS VIP (allwaysvip.com) is a Drupal-based airport concierge services platform "
        "operated by Plaza Premium Group. The site enables travelers to discover and book "
        "airport services including Meet & Assist, Concierge, Fast Track, Porter, Limousine, "
        "and more across 113+ airports worldwide."
    )
    doc.add_paragraph(
        "This document provides a comprehensive analysis of the site's structure, templates, "
        "components, integrations, and complexity to inform migration planning to Adobe Edge "
        "Delivery Services (AEM Sites)."
    )

    doc.add_paragraph()
    summary_data = [
        ["Total Pages", "~54"],
        ["Unique Templates", "10"],
        ["Reusable Blocks/Components", "19"],
        ["Third-Party Integrations", "12"],
        ["Languages Supported", "3 (EN, ZH-CN, ZH-TW)"],
        ["Currencies Supported", "20+"],
        ["Airport Locations", "113+"],
        ["CMS Platform", "Drupal (custom theme)"],
        ["Booking Platform", "Separate subdomain (booking.allwaysvip.com)"],
        ["Estimated Migration Effort", "31–47 working days"],
    ]
    add_styled_table(doc, ["Metric", "Value"], summary_data, col_widths=[6, 11])

    doc.add_page_break()

    # ========== 1. TEMPLATES INVENTORY ==========
    h = doc.add_heading('1. Templates Inventory', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    doc.add_paragraph(
        "The site uses 10 distinct page templates. Each template defines a unique layout "
        "structure, content model, and interaction pattern."
    )

    templates = [
        ["T1", "Homepage", "High",
         "Unique layout with hero image, interactive booking search widget (location autocomplete, date/time pickers, service type radios, guest selector, promo code), embedded YouTube video, service cards grid. Heavy JavaScript interactivity.",
         "https://www.allwaysvip.com/"],
        ["T2", "Service Listing", "Medium",
         "Card grid layout with search/filter. Cards dynamically filtered by location (?loc= param). Breadcrumb + title + search bar + card grid.",
         "https://www.allwaysvip.com/our-services"],
        ["T3", "Service Detail", "Medium",
         "Hero image + rich text body + dynamic airport availability list (up to 113 airports). Includes 'Other Services' related content section.",
         "https://www.allwaysvip.com/our-services/meet-and-assist"],
        ["T4", "Offers Listing", "Medium",
         "Hero banner with geometric design + card grid with search. Same card pattern as Service Listing but different content type. Location-filtered.",
         "https://www.allwaysvip.com/offers"],
        ["T5", "Offer Detail", "High",
         "Two-column layout: left with promotional image, language tabs (EN/ES), rich text T&C; right with embedded booking widget. 'Other Offers' section at bottom.",
         "https://www.allwaysvip.com/offers/madrid-opening-offer"],
        ["T6", "Content Page", "Low",
         "Simple single-column: breadcrumb, title, optional YouTube video, body text paragraphs.",
         "https://www.allwaysvip.com/about-us"],
        ["T7", "Contact Form Page", "High",
         "Multi-field Drupal webform (enquiry type, personal info, booking ref, location, date/time) + 116-row contact directory table + headquarters info.",
         "https://www.allwaysvip.com/contact_us"],
        ["T8", "Registration Form", "High",
         "Smart Traveller registration: personal details, GDPR toggle, data privacy accordion, T&C checkbox. External rewards system integration.",
         "https://www.allwaysvip.com/form/arrture-register"],
        ["T9", "Legal / Policy Page", "Low",
         "Long-form text with numbered sections, bullet lists, hyperlinks. No interactive elements.",
         "https://www.allwaysvip.com/terms-of-use"],
        ["T10", "FAQ Page", "Low",
         "Q&A list with styled question headings and answer paragraphs separated by horizontal rules.",
         "https://www.allwaysvip.com/faqs"],
    ]

    add_styled_table(
        doc,
        ["#", "Template Name", "Complexity", "Description", "Reference URL"],
        templates,
        col_widths=[1, 3, 2, 7, 4]
    )

    # Template screenshots
    doc.add_paragraph()
    doc.add_heading('Template Screenshots', level=2)

    screenshot_map = [
        ("homepage-clean.png", "T1 — Homepage with booking widget and service cards"),
        ("our-services-listing.png", "T2 — Service Listing with card grid"),
        ("service-detail-meet-assist.png", "T3 — Service Detail (Meet & Assist)"),
        ("offers-listing.png", "T4 — Offers Listing with hero banner"),
        ("offer-detail.png", "T5 — Offer Detail with two-column layout and booking sidebar"),
        ("about-us.png", "T6 — Content Page (About Us) with video embed"),
        ("contact-us.png", "T7 — Contact Form Page with directory table"),
        ("register-form.png", "T8 — Registration Form (Smart Traveller)"),
        ("terms-of-use.png", "T9 — Legal Page (Terms of Use)"),
        ("faqs.png", "T10 — FAQ Page"),
    ]

    for filename, caption in screenshot_map:
        add_screenshot(doc, filename, caption, width=4.5)
        doc.add_paragraph()

    doc.add_page_break()

    # ========== 2. BLOCKS / COMPONENTS CATALOG ==========
    h = doc.add_heading('2. Blocks / Components Catalog', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    doc.add_paragraph(
        "The site uses 19 reusable blocks/components. Where blocks share the same content model "
        "but differ visually, they are cataloged as design variations rather than separate blocks."
    )

    blocks = [
        ["B1", "Header / Navigation", "High",
         "Sticky dark purple header: ALLWAYS SVG logo, 3-item main nav (Our Services, Latest Offers, Explore dropdown), Login/Register, language/currency selector (3 langs, 17+ currencies), off-canvas mobile hamburger menu.",
         "All pages"],
        ["B2", "Footer — Brand Showcase", "Medium",
         "Dark purple section: Plaza Premium Group master logo, categorized brand logos (Lounge, Concierge, Hotel, F&B, Rewards). External links to sibling brands.",
         "All pages"],
        ["B3", "Footer — Navigation Bar", "Low",
         "Dark gray bar: Terms of Use, Privacy Policy, Contact Us links + copyright notice.",
         "All pages"],
        ["B4", "Hero Banner with Booking Widget", "High",
         "Large hero image with diagonal geometric overlay. Interactive booking form: location autocomplete, service type tabs, journey type radios, date/time pickers, guest selector, promo code, Search CTA. AJAX-powered.",
         "Homepage, Offer Detail sidebar"],
        ["B5", "Service / Offer Card", "Low",
         "Reusable card: thumbnail image, H4 title, excerpt text, teal 'Explore >' CTA. Same model for services and offers. Responsive grid layout.",
         "Services Listing, Offers Listing, Related sections"],
        ["B6", "Breadcrumb", "Low",
         "Horizontal breadcrumb: Home icon > Section > Page. All interior pages.",
         "All interior pages"],
        ["B7", "YouTube Video Embed", "Low",
         "Standard YouTube iframe embed within content area.",
         "Homepage, About Us"],
        ["B8", "Section Heading with Subtitle", "Low",
         "Centered H2 heading with optional styled subtitle. Used for content section introductions.",
         "Homepage, Listing pages"],
        ["B9", "Related Content Section", "Medium",
         "Dynamic section showing related items using card component (B5). Two variants: location-aware and non-location-aware. Drupal Views-powered.",
         "Service Detail, Offer Detail"],
        ["B10", "Contact Form (Webform)", "High",
         "Drupal webform: enquiry dropdown, name fields, email, phone with country code, booking ref, location, journey type, IATA code, date/time, message, submit. Includes validation.",
         "Contact Us"],
        ["B11", "Contact Directory Table", "Medium",
         "Large data table (116 rows × 4 cols): Service Location, Airport, Contact Number, Operating Hours. Alphabetical. Headquarters section below.",
         "Contact Us"],
        ["B12", "Registration Form", "High",
         "Smart Traveller registration: personal details, email confirmation, phone with country code, password, GDPR toggle, data privacy accordion, T&C checkbox, Reset/Submit.",
         "Register page"],
        ["B13", "FAQ List", "Low",
         "Q&A list: styled teal question headings, answer paragraphs, horizontal rule separators. Static content.",
         "FAQs"],
        ["B14", "Cookie Consent Banner", "Low",
         "Fixed dialog: privacy notice text, 'here' details button, 'Understood' dismiss. ConsenTag + Drupal EU Cookie Compliance.",
         "All pages (first visit)"],
        ["B15", "Login Modal", "Medium",
         "Hidden popup: email/password fields, Forgot Password, Register link. Triggered by header Login link. Arrture auth integration.",
         "All pages (via header)"],
        ["B16", "Language / Currency Selector", "Medium",
         "Header dropdown: 3 languages (EN, ZH-CN, ZH-TW), 17+ currencies. sessionStorage persistence. Page reload on change.",
         "All pages (header)"],
        ["B17", "Offers Hero Banner", "Low",
         "Decorative banner: diagonal turquoise geometric shapes on purple + overlaid text. Design variation of B4 without booking widget.",
         "Offers Listing"],
        ["B18", "Search / Filter Bar", "Low",
         "Inline search input with magnifier icon, right-aligned next to page title. Text-based card filtering.",
         "Services Listing, Offers Listing"],
        ["B19", "Language Tabs (Content)", "Low",
         "Horizontal tab switcher (e.g., English | Español) within offer content to toggle translated versions.",
         "Offer Detail pages"],
    ]

    add_styled_table(
        doc,
        ["#", "Block Name", "Complexity", "Description & Behaviour", "Reference URL(s)"],
        blocks,
        col_widths=[1, 3.5, 2, 7, 3.5]
    )

    doc.add_page_break()

    # ========== 3. PAGE COUNTS BY TEMPLATE ==========
    h = doc.add_heading('3. Page Counts by Template', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    page_counts = [
        ["T1 — Homepage", "1", "Manual — Unique layout, complex booking widget, video, dynamic location detection"],
        ["T2 — Service Listing", "1", "Semi-Auto — Standardized card grid, but location-aware dynamic filtering requires custom logic"],
        ["T3 — Service Detail", "12", "Automatic — Consistent template: image + body text + airport list. Standardized structure."],
        ["T4 — Offers Listing", "1", "Semi-Auto — Same pattern as Service Listing. Location-filtered card grid."],
        ["T5 — Offer Detail", "26", "Semi-Auto — Consistent two-column layout, content varies (language tabs, T&Cs, images). Booking widget sidebar is dynamic."],
        ["T6 — Content Page", "3", "Automatic — Simple text + optional video. Low complexity."],
        ["T7 — Contact Form", "1–2", "Manual — Complex webform + 116-row contact directory. Form logic needs reimplementation."],
        ["T8 — Registration Form", "2", "Manual — External Smart Traveller/Arrture auth system. Cannot be static content."],
        ["T9 — Legal / Policy", "3", "Automatic — Pure long-form text. No interactivity."],
        ["T10 — FAQ", "1", "Automatic — Simple Q&A content. No dynamic behavior."],
    ]

    add_styled_table(
        doc,
        ["Template", "Page Count", "Migration Approach"],
        page_counts,
        col_widths=[4, 2, 11]
    )

    doc.add_paragraph()

    h2 = doc.add_heading('Migration Category Summary', level=2)
    summary_rows = [
        ["Automatically migratable", "19", "12 service details + 3 content pages + 3 legal + 1 FAQ"],
        ["Semi-automatically migratable", "28", "26 offer details + 1 service listing + 1 offer listing"],
        ["Manual migration required", "5–7", "Homepage + contact form + 2 registration/auth forms + booking pages"],
        ["Total", "~54", "All pages from sitemap and navigation"],
    ]
    add_styled_table(
        doc,
        ["Category", "Page Count", "Details"],
        summary_rows,
        col_widths=[5, 2, 10]
    )

    doc.add_page_break()

    # ========== 4. INTEGRATIONS ANALYSIS ==========
    h = doc.add_heading('4. Integrations Analysis', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    integrations = [
        ["I1", "Booking Platform\n(booking.allwaysvip.com)", "Custom Application", "High",
         "Separate subdomain with custom booking/reservation system. Bootstrap + Swiper.js. Full booking workflow: service selection, date/time, guest details, payment.",
         "booking.allwaysvip.com"],
        ["I2", "Stripe Payment Gateway", "API / Embed", "High",
         "Stripe Elements for PCI-compliant credit card processing on booking checkout.",
         "booking.allwaysvip.com (checkout)"],
        ["I3", "Smart Traveller / Arrture Auth", "API / Custom", "High",
         "User registration, login, profile management. GDPR consent collection. Separate profile at booking.allwaysvip.com/profile.",
         "/form/arrture-register, login modal"],
        ["I4", "Google Tag Manager", "Embed", "Medium",
         "Container GTM-N6GBBND. Orchestrates all analytics/marketing tags. Likely deploys GA4, Facebook Pixel via server config.",
         "All pages"],
        ["I5", "ConsenTag (Cookie Consent)", "Embed / Plugin", "Medium",
         "Container 67407556. Cookie consent management controlling tracking scripts. silentMode: true. Coordinates with GTM.",
         "All pages"],
        ["I6", "Drupal EU Cookie Compliance", "Plugin", "Low",
         "Drupal module providing cookie consent banner UI with dismiss button.",
         "All pages"],
        ["I7", "AffiliRed (Affiliate Tracking)", "Embed", "Medium",
         "Merchant ID 4805. Affiliate marketing network for partner referral tracking and conversion attribution.",
         "All pages (global script)"],
        ["I8", "Currency Conversion API", "API (Internal)", "Medium",
         "AJAX POST endpoint at /currency_convert. 20+ currencies (AED, AUD, BRL, CAD, CNY, EUR, GBP, HKD, JPY, USD, etc.).",
         "Homepage, booking widget"],
        ["I9", "Location Autoselect", "Custom Code", "Medium",
         "Custom Drupal module detecting user location, auto-selecting nearest airport. Drives ?loc= URL parameter. Partner/white-label detection.",
         "All pages"],
        ["I10", "YouTube Embeds", "Embed", "Low",
         "Standard YouTube iframe player for promotional videos.",
         "Homepage, About Us"],
        ["I11", "Netcore Smartech", "Embed", "Low",
         "Marketing automation / push notification service from osjs.netcoresmartech.com.",
         "All pages"],
        ["I12", "White-Label Partner System", "Custom Code", "High",
         "Partner code detection via sessionStorage (e.g., 'VISACHINA' triggers rebrand to 'YQNow'). B2B distribution with custom branding per partner.",
         "Booking platform, offer pages"],
    ]

    add_styled_table(
        doc,
        ["#", "Integration", "Type", "Complexity", "Description", "Reference URL(s)"],
        integrations,
        col_widths=[1, 3, 2, 1.5, 6, 3.5]
    )

    doc.add_page_break()

    # ========== 5. COMPLEX USE CASES ==========
    h = doc.add_heading('5. Complex Use Cases & Observations', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    complex_cases = [
        ["C1", "Location-Aware Content Filtering", "All pages",
         "Every page uses ?loc= parameter; services/offers filtered by airport. ~113 airports supported.",
         "Content varies dynamically per airport. Migration must replicate logic or create location-specific pages."],
        ["C2", "Booking Search Widget", "2 instances + booking platform",
         "Homepage (full-page), Offer Detail (sidebar), booking.allwaysvip.com",
         "Complex interactive form: location autocomplete (AJAX), date/time pickers, guest counter. Cannot be replicated as static EDS content."],
        ["C3", "Multi-Currency Display", "All priced pages",
         "Currency selector in header, AJAX conversion API",
         "Real-time conversion across 20+ currencies without page reload. Requires client-side API or pre-computed variants."],
        ["C4", "Multi-Language Content", "All pages + offer details",
         "Header switcher (EN/ZH-CN/ZH-TW), content language tabs on offers",
         "3 languages. Some offer pages have inline tabs (English/Español). Session-based language persistence."],
        ["C5", "White-Label Partner Branding", "10+ partners",
         "Booking platform, offer pages (Perkopolis, MTA Travel, VIP Ontario Tours, etc.)",
         "Partner-specific branding via session storage. Changes logos, colors, services per partner code."],
        ["C6", "Authentication & User Profiles", "3 pages + modal",
         "/form/arrture-register, /form/forgot-password, /view-profile, login modal",
         "External Arrture/Smart Traveller auth system. Session-dependent content (logged-in vs. guest)."],
        ["C7", "Location Redirect Modal", "All pages (conditional)",
         "Triggered when user is at location without digital booking",
         "Modal warns of redirect to 'original booking site'. Conditional logic based on airport capabilities."],
        ["C8", "Dynamic Airport Lists", "12 service detail pages",
         "Each service detail page",
         "Airport lists vary per service (Meet & Assist: 113, Limousine: 3). Centrally managed, filtered per service."],
    ]

    add_styled_table(
        doc,
        ["#", "Complex Behaviour", "Instances", "Where Found", "Why It's Complex"],
        complex_cases,
        col_widths=[1, 3, 2.5, 4, 6.5]
    )

    doc.add_page_break()

    # ========== 6. MIGRATION ESTIMATES ==========
    h = doc.add_heading('6. Migration Estimates', level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    doc.add_heading('Effort Breakdown', level=2)

    # Phase 1
    p = doc.add_paragraph()
    run = p.add_run("Phase 1: Foundation Setup")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    phase1 = [
        ["Project scaffolding & design tokens", "Global styles, fonts, colors, CSS custom properties", "2–3 days",
         "Extract from Drupal theme, establish EDS global styles"],
        ["Header / Navigation block", "Responsive header, mobile menu, language selector", "3–4 days",
         "Complex: dropdown, off-canvas mobile, language/currency switcher needs custom JS"],
        ["Footer block", "Brand showcase + nav footer", "2 days",
         "Multi-section footer with logo grid, categorized brands"],
    ]
    add_styled_table(doc, ["Task", "Scope", "Effort", "Notes"], phase1, col_widths=[4, 4.5, 2, 6.5])

    doc.add_paragraph()

    # Phase 2
    p = doc.add_paragraph()
    run = p.add_run("Phase 2: Content Templates")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    phase2 = [
        ["Content pages (T6, T9, T10)", "About, Legal, FAQ — 7 pages", "1–2 days",
         "Simple text, auto-migratable with import script"],
        ["Service Detail (T3)", "12 pages", "2–3 days",
         "Consistent template. Build once, import script handles rest."],
        ["Service Listing (T2)", "1 page", "1–2 days",
         "Card grid with search filter. Custom block for card rendering."],
        ["Offer Detail (T5)", "26 pages", "3–4 days",
         "Two-column layout, language tabs, T&C. Semi-automated after template."],
        ["Offer Listing (T4)", "1 page", "1 day",
         "Same pattern as Service Listing"],
        ["Homepage (T1)", "1 page", "3–5 days",
         "Unique layout, hero, video, cards. Booking widget = link/redirect."],
    ]
    add_styled_table(doc, ["Task", "Scope", "Effort", "Notes"], phase2, col_widths=[4, 4.5, 2, 6.5])

    doc.add_paragraph()

    # Phase 3
    p = doc.add_paragraph()
    run = p.add_run("Phase 3: Complex Components")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    phase3 = [
        ["Booking Widget integration", "Homepage + Offer Detail sidebar", "2–3 days",
         "Options: iframe embed, redirect to booking.allwaysvip.com, or simplified CTA."],
        ["Contact Form (T7)", "1 page", "2–3 days",
         "Webform replacement (AEM Forms or third-party). Contact directory table."],
        ["Registration / Auth (T8)", "2 pages", "1 day",
         "Likely remain on Arrture platform. Migration = redirect links."],
    ]
    add_styled_table(doc, ["Task", "Scope", "Effort", "Notes"], phase3, col_widths=[4, 4.5, 2, 6.5])

    doc.add_paragraph()

    # Phase 4
    p = doc.add_paragraph()
    run = p.add_run("Phase 4: Dynamic Features")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    phase4 = [
        ["Location-aware content", "All pages", "3–5 days",
         "Strategy: static pages per location vs. client-side filtering. Major decision."],
        ["Multi-language support", "All pages", "3–4 days",
         "EDS multi-language setup. Content authoring in 3 languages."],
        ["Currency conversion", "Booking-related pages", "1–2 days",
         "Client-side API call or redirect to booking platform for pricing."],
    ]
    add_styled_table(doc, ["Task", "Scope", "Effort", "Notes"], phase4, col_widths=[4, 4.5, 2, 6.5])

    doc.add_paragraph()

    # Phase 5
    p = doc.add_paragraph()
    run = p.add_run("Phase 5: QA & Testing")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x3C, 0x10, 0x53)

    phase5 = [
        ["Visual regression testing", "All 54 pages", "3–4 days",
         "Compare EDS output against original site screenshots"],
        ["Functional testing", "Forms, navigation, links, responsive", "2–3 days",
         "Cross-browser, mobile responsiveness, link validation"],
        ["Performance testing", "Lighthouse, Core Web Vitals", "1–2 days",
         "Target Lighthouse score of 100"],
        ["Content review", "All migrated content", "2–3 days",
         "Accuracy check, missing content, broken images"],
    ]
    add_styled_table(doc, ["Task", "Scope", "Effort", "Notes"], phase5, col_widths=[4, 4.5, 2, 6.5])

    doc.add_paragraph()

    # Summary table
    doc.add_heading('Total Effort Summary', level=2)

    effort_summary = [
        ["Automated Migration (content, service details, legal)", "3–5 days"],
        ["Semi-Automated Migration (offers, listings)", "5–7 days"],
        ["Manual / Custom Migration (homepage, forms, booking)", "8–12 days"],
        ["Dynamic Features (location, language, currency)", "7–11 days"],
        ["QA & Testing", "8–12 days"],
        ["Total Estimated Effort", "31–47 days (~6–9 weeks, 1 developer)"],
    ]
    add_styled_table(doc, ["Category", "Effort"], effort_summary, col_widths=[10, 7])

    doc.add_paragraph()

    # Key Decisions
    doc.add_heading('Key Migration Decisions Required', level=2)

    decisions = [
        ["Booking Widget", "(A) Embed as iframe from booking.allwaysvip.com\n(B) Redirect to booking platform\n(C) Rebuild simplified version in EDS",
         "High — central to user journey"],
        ["Location-Aware Content", "(A) Static pages per airport (113 × N pages)\n(B) Client-side JS filtering\n(C) Single content with location API",
         "High — affects page count and content model"],
        ["Authentication", "(A) Keep on Arrture platform (redirect)\n(B) Integrate via client-side SDK\n(C) Rebuild in AEM",
         "High — affects user session handling"],
        ["Multi-Language", "(A) EDS folder-based i18n (/en/, /zh-cn/, /zh-tw/)\n(B) Client-side language switching",
         "Medium — triples content authoring effort"],
        ["Contact Form", "(A) AEM Forms\n(B) Third-party service (Typeform, JotForm)\n(C) Simple mailto link",
         "Medium — functional replacement needed"],
    ]
    add_styled_table(doc, ["Decision", "Options", "Impact"], decisions, col_widths=[3, 8.5, 5.5])

    doc.add_paragraph()

    # Risk Factors
    doc.add_heading('Risk Factors', level=2)

    risks = [
        "Booking platform dependency: The booking flow at booking.allwaysvip.com is a separate application and is out of scope for EDS migration. Integration strategy (iframe vs. redirect) is the highest-risk decision.",
        "Location logic complexity: 113 airports with location-specific service availability creates significant content modeling challenges.",
        "White-label partner system: Partner-specific branding may require custom EDS middleware or separate microsites.",
        "Content freshness: Offer pages are time-sensitive (flash sales, seasonal promotions). Authoring workflow must support rapid content updates.",
    ]
    for risk in risks:
        p = doc.add_paragraph(risk, style='List Bullet')
        for run in p.runs:
            run.font.size = Pt(9.5)

    # ========== SAVE ==========
    output_path = "/workspace/ALLWAYS_VIP_Site_Analysis.docx"
    doc.save(output_path)
    print(f"Document saved to {output_path}")
    print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

if __name__ == "__main__":
    main()
