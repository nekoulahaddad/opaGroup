import os
import re
import tempfile
from datetime import date
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Image,
    Table,
    TableStyle,
    PageBreak,
    ListFlowable,
    ListItem,
    KeepTogether,
    NextPageTemplate,
)
from PIL import Image as PILImage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
OUTPUT_PDF = os.path.join(BASE_DIR, "OPA_GROUP_CO_LTD_Company_Presentation_v2.pdf")

PRIMARY = colors.HexColor("#032E82")
PRIMARY_DARK = colors.HexColor("#021B4A")
LIGHT_BG = colors.HexColor("#F4F7FB")
DARK = colors.HexColor("#0b0f1a")
MUTED = colors.HexColor("#4a5568")
WHITE = colors.white


def _convert_image(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in {".jpg", ".jpeg", ".png"}:
        return path, None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    tmp.close()
    with PILImage.open(path) as im:
        im = im.convert("RGB")
        im.save(tmp.name, format="PNG")
    return tmp.name, tmp.name


def _image_flowable(path, max_w, max_h):
    if not path or not os.path.exists(path):
        return None
    with PILImage.open(path) as im:
        w, h = im.size
    scale = min(max_w / w, max_h / h)
    return Image(path, width=w * scale, height=h * scale)


def _load_text():
    app_js = os.path.join(BASE_DIR, "app.js")
    with open(app_js, "r", encoding="utf-8") as f:
        return f.read()


def _extract_value(text, key):
    pattern = rf"{re.escape(key)}:\s*\"([^\"]+)\""
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""


def _extract_list(text, key):
    pattern = rf"{re.escape(key)}:\s*\[(.*?)\]"
    match = re.search(pattern, text, re.S)
    if not match:
        return []
    items = re.findall(r"\"([^\"]+)\"", match.group(1))
    return [i.strip() for i in items]


def _extract_services(text, key):
    pattern = rf"{re.escape(key)}:\s*\[(.*?)\]"
    match = re.search(pattern, text, re.S)
    if not match:
        return []
    block = match.group(1)
    entries = re.findall(r"\{\s*title:\s*\"([^\"]+)\",\s*text:\s*\"([^\"]+)\"\s*\}", block)
    return [(t.strip(), d.strip()) for t, d in entries]


def _extract_offices(text):
    pattern = r"const offices = \[(.*?)\];"
    match = re.search(pattern, text, re.S)
    if not match:
        return []
    block = match.group(1)
    items = []
    for chunk in re.findall(r"\{(.*?)\}", block, re.S):
        country = re.search(r"country:\s*\"([^\"]+)\"", chunk)
        address = re.search(r"address:\s*\"([^\"]+)\"", chunk)
        phones = re.findall(r"\"([^\"]+)\"", re.search(r"phones:\s*\[(.*?)\]", chunk, re.S).group(1))
        if country and address:
            items.append({
                "country": country.group(1).strip(),
                "address": address.group(1).strip(),
                "phones": [p.strip() for p in phones],
            })
    return items


def build_pdf():
    text = _load_text()

    about_p1 = _extract_value(text, "about_p1")
    about_p2 = _extract_value(text, "about_p2")
    vision_text = _extract_value(text, "vision_text")
    mission_text = _extract_value(text, "mission_text")
    values_list = _extract_list(text, "values_list")
    ceo_text = _extract_value(text, "ceo_text")

    services_core = _extract_services(text, "services_core")
    services_market = _extract_services(text, "services_market")
    services_extended = _extract_services(text, "services_extended")

    products_list = _extract_list(text, "products_list")
    sectors_list = _extract_list(text, "sectors_list")

    presence_countries = _extract_value(text, "presence_countries_text")
    presence_note = _extract_value(text, "presence_note_text")

    offices = _extract_offices(text)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleXL", parent=styles["Title"], fontSize=34, leading=38, textColor=WHITE, spaceAfter=8))
    styles.add(ParagraphStyle(name="CoverSub", parent=styles["BodyText"], fontSize=13, leading=18, textColor=WHITE))
    styles.add(ParagraphStyle(name="H1", parent=styles["Heading1"], fontSize=20, leading=24, textColor=PRIMARY, spaceAfter=10))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontSize=14, leading=18, textColor=PRIMARY, spaceAfter=6))
    styles.add(ParagraphStyle(name="Body", parent=styles["BodyText"], fontSize=10.5, leading=15, textColor=DARK))
    styles.add(ParagraphStyle(name="Muted", parent=styles["BodyText"], fontSize=9.5, leading=13, textColor=MUTED))
    styles.add(ParagraphStyle(name="Link", parent=styles["BodyText"], fontSize=10.5, leading=14, textColor=PRIMARY))

    doc = BaseDocTemplate(
        OUTPUT_PDF,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
        title="OPA GROUP CO LTD Company Presentation",
        author="OPA GROUP CO LTD",
    )

    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")

    temp_files = []
    hero_path = os.path.join(ASSETS_DIR, "hero.webp")
    distribution_path = os.path.join(ASSETS_DIR, "distribution.webp")
    logistics_path = os.path.join(ASSETS_DIR, "logistics.webp")
    products_path = os.path.join(ASSETS_DIR, "products.webp")
    network_path = os.path.join(ASSETS_DIR, "network.webp")
    contact_path = os.path.join(ASSETS_DIR, "contact2.webp")

    hero_img, hero_tmp = _convert_image(hero_path) if os.path.exists(hero_path) else (None, None)
    if hero_tmp:
        temp_files.append(hero_tmp)
    distribution_img, dist_tmp = _convert_image(distribution_path) if os.path.exists(distribution_path) else (None, None)
    if dist_tmp:
        temp_files.append(dist_tmp)
    logistics_img, log_tmp = _convert_image(logistics_path) if os.path.exists(logistics_path) else (None, None)
    if log_tmp:
        temp_files.append(log_tmp)
    products_img, prod_tmp = _convert_image(products_path) if os.path.exists(products_path) else (None, None)
    if prod_tmp:
        temp_files.append(prod_tmp)
    network_img, net_tmp = _convert_image(network_path) if os.path.exists(network_path) else (None, None)
    if net_tmp:
        temp_files.append(net_tmp)
    contact_img, contact_tmp = _convert_image(contact_path) if os.path.exists(contact_path) else (None, None)
    if contact_tmp:
        temp_files.append(contact_tmp)

    logo_path = os.path.join(ASSETS_DIR, "logo.jpeg")

    def draw_cover(canvas, _doc):
        if hero_img and os.path.exists(hero_img):
            with PILImage.open(hero_img) as im:
                w, h = im.size
            scale = max(doc.pagesize[0] / w, doc.pagesize[1] / h)
            canvas.drawImage(hero_img, 0, 0, width=w * scale, height=h * scale, mask="auto")
        canvas.setFillColor(PRIMARY_DARK)
        canvas.rect(0, 0, doc.pagesize[0], 140, fill=1, stroke=0)
        if os.path.exists(logo_path):
            canvas.drawImage(logo_path, doc.leftMargin, doc.pagesize[1] - 130, width=70, height=70, mask="auto")

    def draw_footer(canvas, _doc):
        canvas.saveState()
        canvas.setFillColor(MUTED)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(doc.leftMargin, 20, f"OPA GROUP CO LTD | Company Presentation | {date.today().isoformat()}")
        canvas.restoreState()

    cover_template = PageTemplate(id="Cover", frames=[frame], onPage=draw_cover)
    content_template = PageTemplate(id="Content", frames=[frame], onPage=draw_footer)
    doc.addPageTemplates([cover_template, content_template])

    def section_header(title):
        return Table(
            [[Paragraph(title, ParagraphStyle("SecTitle", parent=styles["H1"], textColor=WHITE))]],
            colWidths=[doc.width],
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), PRIMARY),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]),
        )

    def callout_box(title, body):
        return Table(
            [[Paragraph(f"<b>{title}</b>", styles["Body"])], [Paragraph(body, styles["Body"]) ]],
            colWidths=[doc.width],
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BG),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.lightgrey),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]),
        )

    story = []

    story.append(Spacer(1, 360))
    story.append(Paragraph("OPA GROUP CO LTD", styles["TitleXL"]))
    story.append(Paragraph("International Distribution, Logistics & Trade", styles["CoverSub"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Connecting business, empowering trade across Africa, Asia, and the Middle East.", styles["CoverSub"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("<link href=\"https://opagroupafrica.site/\">https://opagroupafrica.site/</link>", styles["CoverSub"]))

    story.append(NextPageTemplate("Content"))
    story.append(PageBreak())

    story.append(section_header("About OPA Group"))
    story.append(Spacer(1, 10))

    about_image = _image_flowable(distribution_img, max_w=2.6 * inch, max_h=3.3 * inch) if distribution_img else None
    about_left = [Paragraph(about_p1, styles["Body"]), Spacer(1, 6), Paragraph(about_p2, styles["Body"])]
    about_right = [about_image] if about_image else []
    about_table = Table(
        [[about_left, about_right]],
        colWidths=[doc.width * 0.62, doc.width * 0.38],
        style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]),
    )
    story.append(about_table)
    story.append(Spacer(1, 12))

    vision_mission = Table(
        [[
            Paragraph(f"<b>Vision</b><br/>{vision_text}", styles["Body"]),
            Paragraph(f"<b>Mission</b><br/>{mission_text}", styles["Body"]),
        ]],
        colWidths=[doc.width / 2 - 6, doc.width / 2 - 6],
        style=TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BG),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]),
    )
    story.append(vision_mission)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Core Values", styles["H2"]))
    story.append(ListFlowable(
        [ListItem(Paragraph(v, styles["Body"])) for v in values_list],
        bulletType="bullet",
        leftIndent=14,
        bulletColor=PRIMARY,
    ))
    story.append(Spacer(1, 12))

    story.append(callout_box("Message from the CEO", ceo_text))

    story.append(PageBreak())

    story.append(section_header("Core Services"))
    story.append(Spacer(1, 10))

    services_image = _image_flowable(logistics_img, max_w=6.9 * inch, max_h=2.9 * inch) if logistics_img else None
    if services_image:
        story.append(services_image)
        story.append(Spacer(1, 10))

    def services_list(items):
        bullets = [ListItem(Paragraph(f"<b>{t}</b> — {d}", styles["Body"])) for t, d in items]
        return ListFlowable(bullets, bulletType="bullet", leftIndent=14, bulletColor=PRIMARY)

    story.append(Paragraph("Logistics & Supply Chain", styles["H2"]))
    story.append(services_list(services_core))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Market Support", styles["H2"]))
    story.append(services_list(services_market))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Extended Solutions", styles["H2"]))
    story.append(services_list(services_extended))

    story.append(PageBreak())

    story.append(section_header("Products & Market Sectors"))
    story.append(Spacer(1, 10))

    products_image = _image_flowable(products_img, max_w=6.9 * inch, max_h=2.9 * inch) if products_img else None
    if products_image:
        story.append(products_image)
        story.append(Spacer(1, 10))

    products_list_flow = ListFlowable(
        [ListItem(Paragraph(p, styles["Body"])) for p in products_list],
        bulletType="bullet",
        leftIndent=14,
        bulletColor=PRIMARY,
    )

    sectors_list_flow = ListFlowable(
        [ListItem(Paragraph(s, styles["Body"])) for s in sectors_list],
        bulletType="bullet",
        leftIndent=14,
        bulletColor=PRIMARY,
    )

    products_table = Table(
        [[
            [Paragraph("Products", styles["H2"]), products_list_flow],
            [Paragraph("Market Sectors Served", styles["H2"]), sectors_list_flow],
        ]],
        colWidths=[doc.width * 0.52, doc.width * 0.48],
        style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]),
    )

    story.append(products_table)

    story.append(PageBreak())

    story.append(section_header("Geographical Presence"))
    story.append(Spacer(1, 10))

    presence_image = _image_flowable(network_img, max_w=2.6 * inch, max_h=3.2 * inch) if network_img else None

    presence_left = [
        Paragraph("A growing network across Africa, Asia, and the Middle East.", styles["Body"]),
        Spacer(1, 6),
        Paragraph(f"<b>Countries of Operation:</b> {presence_countries}", styles["Body"]),
        Spacer(1, 6),
        Paragraph(presence_note, styles["Body"]),
    ]
    presence_right = [presence_image] if presence_image else []

    presence_table = Table(
        [[presence_left, presence_right]],
        colWidths=[doc.width * 0.62, doc.width * 0.38],
        style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]),
    )
    story.append(presence_table)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Our Head offices", styles["H2"]))
    office_rows = [["Country", "Address", "Phone"]]
    for o in offices:
        office_rows.append([o["country"], o["address"], " | ".join(o["phones"])])

    office_table = Table(office_rows, colWidths=[1.6 * inch, 3.5 * inch, 1.6 * inch], repeatRows=1)
    office_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(office_table)

    story.append(PageBreak())

    story.append(section_header("Contact"))
    story.append(Spacer(1, 10))

    contact_lines = [
        "Website: <link href=\"https://opagroupafrica.site/\">opagroupafrica.site</link>",
        "Email: <link href=\"mailto:info@opagroupafrica.com\">info@opagroupafrica.com</link>",
        "Email: <link href=\"mailto:opagroup66@gmail.com\">opagroup66@gmail.com</link>",
        "Email: <link href=\"mailto:ceo@opagroupafrica.com\">ceo@opagroupafrica.com</link>",
        "Headquarter Phone: <link href=\"tel:+27842600556\">+27 84 260 0556</link>",
        "WhatsApp: <link href=\"https://wa.me/254758896370\">+254 758 896 370</link>",
        "Facebook: <link href=\"https://www.facebook.com/opagroupcompamy\">OPA GROUP CO. LTD</link>",
        "TikTok: <link href=\"https://www.tiktok.com/@opa.group.co.ltd\">@opa.group.co.ltd</link>",
        "Instagram: <link href=\"https://www.instagram.com/opa_group_co_ltd/\">@opa_group_co_ltd</link>",
        "LinkedIn: <link href=\"https://www.linkedin.com/in/opa-group-co-ltd-251bb3396/\">OPA GROUP CO. LTD</link>",
    ]

    contact_left = [Paragraph("We respond quickly with next steps for partners and clients.", styles["Body"]), Spacer(1, 8)]
    for line in contact_lines:
        contact_left.append(Paragraph(line, styles["Body"]))
        contact_left.append(Spacer(1, 4))

    contact_image = _image_flowable(contact_img, max_w=2.8 * inch, max_h=3.2 * inch) if contact_img else None
    contact_right = [contact_image] if contact_image else []

    contact_table = Table(
        [[contact_left, contact_right]],
        colWidths=[doc.width * 0.62, doc.width * 0.38],
        style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]),
    )
    story.append(contact_table)

    doc.build(story)

    for tmp in temp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass


if __name__ == "__main__":
    build_pdf()
    print(f"Generated: {OUTPUT_PDF}")
