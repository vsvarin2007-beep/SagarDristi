import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette (Consistent Dark Tactical)
    BG_DARK = RGBColor(8, 14, 24)        # #080E18
    CARD_BG = RGBColor(14, 24, 38)       # #0E1826
    BORDER_COLOR = RGBColor(36, 56, 82)  # #243852
    BORDER_CYAN = RGBColor(91, 192, 190) # #5BC0BE
    BORDER_GREEN = RGBColor(95, 227, 138)# #5FE38A
    BORDER_AMBER = RGBColor(245, 158, 11)# #F59E0B
    TEXT_WHITE = RGBColor(255, 255, 255)
    TEXT_SLATE = RGBColor(148, 163, 184) # #94A3B8
    TEXT_MUTED = RGBColor(100, 116, 139) # #64748B
    ACCENT_CYAN = RGBColor(91, 192, 190) # #5BC0BE
    ACCENT_GREEN = RGBColor(95, 227, 138)# #5FE38A
    ACCENT_AMBER = RGBColor(245, 158, 11)# #F59E0B
    ACCENT_RED = RGBColor(239, 68, 68)   # #EF4444

    logo_path = os.path.abspath("public/graviton-cropped.png")
    has_logo = os.path.exists(logo_path)

    def add_base_slide(category, title, subtitle):
        slide = prs.slides.add_slide(blank_layout)

        # Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()

        # Top Accent Line
        top_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
        top_line.fill.solid()
        top_line.fill.fore_color.rgb = ACCENT_CYAN
        top_line.line.fill.background()

        # Header Container
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.2))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        # Category
        p0 = tf.paragraphs[0]
        p0.text = f"// {category.upper()}"
        p0.font.name = "Consolas"
        p0.font.size = Pt(10)
        p0.font.bold = True
        p0.font.color.rgb = ACCENT_CYAN
        p0.space_after = Pt(2)

        # Title
        p1 = tf.add_paragraph()
        p1.text = title
        p1.font.name = "Segoe UI"
        p1.font.size = Pt(23)
        p1.font.bold = True
        p1.font.color.rgb = TEXT_WHITE
        p1.space_after = Pt(2)

        # Subtitle
        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "Segoe UI"
            p2.font.size = Pt(11)
            p2.font.color.rgb = TEXT_SLATE

        # Footer
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(8.5), Inches(0.3))
        ftf = footer_box.text_frame
        ftf.word_wrap = True
        ftf.margin_left = ftf.margin_top = ftf.margin_right = ftf.margin_bottom = 0
        fp = ftf.paragraphs[0]
        fp.text = "SagarDrishti  |  Maritime Environmental Monitoring  ·  Team ByteSquad"
        fp.font.name = "Segoe UI"
        fp.font.size = Pt(9)
        fp.font.color.rgb = TEXT_MUTED

        if has_logo:
            slide.shapes.add_picture(logo_path, Inches(11.5), Inches(6.92), height=Inches(0.32))

        return slide

    def add_card(slide, left, top, width, height, title=None, border_color=BORDER_COLOR, bg_color=CARD_BG):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)

        if title:
            tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.18), Inches(width - 0.5), Inches(0.35))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = title.upper()
            p.font.name = "Consolas"
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = ACCENT_CYAN
        return card

    # =========================================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = RGBColor(5, 9, 15)
    bg1.line.fill.background()

    accent_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.1), Inches(3.6))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = ACCENT_CYAN
    accent_bar.line.fill.background()

    tbox = slide1.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(10.5), Inches(4.0))
    tf1 = tbox.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "PROJECT PRESENTATION"
    p.font.name = "Consolas"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.space_after = Pt(12)

    p = tf1.add_paragraph()
    p.text = "SagarDrishti"
    p.font.name = "Segoe UI"
    p.font.size = Pt(52)
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE
    p.space_after = Pt(8)

    p = tf1.add_paragraph()
    p.text = "AI-Powered Maritime Oil Spill Detection and Vessel Attribution"
    p.font.name = "Segoe UI"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN
    p.space_after = Pt(12)

    p = tf1.add_paragraph()
    p.text = "Satellite-based oil spill detection, drift analysis, and vessel movement correlation for maritime environmental monitoring."
    p.font.name = "Segoe UI"
    p.font.size = Pt(13)
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.space_after = Pt(18)

    p = tf1.add_paragraph()
    p.text = "Presented by: Team ByteSquad  |  Maritime Environmental Intelligence"
    p.font.name = "Consolas"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    if has_logo:
        slide1.shapes.add_picture(logo_path, Inches(9.8), Inches(2.1), width=Inches(2.5))

    tech_pills = [
        ("ARTIFICIAL INTELLIGENCE", "Deep Learning & U-Net++"),
        ("SATELLITE IMAGERY", "Sentinel-1 Synthetic Aperture Radar"),
        ("GEOSPATIAL ANALYSIS", "Hydrodynamic Drift & Coastal Bounds"),
        ("MARITIME INTELLIGENCE", "AIS Vessel Trajectory Tracking")
    ]
    for i, (pill_head, pill_sub) in enumerate(tech_pills):
        card_x = 1.2 + i * 2.75
        add_card(slide1, card_x, 5.75, 2.55, 1.1, None, border_color=BORDER_COLOR)
        stb = slide1.shapes.add_textbox(Inches(card_x + 0.15), Inches(5.85), Inches(2.25), Inches(0.9))
        stf = stb.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp1 = stf.paragraphs[0]
        sp1.text = pill_head
        sp1.font.name = "Consolas"
        sp1.font.size = Pt(9.5)
        sp1.font.bold = True
        sp1.font.color.rgb = ACCENT_CYAN
        sp2 = stf.add_paragraph()
        sp2.text = pill_sub
        sp2.font.name = "Segoe UI"
        sp2.font.size = Pt(9)
        sp2.font.color.rgb = TEXT_SLATE

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT
    # =========================================================================
    s2 = add_base_slide(
        "PROBLEM STATEMENT",
        "The Challenge of Maritime Oil Spill Detection",
        "Oil spills in maritime regions pose serious environmental and economic challenges. Detecting spills early and identifying their possible sources is difficult because of changing ocean conditions and incomplete vessel tracking information."
    )

    challenges = [
        ("01. DIFFICULT OIL SPILL DETECTION",
         "• Oil spills can be difficult to identify in satellite imagery.\n"
         "• Natural phenomena such as low-wind areas and biogenic films can resemble oil slicks.\n"
         "• Traditional image thresholding may produce false detections.",
         BORDER_CYAN),
        ("02. LIMITED VESSEL TRACKING",
         "• AIS provides information about vessel positions, speed, and direction.\n"
         "• AIS data may be unavailable, incomplete, or inconsistent.\n"
         "• Vessel tracking alone cannot confirm whether a ship caused an oil spill.",
         BORDER_AMBER),
        ("03. CHANGING OCEAN CONDITIONS",
         "• Wind, currents, and waves influence oil spill movement.\n"
         "• A spill may move far from its original location over time.\n"
         "• Simple backward tracking may produce inaccurate source estimates.",
         BORDER_CYAN),
        ("04. DIFFICULT SOURCE ATTRIBUTION",
         "• Multiple vessels may pass near a detected spill.\n"
         "• Identifying a potential source requires combining spatial, temporal, and movement information.\n"
         "• Multi-factor cross-analysis is necessary to prioritize candidates.",
         BORDER_GREEN)
    ]

    for i, (chead, cbody, col) in enumerate(challenges):
        x = 0.8 + (i % 2) * 5.9
        y = 1.75 + (i // 2) * 2.2
        add_card(s2, x, y, 5.7, 2.05, chead, border_color=col)
        tb = s2.shapes.add_textbox(Inches(x + 0.25), Inches(y + 0.55), Inches(5.2), Inches(1.35))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = cbody
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_SLATE

    # Objective Banner
    obj_card = add_card(s2, 0.8, 6.25, 11.6, 0.65, None, border_color=ACCENT_GREEN)
    otb = s2.shapes.add_textbox(Inches(1.0), Inches(6.32), Inches(11.2), Inches(0.5))
    otf = otb.text_frame
    otf.word_wrap = True
    otf.margin_left = otf.margin_top = otf.margin_right = otf.margin_bottom = 0
    op = otf.paragraphs[0]
    op.text = "Project Objective: Develop an intelligent system that detects oil spill candidates from satellite imagery and correlates their movement with maritime vessel data."
    op.font.name = "Segoe UI"
    op.font.size = Pt(11)
    op.font.bold = True
    op.font.color.rgb = TEXT_WHITE

    # =========================================================================
    # SLIDE 3: PROPOSED SOLUTION & ARCHITECTURE
    # =========================================================================
    s3 = add_base_slide(
        "PROPOSED SOLUTION & SYSTEM ARCHITECTURE",
        "SagarDrishti: End-to-End Intelligence Pipeline",
        "SagarDrishti combines satellite imagery analysis, ocean drift modeling, and vessel movement data to support oil spill investigation."
    )

    steps = [
        ("01. DATA INGESTION",
         "• Sentinel-1 SAR satellite imagery\n"
         "• AIS vessel information\n"
         "• Ocean current and wind data\n"
         "• Geographic and coastal information",
         BORDER_CYAN),
        ("02. SPILL DETECTION",
         "• Preprocessing of satellite imagery\n"
         "• Deep learning image segmentation\n"
         "• Identification of potential spill regions\n"
         "• Extraction of boundaries and coordinates",
         BORDER_GREEN),
        ("03. DRIFT ANALYSIS",
         "• Incorporation of wind & current data\n"
         "• Modeling of possible spill movement\n"
         "• Forward drift estimation\n"
         "• Backtracking to estimate source regions",
         BORDER_CYAN),
        ("04. VESSEL CORRELATION",
         "• Comparison of vessel tracks with origin\n"
         "• Analysis of vessel proximity and timing\n"
         "• Evaluation of speed/course changes\n"
         "• Ranking candidates for investigation",
         BORDER_AMBER),
        ("05. VISUALIZATION",
         "• Interactive maritime map interface\n"
         "• Spill location & vessel trajectories\n"
         "• Investigation & correlation tools\n"
         "• Structured analytical reporting",
         BORDER_CYAN)
    ]

    for i, (shead, sbody, col) in enumerate(steps):
        x = 0.8 + i * 2.36
        add_card(s3, x, 1.8, 2.25, 4.8, shead, border_color=col)
        tb = s3.shapes.add_textbox(Inches(x + 0.15), Inches(2.35), Inches(1.95), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = sbody
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_SLATE

    # =========================================================================
    # SLIDE 4: AI-BASED DETECTION
    # =========================================================================
    s4 = add_base_slide(
        "AI-BASED DETECTION",
        "Deep Learning-Based Satellite Image Segmentation",
        "SagarDrishti uses deep learning-based segmentation (U-Net++) to identify potential oil spill regions in Sentinel-1 SAR imagery."
    )

    det_cards = [
        ("IMAGE PREPROCESSING",
         "• Prepare satellite imagery for model input.\n"
         "• Apply normalization and image filtering.\n"
         "• Utilize available SAR polarizations (VV/VH) and derived radiometric features.\n"
         "• Remove land and non-marine areas using coastal masks.",
         BORDER_CYAN),
        ("IMAGE SEGMENTATION (U-Net++)",
         "• U-Net++ encoder-decoder architecture classifies image pixels.\n"
         "• Nested dense skip connections capture subtle low-backscatter features.\n"
         "• Separates candidate spill regions from surrounding sea surface.\n"
         "• Generates pixel-level probability masks.",
         BORDER_GREEN),
        ("SPILL REGION EXTRACTION",
         "• Identify connected components in the segmentation output.\n"
         "• Delineate spill boundaries and morphological contours.\n"
         "• Calculate geographic location (latitude/longitude coordinates).\n"
         "• Estimate approximate affected spill area when georeferencing is available.",
         BORDER_AMBER)
    ]

    for i, (dhead, dbody, col) in enumerate(det_cards):
        x = 0.8 + i * 3.93
        add_card(s4, x, 1.8, 3.8, 3.6, dhead, border_color=col)
        tb = s4.shapes.add_textbox(Inches(x + 0.2), Inches(2.35), Inches(3.4), Inches(2.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = dbody
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_SLATE

    # Detection Output Box
    add_card(s4, 0.8, 5.6, 11.6, 1.2, "DETECTION OUTPUT PIPELINE", border_color=BORDER_COLOR)
    otb4 = s4.shapes.add_textbox(Inches(1.05), Inches(6.0), Inches(11.1), Inches(0.7))
    otf4 = otb4.text_frame
    otf4.word_wrap = True
    otf4.margin_left = otf4.margin_top = otf4.margin_right = otf4.margin_bottom = 0
    p = otf4.paragraphs[0]
    p.text = "Satellite Image (Sentinel-1 SAR)   ➔   Predicted Spill Mask   ➔   Spill Candidate Coordinates   ➔   Estimated Spill Area"
    p.font.name = "Consolas"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN

    # =========================================================================
    # SLIDE 5: DRIFT ANALYSIS
    # =========================================================================
    s5 = add_base_slide(
        "DRIFT ANALYSIS",
        "Understanding Spill Movement Through Ocean Dynamics",
        "Oil spills move under the influence of environmental conditions. SagarDrishti uses drift analysis to estimate how a detected spill may have moved and to investigate possible source locations."
    )

    # Left: Environmental Factors
    add_card(s5, 0.8, 1.8, 5.7, 4.2, "KEY ENVIRONMENTAL FACTORS", border_color=BORDER_CYAN)
    tb_ef = s5.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_ef = tb_ef.text_frame
    tf_ef.word_wrap = True
    tf_ef.margin_left = tf_ef.margin_top = tf_ef.margin_right = tf_ef.margin_bottom = 0
    p = tf_ef.paragraphs[0]
    p.text = (
        "• Ocean Currents: Surface tidal currents and circulation govern primary bulk drift advection.\n\n"
        "• Wind Direction & Speed: Surface windage directly influences surface slick velocity and elongation.\n\n"
        "• Wave-Induced Movement: Wave action and Stokes drift contribute to dispersion and spreading.\n\n"
        "• Spill Observation Time: Elapsed time between initial discharge and satellite image capture.\n\n"
        "• Geographic Constraints: Coastlines, shallow reefs, and islands that confine or alter possible drift paths."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.color.rgb = TEXT_SLATE

    # Right: Drift Analysis Workflow
    add_card(s5, 6.8, 1.8, 5.7, 4.2, "DRIFT ANALYSIS WORKFLOW", border_color=BORDER_GREEN)
    tb_dw = s5.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_dw = tb_dw.text_frame
    tf_dw.word_wrap = True
    tf_dw.margin_left = tf_dw.margin_top = tf_dw.margin_right = tf_dw.margin_bottom = 0
    p = tf_dw.paragraphs[0]
    p.text = (
        "1. Spill Observation: Identify the spill location from satellite imagery.\n\n"
        "2. Environmental Data Integration: Ingest available wind and ocean current data for surrounding conditions.\n\n"
        "3. Forward Drift Modeling: Estimate possible spill movement under environmental forces.\n\n"
        "4. Backward Drift Estimation: Trace possible movement paths backward from observed spill region.\n\n"
        "5. Source Region Estimation: Identify a candidate region from which the spill may have originated."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.color.rgb = TEXT_SLATE

    # Key Point Banner
    add_card(s5, 0.8, 6.2, 11.7, 0.7, None, border_color=BORDER_AMBER)
    kptb = s5.shapes.add_textbox(Inches(1.0), Inches(6.32), Inches(11.3), Inches(0.5))
    kptf = kptb.text_frame
    kptf.word_wrap = True
    kptf.margin_left = kptf.margin_top = kptf.margin_right = kptf.margin_bottom = 0
    p = kptf.paragraphs[0]
    p.text = "Key Point: Drift analysis estimates possible source regions; it does not independently prove the exact discharge location."
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = ACCENT_AMBER

    # =========================================================================
    # SLIDE 6: VESSEL ATTRIBUTION
    # =========================================================================
    s6 = add_base_slide(
        "VESSEL ATTRIBUTION",
        "Identifying Potential Vessels Associated with an Oil Spill",
        "SagarDrishti correlates estimated spill source regions with available AIS vessel movement data to identify vessels that may have been present during the incident."
    )

    # Left: Correlation Factors
    add_card(s6, 0.8, 1.8, 5.7, 4.2, "VESSEL CORRELATION FACTORS", border_color=BORDER_CYAN)
    tb_vf = s6.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_vf = tb_vf.text_frame
    tf_vf.word_wrap = True
    tf_vf.margin_left = tf_vf.margin_top = tf_vf.margin_right = tf_vf.margin_bottom = 0
    p = tf_vf.paragraphs[0]
    p.text = (
        "1. Spatial Proximity:\n"
        "   Compare vessel positions with the estimated spill source region to identify vessels that passed near the candidate location.\n\n"
        "2. Temporal Alignment:\n"
        "   Compare vessel movement times with estimated spill observation and drift periods.\n\n"
        "3. Vessel Movement:\n"
        "   Analyze vessel speed and course changes to identify unusual movement patterns for investigation.\n\n"
        "4. AIS Data Availability:\n"
        "   Check availability and continuity of vessel tracking data, noting gaps or inconsistencies."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_SLATE

    # Right: Attribution Workflow & Output
    add_card(s6, 6.8, 1.8, 5.7, 4.2, "ATTRIBUTION WORKFLOW & OUTPUT", border_color=BORDER_GREEN)
    tb_ao = s6.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_ao = tb_ao.text_frame
    tf_ao.word_wrap = True
    tf_ao.margin_left = tf_ao.margin_top = tf_ao.margin_right = tf_ao.margin_bottom = 0
    p = tf_ao.paragraphs[0]
    p.text = (
        "Attribution Pipeline:\n"
        "Satellite Spill Detection  ➔  Source Region Estimation  ➔  AIS Track Comparison  ➔  Candidate Vessel Ranking\n\n"
        "Expected Outputs:\n"
        "• Potentially associated vessels list\n"
        "• Vessel trajectory and track visualization\n"
        "• Spatial and temporal correlation indicators\n"
        "• Ranked candidates to prioritize further inspection\n\n"
        "Analysis results are organized in structured tables for forensic review."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_SLATE

    # Important Disclaimer
    add_card(s6, 0.8, 6.2, 11.7, 0.7, None, border_color=BORDER_AMBER)
    dtb = s6.shapes.add_textbox(Inches(1.0), Inches(6.32), Inches(11.3), Inches(0.5))
    dtf = dtb.text_frame
    dtf.word_wrap = True
    dtf.margin_left = dtf.margin_top = dtf.margin_right = dtf.margin_bottom = 0
    p = dtf.paragraphs[0]
    p.text = "Important: A vessel appearing near a spill is not proof of responsibility. The system supports investigation and prioritization, not legal determination."
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = ACCENT_AMBER

    # =========================================================================
    # SLIDE 7: INTERACTIVE MARITIME VISUALIZATION
    # =========================================================================
    s7 = add_base_slide(
        "INTERACTIVE MARITIME VISUALIZATION",
        "Geospatial Monitoring & Vessel Tracking Dashboard",
        "SagarDrishti provides a visual interface for exploring detected oil spills, vessel movement, and possible spill trajectories."
    )

    dash_features = [
        ("INTERACTIVE MAP",
         "• Display maritime regions and coastlines.\n"
         "• Visualize detected spill locations.\n"
         "• Explore geographic areas of interest with pan, zoom, and layer controls.",
         BORDER_CYAN),
        ("VESSEL TRACKING",
         "• Display available AIS vessel positions.\n"
         "• Show vessel movement trajectories.\n"
         "• Present vessel information such as speed, heading, and identification when available.",
         BORDER_GREEN),
        ("SPILL VISUALIZATION",
         "• Display detected spill regions.\n"
         "• Show estimated source regions.\n"
         "• Visualize candidate drift trajectories and dispersion paths.",
         BORDER_CYAN),
        ("INVESTIGATION INTERFACE",
         "• Compare spill locations with vessel movement.\n"
         "• Explore potential vessel correlations.\n"
         "• Display analysis results in a structured, actionable dashboard interface.",
         BORDER_AMBER)
    ]

    for i, (fhead, fbody, col) in enumerate(dash_features):
        x = 0.8 + (i % 2) * 5.9
        y = 1.75 + (i // 2) * 2.15
        add_card(s7, x, y, 5.7, 2.0, fhead, border_color=col)
        tb = s7.shapes.add_textbox(Inches(x + 0.25), Inches(y + 0.55), Inches(5.2), Inches(1.35))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = fbody
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_SLATE

    # Tech Stack Banner
    add_card(s7, 0.8, 6.2, 11.7, 0.7, None, border_color=BORDER_COLOR)
    ttb = s7.shapes.add_textbox(Inches(1.0), Inches(6.32), Inches(11.3), Inches(0.5))
    ttf = ttb.text_frame
    ttf.word_wrap = True
    ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
    p = ttf.paragraphs[0]
    p.text = "Technology Stack: React  ·  TypeScript  ·  Map-based Geospatial Visualization  ·  FastAPI Backend Integration"
    p.font.name = "Consolas"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 8: PREDICTIVE SPILL MONITORING
    # =========================================================================
    s8 = add_base_slide(
        "PREDICTIVE SPILL MONITORING",
        "Spill Movement Prediction & Environmental Risk Analysis",
        "SagarDrishti can use available environmental data to estimate possible future spill movement and support monitoring of affected coastal regions."
    )

    pred_cards = [
        ("FORWARD DRIFT PREDICTION",
         "• Use the detected spill location as the starting point.\n"
         "• Incorporate available wind and ocean current forecast information.\n"
         "• Estimate possible future spill trajectories over operational time horizons.\n"
         "• Visualize predicted movement dynamically on the maritime map.",
         BORDER_CYAN),
        ("COASTAL RISK ASSESSMENT",
         "• Identify coastal regions near predicted spill trajectories.\n"
         "• Highlight shoreline areas that may require further monitoring.\n"
         "• Support the prioritization of environmentally sensitive locations (mangroves, reserves, port zones).",
         BORDER_AMBER),
        ("POTENTIAL APPLICATIONS",
         "• Environmental Monitoring: Track possible movement toward coastal areas.\n"
         "• Response Planning: Provide geographic information to support monitoring and response decisions.\n"
         "• Risk Visualization: Display predicted trajectories and potential areas of concern.",
         BORDER_GREEN)
    ]

    for i, (phead, pbody, col) in enumerate(pred_cards):
        x = 0.8 + i * 3.93
        add_card(s8, x, 1.8, 3.8, 4.2, phead, border_color=col)
        tb = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.35), Inches(3.4), Inches(3.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = pbody
        p.font.name = "Segoe UI"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_SLATE

    # Note Banner
    add_card(s8, 0.8, 6.2, 11.7, 0.7, None, border_color=BORDER_COLOR)
    ntb = s8.shapes.add_textbox(Inches(1.0), Inches(6.32), Inches(11.3), Inches(0.5))
    ntf = ntb.text_frame
    ntf.word_wrap = True
    ntf.margin_left = ntf.margin_top = ntf.margin_right = ntf.margin_bottom = 0
    p = ntf.paragraphs[0]
    p.text = "Note: Prediction accuracy depends on the quality of environmental data and the drift model. Forecasts are estimates, not guaranteed spill paths."
    p.font.name = "Segoe UI"
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 9: TECHNOLOGY STACK & PROJECT EVALUATION
    # =========================================================================
    s9 = add_base_slide(
        "TECHNOLOGY STACK & EVALUATION",
        "Technologies Used & Structured Project Evaluation Approach",
        "A clear breakdown of technologies applied across the pipeline and the methodology for evaluating system outputs."
    )

    # Left: Technologies Used
    add_card(s9, 0.8, 1.8, 5.2, 4.8, "TECHNOLOGIES USED", border_color=BORDER_CYAN)
    tb_tech = s9.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(4.7), Inches(4.1))
    tf_tech = tb_tech.text_frame
    tf_tech.word_wrap = True
    tf_tech.margin_left = tf_tech.margin_top = tf_tech.margin_right = tf_tech.margin_bottom = 0
    p = tf_tech.paragraphs[0]
    p.text = (
        "• Artificial Intelligence:\n"
        "  Deep learning  ·  U-Net++ segmentation  ·  Image analysis\n\n"
        "• Satellite & Environmental Data:\n"
        "  Sentinel-1 SAR  ·  AIS telemetry  ·  Wind & ocean current data\n\n"
        "• Backend Infrastructure:\n"
        "  Python  ·  FastAPI  ·  PyTorch  ·  Scikit-learn\n\n"
        "• Frontend Interface:\n"
        "  React  ·  Vite  ·  TypeScript  ·  Tailwind CSS  ·  MapLibre"
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.color.rgb = TEXT_SLATE

    # Right: Evaluation Approach Table
    add_card(s9, 6.3, 1.8, 6.2, 4.8, "EVALUATION APPROACH & METRICS", border_color=BORDER_GREEN)
    tb_eval = s9.shapes.add_textbox(Inches(6.55), Inches(2.35), Inches(5.7), Inches(4.1))
    tf_eval = tb_eval.text_frame
    tf_eval.word_wrap = True
    tf_eval.margin_left = tf_eval.margin_top = tf_eval.margin_right = tf_eval.margin_bottom = 0
    p = tf_eval.paragraphs[0]
    p.text = (
        "The project should be evaluated using measurable outputs:\n\n"
        "• Spill Detection:\n"
        "  Measure Precision, Recall, and F1-score across validation scenes.\n\n"
        "• Image Segmentation:\n"
        "  Measure Intersection over Union (IoU) and Dice coefficient.\n\n"
        "• Drift Analysis:\n"
        "  Measure spatial error between estimated and reference spill locations.\n\n"
        "• Vessel Correlation:\n"
        "  Measure accuracy of candidate ranking against verified benchmark cases.\n\n"
        "• System Performance:\n"
        "  Measure actual end-to-end inference and processing latency."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_SLATE

    # =========================================================================
    # SLIDE 10: CONCLUSION & FUTURE SCOPE
    # =========================================================================
    s10 = add_base_slide(
        "CONCLUSION & FUTURE SCOPE",
        "Towards Intelligent Maritime Environmental Monitoring",
        "SagarDrishti aims to improve maritime oil spill monitoring by combining satellite image analysis, environmental drift modeling, and vessel movement correlation."
    )

    # Left: Project Contributions
    add_card(s10, 0.8, 1.8, 5.7, 4.1, "PROJECT CONTRIBUTIONS", border_color=BORDER_GREEN)
    tb_pc = s10.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_pc = tb_pc.text_frame
    tf_pc.word_wrap = True
    tf_pc.margin_left = tf_pc.margin_top = tf_pc.margin_right = tf_pc.margin_bottom = 0
    p = tf_pc.paragraphs[0]
    p.text = (
        "✔ AI-based detection of potential oil spill regions from satellite imagery.\n\n"
        "✔ Integration of satellite, vessel, and environmental information.\n\n"
        "✔ Drift analysis to estimate possible spill source regions.\n\n"
        "✔ AIS-based vessel correlation to support investigation.\n\n"
        "✔ Interactive visualization of spill locations and vessel trajectories."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(10.5)
    p.font.color.rgb = TEXT_SLATE

    # Right: Future Scope
    add_card(s10, 6.8, 1.8, 5.7, 4.1, "FUTURE SCOPE & ENHANCEMENTS", border_color=BORDER_CYAN)
    tb_fs = s10.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(3.4))
    tf_fs = tb_fs.text_frame
    tf_fs.word_wrap = True
    tf_fs.margin_left = tf_fs.margin_top = tf_fs.margin_right = tf_fs.margin_bottom = 0
    p = tf_fs.paragraphs[0]
    p.text = (
        "1. Improved Detection:\n"
        "   Train and evaluate the segmentation model on larger and more diverse SAR datasets.\n\n"
        "2. Enhanced Environmental Modeling:\n"
        "   Integrate more accurate and frequently updated wind and ocean current data.\n\n"
        "3. Better Vessel Correlation:\n"
        "   Improve attribution using additional verified vessel and incident information.\n\n"
        "4. Real-Time Data Integration:\n"
        "   Explore automated ingestion of satellite and AIS data sources.\n\n"
        "5. Environmental Response Support:\n"
        "   Extend the system to help monitor coastal risk and support response planning."
    )
    p.font.name = "Segoe UI"
    p.font.size = Pt(9.5)
    p.font.color.rgb = TEXT_SLATE

    # Final Statement Banner
    add_card(s10, 0.8, 6.15, 11.7, 0.75, None, border_color=BORDER_GREEN)
    stb10 = s10.shapes.add_textbox(Inches(1.0), Inches(6.25), Inches(11.3), Inches(0.55))
    stf10 = stb10.text_frame
    stf10.word_wrap = True
    stf10.margin_left = stf10.margin_top = stf10.margin_right = stf10.margin_bottom = 0
    p = stf10.paragraphs[0]
    p.text = "Final Statement: SagarDrishti transforms maritime data into actionable environmental intelligence by connecting oil spill detection, drift analysis, and vessel tracking in a unified platform."
    p.font.name = "Segoe UI"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE

    # Save presentation
    output_path = os.path.abspath("SagarDrishti_ByteSquad_Presentation.pptx")
    prs.save(output_path)
    print("SUCCESS: Updated presentation saved at:", output_path)

    # Also save to public for browser download
    pub_path = os.path.abspath("public/SagarDrishti_ByteSquad_Presentation.pptx")
    prs.save(pub_path)
    print("SUCCESS: Public presentation saved at:", pub_path)
    return output_path

if __name__ == "__main__":
    create_presentation()
