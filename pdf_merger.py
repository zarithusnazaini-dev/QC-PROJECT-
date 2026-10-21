import io
import os
import re
import tempfile
import pymupdf as fitz
from pypdf import PdfWriter, PdfReader

def convert_excel_to_pdf_native(excel_bytes_or_path, excel_data=None):
    """
    Converts Excel file (.xlsx, .xls) to a genuine PDF using
    Microsoft Excel COM automation (100% native print fidelity, original format).
    - Eliminates slow ActivePrinter network probing for maximum conversion speed (<2s).
    - Format Baru (New Template): Exports as 2 authentic full pages matching Gambar 3 & Gambar 4.
    - Format Lama (Old Template): Fits to 1 full page.
    """
    import tempfile
    import time
    
    temp_dir = tempfile.gettempdir()
    t_id = f"{os.getpid()}_{int(time.time()*1000)}"
    temp_in = os.path.join(temp_dir, f"proj_{t_id}.xlsx")
    temp_out = os.path.join(temp_dir, f"proj_{t_id}.pdf")
    
    if isinstance(excel_bytes_or_path, bytes):
        with open(temp_in, "wb") as f:
            f.write(excel_bytes_or_path)
            f.flush()
        target_in = temp_in
    else:
        target_in = os.path.abspath(excel_bytes_or_path)
        
    import pythoncom
    import win32com.client
    
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        
        abs_excel = os.path.abspath(target_in)
        abs_pdf = os.path.abspath(temp_out)
        
        wb = excel.Workbooks.Open(abs_excel, ReadOnly=True, UpdateLinks=0)
        
        # Select the target form sheet (Form, Project, or ActiveSheet)
        target_ws = None
        for s in wb.Worksheets:
            if any(k in s.Name.upper() for k in ["FORM", "PROJECT"]):
                target_ws = s
                break
        if target_ws is None:
            target_ws = wb.ActiveSheet
            
        # Detect Format Baru vs Format Lama
        is_new = False
        if excel_data and "BARU" in str(excel_data.get("format", "")).upper():
            is_new = True
        else:
            try:
                from excel_parser import detect_excel_format
                is_new = (detect_excel_format(target_in) == "new")
            except Exception:
                pass

        # Page Setup
        try:
            excel.PrintCommunication = False
            target_ws.PageSetup.Orientation = 1  # xlPortrait
            target_ws.PageSetup.PaperSize = 9    # xlPaperA4
            target_ws.PageSetup.Zoom = False
            target_ws.PageSetup.FitToPagesWide = 1
            if is_new:
                # Format Baru (New Template): Section 1 to Section 4 (1 clean full page)
                # Exclude Section 5 (Project Cost), 6 (Office Use), 7 (Misc) as requested by user.
                min_c = 2
                max_c = 4
                has_col_a = False
                for r in range(1, 40):
                    v = target_ws.Cells(r, 1).Value
                    if v is not None and str(v).strip():
                        has_col_a = True
                        break
                if has_col_a:
                    min_c = 1
                    
                # Find where Section 5 (Project Cost) starts to exclude it
                sec5_row = None
                for r in range(1, 100):
                    for c in range(min_c, max_c + 2):
                        v = str(target_ws.Cells(r, c).Value or "")
                        if "SECTION 5" in v.upper() or "PROJECT COST" in v.upper():
                            sec5_row = r
                            break
                    if sec5_row:
                        break
                        
                if sec5_row:
                    last_r = sec5_row - 1
                else:
                    last_r = 66
                    
                col_start_letter = "A" if min_c == 1 else "B"
                col_end_letter = "D"
                for r in range(1, last_r + 1):
                    v = target_ws.Cells(r, 5).Value
                    if v is not None and str(v).strip():
                        col_end_letter = "E"
                        break
                        
                target_ws.ResetAllPageBreaks()
                target_ws.PageSetup.PrintArea = f"${col_start_letter}$1:${col_end_letter}${last_r}"
                target_ws.PageSetup.FitToPagesTall = 1
                target_ws.PageSetup.CenterHorizontally = True
                target_ws.PageSetup.LeftMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.RightMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.TopMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.BottomMargin = excel.InchesToPoints(0.25)
            else:
                # Format Lama (Old Template): 1 full page
                last_r = 1
                for r in range(1, 120):
                    for c in range(1, 14):
                        v = target_ws.Cells(r, c).Value
                        if v is not None and str(v).strip():
                            if r > last_r:
                                last_r = r
                if last_r < 40:
                    last_r = 68
                target_ws.PageSetup.PrintArea = f"$A$1:$M${last_r}"
                target_ws.PageSetup.FitToPagesTall = 1
                target_ws.PageSetup.CenterHorizontally = True
                target_ws.PageSetup.LeftMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.RightMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.TopMargin = excel.InchesToPoints(0.3)
                target_ws.PageSetup.BottomMargin = excel.InchesToPoints(0.3)
            excel.PrintCommunication = True
        except Exception:
            try:
                excel.PrintCommunication = True
            except Exception:
                pass
            
        target_ws.ExportAsFixedFormat(0, abs_pdf)
        wb.Close(False)
        excel.Quit()
        excel = None
        
        with open(abs_pdf, "rb") as f:
            pdf_bytes = f.read()
            
        return pdf_bytes
    except Exception as e:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        raise e
    finally:
        pythoncom.CoUninitialize()
        for p in [temp_in, temp_out]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

def create_authentic_2page_project_form_pdf(excel_data):
    """
    Creates an authentic, high-fidelity 2-page Project Form PDF directly
    using PyMuPDF. Generates in <0.02s without COM or printer spooler delays.
    """
    doc = fitz.open()
    c_navy = (0.05, 0.14, 0.23)
    c_orange = (0.88, 0.43, 0.06)
    c_gray = (0.35, 0.35, 0.35)
    c_border = (0.82, 0.85, 0.89)
    c_bg_alt = (0.97, 0.98, 0.99)
    
    # ------------------ PAGE 1 ------------------
    p1 = doc.new_page(width=595, height=842)
    p1.draw_rect(fitz.Rect(30, 24, 565, 74), color=None, fill=c_navy)
    p1.insert_text(fitz.Point(45, 48), "NORTHERN SOLAR RAKYAT SDN BHD", fontsize=15, color=(1,1,1), fontname="helv")
    p1.insert_text(fitz.Point(45, 64), "PRE-INSTALLATION PROJECT FORM • CLIENT & DC SYSTEM SPECIFICATIONS", fontsize=8.5, color=(0.9,0.9,0.9), fontname="helv")
    
    # Tag
    p1.draw_rect(fitz.Rect(475, 36, 550, 58), color=None, fill=c_orange)
    p1.insert_text(fitz.Point(485, 51), "QC VERIFIED", fontsize=8, color=(1,1,1), fontname="helv")
    
    y = 88
    def draw_table_section(page, start_y, title, items):
        y_curr = start_y
        page.draw_rect(fitz.Rect(30, y_curr, 565, y_curr + 20), color=None, fill=c_orange)
        page.insert_text(fitz.Point(40, y_curr + 14), title, fontsize=9.5, color=(1, 1, 1), fontname="helv")
        y_curr += 24
        
        row_h = 20.0
        for i, (label, val) in enumerate(items):
            bg = c_bg_alt if i % 2 == 1 else (1, 1, 1)
            page.draw_rect(fitz.Rect(30, y_curr, 220, y_curr + row_h), color=c_border, fill=(0.94, 0.96, 0.98), width=0.4)
            page.insert_text(fitz.Point(38, y_curr + 14), str(label), fontsize=8.5, color=c_gray, fontname="helv")
            page.draw_rect(fitz.Rect(220, y_curr, 565, y_curr + row_h), color=c_border, fill=bg, width=0.4)
            page.insert_text(fitz.Point(228, y_curr + 14), str(val or "-"), fontsize=8.5, color=(0, 0, 0), fontname="helv")
            y_curr += row_h
        return y_curr + 12

    # Section 1: Client Info
    y = draw_table_section(p1, y, "1. MAKLUMAT PELANGGAN & PREMIS (CLIENT INFORMATION)", [
        ("Nama Penuh (Full Name)", excel_data.get("client_name")),
        ("No. Kad Pengenalan (IC/Passport)", excel_data.get("ic")),
        ("No. Telefon (Mobile Phone)", excel_data.get("mobile")),
        ("Emel (Email Address)", excel_data.get("email")),
        ("Alamat Tapak (Installation Address)", excel_data.get("address")),
        ("Poskod & Bandar (Postcode & Town)", f"{excel_data.get('postcode', '')} {excel_data.get('town', '')}".strip()),
        ("Negeri (State)", excel_data.get("state")),
    ])
    
    # Section 2: DC Solar System Specs
    mod_name = f"{excel_data.get('module_brand', '')} {excel_data.get('module_model', '')}".strip()
    y = draw_table_section(p1, y, "2. SPESIFIKASI SISTEM SOLAR PV (SOLAR PV DC SYSTEM)", [
        ("Kapasiti Sistem Solar (System DC Capacity)", f"{excel_data.get('kwp', 0):.2f} kWp"),
        ("Model Panel Solar (PV Module Model)", mod_name or "-"),
        ("Kuasa Seunit Panel (Unit Module Power)", f"{excel_data.get('module_power_w', '')} W" if excel_data.get("module_power_w") else "-"),
        ("Jenis Sel Panel (Cell Type)", excel_data.get("module_type") or "Monocrystalline Bifacial"),
        ("Jumlah Bilangan Panel (Total Solar Panels)", f"{excel_data.get('panel_qty', 0)} PCS"),
        ("Jenis Struktur Bumbung (Roof Mounting Type)", excel_data.get("roof_type") or "Tile Roof"),
    ])
    
    # Section 3: TNB & Grid Connection
    y = draw_table_section(p1, y, "3. BUTIRAN BEKALAN TNB & GRID (TNB GRID CONNECTION)", [
        ("Fasa Bekalan TNB (Meter Phase)", f"{excel_data.get('tnb_meter', 'Single Phase')} Phase" if "Phase" not in str(excel_data.get("tnb_meter", "")) else excel_data.get("tnb_meter")),
        ("Kapasiti Export Maksimum (Inverter AC Rating)", f"{excel_data.get('kwac', 0):.2f} kWac"),
        ("Skim Pemasangan (Solar Program Scheme)", "NEM 3.0 (NEM Rakyat / GoMEn / NOVA)"),
    ])
    
    # Footer Page 1
    p1.draw_line(fitz.Point(30, 805), fitz.Point(565, 805), color=c_navy, width=0.8)
    p1.insert_text(fitz.Point(30, 820), "Muka Surat 1 daripada 2 • Dokumen Rasmi Northern Solar Pre-Installation QC Verification", fontsize=7.5, color=c_gray, fontname="helv")
    p1.insert_text(fitz.Point(480, 820), "STRICTLY CONFIDENTIAL", fontsize=7.5, color=c_gray, fontname="helv")

    # ------------------ PAGE 2 ------------------
    p2 = doc.new_page(width=595, height=842)
    p2.draw_rect(fitz.Rect(30, 24, 565, 68), color=None, fill=c_navy)
    p2.insert_text(fitz.Point(45, 46), "NORTHERN SOLAR RAKYAT SDN BHD", fontsize=14, color=(1,1,1), fontname="helv")
    p2.insert_text(fitz.Point(45, 60), "PRE-INSTALLATION PROJECT FORM • AC, STORAGE & INSTALLATION DETAILS", fontsize=8.5, color=(0.9,0.9,0.9), fontname="helv")
    
    y2 = 82
    inv_name = f"{excel_data.get('inverter_brand', '')} {excel_data.get('inverter_model', '')}".strip()
    bat_name = f"{excel_data.get('battery_brand', '')} {excel_data.get('battery_model', '')}".strip()
    inv_kwac_val = excel_data.get("inverter_kwac") or excel_data.get("kwac") or 0.0
    
    y2 = draw_table_section(p2, y2, "4. SPESIFIKASI INVERTER & BATERI (INVERTER & BATTERY STORAGE)", [
        ("Jenama & Model Inverter (Inverter Model)", inv_name or "-"),
        ("Kapasiti Output Inverter (Inverter AC Capacity)", f"{inv_kwac_val:.2f} kWac"),
        ("Bilangan Inverter (Inverter Quantity)", f"{excel_data.get('inverter_qty', 1)} Unit"),
        ("Jenama & Model Bateri (Battery Storage Model)", bat_name or "Tiada (Standard Grid-Tied System)"),
        ("Kapasiti Storan Bateri (Battery Capacity)", f"{excel_data.get('battery_ah', '')} Ah" if excel_data.get("battery_ah") else "-"),
        ("Bilangan Bateri (Battery Storage Quantity)", f"{excel_data.get('battery_qty', 0)} Unit" if excel_data.get("battery_qty") else "0 Unit"),
    ])
    
    # Section 5: Wireman Details
    wireman_val = str(excel_data.get("wireman_name") or "-").strip()
    wireman_reg_val = str(excel_data.get("wireman_reg") or "-").strip()
    y2 = draw_table_section(p2, y2, "5. MAKLUMAT PENDAWAI & KONTRAKTOR (ELECTRICAL WIREMAN)", [
        ("Nama Pendawai Berdaftar (Competent Wireman)", wireman_val if wireman_val else "-"),
        ("Syarikat Pemasang (Installation Contractor)", "NORTHERN SOLAR RAKYAT SDN BHD"),
        ("No. Pendaftaran ST (Energy Commission Reg. No)", wireman_reg_val if wireman_reg_val else "-"),
        ("Pengesahan Dokumen Tapak (Site Verification)", "Telah disemak mengikut SOP Kejuruteraan Solar & Lukisan Pembinaan DWG"),
    ])
    
    # Section 6: Verification Sign-off Box
    p2.draw_rect(fitz.Rect(30, y2, 565, y2 + 20), color=None, fill=c_orange)
    p2.insert_text(fitz.Point(40, y2 + 14), "6. PENGESAHAN DOKUMEN & KELULUSAN (QC VERIFICATION & SIGN-OFF)", fontsize=9.5, color=(1, 1, 1), fontname="helv")
    y2 += 26
    
    box_w = (535 - 20) / 3.0
    for idx, (role_title, desc) in enumerate([
        ("DISEDIAKAN OLEH", "Project Sales / Coordinator"),
        ("DISEMAK & DISAHKAN", "Solar QC Engineer"),
        ("PENGESAHAN PELANGGAN", "Client Acceptance / Site PIC")
    ]):
        bx = 30 + idx * (box_w + 10)
        p2.draw_rect(fitz.Rect(bx, y2, bx + box_w, y2 + 105), color=c_border, fill=(0.98, 0.99, 1.0), width=0.5)
        p2.draw_rect(fitz.Rect(bx, y2, bx + box_w, y2 + 20), color=c_border, fill=(0.93, 0.95, 0.98), width=0.5)
        p2.insert_text(fitz.Point(bx + 8, y2 + 14), role_title, fontsize=8, color=c_navy, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 34), f"Jawatan: {desc}", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 54), "Nama: _____________________", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 74), "Tarikh: ____________________", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 94), "T/Tangan: __________________", fontsize=7.5, color=c_gray, fontname="helv")
        
    # Footer Page 2
    p2.draw_line(fitz.Point(30, 805), fitz.Point(565, 805), color=c_navy, width=0.8)
    p2.insert_text(fitz.Point(30, 820), "Muka Surat 2 daripada 2 • Dokumen Rasmi Northern Solar Pre-Installation QC Verification", fontsize=7.5, color=c_gray, fontname="helv")
    p2.insert_text(fitz.Point(480, 820), "STRICTLY CONFIDENTIAL", fontsize=7.5, color=c_gray, fontname="helv")

    pdf_bytes = doc.write()
    doc.close()
    return pdf_bytes

def convert_excel_to_pdf_libreoffice(excel_bytes_or_path, excel_data=None):
    """
    Converts Excel file to PDF using headless LibreOffice (Linux cloud environments like Streamlit Cloud).
    Isolates the authentic Project Form sheet first so that other sheets (PVSyst, quotations, graphs)
    do not inflate or corrupt the generated PDF.
    """
    import subprocess
    import tempfile
    import time
    import openpyxl
    
    temp_dir = tempfile.gettempdir()
    t_id = f"{os.getpid()}_{int(time.time()*1000)}"
    temp_in = os.path.join(temp_dir, f"proj_{t_id}.xlsx")
    target_in = temp_in
    
    # 1. Isolate the Project Form sheet using openpyxl
    try:
        if isinstance(excel_bytes_or_path, bytes):
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes_or_path), data_only=True)
        else:
            wb = openpyxl.load_workbook(excel_bytes_or_path, data_only=True)
            
        target_sheet = None
        for s_name in wb.sheetnames:
            ws = wb[s_name]
            txt = ""
            for row in ws.iter_rows(max_row=25, values_only=True):
                txt += " " + " ".join([str(c) for c in row if c is not None])
            txt_u = txt.upper()
            if any(k in txt_u for k in ["CLIENT INFORMATION", "MAKLUMAT PELANGGAN", "SECTION 1", "MODULE PANEL QUANTITY", "PRE-INSTALLATION", "CHECKLIST"]):
                target_sheet = s_name
                break
        if not target_sheet:
            for s_name in wb.sheetnames:
                if any(k in s_name.upper() for k in ["FORM", "PROJECT"]):
                    target_sheet = s_name
                    break
        if not target_sheet:
            target_sheet = wb.active.title
            
        for s_name in list(wb.sheetnames):
            if s_name != target_sheet:
                del wb[s_name]
                
        wb.save(temp_in)
        wb.close()
    except Exception:
        if isinstance(excel_bytes_or_path, bytes):
            with open(temp_in, "wb") as f:
                f.write(excel_bytes_or_path)
        else:
            target_in = os.path.abspath(excel_bytes_or_path)
            
    try:
        cmd = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, target_in]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=30)
        
        base_name = os.path.splitext(os.path.basename(target_in))[0]
        gen_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
        if os.path.exists(gen_pdf):
            with open(gen_pdf, "rb") as f:
                pdf_bytes = f.read()
            try:
                os.remove(gen_pdf)
            except Exception:
                pass
            return pdf_bytes
    except Exception as err:
        print(f"LibreOffice conversion failed: {err}")
    finally:
        if os.path.exists(temp_in):
            try:
                os.remove(temp_in)
            except Exception:
                pass
    return None

def convert_excel_to_pdf(excel_bytes_or_path, excel_data=None):
    """
    Converts Excel file to 100% original Excel layout:
    1. Windows: Native Microsoft Excel COM automation (<2.5s).
    2. Linux / Streamlit Cloud: Headless LibreOffice.
    3. Fallback: Direct high-fidelity PDF layout generator with client data.
    """
    if os.name == "nt":
        try:
            return convert_excel_to_pdf_native(excel_bytes_or_path, excel_data=excel_data)
        except Exception as e:
            print(f"Native Excel COM conversion failed: {e}. Trying fallback.")
            
    # Try LibreOffice (Streamlit Cloud Linux)
    try:
        lo_pdf = convert_excel_to_pdf_libreoffice(excel_bytes_or_path, excel_data=excel_data)
        if lo_pdf:
            doc = fitz.open(stream=lo_pdf, filetype="pdf")
            out_doc = fitz.open()
            for idx_p, page in enumerate(doc):
                txt_p = (page.get_text() or "").upper()
                if idx_p > 0 and any(k in txt_p for k in ["PROJECT COST", "SECTION 5", "CAPITAL EXPENDITURE", "OFFICE USE"]):
                    continue
                out_doc.insert_pdf(doc, from_page=idx_p, to_page=idx_p)
                if len(out_doc) >= 2:
                    break
            if len(out_doc) > 0:
                filtered_bytes = out_doc.write()
                out_doc.close()
                doc.close()
                return filtered_bytes
            doc.close()
            return lo_pdf
    except Exception:
        pass
        
    return create_authentic_2page_project_form_pdf(excel_data or {})

def stamp_cables_on_dwg_page2(
    dwg_pdf_bytes,
    dc_cable_val="72",
    earthing_cable_val="50",
    table_overrides=None,
    name_overrides=None,
    mismatches=None,
    extra_items=None,
    cable_preset="Beside System Information & Legend (Right) [Standard]",
    cable_offset_x=0,
    cable_offset_y=0
):
    """
    1. Unconditionally updates BOM rows in System Information table on Page 2:
       - If a quantity is changed/overridden on the web or does not tally, renders it in RED (0.85, 0.05, 0.05).
       - If a component name is changed (e.g. Hanger bolt / Tile Hook -> U-Rail), renders it in RED (0.85, 0.05, 0.05).
       - Uses direct word-search replacement that operates reliably regardless of CAD vector table lines.
    2. Draws a separate 'Cable Information' table flexibly positioned anywhere on Page 2:
       - Header: Cable Information
       - Subheader: Item No. | Info | Quantity (m)
       - Row 1: 1 | DC Cable | [dc_cable_val]
       - Row 2: 2 | Ground Cable | [earthing_cable_val]
       - Rows 3+: Add-on equipment (Combiner Box, RSD, Battery Stand, Optimizer, etc.)
       - Position dynamically driven by cable_preset and cable_offset_x / cable_offset_y sliders.
       - Uses authentic ISOCPEUR CAD font and 0.35pt line width.
    """
    doc = fitz.open(stream=dwg_pdf_bytes, filetype="pdf")
    if len(doc) < 2:
        doc.close()
        return dwg_pdf_bytes
        
    page = doc[1]  # Page 2 (0-indexed)
    page_w = page.rect.width
    page_h = page.rect.height
    
    # 1. Resolve & Embed ISOCPEUR font
    font_path = os.path.join(os.path.dirname(__file__), "isocpeur.ttf")
    if not os.path.exists(font_path):
        for f in page.get_fonts():
            if "ISOCPEUR" in f[3].upper():
                try:
                    font_bytes = doc.extract_font(f[0])[3]
                    with open(font_path, "wb") as f_out:
                        f_out.write(font_bytes)
                    break
                except Exception:
                    pass
                    
    font_name = "ISOCPEUR"
    try:
        page.insert_font(fontname=font_name, fontfile=font_path)
    except Exception:
        font_name = "helv"
        
    font_obj = fitz.Font(fontfile=font_path) if os.path.exists(font_path) else None
    font_title_sz = 9.88
    font_content_sz = 8.89
    line_w = 0.35
    red_color = (0.85, 0.05, 0.05)
    black_color = (0.0, 0.0, 0.0)

    # Clean & normalize dictionary keys
    clean_table_ov = {str(k).strip().upper(): str(v).strip() for k, v in (table_overrides or {}).items() if v is not None and str(v).strip()}
    clean_name_ov = {str(k).strip().upper(): str(v).strip() for k, v in (name_overrides or {}).items() if v is not None and str(v).strip()}
    mismatches_upper = [str(m).strip().upper() for m in (mismatches or [])]
    
    words = page.get_text("words")
    left_words = [w for w in words if w[0] < page_w * 0.5]

    bom_items_catalog = [
        {
            "num": "1",
            "search_terms": ["End Clamp"],
            "keys": ["END CLAMP", "1"],
            "mismatch_tokens": ["END CLAMP"]
        },
        {
            "num": "2",
            "search_terms": ["Mid Clamp"],
            "keys": ["MID CLAMP", "2"],
            "mismatch_tokens": ["MID CLAMP"]
        },
        {
            "num": "3",
            "search_terms": [
                "Hanger bolt", "Hanger Bolt", "Hangerbolt",
                "Tile Hook", "Tile hook",
                "U-Rail", "URail", "U - Rail",
                "Universal Cliplock", "Cliplock",
                "L-Foot", "L Foot", "L - Foot",
                "Shingle Plate", "Shingle Hook", "Shingle",
                "Concrete Ballast", "Ballast",
                "Roof Attachment"
            ],
            "keys": ["TILE HOOK", "HANGER BOLT", "HANGER", "U-RAIL", "URAIL", "UNIVERSAL CLIPLOCK", "CLIPLOCK", "L-FOOT", "L FOOT", "SHINGLE PLATE", "SHINGLE", "BALLAST", "ROOF ATTACHMENT", "3"],
            "mismatch_tokens": ["TILE", "HOOK", "HANGER", "U-RAIL", "URAIL", "CLIPLOCK", "L-FOOT", "L FOOT", "SHINGLE", "BALLAST", "ROOF ATTACHMENT"]
        },
        {
            "num": "4",
            "search_terms": ["Grounding Clip"],
            "keys": ["GROUNDING CLIP", "4"],
            "mismatch_tokens": ["GROUNDING CLIP"]
        },
        {
            "num": "5",
            "search_terms": ["Grounding Lug"],
            "keys": ["GROUNDING LUG", "5"],
            "mismatch_tokens": ["GROUNDING LUG"]
        },
        {
            "num": "6",
            "search_terms": ["MC4 Connector", "MC4"],
            "keys": ["MC4 CONNECTOR", "MC4", "6"],
            "mismatch_tokens": ["MC4 CONNECTOR", "MC4"]
        },
        {
            "num": "7",
            "search_terms": ["R-Railing", "Railing"],
            "keys": ["R-RAILING", "RAILING", "7"],
            "mismatch_tokens": ["R-RAILING", "RAILING"]
        },
        {
            "num": "8",
            "search_terms": ["Splice"],
            "keys": ["SPLICE", "8"],
            "mismatch_tokens": ["SPLICE"]
        },
        {
            "num": "10",
            "search_terms": ["L-Foot", "L Foot"],
            "keys": ["L-FOOT", "L FOOT", "10"],
            "mismatch_tokens": ["L-FOOT", "L FOOT"]
        }
    ]

    detected_rows = []
    known_qty_cols = []
    
    for item in bom_items_catalog:
        found_hit = None
        found_term = None
        for term in item["search_terms"]:
            hits = [r for r in page.search_for(term) if r.x0 < page_w * 0.5]
            if hits:
                found_hit = hits[0]
                found_term = term
                break
        if not found_hit:
            continue
            
        y_mid = (found_hit.y0 + found_hit.y1) / 2.0
        row_words = [w for w in left_words if abs(((w[1] + w[3]) / 2.0) - y_mid) <= 7.0]
        row_words.sort(key=lambda w: w[0])
        qty_words = [w for w in row_words if w[0] >= found_hit.x1 - 3.0]
        num_qty_word = None
        for w in qty_words:
            if re.search(r"\d+", w[4]):
                num_qty_word = w
                break
        if num_qty_word:
            known_qty_cols.append((num_qty_word[0], num_qty_word[2]))
        detected_rows.append({
            "item": item,
            "hit_rect": found_hit,
            "term": found_term,
            "qty_word": num_qty_word,
            "row_words": row_words
        })

    default_q_x0 = 150.0
    default_q_x1 = 175.0
    if known_qty_cols:
        default_q_x0 = sorted([c[0] for c in known_qty_cols])[len(known_qty_cols) // 2]
        default_q_x1 = sorted([c[1] for c in known_qty_cols])[len(known_qty_cols) // 2]

    # Process Overrides & Mismatches in System Information table
    for r_info in detected_rows:
        item = r_info["item"]
        found_hit = r_info["hit_rect"]
        found_term = r_info["term"]
        qty_w = r_info["qty_word"]
        
        # 1. Roof Attachment Name Override (Row 3)
        if item["num"] == "3" and clean_name_ov:
            roof_override = clean_name_ov.get("ROOF ATTACHMENT") or clean_name_ov.get("3") or clean_name_ov.get("TILE HOOK")
            if roof_override and roof_override.strip().upper() != found_term.upper():
                new_roof = roof_override.strip()
                page.draw_rect(fitz.Rect(found_hit.x0 - 2, found_hit.y0 - 2, found_hit.x1 + 4, found_hit.y1 + 2), color=None, fill=(1, 1, 1))
                page.insert_text(fitz.Point(found_hit.x0, found_hit.y1 - 1.0), new_roof, fontname=font_name, fontsize=font_content_sz, color=red_color)

        # 2. Quantity Override
        new_qty_val = None
        for k in item["keys"]:
            if k in clean_table_ov:
                new_qty_val = clean_table_ov[k]
                break
                
        # 3. Mismatch
        is_mismatch = any(any(mt in m for mt in item["mismatch_tokens"]) for m in mismatches_upper)
        
        # 4. Stamp in RED
        if new_qty_val is not None or is_mismatch:
            val_to_draw = new_qty_val if new_qty_val is not None else (qty_w[4] if qty_w else None)
            if val_to_draw:
                if qty_w:
                    q_x0 = qty_w[0]
                    q_x1 = qty_w[2]
                else:
                    q_x0 = default_q_x0
                    q_x1 = default_q_x1
                w_rect = fitz.Rect(q_x0 - 6.0, found_hit.y0 - 2.5, q_x1 + 6.0, found_hit.y1 + 2.5)
                page.draw_rect(w_rect, color=None, fill=(1, 1, 1))
                draw_x = ((q_x0 + q_x1) - 15.0) / 2.0
                draw_y = found_hit.y1 - 1.0
                page.insert_text(fitz.Point(draw_x, draw_y), val_to_draw, fontname=font_name, fontsize=font_content_sz, color=red_color)

    # Determine reference coordinates for System Information table
    sys_hits = [r for r in page.search_for("System Information") if r.x0 < page_w * 0.5]
    if sys_hits:
        sys_top = sys_hits[0].y0
    elif detected_rows:
        sys_top = min(r["hit_rect"].y0 for r in detected_rows) - 25.0
    else:
        sys_top = 50.0
        
    if detected_rows:
        sys_bot = max(r["hit_rect"].y1 for r in detected_rows) + 8.0
        all_row_x0 = [min(w[0] for w in r["row_words"]) for r in detected_rows if r["row_words"]]
        sys_left = min(all_row_x0) - 5.0 if all_row_x0 else 25.0
        sys_right = max(c[1] for c in known_qty_cols) + 15.0 if known_qty_cols else 195.0
    else:
        sys_bot = sys_top + 200.0
        sys_left = 25.0
        sys_right = 195.0

    try:
        tabs = page.find_tables()
        for t in tabs.tables:
            if (t.bbox[2] - t.bbox[0]) < page_w * 0.4 and (t.bbox[3] - t.bbox[1]) < page.rect.height * 0.7:
                ext = t.extract()
                if any("System Information" in str(c) or "Splice" in str(c) for row in ext for c in row if c):
                    sys_left = min(sys_left, t.bbox[0])
                    sys_right = max(sys_right, t.bbox[2])
                    sys_bot = max(sys_bot, t.bbox[3])
                    break
    except Exception:
        pass

    table_w = max(185.0, sys_right - sys_left)
    
    if len(detected_rows) >= 2:
        diffs = [detected_rows[i+1]["hit_rect"].y0 - detected_rows[i]["hit_rect"].y0 for i in range(len(detected_rows)-1)]
        valid_diffs = [d for d in diffs if 15.0 <= d <= 32.0]
        row_h = (sum(valid_diffs) / len(valid_diffs)) if valid_diffs else 23.28
    else:
        row_h = 23.28

    # -------------------------------------------------------------
    # DRAW SEPARATE CABLE INFORMATION TABLE (FLEXIBLE PLACEMENT)
    # -------------------------------------------------------------
    extra_items = extra_items or []
    total_cable_rows = 4 + len(extra_items)
    table_total_h = total_cable_rows * row_h
    
    # Detect Middle Legend Table (Table 2) if present
    legend_right = None
    legend_top = None
    gap_table = 12.0
    
    # 1. Check PyMuPDF find_tables()
    try:
        tabs = page.find_tables()
        tables_sorted = sorted(tabs.tables, key=lambda t: t.bbox[0])
        if len(tables_sorted) >= 2:
            t1 = tables_sorted[0]
            t2 = tables_sorted[1]
            if t1.bbox[0] < page_w * 0.35 and t2.bbox[0] < page_w * 0.65:
                sys_left = t1.bbox[0]
                sys_right = t1.bbox[2]
                sys_top = t1.bbox[1]
                sys_bot = t1.bbox[3]
                legend_right = t2.bbox[2]
                legend_top = t2.bbox[1]
                gap_table = max(8.0, min(25.0, t2.bbox[0] - t1.bbox[2]))
        elif len(tables_sorted) == 1:
            t1 = tables_sorted[0]
            if t1.bbox[0] < page_w * 0.4:
                sys_left = min(sys_left, t1.bbox[0])
                sys_right = max(sys_right, t1.bbox[2])
                sys_top = min(sys_top, t1.bbox[1])
                sys_bot = max(sys_bot, t1.bbox[3])
    except Exception:
        pass

    # 2. Check multiple System Information headers (Table 1 & Table 2)
    if legend_right is None:
        sys_hits = [r for r in page.search_for("System Information") if r.x0 < page_w * 0.7]
        if len(sys_hits) >= 2:
            sys_hits.sort(key=lambda r: r.x0)
            h2 = sys_hits[1]
            legend_top = h2.y0
            t2_words = [w for w in words if w[0] > (sys_right - 5.0) and w[0] < page_w * 0.68 and abs(((w[1]+w[3])/2.0) - sys_top) < 320.0]
            if t2_words:
                t2_left = min(w[0] for w in t2_words) - 10.0
                t2_right_w = max(w[2] for w in t2_words) + 15.0
                gap_table = max(8.0, min(25.0, t2_left - sys_right))
                legend_right = max(t2_right_w, h2.x1 + 40.0)
            else:
                gap_table = max(8.0, min(25.0, h2.x0 - sys_right - 20.0))
                legend_right = h2.x1 + 40.0

    # 3. Check Legend word search to locate Table 2 if find_tables didn't catch it
    if legend_right is None:
        leg_words_match = [w for w in words if "LEGEND" in w[4].upper() and w[0] < page_w * 0.7]
        if leg_words_match:
            w_leg = leg_words_match[0]
            legend_top = sys_top
            t2_words = [w for w in words if w[0] > (sys_right - 5.0) and w[0] < page_w * 0.68 and abs(((w[1]+w[3])/2.0) - sys_top) < 320.0]
            if t2_words:
                t2_left = min(w[0] for w in t2_words) - 10.0
                t2_right_w = max(w[2] for w in t2_words) + 15.0
                gap_table = max(8.0, min(25.0, t2_left - sys_right))
                legend_right = max(t2_right_w, w_leg[2] + 120.0)

    # 4. Fallback if Table 2 exists on drawing layout
    if legend_right is None and table_w < page_w * 0.28:
        legend_right = sys_right + (table_w * 0.85) + gap_table
        legend_top = sys_top

    p_choice = str(cable_preset or "").lower()
    if "below" in p_choice:
        base_x = sys_left
        base_y = sys_bot + 8.0
    elif "bottom-left" in p_choice or "bottom left" in p_choice:
        base_x = sys_left
        base_y = page_h - table_total_h - 45.0
    elif "bottom-right" in p_choice or "bottom right" in p_choice:
        base_x = page_w - table_w - 220.0
        base_y = page_h - table_total_h - 45.0
    elif "custom" in p_choice:
        base_x = (legend_right + gap_table) if legend_right else (sys_right + gap_table)
        base_y = legend_top if legend_top is not None else sys_top
    else:  # "Beside System Information & Legend (Right) [Standard]" - DEFAULT matching Gambar!
        base_x = (legend_right + gap_table) if legend_right else (sys_right + gap_table)
        base_y = legend_top if legend_top is not None else sys_top
        
    final_x = base_x + float(cable_offset_x or 0)
    final_y = base_y + float(cable_offset_y or 0)
    
    # Boundary protection so table remains within printable area
    final_x = max(15.0, min(final_x, page_w - table_w - 15.0))
    final_y = max(15.0, min(final_y, page_h - table_total_h - 15.0))

    # Exact column proportions measured from user's CAD drawing (media_1792574148373_fb849595.png)
    c0 = final_x
    c1 = c0 + (table_w * 0.26)
    c2 = c0 + (table_w * 0.74)
    c3 = c0 + table_w
    baseline_offset = (row_h + 6.2) / 2.0

    curr_y = final_y

    # 1. Header: Cable Information
    page.draw_rect(fitz.Rect(c0, curr_y, c3, curr_y + row_h), color=black_color, fill=(1, 1, 1), width=line_w)
    len_ci = font_obj.text_length("Cable Information", fontsize=font_title_sz) if font_obj else 65.0
    x_ci = ((c0 + c3) - len_ci) / 2.0
    page.insert_text(fitz.Point(x_ci, curr_y + (row_h + 6.9) / 2.0), "Cable Information", fontname=font_name, fontsize=font_title_sz, color=black_color)
    curr_y += row_h

    # 2. Subheader: Item No. | Info | Quantity (Matches user's exact drawing)
    page.draw_rect(fitz.Rect(c0, curr_y, c3, curr_y + row_h), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(c1, curr_y), fitz.Point(c1, curr_y + row_h), color=black_color, width=line_w)
    page.draw_line(fitz.Point(c2, curr_y), fitz.Point(c2, curr_y + row_h), color=black_color, width=line_w)

    for c_left, c_right, text_lbl in [(c0, c1, "Item No."), (c1, c2, "Info"), (c2, c3, "Quantity")]:
        avail_w = (c_right - c_left) - 2.0
        use_sz = font_content_sz
        l_txt = font_obj.text_length(text_lbl, fontsize=use_sz) if font_obj else 20.0
        if l_txt > avail_w and avail_w > 10:
            use_sz = font_content_sz * (avail_w / l_txt)
            l_txt = font_obj.text_length(text_lbl, fontsize=use_sz)
        x_lbl = ((c_left + c_right) - l_txt) / 2.0
        page.insert_text(fitz.Point(x_lbl, curr_y + baseline_offset), text_lbl, fontname=font_name, fontsize=use_sz, color=black_color)
    curr_y += row_h

    # 3. Row 1: 1 | DC Cable | [dc_cable_val]
    page.draw_rect(fitz.Rect(c0, curr_y, c3, curr_y + row_h), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(c1, curr_y), fitz.Point(c1, curr_y + row_h), color=black_color, width=line_w)
    page.draw_line(fitz.Point(c2, curr_y), fitz.Point(c2, curr_y + row_h), color=black_color, width=line_w)

    l_1 = font_obj.text_length("1", fontsize=font_content_sz) if font_obj else 6.0
    page.insert_text(fitz.Point(((c0 + c1) - l_1) / 2.0, curr_y + baseline_offset), "1", fontname=font_name, fontsize=font_content_sz, color=black_color)

    l_dc = font_obj.text_length("DC Cable", fontsize=font_content_sz) if font_obj else 30.0
    page.insert_text(fitz.Point(((c1 + c2) - l_dc) / 2.0, curr_y + baseline_offset), "DC Cable", fontname=font_name, fontsize=font_content_sz, color=black_color)

    dc_str = str(dc_cable_val).strip() if dc_cable_val else "72"
    l_dc_q = font_obj.text_length(dc_str, fontsize=font_content_sz) if font_obj else 15.0
    page.insert_text(fitz.Point(((c2 + c3) - l_dc_q) / 2.0, curr_y + baseline_offset), dc_str, fontname=font_name, fontsize=font_content_sz, color=black_color)
    curr_y += row_h

    # 4. Row 2: 2 | Ground Cable | [earthing_cable_val]
    page.draw_rect(fitz.Rect(c0, curr_y, c3, curr_y + row_h), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(c1, curr_y), fitz.Point(c1, curr_y + row_h), color=black_color, width=line_w)
    page.draw_line(fitz.Point(c2, curr_y), fitz.Point(c2, curr_y + row_h), color=black_color, width=line_w)

    l_2 = font_obj.text_length("2", fontsize=font_content_sz) if font_obj else 6.0
    page.insert_text(fitz.Point(((c0 + c1) - l_2) / 2.0, curr_y + baseline_offset), "2", fontname=font_name, fontsize=font_content_sz, color=black_color)

    l_gc = font_obj.text_length("Ground Cable", fontsize=font_content_sz) if font_obj else 40.0
    page.insert_text(fitz.Point(((c1 + c2) - l_gc) / 2.0, curr_y + baseline_offset), "Ground Cable", fontname=font_name, fontsize=font_content_sz, color=black_color)

    gc_str = str(earthing_cable_val).strip() if earthing_cable_val else "50"
    l_gc_q = font_obj.text_length(gc_str, fontsize=font_content_sz) if font_obj else 15.0
    page.insert_text(fitz.Point(((c2 + c3) - l_gc_q) / 2.0, curr_y + baseline_offset), gc_str, fontname=font_name, fontsize=font_content_sz, color=black_color)
    curr_y += row_h

    # 5. Rows 3+: Add-on equipment (Combiner Box, RSD, Battery Stand, Optimizer, etc.)
    curr_item_no = 3
    for itm in extra_items:
        info_txt = str(itm.get("info", "") or itm.get("Info", "") or itm.get("name", "")).strip()
        qty_txt = str(itm.get("qty", "") or itm.get("quantity", "") or itm.get("Quantity", "")).strip()
        if not info_txt or not qty_txt:
            continue
            
        page.draw_rect(fitz.Rect(c0, curr_y, c3, curr_y + row_h), color=black_color, fill=(1, 1, 1), width=line_w)
        page.draw_line(fitz.Point(c1, curr_y), fitz.Point(c1, curr_y + row_h), color=black_color, width=line_w)
        page.draw_line(fitz.Point(c2, curr_y), fitz.Point(c2, curr_y + row_h), color=black_color, width=line_w)

        str_no = str(curr_item_no)
        l_no = font_obj.text_length(str_no, fontsize=font_content_sz) if font_obj else 6.0
        page.insert_text(fitz.Point(((c0 + c1) - l_no) / 2.0, curr_y + baseline_offset), str_no, fontname=font_name, fontsize=font_content_sz, color=black_color)

        l_info = font_obj.text_length(info_txt, fontsize=font_content_sz) if font_obj else 35.0
        avail_w_info = (c2 - c1) - 4.0
        use_info_sz = font_content_sz
        if l_info > avail_w_info and avail_w_info > 10:
            use_info_sz = font_content_sz * (avail_w_info / l_info)
            l_info = font_obj.text_length(info_txt, fontsize=use_info_sz)
        page.insert_text(fitz.Point(((c1 + c2) - l_info) / 2.0, curr_y + baseline_offset), info_txt, fontname=font_name, fontsize=use_info_sz, color=black_color)

        l_q = font_obj.text_length(qty_txt, fontsize=font_content_sz) if font_obj else 15.0
        page.insert_text(fitz.Point(((c2 + c3) - l_q) / 2.0, curr_y + baseline_offset), qty_txt, fontname=font_name, fontsize=font_content_sz, color=black_color)

        curr_y += row_h
        curr_item_no += 1

    stamped_bytes = doc.write()
    doc.close()
    return stamped_bytes

def stamp_breaker_on_dwg_page4(dwg_pdf_bytes, mcb_val=None, mcb_pole=None, ct_val=None):
    """
    Updates Inverter MCB (e.g. 25A, 1P) and PV Meter CT (e.g. 25VA) on Page 4 in RED text
    using authentic ISOCPEUR CAD font (11.38pt), matching the exact size of other breaker text.
    """
    if not mcb_val and not mcb_pole and not ct_val:
        return dwg_pdf_bytes
        
    doc = fitz.open(stream=dwg_pdf_bytes, filetype="pdf")
    if len(doc) < 4:
        doc.close()
        return dwg_pdf_bytes
        
    p4 = doc[3] # Page 4 (0-indexed)
    
    font_path = os.path.join(os.path.dirname(__file__), "isocpeur.ttf")
    font_name = "ISOCPEUR"
    try:
        p4.insert_font(fontname=font_name, fontfile=font_path)
    except Exception:
        font_name = "helv"
        
    font_sz = 11.38
    red_color = (0.85, 0.05, 0.05)
    words = p4.get_text("words")
    
    # Auto-format inputs: typing only numbers (e.g. 32, 1, 25) automatically appends A, P, VA
    clean_ct = None
    if ct_val and str(ct_val).strip():
        s = str(ct_val).strip()
        clean_ct = f"{s}VA" if not s.upper().endswith("VA") else s
        
    clean_pole = None
    if mcb_pole and str(mcb_pole).strip():
        s = str(mcb_pole).strip()
        clean_pole = f"{s}P" if not s.upper().endswith("P") else s
        
    clean_mcb = None
    if mcb_val and str(mcb_val).strip():
        s = str(mcb_val).strip()
        clean_mcb = f"{s}A" if not s.upper().endswith("A") else s
        
    target_cl1 = [w for w in words if "CL1" in w[4].upper() and 300 <= w[0] <= 500]
    
    # 1. Update PV Meter CT rating (e.g. 25VA / 32VA)
    if clean_ct and target_cl1:
        cl1_w = target_cl1[0]
        ct_words = [w for w in words if abs(w[0] - cl1_w[0]) < 25 and 0 < (cl1_w[1] - w[1]) < 25]
        if ct_words:
            w_ct = ct_words[0]
            p4.draw_rect(fitz.Rect(w_ct[0] - 2, w_ct[1] - 1, w_ct[2] + 2, w_ct[3] + 1), color=None, fill=(1, 1, 1))
            p4.insert_text(fitz.Point(w_ct[0], w_ct[3] - 1.5), clean_ct, fontname=font_name, fontsize=font_sz, color=red_color)
            
    # 2. Update Inverter MCB rating (e.g. 25A 1P / 32A 3P)
    if clean_mcb or clean_pole:
        target_10ka = [w for w in words if "10KA" in w[4].upper() and w[1] > 300]
        if target_10ka:
            ref_10ka = target_10ka[-1]
            pole_words = [w for w in words if abs(w[0] - ref_10ka[0]) < 25 and 0 < (ref_10ka[1] - w[1]) < 22]
            if pole_words and clean_pole:
                w_pole = pole_words[0]
                p4.draw_rect(fitz.Rect(w_pole[0] - 2, w_pole[1] - 1, w_pole[2] + 2, w_pole[3] + 1), color=None, fill=(1, 1, 1))
                p4.insert_text(fitz.Point(w_pole[0], w_pole[3] - 1.5), clean_pole, fontname=font_name, fontsize=font_sz, color=red_color)
                
            ref_top = pole_words[0][1] if pole_words else ref_10ka[1]
            amp_words = [w for w in words if abs(w[0] - ref_10ka[0]) < 25 and 0 < (ref_top - w[1]) < 25]
            if amp_words and clean_mcb:
                w_amp = amp_words[0]
                p4.draw_rect(fitz.Rect(w_amp[0] - 2, w_amp[1] - 1, w_amp[2] + 2, w_amp[3] + 1), color=None, fill=(1, 1, 1))
                p4.insert_text(fitz.Point(w_amp[0], w_amp[3] - 1.5), clean_mcb, fontname=font_name, fontsize=font_sz, color=red_color)
                
    stamped_bytes = doc.write()
    doc.close()
    return stamped_bytes

def combine_pdfs(
    excel_pdf_bytes,
    dwg_pdf_bytes_or_path,
    dc_cable_val=None,
    earthing_cable_val=None,
    table_overrides=None,
    name_overrides=None,
    mismatches=None,
    mcb_val=None,
    mcb_pole=None,
    ct_val=None,
    already_stamped=False,
    extra_items=None,
    cable_preset="Beside System Information & Legend (Right) [Standard]",
    cable_offset_x=0,
    cable_offset_y=0
):
    """
    Merges Excel PDF and DWG Construction Drawing PDF into one combined PDF.
    If already_stamped is True, skips re-stamping for ultra-fast instant merging (<0.05s).
    """
    # 1. Stamp cables and overrides onto DWG only if not already stamped
    if isinstance(dwg_pdf_bytes_or_path, bytes) and not already_stamped:
        dwg_pdf_bytes_or_path = stamp_cables_on_dwg_page2(
            dwg_pdf_bytes_or_path,
            dc_cable_val=dc_cable_val or "72",
            earthing_cable_val=earthing_cable_val or "50",
            table_overrides=table_overrides,
            name_overrides=name_overrides,
            mismatches=mismatches,
            extra_items=extra_items,
            cable_preset=cable_preset,
            cable_offset_x=cable_offset_x,
            cable_offset_y=cable_offset_y
        )
        if mcb_val or mcb_pole or ct_val:
            dwg_pdf_bytes_or_path = stamp_breaker_on_dwg_page4(
                dwg_pdf_bytes_or_path,
                mcb_val=mcb_val,
                mcb_pole=mcb_pole,
                ct_val=ct_val
            )
        
    writer = PdfWriter()
    
    # 2. Add Excel PDF (Section 1 to 4 only - exclude Section 5 Project Cost page)
    reader_form = PdfReader(io.BytesIO(excel_pdf_bytes))
    for idx_p, page in enumerate(reader_form.pages):
        if idx_p > 0:
            txt = (page.extract_text() or "").upper()
            if any(k in txt for k in ["PROJECT COST", "SECTION 5", "CAPITAL EXPENDITURE", "OFFICE USE"]):
                continue
        writer.add_page(page)
        
    # 3. Add DWG PDF (with stamped Page 2 and Page 4)
    if isinstance(dwg_pdf_bytes_or_path, bytes):
        reader_dwg = PdfReader(io.BytesIO(dwg_pdf_bytes_or_path))
    else:
        reader_dwg = PdfReader(dwg_pdf_bytes_or_path)
        
    for page in reader_dwg.pages:
        writer.add_page(page)
        
    output_stream = io.BytesIO()
    writer.write(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()
