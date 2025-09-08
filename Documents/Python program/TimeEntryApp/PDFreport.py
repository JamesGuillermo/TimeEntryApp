from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import sys
import os
from tkinter import Tk, filedialog, messagebox


def build_pdf(pdf_path):
    """Build the PDF report and save to pdf_path."""
    # Define styles
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    heading_style = styles['Heading2']
    normal_style = styles['Normal']

    # Build document elements
    elements = []

    # Title
    elements.append(Paragraph("Step Duration Analysis Report", title_style))
    elements.append(Spacer(1, 20))

    # Per-month analysis text
    report_text = """
    <h2>Monthly Analysis</h2>
    <b>June 2025:</b><br/>
    - Longest delays: Step 3 (74 min) and Step 5 (96 min).<br/>
    - These are the main bottlenecks.<br/><br/>

    <b>July 2025:</b><br/>
    - Improvements in Step 3 (69 min) and Step 5 (92 min).<br/>
    - Other steps slightly slower.<br/><br/>

    <b>August 2025:</b><br/>
    - Major improvements in Step 1, 2, 3, 4, and 6.<br/>
    - Step 5 became much slower (120 min).<br/><br/>
    """
    elements.append(Paragraph(report_text, normal_style))

    # Add summary table
    data_table = [
        ["Month", "Duration 1", "Duration 2", "Duration 3", "Duration 4", "Duration 5", "Duration 6"],
        ["2025-06", 13.62, 25.31, 74.33, 11.92, 96.32, 5.40],
        ["2025-07", 14.75, 27.04, 69.28, 13.34, 92.30, 8.02],
        ["2025-08", 10.17, 22.85, 30.55, 4.00, 120.52, 3.89]
    ]

    table = Table(data_table, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
        ('TEXTCOLOR',(0,0),(-1,0),colors.black),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')
    ]))
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("<b>Average Step Durations (minutes)</b>", heading_style))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Overall conclusion
    conclusion_text = """
    <h2>Conclusion</h2>
    - Across 3 months, most steps show clear improvements.<br/>
    - Step 3 improved the most (from 74 min to 30 min).<br/>
    - Step 5 worsened (from 96 min to 120 min).<br/>
    - Overall: process is getting faster, but Step 5 is now the main bottleneck.<br/>
    """
    elements.append(Paragraph(conclusion_text, normal_style))

    # Ensure target directory exists
    target_dir = os.path.dirname(os.path.abspath(pdf_path))
    if target_dir and not os.path.exists(target_dir):
        os.makedirs(target_dir, exist_ok=True)

    # Build PDF
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    doc.build(elements)

    return pdf_path


if __name__ == "__main__":
    # If a path is provided as a CLI argument, use it (non-interactive)
    if len(sys.argv) > 1:
        out_path = sys.argv[1]
        try:
            result = build_pdf(out_path)
            print(f"PDF saved to: {result}")
        except Exception as e:
            print(f"Failed to build PDF: {e}")
        sys.exit(0)

    # Otherwise, open a Save As dialog for user to choose location
    root = Tk()
    root.withdraw()
    default_name = "Step_Duration_Report.pdf"
    # Prefer current working directory as default
    initial_dir = os.getcwd()
    save_path = filedialog.asksaveasfilename(
        title="Save PDF report as...",
        initialdir=initial_dir,
        initialfile=default_name,
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    root.destroy()

    if not save_path:
        # User cancelled
        print("No file selected. Exiting.")
        sys.exit(0)

    try:
        out = build_pdf(save_path)
        print(f"PDF saved to: {out}")
    except Exception as e:
        # Show messagebox if possible
        try:
            root = Tk(); root.withdraw(); messagebox.showerror("Error", f"Failed to build PDF:\n{e}"); root.destroy()
        except:
            print(f"Failed to build PDF: {e}")
        sys.exit(1)
