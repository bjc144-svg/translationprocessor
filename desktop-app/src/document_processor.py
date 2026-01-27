"""
Document Processor
Handles Word document processing, header/footer addition, certificate generation, and PDF conversion
"""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from pathlib import Path
import tempfile
import shutil


class DocumentProcessor:
    def __init__(self):
        """Initialize document processor"""

        # Get templates directory
        self.templates_dir = Path(__file__).parent.parent / 'templates'
        self.templates_dir.mkdir(parents=True, exist_ok=True)

        # Get assets directory
        self.assets_dir = Path(__file__).parent.parent / 'assets'
        self.assets_dir.mkdir(parents=True, exist_ok=True)

    def process_translation(self, input_file, output_file, metadata):
        """
        Process translation document

        Args:
            input_file: Path to input Word document
            output_file: Path to output PDF file
            metadata: Dictionary containing:
                - case_number
                - language_pair
                - source_language
                - target_language
                - translator_name
                - translator_signature (path to signature image)
                - date

        Returns:
            True if successful, False otherwise
        """

        try:
            print("Starting document processing...")
            print(f"Input: {input_file}")
            print(f"Output: {output_file}")
            print(f"Metadata: {metadata}")

            # Create temporary directory for processing
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)

                # Step 1: Add header and footer to translation document
                print("Step 1: Adding header and footer to translation...")
                translation_docx = temp_path / "translation_with_header.docx"
                self.add_header_footer(input_file, translation_docx, metadata)

                # Step 2: Create translator certificate
                print("Step 2: Creating translator certificate...")
                translator_cert_docx = temp_path / "translator_certificate.docx"
                self.create_translator_certificate(translator_cert_docx, metadata)

                # Step 3: Create Park certificate
                print("Step 3: Creating Park certificate...")
                park_cert_docx = temp_path / "park_certificate.docx"
                self.create_park_certificate(park_cert_docx, metadata)

                # Step 4: Combine documents
                print("Step 4: Combining documents...")
                combined_docx = temp_path / "combined.docx"
                self.combine_documents(
                    [translation_docx, translator_cert_docx, park_cert_docx],
                    combined_docx
                )

                # Step 5: Convert to PDF
                print("Step 5: Converting to PDF...")
                self.convert_to_pdf(combined_docx, output_file)

                print("Processing complete!")
                return True

        except Exception as e:
            print(f"Error processing document: {e}")
            import traceback
            traceback.print_exc()
            return False

    def add_header_footer(self, input_file, output_file, metadata):
        """Add header and footer to translation document"""

        # Open document
        doc = Document(input_file)

        # Add header
        section = doc.sections[0]
        header = section.header

        # Clear existing header
        header.is_linked_to_previous = False
        for paragraph in header.paragraphs:
            paragraph.clear()

        # Create header table (2 columns)
        header_table = header.add_table(1, 2, Inches(6.5))
        header_table.autofit = False

        # Left cell - Logo placeholder
        left_cell = header_table.rows[0].cells[0]
        left_para = left_cell.paragraphs[0]
        left_para.text = "[LOGO HERE]"
        left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Right cell - Contact info
        right_cell = header_table.rows[0].cells[1]
        right_cell.vertical_alignment = 1  # Center vertically

        # Add contact info (right-aligned)
        contact_lines = [
            "212-581-8877",
            "eval@parkeval.com",
            "www.parkeval.com"
        ]

        for i, line in enumerate(contact_lines):
            if i > 0:
                right_cell.add_paragraph()
            para = right_cell.paragraphs[i] if i == 0 else right_cell.paragraphs[-1]
            para.text = line
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = para.runs[0]
            run.font.size = Pt(10)

        # Add horizontal line after header using border instead of underscores
        line_para = header.add_paragraph()
        pPr = line_para._element.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '000000')
        pBdr.append(bottom)
        pPr.append(pBdr)

        # Add footer
        footer = section.footer
        footer.is_linked_to_previous = False

        # Clear existing footer
        for paragraph in footer.paragraphs:
            paragraph.clear()

        # Add horizontal line before footer using border instead of underscores
        line_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        pPr = line_para._element.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '6')
        top.set(qn('w:space'), '1')
        top.set(qn('w:color'), '000000')
        pBdr.append(top)
        pPr.append(pBdr)

        # Add "CERTIFIED TRANSLATION" (centered, all caps)
        cert_para = footer.add_paragraph("CERTIFIED TRANSLATION")
        cert_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cert_para.runs[0].font.bold = True
        cert_para.runs[0].font.size = Pt(12)

        # Add footer info table (3 columns)
        footer_table = footer.add_table(1, 3, Inches(6.5))
        footer_table.autofit = False

        # Left - Case number
        left_cell = footer_table.rows[0].cells[0]
        left_para = left_cell.paragraphs[0]
        left_para.text = f"Park Case #{metadata['case_number']}"
        left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        left_para.runs[0].font.size = Pt(9)

        # Center - Page number (using field code)
        center_cell = footer_table.rows[0].cells[1]
        center_para = center_cell.paragraphs[0]
        center_para.text = "Page "
        center_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        center_para.runs[0].font.size = Pt(9)

        # Add page number field
        # Note: This is a simplified version - actual page numbering requires field codes
        run = center_para.add_run()
        run.text = "X of Y"  # Placeholder - proper page numbers need field codes
        run.font.size = Pt(9)

        # Right - Language pair
        right_cell = footer_table.rows[0].cells[2]
        right_para = right_cell.paragraphs[0]
        right_para.text = metadata['language_pair']
        right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        right_para.runs[0].font.size = Pt(9)

        # Save document
        doc.save(output_file)
        print(f"Saved document with header/footer to: {output_file}")

    def create_translator_certificate(self, output_file, metadata):
        """Create translator certificate"""

        doc = Document()

        # Remove default header/footer
        section = doc.sections[0]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Add logo placeholder
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = logo_para.add_run("[PARK EVALUATION SERVICES LOGO]")
        run.font.size = Pt(14)
        run.font.bold = True

        doc.add_paragraph()  # Spacing

        # Title
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("TRANSLATOR CERTIFICATION")
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.all_caps = True

        doc.add_paragraph()  # Spacing

        # Certificate text
        cert_text = f"""I, {metadata['translator_name']}, hereby certify that I am competent to translate from {metadata['source_language']} into {metadata['target_language']}, and that the attached translation of the document is accurate and complete to the best of my knowledge and belief.

I further certify that I am not a party to this action and do not have a financial interest in the outcome of this matter.

This certification is made in accordance with the requirements of the Federal Rules of Civil Procedure and applicable state rules.

Date: {metadata['date']}

Translator Name: {metadata['translator_name']}

Language Pair: {metadata['language_pair']}

Case Number: Park Case #{metadata['case_number']}
"""

        para = doc.add_paragraph(cert_text)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for run in para.runs:
            run.font.size = Pt(11)

        doc.add_paragraph()  # Spacing

        # Signature section
        sig_para = doc.add_paragraph("_" * 40)
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        sig_label = doc.add_paragraph("Translator Signature")
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

        # Add signature image if available
        if metadata.get('translator_signature') and os.path.exists(metadata['translator_signature']):
            try:
                # Insert signature image
                sig_img_para = doc.add_paragraph()
                sig_img_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = sig_img_para.add_run()
                run.add_picture(metadata['translator_signature'], width=Inches(2))
            except Exception as e:
                print(f"Error adding signature image: {e}")

        doc.add_paragraph()  # Spacing
        doc.add_paragraph()  # Spacing

        # Footer with contact info
        footer_para = doc.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer_para.add_run("212-581-8877 • www.parkeval.com")
        run.font.size = Pt(10)

        # Save certificate
        doc.save(output_file)
        print(f"Saved translator certificate to: {output_file}")

    def create_park_certificate(self, output_file, metadata):
        """Create Park employee certificate"""

        doc = Document()

        # Remove default header/footer
        section = doc.sections[0]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Add logo placeholder
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = logo_para.add_run("[PARK EVALUATION SERVICES LOGO]")
        run.font.size = Pt(14)
        run.font.bold = True

        doc.add_paragraph()  # Spacing

        # Title
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("CERTIFICATION OF TRANSLATION SERVICES")
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.all_caps = True

        doc.add_paragraph()  # Spacing

        # Certificate text
        cert_text = f"""Park Evaluation Services hereby certifies that the attached translation has been reviewed and verified for accuracy and completeness.

This translation was performed by a qualified translator competent in {metadata['language_pair']}.

The translation services provided meet professional standards and are suitable for official use.

Date: {metadata['date']}

Language Pair: {metadata['language_pair']}

Case Number: Park Case #{metadata['case_number']}

This certification is provided by Park Evaluation Services in the regular course of business.
"""

        para = doc.add_paragraph(cert_text)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for run in para.runs:
            run.font.size = Pt(11)

        doc.add_paragraph()  # Spacing

        # Signature section (Park employee - fixed)
        sig_para = doc.add_paragraph("_" * 40)
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        sig_label = doc.add_paragraph("Authorized Signature - Park Evaluation Services")
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

        # Placeholder for Park employee signature (you can add actual signature image here)
        # doc.add_picture('path/to/park_signature.png', width=Inches(2))

        doc.add_paragraph()  # Spacing
        doc.add_paragraph()  # Spacing

        # Footer with contact info
        footer_para = doc.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer_para.add_run("212-581-8877 • www.parkeval.com")
        run.font.size = Pt(10)

        # Save certificate
        doc.save(output_file)
        print(f"Saved Park certificate to: {output_file}")

    def combine_documents(self, doc_paths, output_path):
        """Combine multiple Word documents into one"""
        from copy import deepcopy

        # Create new document starting with first document
        combined = Document(doc_paths[0])

        # Add remaining documents
        for doc_path in doc_paths[1:]:
            # Add page break
            combined.add_page_break()

            # Read document to append
            sub_doc = Document(doc_path)

            # Copy all elements using deepcopy to prevent XML corruption
            for element in sub_doc.element.body:
                # Create a deep copy of the element to preserve all attributes and children
                element_copy = deepcopy(element)
                combined.element.body.append(element_copy)

        # Save combined document
        combined.save(output_path)
        print(f"Saved combined document to: {output_path}")

    def convert_to_pdf(self, docx_path, pdf_path):
        """Convert Word document to PDF"""

        try:
            # Try using docx2pdf (works on Windows with Word installed)
            from docx2pdf import convert
            convert(str(docx_path), str(pdf_path))
            print(f"Converted to PDF using docx2pdf: {pdf_path}")
            return True
        except Exception as e:
            print(f"docx2pdf conversion failed: {e}")

            # Fallback: Try using win32com (Windows only)
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False

                doc = word.Documents.Open(str(Path(docx_path).absolute()))
                doc.SaveAs(str(Path(pdf_path).absolute()), FileFormat=17)  # 17 = PDF
                doc.Close()
                word.Quit()

                print(f"Converted to PDF using Word COM: {pdf_path}")
                return True
            except Exception as e2:
                print(f"Word COM conversion also failed: {e2}")

                # If both methods fail, save as DOCX with note
                fallback_path = str(pdf_path).replace('.pdf', '_CONVERT_TO_PDF.docx')
                shutil.copy(docx_path, fallback_path)
                print(f"Could not convert to PDF. Saved as Word document: {fallback_path}")
                print("Please manually convert to PDF using Microsoft Word.")
                return False
