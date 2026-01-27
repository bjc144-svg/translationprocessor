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

        # Add header and footer to all sections
        for section in doc.sections:
            # Add header
            header = section.header
            header.is_linked_to_previous = False

            # Clear existing header
            for paragraph in header.paragraphs:
                paragraph.clear()

            # Create header table (2 columns) with proper widths
            header_table = header.add_table(rows=1, cols=2)
            header_table.autofit = False
            header_table.allow_autofit = False

            # Set column widths
            header_table.columns[0].width = Inches(3.0)  # Logo column
            header_table.columns[1].width = Inches(3.5)  # Contact info column

            # Left cell - Logo
            left_cell = header_table.rows[0].cells[0]
            left_cell.width = Inches(3.0)
            left_para = left_cell.paragraphs[0]

            # Try to add logo image
            logo_path = self.assets_dir / 'logo.png'
            if logo_path.exists():
                try:
                    run = left_para.add_run()
                    run.add_picture(str(logo_path), width=Inches(2.0))
                    left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                except Exception as e:
                    print(f"Error adding logo: {e}")
                    left_para.text = "Park Evaluation Services"
                    left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                left_para.text = "Park Evaluation Services"
                left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = left_para.runs[0]
                run.font.bold = True
                run.font.size = Pt(12)

            # Right cell - Contact info
            right_cell = header_table.rows[0].cells[1]
            right_cell.width = Inches(3.5)
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

            # Add horizontal line after header using border
            line_para = header.add_paragraph()
            self._add_horizontal_line(line_para)

            # Add footer
            footer = section.footer
            footer.is_linked_to_previous = False

            # Clear existing footer
            for paragraph in footer.paragraphs:
                paragraph.clear()

            # Add horizontal line before footer
            footer_line = footer.paragraphs[0]
            self._add_horizontal_line(footer_line)

            # Add "CERTIFIED TRANSLATION" (centered, all caps)
            cert_para = footer.add_paragraph("CERTIFIED TRANSLATION")
            cert_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cert_para.runs[0].font.bold = True
            cert_para.runs[0].font.size = Pt(12)

            # Add footer info table (3 columns)
            footer_table = footer.add_table(rows=1, cols=3)
            footer_table.autofit = False

            # Set column widths
            footer_table.columns[0].width = Inches(2.0)
            footer_table.columns[1].width = Inches(2.5)
            footer_table.columns[2].width = Inches(2.0)

            # Left - Case number
            left_cell = footer_table.rows[0].cells[0]
            left_para = left_cell.paragraphs[0]
            left_para.text = f"Park Case #{metadata['case_number']}"
            left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            left_para.runs[0].font.size = Pt(9)

            # Center - Page number (using field codes)
            center_cell = footer_table.rows[0].cells[1]
            center_para = center_cell.paragraphs[0]
            center_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Add page number field codes
            self._add_page_number_field(center_para)

            # Right - Language pair
            right_cell = footer_table.rows[0].cells[2]
            right_para = right_cell.paragraphs[0]
            right_para.text = metadata['language_pair']
            right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            right_para.runs[0].font.size = Pt(9)

        # Save document
        doc.save(output_file)
        print(f"Saved document with header/footer to: {output_file}")

    def _add_horizontal_line(self, paragraph):
        """Add a horizontal line to a paragraph using border"""
        p = paragraph._element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')  # Line thickness
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '000000')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def _add_page_number_field(self, paragraph):
        """Add page number field (Page X of Y) to paragraph"""
        run = paragraph.add_run("Page ")
        run.font.size = Pt(9)

        # Add PAGE field for current page
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run_element = run._element
        run_element.append(fldChar1)
        run_element.append(instrText)
        run_element.append(fldChar2)

        # Add " of " text
        run2 = paragraph.add_run(" of ")
        run2.font.size = Pt(9)

        # Add NUMPAGES field for total pages
        fldChar3 = OxmlElement('w:fldChar')
        fldChar3.set(qn('w:fldCharType'), 'begin')

        instrText2 = OxmlElement('w:instrText')
        instrText2.set(qn('xml:space'), 'preserve')
        instrText2.text = "NUMPAGES"

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        run2_element = run2._element
        run2_element.append(fldChar3)
        run2_element.append(instrText2)
        run2_element.append(fldChar4)

    def create_translator_certificate(self, output_file, metadata):
        """Create translator certificate"""

        doc = Document()

        # Configure section to have no header/footer
        section = doc.sections[0]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Clear all header content
        for paragraph in section.header.paragraphs:
            paragraph.clear()

        # Clear all footer content
        for paragraph in section.footer.paragraphs:
            paragraph.clear()

        # Set different first page (to ensure headers/footers don't appear)
        section.different_first_page_header_footer = True

        # Add logo
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        logo_path = self.assets_dir / 'logo.png'
        if logo_path.exists():
            try:
                run = logo_para.add_run()
                run.add_picture(str(logo_path), width=Inches(2.5))
            except Exception as e:
                print(f"Error adding logo: {e}")
                run = logo_para.add_run("Park Evaluation Services")
                run.font.size = Pt(14)
                run.font.bold = True
        else:
            run = logo_para.add_run("Park Evaluation Services")
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

        # Signature section - line above signature
        sig_line_para = doc.add_paragraph()
        sig_line_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_line_run = sig_line_para.add_run("_" * 40)
        sig_line_run.font.size = Pt(11)

        # Add signature image if available (ABOVE the line)
        if metadata.get('translator_signature') and os.path.exists(metadata['translator_signature']):
            try:
                # Insert signature image BEFORE the line
                sig_img_para = doc.paragraphs[-2]  # Get paragraph before the line
                sig_img_para.insert_paragraph_before()
                inserted_para = doc.paragraphs[-3]
                inserted_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = inserted_para.add_run()

                # Use absolute path for signature
                sig_path = Path(metadata['translator_signature']).absolute()
                run.add_picture(str(sig_path), width=Inches(2))
                print(f"Added signature image: {sig_path}")
            except Exception as e:
                print(f"Error adding signature image: {e}")
                import traceback
                traceback.print_exc()

        sig_label = doc.add_paragraph("Translator Signature")
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

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

        # Configure section to have no header/footer
        section = doc.sections[0]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Clear all header content
        for paragraph in section.header.paragraphs:
            paragraph.clear()

        # Clear all footer content
        for paragraph in section.footer.paragraphs:
            paragraph.clear()

        # Set different first page (to ensure headers/footers don't appear)
        section.different_first_page_header_footer = True

        # Add logo
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        logo_path = self.assets_dir / 'logo.png'
        if logo_path.exists():
            try:
                run = logo_para.add_run()
                run.add_picture(str(logo_path), width=Inches(2.5))
            except Exception as e:
                print(f"Error adding logo: {e}")
                run = logo_para.add_run("Park Evaluation Services")
                run.font.size = Pt(14)
                run.font.bold = True
        else:
            run = logo_para.add_run("Park Evaluation Services")
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

        # Signature section (Park employee signature)
        # Check for park_signature.png in assets folder
        park_sig_path = self.assets_dir / 'park_signature.png'
        if park_sig_path.exists():
            try:
                sig_img_para = doc.add_paragraph()
                sig_img_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = sig_img_para.add_run()
                run.add_picture(str(park_sig_path), width=Inches(2))
                print(f"Added Park signature: {park_sig_path}")
            except Exception as e:
                print(f"Error adding Park signature: {e}")

        sig_para = doc.add_paragraph("_" * 40)
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        sig_label = doc.add_paragraph("Authorized Signature - Park Evaluation Services")
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

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

        # Create new document starting with first document
        combined = Document(doc_paths[0])

        # Add remaining documents
        for doc_path in doc_paths[1:]:
            # Add page break before next document
            combined.add_page_break()

            # Read document to append
            sub_doc = Document(doc_path)

            # Copy paragraphs and other elements (excluding section properties)
            for element in sub_doc.element.body:
                # Skip section properties to avoid blank pages
                if element.tag.endswith('}sectPr'):
                    continue
                combined.element.body.append(element)

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
