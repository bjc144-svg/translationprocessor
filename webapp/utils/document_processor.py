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

    def convert_doc_to_docx(self, doc_path):
        """Convert .doc file to .docx using Word COM"""
        try:
            import pythoncom
            import win32com.client

            # Initialize COM
            pythoncom.CoInitialize()

            try:
                # Create output path
                docx_path = str(doc_path).replace('.doc', '.docx')

                # Open Word
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False

                # Open the .doc file
                doc = word.Documents.Open(str(Path(doc_path).absolute()))

                # Save as .docx (FileFormat 16 = .docx)
                doc.SaveAs2(str(Path(docx_path).absolute()), FileFormat=16)
                doc.Close()
                word.Quit()

                print(f"Converted .doc to .docx: {docx_path}")
                return docx_path

            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            print(f"Error converting .doc to .docx: {e}")
            import traceback
            traceback.print_exc()
            return None

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

            # Convert .doc to .docx if needed
            input_path = Path(input_file)
            if input_path.suffix.lower() == '.doc':
                print("Detected .doc file, converting to .docx...")
                converted_path = self.convert_doc_to_docx(input_file)
                if converted_path:
                    input_file = converted_path
                else:
                    print("Error: Failed to convert .doc to .docx")
                    return False

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
        original_doc = Document(input_file)

        # IMPORTANT: Create a new document with a single section to ensure correct page numbering
        # If the document has multiple sections, SECTIONPAGES will only count pages in each individual section
        # By creating a new document with all content in one section, SECTIONPAGES will count all pages
        print(f"Original document has {len(original_doc.sections)} section(s)")

        # Create a new document to consolidate content
        doc = Document()

        # Copy all paragraphs from the original document (regardless of sections)
        from io import BytesIO
        for para in original_doc.paragraphs:
            new_para = doc.add_paragraph()
            new_para.alignment = para.alignment
            new_para.style = para.style

            # Copy runs (text with formatting and images)
            for run in para.runs:
                # Check if run contains an image
                has_image = False
                drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing', run._element.nsmap if hasattr(run._element, 'nsmap') else None)

                if not drawings:
                    for child in run._element:
                        if 'drawing' in child.tag or 'pict' in child.tag:
                            drawings = [child]
                            break

                if drawings:
                    has_image = True
                    # Copy images
                    for drawing in drawings:
                        try:
                            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                            blip = drawing.find('.//a:blip', ns)

                            if blip is None:
                                for elem in drawing.iter():
                                    if 'blip' in elem.tag.lower():
                                        blip = elem
                                        break

                            if blip is not None:
                                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                if not rId:
                                    rId = blip.get('embed')

                                if rId and rId in original_doc.part.related_parts:
                                    image_part = original_doc.part.related_parts[rId]
                                    image_bytes = image_part.blob

                                    # Get image dimensions
                                    extent = drawing.find('.//wp:extent', {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
                                    if extent is None:
                                        for elem in drawing.iter():
                                            if 'extent' in elem.tag.lower():
                                                extent = elem
                                                break

                                    new_run = new_para.add_run()
                                    if extent is not None and extent.get('cx'):
                                        try:
                                            cx = int(extent.get('cx'))
                                            width_inches = cx / 914400.0
                                            new_run.add_picture(BytesIO(image_bytes), width=Inches(width_inches))
                                        except:
                                            new_run.add_picture(BytesIO(image_bytes))
                                    else:
                                        new_run.add_picture(BytesIO(image_bytes))
                        except Exception as e:
                            print(f"Warning: Could not copy image: {e}")

                # Copy text if no image
                if not has_image:
                    new_run = new_para.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.name:
                        new_run.font.name = run.font.name

        # Copy tables from the original document
        for table in original_doc.tables:
            new_table = doc.add_table(rows=len(table.rows), cols=len(table.columns))
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    new_table.rows[i].cells[j].text = cell.text

        print(f"Consolidated document has {len(doc.sections)} section(s)")

        # Apply header/footer to ALL sections
        for section in doc.sections:
            # Set up header
            header = section.header
            header.is_linked_to_previous = False

            # Clear existing header
            for paragraph in header.paragraphs:
                paragraph.clear()

            # Create header table (2 columns) - span full width
            header_table = header.add_table(1, 2, Inches(6.5))
            header_table.autofit = False
            header_table.allow_autofit = False

            # Set table to full width with no spacing
            tbl = header_table._element
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)

            # Set table width to 100% (5000 = 100% in Word's measurement)
            tblW = OxmlElement('w:tblW')
            tblW.set(qn('w:w'), '5000')
            tblW.set(qn('w:type'), 'pct')
            tblPr.append(tblW)

            # Remove table cell spacing
            tblCellSpacing = OxmlElement('w:tblCellSpacing')
            tblCellSpacing.set(qn('w:w'), '0')
            tblCellSpacing.set(qn('w:type'), 'dxa')
            tblPr.append(tblCellSpacing)

            # Set column widths (proportional)
            header_table.rows[0].cells[0].width = Inches(2.75)
            header_table.rows[0].cells[1].width = Inches(3.75)

            # Left cell - Logo
            left_cell = header_table.rows[0].cells[0]
            left_para = left_cell.paragraphs[0]
            left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Try to add logo image
            logo_path = self.assets_dir / 'park_logo.png'
            logo_added = False

            if logo_path.exists():
                try:
                    run = left_para.add_run()
                    run.add_picture(str(logo_path), height=Inches(0.5))
                    logo_added = True
                    print(f"Successfully added Park logo from: {logo_path}")
                except Exception as e:
                    print(f"Error adding logo image: {e}")
                    import traceback
                    traceback.print_exc()

            # If logo couldn't be loaded, use text fallback
            if not logo_added:
                left_para.text = "PARK EVALUATION SERVICES"
                left_para.runs[0].font.bold = True
                left_para.runs[0].font.size = Pt(10)
                print(f"Warning: Park logo file not found at: {logo_path}")

            # Right cell - Contact info
            right_cell = header_table.rows[0].cells[1]
            right_cell.vertical_alignment = 1  # Center vertically

            # Remove cell margins for proper right alignment
            tc = right_cell._element
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for margin_name in ['top', 'left', 'bottom', 'right']:
                node = OxmlElement(f'w:{margin_name}')
                node.set(qn('w:w'), '0')
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)

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
                # Remove paragraph spacing for tight alignment
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                run = para.runs[0]
                run.font.size = Pt(10)

            # Add horizontal line after header using border
            line_para = header.add_paragraph()
            pPr = line_para._element.get_or_add_pPr()
            # Remove paragraph spacing
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            pPr.append(spacing)
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), '000000')
            pBdr.append(bottom)
            pPr.append(pBdr)

            # Set up footer
            footer = section.footer
            footer.is_linked_to_previous = False

            # Clear existing footer
            for paragraph in footer.paragraphs:
                paragraph.clear()

            # Add horizontal line before footer
            line_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            pPr = line_para._element.get_or_add_pPr()
            # Remove paragraph spacing
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            pPr.append(spacing)
            pBdr = OxmlElement('w:pBdr')
            top = OxmlElement('w:top')
            top.set(qn('w:val'), 'single')
            top.set(qn('w:sz'), '6')
            top.set(qn('w:space'), '0')
            top.set(qn('w:color'), '000000')
            pBdr.append(top)
            pPr.append(pBdr)

            # Add "CERTIFIED TRANSLATION" (centered, all caps)
            cert_para = footer.add_paragraph("CERTIFIED TRANSLATION")
            cert_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Remove paragraph spacing
            cert_para.paragraph_format.space_before = Pt(0)
            cert_para.paragraph_format.space_after = Pt(0)
            cert_para.runs[0].font.bold = True
            cert_para.runs[0].font.size = Pt(12)

            # Add footer info table (3 columns) - span full width like header
            footer_table = footer.add_table(1, 3, Inches(6.5))
            footer_table.autofit = False
            footer_table.allow_autofit = False

            # Set table to full width with no spacing
            tbl = footer_table._element
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)

            # Set table width to 100% (5000 = 100% in Word's measurement)
            tblW = OxmlElement('w:tblW')
            tblW.set(qn('w:w'), '5000')
            tblW.set(qn('w:type'), 'pct')
            tblPr.append(tblW)

            # Remove table cell spacing
            tblCellSpacing = OxmlElement('w:tblCellSpacing')
            tblCellSpacing.set(qn('w:w'), '0')
            tblCellSpacing.set(qn('w:type'), 'dxa')
            tblPr.append(tblCellSpacing)

            # Left - Case number
            left_cell = footer_table.rows[0].cells[0]
            left_para = left_cell.paragraphs[0]
            left_para.text = f"Park Case #{metadata['case_number']}"
            left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            left_para.runs[0].font.size = Pt(9)

            # Center - Page number with field codes
            center_cell = footer_table.rows[0].cells[1]
            center_para = center_cell.paragraphs[0]
            center_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Add "Page " text
            run = center_para.add_run("Page ")
            run.font.size = Pt(9)

            # Add PAGE field (current page number)
            self._add_page_number(center_para)

            # Add " of " text
            run = center_para.add_run(" of ")
            run.font.size = Pt(9)

            # Add NUMPAGES field (total pages)
            self._add_num_pages(center_para)

            # Right - Language pair
            right_cell = footer_table.rows[0].cells[2]
            right_para = right_cell.paragraphs[0]
            right_para.text = metadata['language_pair']
            right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            right_para.runs[0].font.size = Pt(9)

        # Save document
        doc.save(output_file)
        print(f"Saved document with header/footer to: {output_file}")

    def _add_page_number(self, paragraph):
        """Add PAGE field to paragraph for current page number"""
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = 'PAGE'

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.size = Pt(9)

    def _add_num_pages(self, paragraph):
        """Add SECTIONPAGES field to count only pages in the current section (translation, not certificates)"""
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = 'SECTIONPAGES'

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.size = Pt(9)

    def create_translator_certificate(self, output_file, metadata):
        """Create translator certificate"""

        doc = Document()

        # Configure section - no header/footer, Portrait orientation, Letter size
        section = doc.sections[0]
        section.page_height = Inches(11)  # Letter height
        section.page_width = Inches(8.5)  # Letter width
        section.orientation = 0  # 0 = Portrait, 1 = Landscape
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Clear header and footer completely
        for paragraph in section.header.paragraphs:
            paragraph.clear()
        for paragraph in section.footer.paragraphs:
            paragraph.clear()

        # Add logo
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Try to add logo image
        logo_path = self.assets_dir / 'park_logo.png'
        logo_added = False

        if logo_path.exists():
            try:
                run = logo_para.add_run()
                run.add_picture(str(logo_path), height=Inches(0.8))
                logo_added = True
                print(f"Successfully added Park logo to translator certificate from: {logo_path}")
            except Exception as e:
                print(f"Error adding logo image to translator certificate: {e}")
                import traceback
                traceback.print_exc()

        # If logo couldn't be loaded, use text fallback
        if not logo_added:
            run = logo_para.add_run("PARK EVALUATION SERVICES")
            run.font.size = Pt(14)
            run.font.bold = True
            print(f"Warning: Park logo file not found at: {logo_path}")

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

        sig_label = doc.add_paragraph(metadata['translator_name'])
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

        # Add signature image if available
        sig_added = False
        if metadata.get('translator_signature'):
            sig_path = metadata['translator_signature']
            print(f"Looking for translator signature at: {sig_path}")

            if os.path.exists(sig_path):
                try:
                    # Insert signature image ABOVE the line
                    sig_img_para = doc.paragraphs[-2]  # Get the signature line paragraph
                    run = sig_img_para.insert_paragraph_before().add_run()
                    run.add_picture(sig_path, width=Inches(2))
                    sig_added = True
                    print(f"Successfully added translator signature from: {sig_path}")
                except Exception as e:
                    print(f"Error adding signature image: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                print(f"Warning: Translator signature file not found: {sig_path}")

        if not sig_added:
            # Add placeholder text if signature couldn't be loaded
            sig_placeholder = doc.add_paragraph("[Signature image not available]")
            sig_placeholder.alignment = WD_ALIGN_PARAGRAPH.LEFT
            sig_placeholder.runs[0].font.size = Pt(9)
            sig_placeholder.runs[0].font.italic = True

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

        # Configure section - no header/footer, Portrait orientation, Letter size
        section = doc.sections[0]
        section.page_height = Inches(11)  # Letter height
        section.page_width = Inches(8.5)  # Letter width
        section.orientation = 0  # 0 = Portrait, 1 = Landscape
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        # Clear header and footer completely
        for paragraph in section.header.paragraphs:
            paragraph.clear()
        for paragraph in section.footer.paragraphs:
            paragraph.clear()

        # Add logo
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Try to add logo image
        logo_path = self.assets_dir / 'park_logo.png'
        logo_added = False

        if logo_path.exists():
            try:
                run = logo_para.add_run()
                run.add_picture(str(logo_path), height=Inches(0.8))
                logo_added = True
                print(f"Successfully added Park logo to Park certificate from: {logo_path}")
            except Exception as e:
                print(f"Error adding logo image to Park certificate: {e}")
                import traceback
                traceback.print_exc()

        # If logo couldn't be loaded, use text fallback
        if not logo_added:
            run = logo_para.add_run("PARK EVALUATION SERVICES")
            run.font.size = Pt(14)
            run.font.bold = True
            print(f"Warning: Park logo file not found at: {logo_path}")

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

        # Signature section (Park employee)
        sig_para = doc.add_paragraph("_" * 40)
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        sig_label = doc.add_paragraph("Authorized Signature - Park Evaluation Services")
        sig_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_label.runs[0].font.size = Pt(10)
        sig_label.runs[0].font.italic = True

        # Add Park signature image from assets directory
        park_sig_path = self.assets_dir / 'park_signature.png'
        sig_added = False

        print(f"Looking for Park signature at: {park_sig_path}")

        if park_sig_path.exists():
            try:
                # Insert signature image ABOVE the line
                sig_img_para = doc.paragraphs[-2]  # Get the signature line paragraph
                run = sig_img_para.insert_paragraph_before().add_run()
                run.add_picture(str(park_sig_path), width=Inches(2))
                sig_added = True
                print(f"Successfully added Park signature from: {park_sig_path}")
            except Exception as e:
                print(f"Error adding Park signature image: {e}")
                import traceback
                traceback.print_exc()
        else:
            print(f"Warning: Park signature file not found: {park_sig_path}")

        if not sig_added:
            # Add placeholder text if signature couldn't be loaded
            sig_placeholder = doc.add_paragraph("[Park signature image not available]")
            sig_placeholder.alignment = WD_ALIGN_PARAGRAPH.LEFT
            sig_placeholder.runs[0].font.size = Pt(9)
            sig_placeholder.runs[0].font.italic = True

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
        """Combine multiple Word documents into one using safe section-based approach"""
        from io import BytesIO
        import zipfile
        from lxml import etree

        # Start with the first document (translation with header/footer)
        combined = Document(doc_paths[0])

        # Set the first section to restart page numbering at 1
        # This ensures page numbers in the translation section don't include certificate pages
        first_section = combined.sections[0]
        sectPr = first_section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is None:
            pgNumType = OxmlElement('w:pgNumType')
            sectPr.append(pgNumType)
        pgNumType.set(qn('w:start'), '1')

        # Add remaining documents (certificates)
        for idx, doc_path in enumerate(doc_paths[1:], start=1):
            # Read the document to append
            sub_doc = Document(doc_path)

            # Add a new section for this document (creates page break and allows different header/footer)
            new_section = combined.add_section()

            # Make sure this section doesn't inherit header/footer from previous section
            new_section.header.is_linked_to_previous = False
            new_section.footer.is_linked_to_previous = False

            # Clear header and footer for this section (certificates shouldn't have them)
            for paragraph in new_section.header.paragraphs:
                paragraph.clear()
            for paragraph in new_section.footer.paragraphs:
                paragraph.clear()

            # Restart page numbering for this section at 1
            # This ensures certificates don't continue the page count from translation pages
            sectPr = new_section._sectPr
            pgNumType = sectPr.find(qn('w:pgNumType'))
            if pgNumType is None:
                pgNumType = OxmlElement('w:pgNumType')
                sectPr.append(pgNumType)
            pgNumType.set(qn('w:start'), '1')

            # Copy all paragraphs from sub_doc
            for paragraph in sub_doc.paragraphs:
                # Create a new paragraph in the combined document
                new_para = combined.add_paragraph()
                new_para.alignment = paragraph.alignment
                new_para.style = paragraph.style

                # Copy paragraph formatting
                if paragraph.runs:
                    for run in paragraph.runs:
                        # Check if run contains an image (drawing or picture)
                        has_image = False

                        # Check for drawing elements (modern Word images)
                        drawings = run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing', run._element.nsmap if hasattr(run._element, 'nsmap') else None)

                        if not drawings:
                            # Also check without namespace for older formats
                            for child in run._element:
                                if 'drawing' in child.tag or 'pict' in child.tag:
                                    drawings = [child]
                                    break

                        if drawings:
                            has_image = True
                            # Try to extract and copy the image
                            for drawing in drawings:
                                try:
                                    # Find the blip element that contains the image relationship ID
                                    # Try multiple ways to find it
                                    blip = None

                                    # Method 1: Use find with full namespace
                                    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                                          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                                    blip = drawing.find('.//a:blip', ns)

                                    # Method 2: Search all descendants
                                    if blip is None:
                                        for elem in drawing.iter():
                                            if 'blip' in elem.tag.lower():
                                                blip = elem
                                                break

                                    if blip is not None:
                                        # Get the relationship ID
                                        rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                        if not rId:
                                            # Try without namespace
                                            rId = blip.get('embed')

                                        if rId and rId in sub_doc.part.related_parts:
                                            # Get the image part
                                            image_part = sub_doc.part.related_parts[rId]
                                            image_bytes = image_part.blob

                                            # Try to get image dimensions from the extent element
                                            extent = drawing.find('.//wp:extent', {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
                                            if extent is None:
                                                # Try searching all descendants
                                                for elem in drawing.iter():
                                                    if 'extent' in elem.tag.lower():
                                                        extent = elem
                                                        break

                                            # Add the image to the new paragraph
                                            new_run = new_para.add_run()

                                            if extent is not None and extent.get('cx'):
                                                # Convert EMU to inches (1 inch = 914400 EMUs)
                                                try:
                                                    cx = int(extent.get('cx'))
                                                    width_inches = cx / 914400.0
                                                    new_run.add_picture(BytesIO(image_bytes), width=Inches(width_inches))
                                                except:
                                                    # Fallback to default size
                                                    new_run.add_picture(BytesIO(image_bytes))
                                            else:
                                                # No size info, use default
                                                new_run.add_picture(BytesIO(image_bytes))

                                            print(f"Successfully copied image from certificate")
                                        else:
                                            print(f"Warning: Could not find image relationship {rId}")
                                            new_run = new_para.add_run("[Image not found]")
                                    else:
                                        print(f"Warning: Could not find blip element in drawing")
                                        new_run = new_para.add_run("[Image]")

                                except Exception as e:
                                    print(f"Warning: Could not copy image: {e}")
                                    import traceback
                                    traceback.print_exc()
                                    new_run = new_para.add_run("[Image error]")

                        # If no image, copy text and formatting
                        if not has_image:
                            new_run = new_para.add_run(run.text)
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            if run.font.size:
                                new_run.font.size = run.font.size
                            if run.font.name:
                                new_run.font.name = run.font.name

            # Copy tables
            for table in sub_doc.tables:
                new_table = combined.add_table(rows=len(table.rows), cols=len(table.columns))
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.rows[i].cells[j].text = cell.text

        # Save combined document
        combined.save(output_path)
        print(f"Saved combined document to: {output_path}")

    def convert_to_pdf(self, docx_path, pdf_path):
        """Convert Word document to PDF"""

        try:
            # Try using docx2pdf (works on Windows with Word installed)
            # Initialize COM for the current thread
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_initialized = True
            except:
                com_initialized = False

            try:
                from docx2pdf import convert
                convert(str(docx_path), str(pdf_path))
                print(f"Converted to PDF using docx2pdf: {pdf_path}")
                return True
            finally:
                # Uninitialize COM
                if com_initialized:
                    try:
                        pythoncom.CoUninitialize()
                    except:
                        pass

        except Exception as e:
            print(f"docx2pdf conversion failed: {e}")

            # Fallback: Try using win32com (Windows only)
            try:
                import pythoncom
                import win32com.client

                # Initialize COM for the current thread
                pythoncom.CoInitialize()

                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False

                    doc = word.Documents.Open(str(Path(docx_path).absolute()))
                    doc.SaveAs(str(Path(pdf_path).absolute()), FileFormat=17)  # 17 = PDF
                    doc.Close()
                    word.Quit()

                    print(f"Converted to PDF using Word COM: {pdf_path}")
                    return True
                finally:
                    # Uninitialize COM
                    try:
                        pythoncom.CoUninitialize()
                    except:
                        pass

            except Exception as e2:
                print(f"Word COM conversion also failed: {e2}")

                # If both methods fail, save as DOCX with note
                fallback_path = str(pdf_path).replace('.pdf', '_CONVERT_TO_PDF.docx')
                shutil.copy(docx_path, fallback_path)
                print(f"Could not convert to PDF. Saved as Word document: {fallback_path}")
                print("Please manually convert to PDF using Microsoft Word.")
                return False
