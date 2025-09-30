import os
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import black

class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str = None, font_size: int = 12):
        self.template_image = template_image
        self.mapping_data = mapping_data
        self.output_font_path = font_path # e.g., "NotoSansDevanagari-Regular.ttf" for Hindi
        self.output_font_size = font_size
        
        # Ensure template is in RGB for drawing
        if self.template_image.mode != 'RGB':
            self.template_image = self.template_image.convert('RGB')
        
        # Determine DPI for reportlab if needed, or assume default
        self.dpi = 300 # Assuming templates are scanned at 300 DPI for good quality
        
        # Calculate scaling factor from template pixels to ReportLab points (1 inch = 72 points)
        # Assuming template_image.width is pixels, and we want to map to a standard page size like A4
        # We need to decide a base page size for ReportLab and scale accordingly
        # Let's say we map the template_image width to the width of A4 portrait
        # A4 width = 8.27 inches = 8.27 * 72 points = 595.2 points
        
        # A simple approach: use the template image itself as the base for PDF generation
        # Set PDF page size to match image size in points if DPI is used
        self.page_width_pts = (self.template_image.width / self.dpi) * inch
        self.page_height_pts = (self.template_image.height / self.dpi) * inch
        
        # Font for PIL drawing (if using PIL for text directly on image before converting to PDF)
        # Using a default font, user can specify NotoSansDevanagari-Regular.ttf etc.
        try:
            self.pil_font = ImageFont.truetype(font_path if font_path else "arial.ttf", font_size)
        except IOError:
            print(f"Warning: Font '{font_path if font_path else 'arial.ttf'}' not found. Using default PIL font.")
            self.pil_font = ImageFont.load_default()

    def _draw_text_on_image(self, draw: ImageDraw.Draw, text: str, x: int, y: int, w: int, h: int, color=(0, 0, 0)):
        """Draws text on the image, handling basic wrapping if needed."""
        # This is a simplification; for complex wrapping, we might need a dedicated function
        # For now, just place the text. More advanced wrapping/font fitting can be added.
        
        # Estimate text size for single line
        text_bbox = draw.textbbox((0,0), text, font=self.pil_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        # Simple centering within the box (vertically and horizontally if text fits)
        # Adjust Y to be top of text, not center
        draw_x = x
        draw_y = y + (h - text_height) // 2 # Center vertically if fits, or align to top

        # If text is too wide, just draw it (truncation might occur, or manual wrapping needed)
        # A better approach for wrapping is to break `text` into lines
        if text_width > w:
            # Simple word wrap logic (can be improved)
            words = text.split(' ')
            lines = []
            current_line = ""
            for word in words:
                test_line = f"{current_line} {word}".strip()
                test_bbox = draw.textbbox((0,0), test_line, font=self.pil_font)
                test_width = test_bbox[2] - test_bbox[0]
                if test_width <= w:
                    current_line = test_line
                else:
                    lines.append(current_line)
                    current_line = word
            lines.append(current_line)

            line_height = text_height # Approximation
            start_y = y + (h - len(lines) * line_height) // 2 # Center block of text vertically

            for i, line in enumerate(lines):
                line_bbox = draw.textbbox((0,0), line, font=self.pil_font)
                line_width = line_bbox[2] - line_bbox[0]
                draw_x_line = x + (w - line_width) // 2 # Center horizontally
                draw_y_line = start_y + i * line_height
                if draw_y_line + line_height <= y + h: # Only draw if it fits in the box
                    draw.text((draw_x_line, draw_y_line), line, font=self.pil_font, fill=color)
        else:
            # Text fits on one line, center it horizontally
            draw_x = x + (w - text_width) // 2
            draw.text((draw_x, draw_y), text, font=self.pil_font, fill=color)


    def fill_and_save_pdf(self, output_folder: str, candidate_data: dict, candidate_srno: str, candidate_name: str, photo_path: str = None):
        """
        Fills the form for a single candidate and saves it as a PDF.
        """
        # Create a blank image to draw on
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)

        # Draw fields
        for field_name, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords["x"], coords["y"], coords["w"], coords["h"]

            if field_name.lower() == "photo" and photo_path and os.path.exists(photo_path):
                try:
                    photo = Image.open(photo_path)
                    photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                    filled_image.paste(photo, (x, y))
                except Exception as e:
                    print(f"Error placing photo for {candidate_name}: {e}")
            else:
                # Find data in candidate_data (case-insensitive search for flexibility)
                field_value = None
                for key, value in candidate_data.items():
                    if key.lower().replace("_", " ") == field_name.lower().replace("_", " "):
                        field_value = str(value)
                        break
                
                if field_value:
                    # Draw text using PIL.ImageDraw (more flexible for exact pixel control)
                    self._draw_text_on_image(draw, field_value, x, y, w, h, color=(0, 0, 0)) # Black text

        # Save the filled image to a buffer
        img_byte_arr = BytesIO()
        filled_image.save(img_byte_arr, format='PNG', dpi=(self.dpi, self.dpi)) # Save with target DPI
        img_byte_arr.seek(0)
        
        # Create PDF using ReportLab
        # Define PDF page size based on the image's dimensions at the chosen DPI
        # For A4, use: pagesize=A4
        
        output_filename = f"{candidate_srno}_{candidate_name.replace(' ', '_')}_filled.pdf"
        output_pdf_path = os.path.join(output_folder, output_filename)
        
        c = canvas.Canvas(output_pdf_path, pagesize=(self.page_width_pts, self.page_height_pts))
        
        # Draw the filled image onto the PDF canvas
        # The image is scaled to fit the page size
        c.drawImage(ImageReader(img_byte_arr), 0, 0, width=self.page_width_pts, height=self.page_height_pts)
        
        c.save()
        
        return output_pdf_path
