import os
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str, font_size: int = 48):
        self.template_image = template_image
        self.mapping_data = mapping_data
        self.output_font_path = font_path
        self.output_font_size = font_size
        
        if self.template_image.mode != 'RGB':
            self.template_image = self.template_image.convert('RGB')
        
        self.dpi = 300
        self.page_width_pts = (self.template_image.width / self.dpi) * inch
        self.page_height_pts = (self.template_image.height / self.dpi) * inch
        
        try:
            self.pil_font = ImageFont.truetype(self.output_font_path, self.output_font_size)
        except IOError:
            print(f"FATAL: Font file not found at '{self.output_font_path}'. Please ensure it is in the assets folder.")
            self.pil_font = ImageFont.load_default()

    def _draw_text_on_image(self, draw: ImageDraw.Draw, text: str, x: int, y: int, w: int, h: int, color=(0, 0, 0)):
        text_bbox = draw.textbbox((0,0), text, font=self.pil_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        draw_y = y + (h - text_height) // 2
        
        if text_width > w:
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
            line_height = text_height
            start_y = y + (h - len(lines) * line_height) // 2
            for i, line in enumerate(lines):
                line_bbox = draw.textbbox((0,0), line, font=self.pil_font)
                line_width = line_bbox[2] - line_bbox[0]
                draw_x_line = x + (w - line_width) // 2
                draw_y_line = start_y + i * line_height
                if draw_y_line + line_height <= y + h:
                    draw.text((draw_x_line, draw_y_line), line, font=self.pil_font, fill=color)
        else:
            draw_x = x + (w - text_width) // 2
            draw.text((draw_x, draw_y), text, font=self.pil_font, fill=color)

    def fill_and_save_pdf(self, output_folder: str, candidate_data: dict, candidate_srno: str, candidate_name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)

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
                field_value = None
                for key, value in candidate_data.items():
                    if key.lower().replace("_", " ") == field_name.lower().replace("_", " "):
                        field_value = str(value)
                        break
                if field_value:
                    self._draw_text_on_image(draw, field_value, x, y, w, h, color=(0, 0, 0))

        img_byte_arr = BytesIO()
        filled_image.save(img_byte_arr, format='PNG', dpi=(self.dpi, self.dpi))
        img_byte_arr.seek(0)
        
        output_filename = f"{candidate_srno}_{candidate_name.replace(' ', '_')}_filled.pdf"
        output_pdf_path = os.path.join(output_folder, output_filename)
        
        c = canvas.Canvas(output_pdf_path, pagesize=(self.page_width_pts, self.page_height_pts))
        c.drawImage(ImageReader(img_byte_arr), 0, 0, width=self.page_width_pts, height=self.page_height_pts)
        c.save()
        
        return output_pdf_path
