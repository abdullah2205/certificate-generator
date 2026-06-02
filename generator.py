from PIL import Image, ImageDraw, ImageFont
import os

def generate_certificate(template_path, output_path, data, font_path, font_size, text_coords):
    """
    Generates a certificate by overlaying text on a template image.
    """
    # Open the template
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)
    
    # Load font
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        font = ImageFont.load_default()

    # Overlay text
    for key, text in data.items():
        if key in text_coords:
            x, y = text_coords[key]
            
            # Hitung lebar teks untuk diletakkan persis di tengah (center)
            try:
                text_width = draw.textlength(text, font=font)
            except AttributeError:
                # Fallback jika mengunakan library Pillow versi lama
                text_width = font.getsize(text)[0]
                
            centered_x = x - (text_width / 2)
            draw.text((centered_x, y), text, fill="black", font=font)

    # Save as PDF
    image.save(output_path, "PDF", resolution=100.0)
    return output_path

# Example Usage (for testing)
if __name__ == "__main__":
    # Assuming we have a template image named 'template.png'
    # and a font file named 'arial.ttf'
    # This is just a structural test.
    print("Core logic ready.")
