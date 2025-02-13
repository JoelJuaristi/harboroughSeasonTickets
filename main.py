import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

def create_card(row, template_path, output_folder):
    """
    Creates an individual membership card using data from an Excel row and saves it as a PNG.
    """
    try:
        # Load the image
        image = Image.open(template_path)

        # Prepare to draw on the image
        draw = ImageDraw.Draw(image)

        # Define the text and font
        text = f"{row['nombre']} {row['apellidos']}\nMember number: {row['numero_socio']}"
        font_path = "arial.ttf"  # Replace with the path to your font file
        font_size = 20
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            font = ImageFont.load_default()  # Fallback to default font

        # Calculate the position to center the text
        # Use textbbox to get the bounding box of the text
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]  # Calculate text width
        text_height = bbox[3] - bbox[1]  # Calculate text height

        image_width, image_height = image.size
        x = (image_width - text_width) / 2
        y = (image_height - text_height) / 2

        # Define text color (RGB)
        text_color = (255, 255, 255)  # White color

        # Draw the text on the image
        draw.text((x, y), text, font=font, fill=text_color)

        # Save the modified image
        output_path = os.path.join(output_folder, f"card_{row['numero_socio']}.png")
        image.save(output_path)

        return output_path
    
    except Exception as e:
        print(f"Error creating card: {str(e)}")
        raise  # Re-raise the exception if needed

def main():
    # Configuration
    excel_path = "Docs/socios.xlsx"  # Path to your Excel file
    template_path = "Templates/baseCard.png"  # Path to your template image
    output_folder = "SeasonTickets"  # Output folder for generated cards
    
    # Create output folder if it does not exist
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        # Read data from Excel
        df = pd.read_excel(excel_path)
        
        # Process each member
        for index, row in df.iterrows():
            try:
                # Create the membership card
                card_path = create_card(row, template_path, output_folder)
                print(f"Successfully processed: {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']})")
            except Exception as e:
                print(f"Error processing {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']}): {str(e)}")
                
    except Exception as e:
        print(f"General error: {str(e)}")

if __name__ == "__main__":
    main()