import pandas as pd
import os
import sys
from PIL import Image, ImageDraw, ImageFont
# Import smtplib for the actual sending function
import smtplib
# Here are the email package modules we'll need
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def create_card(row, template_path, output_folder):
    """
    Creates an individual membership card using data from an Excel row and saves it as a PNG.
    """
    try:
        # Load the image
        image = Image.open(template_path)

        # Prepare to draw on the image
        draw = ImageDraw.Draw(image)

        # Check if value is NaN (which is a float) using pandas.isna() or by converting to string if needed
        nombre = str(row['nombre']).strip() if 'nombre' in row and pd.notna(row['nombre']) else ''
        apellidos = str(row['apellidos']).strip() if 'apellidos' in row and pd.notna(row['apellidos']) else ''

        full_name = f"{nombre} {apellidos}".strip()
        member_number = f"{row['numero_socio']:04d}"  # Format the member number as a 6-digit number
        font_path = "arial.ttf"  # Replace with the path to your font file
        font_size = 28
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            font = ImageFont.load_default()  # Fallback to default font

        # Define text color (RGB)
        text_color = (255, 255, 255)  # White color

        # Draw the name on the image
        draw.text((170, 427), full_name, font=font, fill=text_color)
        # Draw the member number on the image
        draw.text((170, 467), member_number, font=font, fill=text_color)

        # Save the modified image
        output_path = os.path.join(output_folder, f"card_{row['numero_socio']}.png")
        image.save(output_path)

        return output_path
    
    except Exception as e:
        print(f"Error creating card: {str(e)}")
        raise  # Re-raise the exception if needed

def wellcome_card(row, template_path, output_folder):
    try:
        # Load the image
        image = Image.open(template_path)

        # Prepare to draw on the image
        draw = ImageDraw.Draw(image)

        # Define the text and font
        text = f"{row['nombre']}"
        font_path = "impact.ttf"  # Replace with the path to your font file
        font_size = 100
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
        x = (image_width - text_width) / 1.1
        y = (image_height - text_height) / 1.2

        # Define text color (RGB)
        text_color = (255, 255, 255)  # White color

        # Draw the text on the image
        draw.text((x, y), text, font=font, fill=text_color)

        # Save the modified image
        output_path = os.path.join(output_folder, f"wellcomecard_{row['numero_socio']}.png")
        image.save(output_path)

        return output_path
    
    except Exception as e:
        print(f"Error creating card: {str(e)}")
        raise  # Re-raise the exception if needed

def main():
    # Configuration
    excel_path = "Docs/sociosExample.xlsx"  # Path to your Excel file
    template_path = "Templates/dorso.png"  # Path to your template image
    wellcome_path = "Templates/wellcoming.png"  # Path to your template image
    output_folder = "SeasonTickets"  # Output folder for generated cards
    
    # Create output folder if it does not exist
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        # Read data from Excel
        df = pd.read_excel(excel_path)

        sys.stdout = open('logs/last_run.txt', 'w')
        
        # Process each member
        for index, row in df.iterrows():
            try:
                # Create the membership card
                card_path = create_card(row, template_path, output_folder)
                wellcome_output_path = wellcome_card(row, wellcome_path, output_folder)
                print(f"Successfully processed: {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']})")
                df.at[index, 'card_path'] = card_path
                # card_path = row['card_path']
                # sendEmail(row['correo'], card_path, wellcome_output_path)
            except Exception as e:
                print(f"Error processing {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']}): {str(e)}")
        df.to_excel(excel_path, index=False)
                
    except Exception as e:
        print(f"General error: {str(e)}")

def sendEmail(email, card_path, wellcome_path):
    # Create the container (outer) email message.
    msg = MIMEMultipart()
    msg['Subject'] = 'Your Harborough Town Season Ticket'
    msg['From'] = 'joeljuaristi90@gmail.com'
    
    # Check if email is a list or a single address
    if isinstance(email, list):
        msg['To'] = ', '.join(email)
        recipients = email
    else:
        msg['To'] = email
        recipients = [email]
        
    texto = MIMEText('¡Bienvenido a la colmena!\n\n¡Gracias por ser miembro LMI FC! Aquí tienes tu carnet de socio que te acredita para acceder a los partidos de local del Harborough Town, entre otras ventajas.\n\nAdjunto al mismo veréis una tarjeta personalizada con el nombre de cada uno que os animamos a compartir en redes como orgullosos socios ‘Bees’.\n\n¡Vamos a por el ascenso, abejorros!', 'plain', 'utf-8')
    msg.attach(texto)

    # Attach front of card
    front_path = r'Templates\frente.png'
    with open(front_path, 'rb') as fp:
        img = MIMEImage(fp.read())
    msg.attach(img)
    # Open the files in binary mode.
    with open(card_path, 'rb') as fp:
        img = MIMEImage(fp.read())
    msg.attach(img)

    # Attach wellcome card
    with open(wellcome_path, 'rb') as fp:
        img = MIMEImage(fp.read())
    msg.attach(img)


    # Get account password - read as text, not binary
    with open('gmailPass.txt', 'r') as fp:
        password = fp.read().strip()
        print(f"Password: {password}")

    # Send the email via our own SMTP server.
    s = smtplib.SMTP('smtp.gmail.com', 587) 
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], recipients, msg.as_string())
    s.quit()

if __name__ == "__main__":
    main()


