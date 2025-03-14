import pandas as pd
import os
import sys
from PIL import Image, ImageDraw, ImageFont
# Import smtplib for the actual sending function
import smtplib
# Here are the email package modules we'll need
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

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
        full_name = f"{row['nombre']} {row['apellidos']}"
        member_number = f"{row['numero_socio']}"
        font_path = "arial.ttf"  # Replace with the path to your font file
        font_size = 15
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            font = ImageFont.load_default()  # Fallback to default font

        # Define text color (RGB)
        text_color = (255, 255, 255)  # White color

        # Draw the name on the image
        draw.text((302.2903225806451, 956.4285714285714), full_name, font=font, fill=text_color)
        # Draw the member number on the image
        draw.text((304.2903225806451, 983.4285714285714), member_number, font=font, fill=text_color)

        # Save the modified image
        output_path = os.path.join(output_folder, f"card_{row['numero_socio']}.png")
        image.save(output_path)

        return output_path
    
    except Exception as e:
        print(f"Error creating card: {str(e)}")
        raise  # Re-raise the exception if needed

def main():
    # Configuration
    excel_path = "Docs/sociosExample.xlsx"  # Path to your Excel file
    template_path = "Templates/winner.png"  # Path to your template image
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
                print(f"Successfully processed: {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']})")
                # sendEmail(row['correo'], card_path)
                df.at[index, 'card_path'] = card_path
            except Exception as e:
                print(f"Error processing {row['nombre']} {row['apellidos']} (Member No: {row['numero_socio']}): {str(e)}")
        df.to_excel(excel_path, index=False)
                
    except Exception as e:
        print(f"General error: {str(e)}")

def sendEmail(email, card_path):
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
        
    msg.preamble = 'Test email with attachment'

    # Open the files in binary mode.
    with open(card_path, 'rb') as fp:
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


