import os
import pickle
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Gmail API Scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def authenticate_gmail():
    """
    Authenticates the user and returns the Gmail API service.
    """
    creds = None
    # Load credentials from file if they exist
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If no valid credentials, prompt the user to log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    # Build the Gmail API service
    service = build('gmail', 'v1', credentials=creds)
    return service

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

def create_message(sender, to, subject, body, attachment_path=None):
    """
    Creates an email message with an optional attachment.
    """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Add the body of the email
    message.attach(MIMEText(body, 'plain'))

    # Add an attachment if provided
    if attachment_path:
        with open(attachment_path, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(attachment_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        message.attach(part)

    # Encode the message in base64
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw_message}

def send_email(service, user_id, message):
    """
    Sends an email using the Gmail API.
    """
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        print(f"Email sent: {message['id']}")
        return message
    except Exception as e:
        print(f"Error sending email: {str(e)}")
        raise e

def main():
    # Configuration
    excel_path = "Docs/socios.xlsx"  # Path to your Excel file
    template_path = "SeasonTickets/baseCard.png"  # Path to your template image
    output_folder = "SeasonTickets"  # Output folder for generated cards
    sender_email = "your-email@gmail.com"  # Replace with your Gmail address

    # Authenticate Gmail API
    service = authenticate_gmail()

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
                print(f"Successfully processed: {row['NOMBRE']} {row['APELLIDOS']} (Member No: {row['NUMERO_SOCIO']})")

                # Send the card via email
                to_email = row['CORREO']  # Ensure your Excel file has an 'email' column
                subject = "Your Membership Card"
                body = f"Dear {row['NOMBRE']} {row['APELLIDOS']},\n\nPlease find your membership card attached.\n\nBest regards,\nYour Organization"

                # Create and send the email
                message = create_message(sender_email, to_email, subject, body, card_path)
                send_email(service, "me", message)

            except Exception as e:
                print(f"Error processing {row['NOMBRE']} {row['APELLIDOS']} (Member No: {row['NUMERO_SOCIO']}): {str(e)}")
                
    except Exception as e:
        print(f"General error: {str(e)}")

if __name__ == "__main__":
    main()