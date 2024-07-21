import pandas as pd
import numpy as np
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog


def read_excel_table(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        return df

    except Exception as e:
        print("An error occurred:", e)
        return None

def create_output_df(input_df):
    headers = ['Name_Wichtel', 'Name_Empfänger']
    data = {
        'Name_Wichtel': input_df['Name'].tolist(),
        'SecondColumn': [None] * input_df.shape[0]  # Empty column initially filled with None
    }
    output_df = pd.DataFrame(data, columns=headers) 
    return output_df

def send_email(sender_email, sender_password, receiver_email, subject, message):
    # Set up the MIME
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    try: 
        # Connect to the SMTP server
        server = smtplib.SMTP('mail.gmx.com', 587)  # For Gmail, use its SMTP server
        server.starttls()  # Start TLS encryption
        server.login(sender_email, sender_password)  # Login to your email account
        # Send the email
        server.sendmail(sender_email, receiver_email, msg.as_string())
    except Exception as e:
        print("An error occurred:", e)
        print("have you enabled IMAP/SMTP in your mail account?")
        return None
    # Quit the server
    server.quit()

def select_folder():
    # Create a root window
    root = tk.Tk()
    # Hide the root window
    root.withdraw()
    # Open the file dialog
    file_selected = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls *.csv")]
    )
    return file_selected

if __name__ == "__main__":
    
    path = select_folder()
    print(path)
    #create dataframes (df):
    input_df = read_excel_table(path)
    candidates_df = pd.DataFrame(input_df.Name)
    output_df = create_output_df(input_df)
    cnt = 0
    leave = False

    while not leave and cnt < 1000:
        cnt += 1
        del candidates_df
        del output_df
        candidates_df = pd.DataFrame(input_df.Name)
        output_df = create_output_df(input_df)

        for i in range(0,input_df.shape[0]):
            check = 0
            cnt2 = 0
            while (not check) and cnt2 < 1000:
                cnt2 += 1
                r = random.randint(0, candidates_df.shape[0]-1)
                opfer = candidates_df.Name.iloc[r]
                if (input_df[input_df.Name == opfer].Fam_ID.values != input_df.Fam_ID[i]) and (input_df[input_df.Name == opfer].Partner_ID.values != input_df.Partner_ID[i]):
                    check = 1
                    candidates_df = candidates_df.drop(candidates_df.index[r])
                    output_df.Name_Empfänger[i] = opfer
        if output_df['Name_Empfänger'].isna().any():
            leave = False
        else:
            leave = True
            
    print("cnt: ", cnt)
    output_df.to_csv('WichtelSpiel.csv', index=False)


    root = tk.Tk()
    root.withdraw()
    sender_email = simpledialog.askstring(title="Email", prompt="please Enter your email address")
    password = simpledialog.askstring(title="password", prompt="please Enter your password")
    root.destroy()
    subject = 'Dein Wichtel'

    for i in range (output_df.shape[0]):
        name = output_df.Name_Wichtel[i]
        receiver_email = input_df[input_df.Name == output_df.Name_Wichtel[i]].Email.values[0]
        opfer = output_df.Name_Empfänger[i]
        text = 'Hallihallo liebe/r ' + name + '\n\nDie (un)glückliche Person, die du beschenken darfst, heeeeisst.... ' + opfer + '!!! Überleg dir etwas Schönes und dann sehen wir uns am 25.12. bei Fabi!\n\nLiebe Grüsse,\nder Wichtelkönig' 
        print(name, receiver_email, opfer)
        #send_email(sender_email, password, receiver_email, subject, text) #uncomment to send email
