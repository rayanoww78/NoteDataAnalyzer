import imaplib
import email
import os
import re
import fitz  # PyMuPDF pour lire les PDFs
import pandas as pd
import matplotlib.pyplot as plt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


EMAIL_HOST = "mail.uphf.fr"
EMAIL_PORT = 993
EMAIL_USER = "TON_MAIL_UPHF"
EMAIL_PASS = "TON_PROPRE_MDP"


SMTP_HOST = "smtp.uphf.fr"
SMTP_PORT = 587
SEND_TO = "EMAIL_RECEPTION"


SAVE_FOLDER = r"CHEMIN_TELECHARGEMENTS_PDF_NOTES"
os.makedirs(SAVE_FOLDER, exist_ok=True)


# 📥 Télécharger les relevés de notes (PDF) depuis Zimbra
def download_pdfs_from_email():
    """Télécharge les PDFs des relevés de notes reçus par mail."""
    mail = imaplib.IMAP4_SSL(EMAIL_HOST, EMAIL_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")


    result, data = mail.search(None, '(SUBJECT "de notes du semestre 1" SEEN)')
    mail_ids = data[0].split()

    for num in mail_ids:
        result, data = mail.fetch(num, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        for part in msg.walk():
            if part.get_content_subtype() == "pdf":
                filename = part.get_filename()
                filepath = os.path.join(SAVE_FOLDER, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"📥 Fichier téléchargé : {filename}")

    mail.logout()


def extract_notes_from_pdf(pdf_path):
    """Extrait les notes d'un bulletin de notes PDF."""
    doc = fitz.open(pdf_path)
    extracted_text = ""


    for page in doc:
        extracted_text += page.get_text() + "\n"

    doc.close()


    ue_pattern = re.compile(r"(UE\d\.\d\s-\sC\d\s.+?)\s(\d+\.\d+)")
    matieres_pattern = re.compile(r"(R\d\.\d{2}\s.+?)\s(\d+\.\d+)")

    data = []
    filename = os.path.basename(pdf_path)


    date_pattern = re.compile(r"(\d{4}-\d{2}-\d{2})")
    date_match = date_pattern.search(filename)
    date = date_match.group(1) if date_match else None


    for ue_name, ue_note in ue_pattern.findall(extracted_text):
        ue_note = float(ue_note)
        if ue_note <= 20:
            data.append({"Fichier": filename, "Date": date, "Type": "UE", "Nom": ue_name.strip(), "Note": ue_note})

   
    for matiere_name, matiere_note in matieres_pattern.findall(extracted_text):
        matiere_note = float(matiere_note)
        if matiere_note <= 20:
            data.append({"Fichier": filename, "Date": date, "Type": "Matière", "Nom": matiere_name.strip(), "Note": matiere_note})

    return data


def process_all_pdfs_in_folder(folder_path):
    """Scanne tous les PDF du dossier et extrait les notes."""
    all_data = []


    if not os.path.exists(folder_path):
        print("🚨 Le dossier spécifié n'existe pas !")
        return None


    for pdf_file in os.listdir(folder_path):
        if pdf_file.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, pdf_file)
            print(f"📄 Traitement de : {pdf_file}")
            all_data.extend(extract_notes_from_pdf(pdf_path))


    df = pd.DataFrame(all_data)


    df = df[df['Note'] <= 20]

    if df.empty:
        print("⚠️ Aucun relevé de notes valide trouvé.")
        return None


    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.sort_values(by='Date')

    return df



def plot_grades_by_ue(df):
    """Crée un graphe montrant l'évolution des moyennes par UE au fil du temps."""
    df_ue = df[df['Type'] == 'UE'].groupby(['Date', 'Nom'])['Note'].mean().reset_index()

    plt.figure(figsize=(10, 5))
    for ue in df_ue['Nom'].unique():
        df_ue_ue = df_ue[df_ue['Nom'] == ue]
        plt.plot(df_ue_ue['Date'], df_ue_ue['Note'], marker="o", linestyle="-", label=ue)

    plt.xlabel("Date")
    plt.ylabel("Moyenne par UE")
    plt.title("📊 Évolution des moyennes par UE")
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid()


    graph_path_ue = "evolution_moyennes_ue.png"
    plt.savefig(graph_path_ue)
    plt.show()

    return graph_path_ue


def plot_grades_by_matiere(df):
    """Crée un graphe montrant l'évolution des moyennes par matière au fil du temps."""
    df_matiere = df[df['Type'] == 'Matière'].groupby(['Date', 'Nom'])['Note'].mean().reset_index()

    plt.figure(figsize=(10, 5))
    for matiere in df_matiere['Nom'].unique():
        df_matiere_matiere = df_matiere[df_matiere['Nom'] == matiere]
        plt.plot(df_matiere_matiere['Date'], df_matiere_matiere['Note'], marker="x", linestyle="--", label=matiere)

    plt.xlabel("Date")
    plt.ylabel("Moyenne par Matière")
    plt.title("📊 Évolution des moyennes par Matière")
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid()


    graph_path_matiere = "evolution_moyennes_matiere.png"
    plt.savefig(graph_path_matiere)
    plt.show()

    return graph_path_matiere




def send_email_with_report(csv_path, graph_path_ue, graph_path_matiere, recipient_email):
    """Envoie un rapport par email avec les fichiers CSV et graphiques en pièce jointe."""


    GMAIL_USER = "EMAIL_D'ENVOIE"
    GMAIL_PASS = "bqkw urbp ifxz bhrr"  #
    SEND_TO = recipient_email

    SMTP_HOST = "smtp.gmail.com"
    SMTP_PORT = 587


    msg = MIMEMultipart()
    msg["From"] = GMAIL_USER
    msg["To"] = SEND_TO
    msg["Subject"] = "📊 Rapport d'évolution des notes"


    body = "Bonjour,\n\nVoici l'évolution de tes notes au fil de ton semestre en pièce jointe ! 📈\n\nBonne analyse !"
    msg.attach(MIMEText(body, "plain"))


    with open(graph_path_ue, "rb") as f:
        attach_ue = email.mime.base.MIMEBase("application", "octet-stream")
        attach_ue.set_payload(f.read())
        email.encoders.encode_base64(attach_ue)
        attach_ue.add_header("Content-Disposition", f"attachment; filename={os.path.basename(graph_path_ue)}")
        msg.attach(attach_ue)


    with open(graph_path_matiere, "rb") as f:
        attach_matiere = email.mime.base.MIMEBase("application", "octet-stream")
        attach_matiere.set_payload(f.read())
        email.encoders.encode_base64(attach_matiere)
        attach_matiere.add_header("Content-Disposition", f"attachment; filename={os.path.basename(graph_path_matiere)}")
        msg.attach(attach_matiere)

    try:

        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, SEND_TO, msg.as_string())
        server.quit()
        print("✅ Rapport envoyé par email !")
    except Exception as e:
        print(f"❌ Erreur lors de l'envoi de l'email : {e}")


print("📨 Récupération des relevés de notes...")
download_pdfs_from_email()

print("📊 Extraction des notes...")
df_notes = process_all_pdfs_in_folder(SAVE_FOLDER)

if df_notes is not None:
    csv_path = "relevés_notes.csv"
    df_notes.to_csv(csv_path, index=False)
    print(f"📂 Notes sauvegardées dans {csv_path}")

    print("📈 Génération des graphes...")
    graph_path_ue = plot_grades_by_ue(df_notes)
    graph_path_matiere = plot_grades_by_matiere(df_notes)

    print("📤 Envoi du rapport par email...")
    send_email_with_report(csv_path, graph_path_ue, graph_path_matiere,"elamjadrayan@gmail.com")
else:
    print("⚠️ Aucune donnée à traiter.")
