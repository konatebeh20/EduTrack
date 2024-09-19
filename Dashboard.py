# pip install sqlalchemy pymysql
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from sqlalchemy import create_engine, Column, Integer, String, ForeignKey, Enum, Date, DECIMAL
from sqlalchemy.orm import relationship, sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# SQLAlchemy configuration for MySQL
DATABASE_URL = "mysql+pymysql://username:password@localhost/universite_db"
engine = create_engine(DATABASE_URL, echo=True)
Session = sessionmaker(bind=engine)
session = Session()
Base = declarative_base()

# Configuration pour l'envoi des emails
DIRECTOR_EMAIL = 'directeur@universite.fr'
SMTP_SERVER = 'smtp.universite.fr'
SMTP_PORT = 587
SMTP_USER = 'votre.email@universite.fr'
SMTP_PASSWORD = 'votre_mot_de_passe'

# Define models
class Mention(Base):
    __tablename__ = 'mentions'
    id = Column(Integer, primary_key=True)
    libelle = Column(String(50), unique=True, nullable=False)

class Domaine(Base):
    __tablename__ = 'domaines'
    id = Column(Integer, primary_key=True)
    libelle = Column(String(100), unique=True, nullable=False)

class Parcours(Base):
    __tablename__ = 'parcours'
    id = Column(Integer, primary_key=True)
    libelle = Column(String(100), unique=True, nullable=False)
    domaine_id = Column(Integer, ForeignKey('domaines.id'))
    domaine = relationship('Domaine')

class Etudiant(Base):
    __tablename__ = 'etudiants'
    id = Column(Integer, primary_key=True)
    nom = Column(String(50), nullable=False)
    prenom = Column(String(50), nullable=False)
    date_naissance = Column(Date)
    lieu_naissance = Column(String(100))
    genre = Column(Enum('Masculin', 'Féminin'))
    matricule = Column(String(20), unique=True)
    parcours_id = Column(Integer, ForeignKey('parcours.id'))
    parcours = relationship('Parcours')
    annees_academiques = relationship('AnneeAcademique', back_populates='etudiant')

class AnneeAcademique(Base):
    __tablename__ = 'annees_academiques'
    id = Column(Integer, primary_key=True)
    annee = Column(String(9), unique=True, nullable=False)
    semestre1_id = Column(Integer, ForeignKey('semestres.id'))
    semestre2_id = Column(Integer, ForeignKey('semestres.id'))
    semestre1 = relationship('Semestre', foreign_keys=[semestre1_id])
    semestre2 = relationship('Semestre', foreign_keys=[semestre2_id])
    etudiant_id = Column(Integer, ForeignKey('etudiants.id'))
    etudiant = relationship('Etudiant', back_populates='annees_academiques')

class Semestre(Base):
    __tablename__ = 'semestres'
    id = Column(Integer, primary_key=True)
    numero = Column(Integer, nullable=False)
    annee_academique_id = Column(Integer, ForeignKey('annees_academiques.id'))
    ues = relationship('UE', back_populates='semestre')

class UE(Base):
    __tablename__ = 'unites_enseignement'
    id = Column(Integer, primary_key=True)
    code = Column(String(10), unique=True, nullable=False)
    libelle = Column(String(100), nullable=False)
    credits = Column(Integer, nullable=False)
    notes = relationship('Note', back_populates='ue')

class Note(Base):
    __tablename__ = 'notes'
    id = Column(Integer, primary_key=True)
    etudiant_id = Column(Integer, ForeignKey('etudiants.id'))
    ue_id = Column(Integer, ForeignKey('unites_enseignement.id'))
    semestre_id = Column(Integer, ForeignKey('semestres.id'))
    note = Column(DECIMAL(4,2))
    mention_id = Column(Integer, ForeignKey('mentions.id'))
    etudiant = relationship('Etudiant')
    ue = relationship('UE', back_populates='notes')
    semestre = relationship('Semestre')
    mention = relationship('Mention')

Base.metadata.create_all(engine)

# Classe pour l'application Tkinter
class DashboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Dashboard Étudiants")
        self.root.geometry("800x600")
        
        # Set icon
        self.root.iconphoto(True, tk.PhotoImage(file='images/icon.png'))
        
        # Frame pour le tableau des étudiants
        self.frame_table = tk.Frame(self.root)
        self.frame_table.pack(fill=tk.BOTH, expand=True)
        
        # Créer une table pour afficher les étudiants
        columns = ('id', 'Nom', 'Prénom', 'Matricule', 'Parcours', 'Anonymat', 'Groupe TD')
        self.tree = ttk.Treeview(self.frame_table, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Frame pour les boutons
        self.frame_buttons = tk.Frame(self.root)
        self.frame_buttons.pack(fill=tk.X)
        
        self.btn_importer = tk.Button(self.frame_buttons, text="Importer données", command=self.importer_donnees)
        self.btn_importer.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.btn_generer_excel = tk.Button(self.frame_buttons, text="Générer Bulletin Excel", command=self.generer_bulletin_excel)
        self.btn_generer_excel.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.btn_generer_pdf = tk.Button(self.frame_buttons, text="Générer Bulletin PDF", command=self.generer_bulletin_pdf)
        self.btn_generer_pdf.pack(side=tk.LEFT, padx=10, pady=10)

        self.btn_envoyer_email = tk.Button(self.frame_buttons, text="Envoyer Email", command=self.envoyer_bulletin)
        self.btn_envoyer_email.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.status_label = tk.Label(self.frame_buttons, text="")
        self.status_label.pack(pady=10)

        self.file_path = None
        self.etudiants = []
        self.notes = []
        self.resultats = []

    # def importer_donnees(self):
    #     # Importer les fichiers Excel
    #     fichier_etudiants = filedialog.askopenfilename(title="Importer la liste des étudiants")
    #     fichier_notes = filedialog.askopenfilename(title="Importer les notes")
        
    #     # Charger les données
    #     df_etudiants = pd.read_excel(fichier_etudiants)
    #     df_notes = pd.read_excel(fichier_notes)
        
    #     # Traitement des étudiants
    #     for _, row in df_etudiants.iterrows():
    #         etudiant = Etudiant(nom=row['Nom'], prenom=row['Prénom'], date_naissance=row['Date de naissance'],
    #                             lieu_naissance=row['Lieu de naissance'], genre=row['Genre'],
    #                             matricule=row['Matricule'], parcours_id=row['Parcours ID'])
    #         session.add(etudiant)
        
    #     session.commit()
        
    #     # Traitement des notes
    #     for _, row in df_notes.iterrows():
    #         note = Note(
    #             etudiant_id=row['Etudiant ID'],
    #             ue_id=row['UE ID'],
    #             semestre_id=row['Semestre ID'],
    #             note=row['Note'],
    #             mention_id=row['Mention ID']
    #         )
    #         session.add(note)
        
    #     session.commit()
        
    #     self.load_students()
    #     self.status_label.config(text="Les données ont été importées avec succès.")

# Main function
def main():
    # Assurez-vous que le répertoire de sortie existe
    output_dir = 'bulletins'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Lire les données et les coefficients
    grades_file = 'notes.xlsx'
    coefficients = {'Mathematiques': 3, 'Statistiques': 2, 'Economie': 2}  # Exemple de coefficients
    df = read_grades(grades_file)
    
    # Calculer le bulletin
    df = calculate_report_card(df, coefficients)

    # Générer les fichiers Excel et PDF
    for student_name in df['Nom'].unique():
        student_data = df[df['Nom'] == student_name]
        generate_excel_report(student_name, student_data, output_dir)
        generate_pdf_report(student_name, student_data, output_dir)
        
        # Envoyer un email à l'étudiant
        send_email(
            to_email=f'{student_name}@universite.fr',
            subject='Votre bulletin est prêt',
            body=f'Bonjour {student_name},\n\nVotre bulletin est prêt. Vous pouvez le récupérer dans une semaine à l\'adresse suivante : [adresse].\n\nCordialement,\nL\'Université',
            attachments=[
                os.path.join(output_dir, f'{student_name}.xlsx'),
                os.path.join(output_dir, f'{student_name}.pdf')
            ]
        )

# Lire les fichiers Excel
def read_grades(file_path):
    return pd.read_excel(file_path)

def importer_donnees(self):
    # Importer les fichiers Excel
    fichier_presence = filedialog.askopenfilename(title="Importer la liste de présence")
    fichier_etudiants = filedialog.askopenfilename(title="Importer la liste des étudiants", filetypes=[("Excel files", "*.xlsx")])
    fichier_notes = filedialog.askopenfilename(title="Importer les notes", filetypes=[("Excel files", "*.xlsx")])
    fichier_resultats = filedialog.askopenfilename(title="Importer les résultats finaux")


    if not fichier_etudiants or not fichier_notes:
        self.status_label.config(text="Veuillez sélectionner les fichiers à importer.")
        return

     # Création des étudiants
        self.etudiants = []
        for idx, row in df_presence.iterrows():
            etudiant = {'id': idx, 'nom': row['NOM ET PRENOM(S)'].split()[0], 
                        'prenom': ' '.join(row['NOM ET PRENOM(S)'].split()[1:]),
                        'numero_anonymat': row['ANONYMAT'],
                        'groupe_td': row['GPE TD']}
            self.etudiants.append(etudiant)
            self.tree.insert('', 'end', values=(etudiant['id'], etudiant['nom'], etudiant['prenom'], etudiant['numero_anonymat'], etudiant['groupe_td']))

        self.status_label.config(text="Les données ont été importées avec succès.")
    
    # Charger les données
    df_etudiants = pd.read_excel(fichier_etudiants)
    df_notes = pd.read_excel(fichier_notes)
    df_presence = pd.read_excel(fichier_presence)
    df_resultats = pd.read_excel(fichier_resultats)
    
    # Traitement des étudiants
    for _, row in df_etudiants.iterrows():
        etudiant = Etudiant(
            nom=row['Nom'],
            prenom=row['Prénom'],
            date_naissance=row['Date de naissance'],
            lieu_naissance=row['Lieu de naissance'],
            genre=row['Genre'],
            matricule=row['Matricule'],
            parcours_id=row['Parcours ID']
        )
        session.add(etudiant)
    
    session.commit()
    
    # Traitement des notes
    for _, row in df_notes.iterrows():
        note = Note(
            etudiant_id=row['Etudiant ID'],
            ue_id=row['UE ID'],
            semestre_id=row['Semestre ID'],
            note=row['Note'],
            mention_id=row['Mention ID']
        )
        session.add(note)
    
    session.commit()
    
    self.load_students()
    self.status_label.config(text="Les données ont été importées avec succès.") 

def load_students(self):
        self.tree.delete(*self.tree.get_children())
        etudiants = session.query(Etudiant).all()
        for etudiant in etudiants:
            self.tree.insert('', 'end', values=(etudiant.id, etudiant.nom, etudiant.prenom, etudiant.matricule, etudiant.parcours.libelle))

 def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.status_label.config(text=f"Fichier chargé : {os.path.basename(self.file_path)}")
        else:
            self.status_label.config(text="Aucun fichier sélectionné")
            
def calculate_report_card(self, df, coefficients):
        df['Total'] = df.apply(lambda row: sum(row[course] * coefficients[course] for course in coefficients), axis=1)
        df['Moyenne'] = df['Total'] / sum(coefficients.values())
        return df


    # def generate_excel_report(self, etudiant):
    #     # Code pour générer le rapport Excel pour l'étudiant
    #     pass
    
def generate_excel_report(self, etudiant):
    file_path = os.path.join(output_dir, f'{student_name}.xlsx')

    df = pd.DataFrame([{
        'Nom': etudiant.nom,
        'Prénom': etudiant.prenom,
        'Matricule': etudiant.matricule,
        # Ajoutez d'autres informations si nécessaire
    }])
    filename = f"rapport_{etudiant.matricule}.xlsx"
    data.to_excel(file_path, index=False)
    df.to_excel(filename, index=False)
    self.status_label.config(text=f"Rapport Excel généré : {filename}")


    # def generate_pdf_report(self, etudiant):
    # # Code pour générer le rapport PDF pour l'étudiant
    #    pass

def generer_bulletin_excel(self):
        selected_item = self.tree.selection()
        if selected_item:
            etudiant_id = self.tree.item(selected_item)['values'][0]
            etudiant = next(e for e in self.etudiants if e['id'] == etudiant_id)
            
            # Simulation des données pour l'exemple
            output_dir = 'bulletins'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            student_data = pd.DataFrame([{'Matiere': 'Mathématiques', 'Note': 15}, {'Matiere': 'Statistiques', 'Note': 14}])
            self.generate_excel_report(f"{etudiant['nom']}_{etudiant['prenom']}", student_data, output_dir)
            messagebox.showinfo("Succès", f"Bulletin Excel généré: {etudiant['nom']}_{etudiant['prenom']}_bulletin.xlsx")
        else:
            messagebox.showwarning("Erreur", "Veuillez sélectionner un étudiant.")
            
def generate_pdf_report(self, etudiant):
    file_path = os.path.join(output_dir, f'{student_name}.pdf')

    c = canvas.Canvas(file_path, pagesize=letter)
    c.drawString(100, 750, f'Bulletin de {student_name}')
    y_position = 700

    filename = f"rapport_{etudiant.matricule}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    
    content = []
    content.append(Paragraph(f"Nom: {etudiant.nom}", styles['Normal']))
    content.append(Paragraph(f"Prénom: {etudiant.prenom}", styles['Normal']))
    content.append(Paragraph(f"Matricule: {etudiant.matricule}", styles['Normal']))
    # Ajoutez d'autres informations si nécessaire

    doc.build(content)
    self.status_label.config(text=f"Rapport PDF généré : {filename}")

    for index, row in data.iterrows():
        c.drawString(100, y_position, f'{row["Matiere"]}: {row["Note"]}')
        y_position -= 20
    c.save()

def send_email(self):
    sender_email = "your_email@example.com"
    receiver_email = "receiver_email@example.com"
    subject = "Bulletin d'Étudiant"
    body = "Veuillez trouver ci-joint le bulletin de l'étudiant."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    # Ajouter des fichiers attachés
    filenames = ["path/to/file1.xlsx", "path/to/file2.pdf"]
    for filename in filenames:
        with open(filename, "rb") as attachment:
            part = MIMEApplication(attachment.read(), Name=os.path.basename(filename))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filename)}"'
            msg.attach(part)

    # def send_email(self):
    # # Envoyer l'email avec les fichiers joints
    #   pass
         
    # Envoyer l'email
    with smtplib.SMTP('smtp.example.com', 587) as server:
        server.starttls()
        server.login("your_email@example.com", "your_password")
        server.sendmail(sender_email, receiver_email, msg.as_string())
    
    self.status_label.config(text="Email envoyé avec succès.")

    for file in attachments:
        with open(file, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file)}"'
            msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
            
def generer_bulletin_pdf(self):
        selected_item = self.tree.selection()
        if selected_item:
            etudiant_id = self.tree.item(selected_item)['values'][0]
            etudiant = next(e for e in self.etudiants if e['id'] == etudiant_id)
            
            # Simulation des données pour l'exemple
            output_dir = 'bulletins'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            student_data = pd.DataFrame([{'Matiere': 'Mathématiques', 'Note': 15}, {'Matiere': 'Statistiques', 'Note': 14}])
            self.generate_pdf_report(f"{etudiant['nom']}_{etudiant['prenom']}", student_data, output_dir)
            messagebox.showinfo("Succès", f"Bulletin PDF généré: {etudiant['nom']}_{etudiant['prenom']}_bulletin.pdf")
        else:
            messagebox.showwarning("Erreur", "Veuillez sélectionner un étudiant.")

def envoyer_bulletin(self):
        selected_item = self.tree.selection()
        if selected_item:
            etudiant_id = self.tree.item(selected_item)['values'][0]
            etudiant = next(e for e in self.etudiants if e['id'] == etudiant_id)
            
            output_dir = 'bulletins'
            excel_file = os.path.join(output_dir, f'{etudiant["nom"]}_{etudiant["prenom"]}_bulletin.xlsx')
            pdf_file = os.path.join(output_dir, f'{etudiant["nom"]}_{etudiant["prenom"]}_bulletin.pdf')
            
            # Envoyer un email à l'étudiant
            self.send_email(
                to_email=f'{etudiant["nom"].lower()}.{etudiant["prenom"].lower()}@universite.fr',
                subject='Votre bulletin',
                body=f'Bonjour {etudiant["prenom"]},\n\nVeuillez trouver ci-joint votre bulletin.',
                attachments=[excel_file, pdf_file]
            )
            messagebox.showinfo("Succès", "Les bulletins ont été envoyés par email.")
        else:
            messagebox.showwarning("Erreur", "Veuillez sélectionner un étudiant.")

def process_data(self):
        if not self.file_path:
            messagebox.showwarning("Avertissement", "Veuillez d'abord charger un fichier Excel.")
            return

        output_dir = 'bulletins'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Lire les données et les coefficients
        coefficients = {'Mathematiques': 3, 'Statistiques': 2, 'Economie': 2}  # Exemple de coefficients
        df = pd.read_excel(self.file_path)
        
        # Calculer le bulletin
        df = self.calculate_report_card(df, coefficients)

        # Générer les fichiers Excel et PDF
        for student_name in df['Nom'].unique():
            student_data = df[df['Nom'] == student_name]
            self.generate_excel_report(student_name, student_data, output_dir)
            self.generate_pdf_report(student_name, student_data, output_dir)
            
            # Envoyer un email à l'étudiant
            self.send_email(
                to_email=f'{student_name}@universite.fr',
                subject='Votre bulletin est prêt',
                body=f'Bonjour {student_name},\n\nVotre bulletin est prêt. Vous pouvez le récupérer dans une semaine à l\'adresse suivante : [adresse].\n\nCordialement,\nL\'Université',
                attachments=[
                    os.path.join(output_dir, f'{student_name}.xlsx'),
                    os.path.join(output_dir, f'{student_name}.pdf')
                ]
            )
        
        # Envoyer un email au Directeur des Études
        self.send_email(
            to_email=DIRECTOR_EMAIL,
            subject='Bulletins des étudiants',
            body='Les bulletins des étudiants ont été générés et envoyés.',
            attachments=[
                os.path.join(output_dir, f'{student_name}.xlsx'),
                os.path.join(output_dir, f'{student_name}.pdf')
            ]
        )

        self.status_label.config(text="Traitement terminé. Les bulletins ont été envoyés.")


if __name__ == "__main__":
    root = tk.Tk()
    app = DashboardApp(root)
    root.mainloop()

