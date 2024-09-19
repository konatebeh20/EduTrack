# pip install sqlalchemy pymysql
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
        columns = ('id', 'Nom', 'Prénom', 'Matricule', 'Parcours')
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

    def importer_donnees(self):
        # Importer les fichiers Excel
        fichier_etudiants = filedialog.askopenfilename(title="Importer la liste des étudiants")
        fichier_notes = filedialog.askopenfilename(title="Importer les notes")
        
        # Charger les données
        df_etudiants = pd.read_excel(fichier_etudiants)
        df_notes = pd.read_excel(fichier_notes)
        
        # Traitement des étudiants
        for _, row in df_etudiants.iterrows():
            etudiant = Etudiant(nom=row['Nom'], prenom=row['Prénom'], date_naissance=row['Date de naissance'],
                                lieu_naissance=row['Lieu de naissance'], genre=row['Genre'],
                                matricule=row['Matricule'], parcours_id=row['Parcours ID'])
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

    def calculate_report_card(self, df, coefficients):
        df['Total'] = df.apply(lambda row: sum(row[course] * coefficients[course] for course in coefficients), axis=1)
        df['Moyenne'] = df['Total'] / sum(coefficients.values())
        return df

    def generate_excel_report(self, etudiant):
        # Code pour générer le rapport Excel pour l'étudiant
        pass

    def generate_pdf_report(self, etudiant):
        # Code pour générer le rapport PDF pour l'étudiant
        pass

    def send_email(self):
        # Envoyer l'email avec les fichiers joints
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = DashboardApp(root)
    root.mainloop()

