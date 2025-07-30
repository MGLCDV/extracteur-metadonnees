import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader
from datetime import datetime
import exifread
import zipfile
import rarfile
import py7zr
import openpyxl
from pptx import Presentation
from docx import Document
from mutagen import File
from pymediainfo import MediaInfo
from tkinter import messagebox
from email import policy
from email.parser import BytesParser

import os


class App:
    META_INFOS = {
        ".pdf": ["Titre", "Auteur", "Sujet", "Producteur", "Créé le", "Modifié le"],
        ".docx": ["Titre", "Auteur", "Sujet", "Commentaires", "Créé le", "Modifié le"],
        ".xlsx": ["Titre", "Sujet", "Auteur", "Commentaires", "Catégorie", "Identifiant", "Langue", "Créé le", "Modifié le"],
        ".pptx": ["Titre", "Sujet", "Auteur", "Commentaires", "Catégorie", "Créé le", "Modifié le", "Dernier modifié par"],
        ".jpg": ["Date prise", "Modèle appareil", "Logiciel", "GPS", "Auteur"],
        ".jpeg": ["Date prise", "Modèle appareil", "Logiciel", "GPS", "Auteur"],
        ".png": ["Date prise", "Logiciel"],
        ".tiff": ["Date prise", "Modèle appareil", "Logiciel", "GPS", "Auteur"],
        ".mp3": ["Titre", "Artiste", "Album", "Année", "Commentaires"],
        ".wav": ["Titre", "Artiste", "Album", "Année", "Commentaires"],
        ".mp4": ["Durée", "Codec vidéo", "Codec audio", "Résolution", "Date création"],
        ".mov": ["Durée", "Codec vidéo", "Codec audio", "Résolution", "Date création"],
        ".avi": ["Durée", "Codec vidéo", "Codec audio", "Résolution", "Date création"],
        ".zip": ["Nom fichiers", "Taille compressée", "Taille décompressée", "Date modification"],
        ".rar": ["Nom fichiers", "Taille compressée", "Taille décompressée", "Date modification"],
        ".7z": ["Nom fichiers", "Taille compressée", "Taille décompressée", "Date modification"]
    }

    
    
    def __init__(self, master):
        self.master = master
        self.master.title("Extracteur de Métadonnées")
        self.master.geometry("1080x720")

        self.chemin_fichier = tk.StringVar()

        self.creer_widgets()

    def creer_widgets(self):
        btn_choisir = tk.Button(self.master, text="Choisir un fichier", command=self.choisir_fichier, font=("Arial", 12))
        btn_choisir.pack(pady=15)

        lbl_fichier = tk.Label(self.master, textvariable=self.chemin_fichier, wraplength=1000, font=("Arial", 11))
        lbl_fichier.pack(pady=10)

        self.txt_metadonnees = tk.Text(self.master, height=25, width=120, font=("Consolas", 11))
        self.txt_metadonnees.pack(pady=15)
        self.txt_metadonnees.config(state=tk.DISABLED)

        btn_copier = tk.Button(self.master, text="Copier dans le presse-papier", command=self.copier_metadonnees, font=("Arial", 11))
        btn_copier.pack(pady=8)

        btn_analyser = tk.Button(self.master, text="Analyser", command=self.analyser, font=("Arial", 16, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10)
        btn_analyser.pack(pady=20)


    def choisir_fichier(self):
        fichier = filedialog.askopenfilename(
            title="Choisir un fichier",
            filetypes=[
                        ("Fichiers supportés", "*.pdf *.jpg *.jpeg *.png *.tiff *.docx *.xlsx *.pptx *.mp3 *.wav *.mp4 *.mov *.avi *.zip *.rar *.7z *.eml"),
                        ("Fichiers PDF", "*.pdf"),
                        ("Images", "*.jpg *.jpeg *.png *.tiff"),
                        ("Documents Word", "*.docx"),
                        ("Fichiers Excel", "*.xlsx"),
                        ("Fichiers PowerPoint", "*.pptx"),
                        ("Fichiers vidéo", "*.mp4 *.mov *.avi"),
                        ("Archives", "*.zip *.rar *.7z"),
                        ("Emails" , "*.eml"),
                        ("Tous les fichiers", "*.*"),
                    ]

        )
        if fichier:
            self.chemin_fichier.set(fichier)
            ext = os.path.splitext(fichier)[1].lower()
            infos = self.META_INFOS.get(ext, ["Aucune information disponible pour ce type de fichier."])
            self.afficher_infos(infos)
            
    def copier_metadonnees(self):
        self.master.clipboard_clear()
        texte = self.txt_metadonnees.get("1.0", tk.END).strip()
        self.master.clipboard_append(texte)
        self.master.update()

    
    def afficher_infos(self, infos):
        self.txt_metadonnees.config(state=tk.NORMAL)
        self.txt_metadonnees.delete(1.0, tk.END)
        self.txt_metadonnees.insert(tk.END, "Métadonnées que le programme va tenter de récupérer :\n\n")
        for info in infos:
            self.txt_metadonnees.insert(tk.END, f" - {info}\n")
        self.txt_metadonnees.config(state=tk.DISABLED)
    
    def afficher_resultat(self, texte):
        if texte is None:
            texte = "Aucun résultat à afficher."
        else:
            texte = str(texte)
        self.txt_metadonnees.config(state=tk.NORMAL)
        self.txt_metadonnees.delete(1.0, tk.END)
        self.txt_metadonnees.insert(tk.END, texte)
        self.txt_metadonnees.config(state=tk.DISABLED)
    
    def analyser(self):
        fichier = self.chemin_fichier.get()
        if not fichier:
            messagebox.showwarning("Attention", "Veuillez sélectionner un fichier.")
            return
        
        ext = os.path.splitext(fichier)[1].lower()
        
        resultat = ""
        if ext == ".pdf":
            resultat = self.extraire_metadonnees_pdf(fichier)
        elif ext in [".jpg", ".jpeg", ".png", ".tiff"]:
            resultat = self.extraire_metadonnees_img(fichier)
        elif ext == ".docx":
            resultat = self.extraire_metadonnees_docx(fichier)
        elif ext == ".xlsx":
            resultat = self.extraire_metadonnees_xlsx(fichier)
        elif ext == ".pptx":
            resultat = self.extraire_metadonnees_pptx(fichier)
        elif ext in [".mp3", ".wav"]:
            resultat = self.extraire_metadonnees_audio(fichier)
        elif ext in [".zip", ".rar", ".7z"]:
            resultat = self.extraire_metadonnees_archive(fichier)
        elif ext == ".eml":
            resultat = self.extraire_metadonnees_eml(fichier)
        else:
            resultat = "Format non pris en charge pour l'instant."
        
        self.afficher_resultat(resultat)

    def extraire_metadonnees_pdf(self, chemin):
        try:
            lecteur = PdfReader(chemin)
            infos = lecteur.metadata

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n\n"
            texte += "-"*40 + "\n"

            def afficher(cle, valeur):
                nonlocal texte
                texte += f"{cle:<20}: {valeur}\n"

            # Accès individuel avec .get()
            titre = infos.get("/Title", "Non défini")
            auteur = infos.get("/Author", "Non défini")
            sujet = infos.get("/Subject", "Non défini")
            producteur = infos.get("/Producer", "Non défini")
            date_creation = infos.get("/CreationDate", None)
            date_modif = infos.get("/ModDate", None)

            afficher("Titre", titre)
            afficher("Auteur", auteur)
            afficher("Sujet", sujet)
            afficher("Producteur", producteur)

            if date_creation:
                afficher("Créé le", self.formater_date_pdf(date_creation))
            if date_modif:
                afficher("Modifié le", self.formater_date_pdf(date_modif))
            print("2")
            print(texte)
            return texte
        except Exception as e:
            texte = f"Erreur lors de l'extraction des métadonnées : {e}"
            return texte
    
    def extraire_metadonnees_img(self, chemin):
        try:
            with open(chemin, 'rb') as f:
                tags = exifread.process_file(f, details=False)

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            if not tags:
                return texte + "Aucune métadonnée EXIF trouvée."

            cles_importantes = [
                "Image Make",
                "Image Model",
                "EXIF DateTimeOriginal",
                "EXIF LensModel",
                "EXIF FNumber",
                "EXIF ExposureTime",
                "GPS GPSLatitude",
                "GPS GPSLongitude",
            ]

            trouve_au_moins = False
            for cle in cles_importantes:
                if cle in tags:
                    valeur = tags[cle]
                    texte += f"{cle:<25}: {valeur}\n"
                    trouve_au_moins = True

            if not trouve_au_moins:
                texte += "Aucune métadonnée EXIF pertinente trouvée."

            return texte

        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées image : {e}"


    def extraire_metadonnees_docx(self, chemin):
        try:
            doc = Document(chemin)
            infos = doc.core_properties

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            def valeur(v): return v if v else "Non défini"

            texte += f"Titre               : {valeur(infos.title)}\n"
            texte += f"Auteur              : {valeur(infos.author)}\n"
            texte += f"Sujet               : {valeur(infos.subject)}\n"
            texte += f"Catégorie           : {valeur(infos.category)}\n"
            texte += f"Commentaires        : {valeur(infos.comments)}\n"
            texte += f"Créé le             : {valeur(infos.created)}\n"
            texte += f"Modifié le          : {valeur(infos.modified)}\n"
            texte += f"Dernier modifié par : {valeur(infos.last_modified_by)}\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées DOCX : {e}"

    def extraire_metadonnees_xlsx(self, chemin):
        try:
            wb = openpyxl.load_workbook(chemin)
            props = wb.properties

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            def valeur(v): return v if v else "Non défini"

            texte += f"Titre        : {valeur(props.title)}\n"
            texte += f"Sujet        : {valeur(props.subject)}\n"
            texte += f"Auteur       : {valeur(props.author)}\n"
            texte += f"Commentaires : {valeur(props.comments)}\n"
            texte += f"Catégorie    : {valeur(props.category)}\n"
            texte += f"Identifiant  : {valeur(props.identifier)}\n"
            texte += f"Langue       : {valeur(props.language)}\n"
            texte += f"Créé le      : {valeur(props.created)}\n"
            texte += f"Modifié le   : {valeur(props.modified)}\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées XLSX : {e}"

    def extraire_metadonnees_pptx(self, chemin):
        try:
            pres = Presentation(chemin)
            props = pres.core_properties

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            def valeur(v): return v if v else "Non défini"

            texte += f"Titre               : {valeur(props.title)}\n"
            texte += f"Sujet               : {valeur(props.subject)}\n"
            texte += f"Auteur              : {valeur(props.author)}\n"
            texte += f"Commentaires        : {valeur(props.comments)}\n"
            texte += f"Catégorie           : {valeur(props.category)}\n"
            texte += f"Créé le             : {valeur(props.created)}\n"
            texte += f"Modifié le          : {valeur(props.modified)}\n"
            texte += f"Dernier modifié par : {valeur(props.last_modified_by)}\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées PPTX : {e}"

    def extraire_metadonnees_audio(self, chemin):
        try:
            audio = File(chemin)
            if audio is None or audio.tags is None:
                return f"Aucune métadonnée audio trouvée pour {os.path.basename(chemin)}"

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            for cle, valeur in audio.tags.items():
                texte += f"{cle:<20}: {valeur}\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées audio : {e}"


    def extraire_metadonnees_video(self, chemin):
        try:
            media_info = MediaInfo.parse(chemin)
            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            for track in media_info.tracks:
                if track.track_type in ["General", "Video", "Audio"]:
                    texte += f"[{track.track_type} Track]\n"
                    for attr, value in track.to_data().items():
                        if value:
                            texte += f"  {attr}: {value}\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées vidéo : {e}"

    def extraire_metadonnees_archive(self, chemin):
        try:
            ext = os.path.splitext(chemin)[1].lower()
            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n"
            texte += "-" * 40 + "\n"

            if ext == ".zip":
                with zipfile.ZipFile(chemin, 'r') as archive:
                    for info in archive.infolist():
                        texte += f"Nom: {info.filename}\n"
                        texte += f"  Taille compressée   : {info.compress_size} bytes\n"
                        texte += f"  Taille décompressée : {info.file_size} bytes\n"
                        texte += f"  Date modif          : {info.date_time}\n"
                        texte += "-" * 20 + "\n"

            elif ext == ".rar":
                with rarfile.RarFile(chemin, 'r') as archive:
                    for info in archive.infolist():
                        texte += f"Nom: {info.filename}\n"
                        texte += f"  Taille compressée   : {info.compress_size} bytes\n"
                        texte += f"  Taille décompressée : {info.file_size} bytes\n"
                        texte += f"  Date modif          : {info.date_time}\n"
                        texte += "-" * 20 + "\n"

            elif ext == ".7z":
                with py7zr.SevenZipFile(chemin, mode='r') as archive:
                    for info in archive.list():
                        texte += f"Nom: {info.filename}\n"
                        texte += f"  Taille compressée   : {info.compressed} bytes\n"
                        texte += f"  Taille décompressée : {info.uncompressed} bytes\n"
                        texte += f"  Date modif          : {info.date_time}\n"
                        texte += "-" * 20 + "\n"

            else:
                texte += "Format d’archive non supporté.\n"

            return texte
        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées archive : {e}"

    def extraire_metadonnees_eml(self, chemin):
        try:
            with open(chemin, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)

            texte = f"\n[MÉTADONNÉES POUR] {os.path.basename(chemin)}\n\n"
            texte += "------------------------------------------------------\n"

            def afficher(cle, valeur):
                nonlocal texte
                texte += f"{cle:<20}: {valeur if valeur else 'Non défini'}\n"

            afficher("De", msg.get('From'))
            afficher("À", msg.get('To'))
            afficher("Sujet", msg.get('Subject'))
            afficher("Date", msg.get('Date'))
            afficher("Répondre à", msg.get('Reply-To'))
            afficher("Client Mail", msg.get('User-Agent') or msg.get('X-Mailer'))

            texte += "\n[Serveurs SMTP traversés]\n"
            received_headers = msg.get_all('Received', [])
            for idx, h in enumerate(received_headers, 1):
                texte += f"  {idx}. {h}\n"

            return texte

        except Exception as e:
            return f"Erreur lors de l'extraction des métadonnées EML : {e}"
        
        
        
        
    
    
    def formater_date_pdf(self, brut):
        """
        Convertit les dates PDF du style 'D:20230402111000Z' en datetime lisible.
        """
        try:
            brut = brut.strip()
            if brut.startswith("D:"):
                brut = brut[2:]
            dt = datetime.strptime(brut[:14], "%Y%m%d%H%M%S")
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            return f"Date non reconnue ({brut})"



if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
