import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors

def verifier_dossier(fichier):
    rapport = []

    try:
        if fichier.endswith('.csv'):
            df = pd.read_csv(fichier)
        elif fichier.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(fichier)
        else:
            return "Format de fichier non supporté. Utilisez un fichier CSV ou Excel."
    except Exception as e:
        return f"Erreur lors de la lecture du fichier : {e}"

    colonnes_requises = [
        'Référence administrative - Demande', 'Numéro',
        'Nature du poste financé', 'Tâches prévues sur le projet',
        'Salaire brut', 'Charges patronales',
        'Aides financières (contrats aidés)', '% temps consacré au projet',
        'Total dépenses éligibles', 'Commentaire',
        'Instruction sur les dépenses de personnel - code',
        'Instruction sur les dépenses de personnel - libelle',
        "Commentaire d'instruction"
    ]

    if not all(col in df.columns for col in colonnes_requises):
        return "Le fichier ne contient pas toutes les colonnes requises."

    colonnes_avant_total = [
        'Référence administrative - Demande', 'Numéro',
        'Nature du poste financé', 'Tâches prévues sur le projet',
        'Salaire brut', 'Charges patronales',
        'Aides financières (contrats aidés)', '% temps consacré au projet'
    ]
    lignes_incompletes = df[df[colonnes_avant_total].isnull().any(axis=1)]
    
    if not lignes_incompletes.empty:
        rapport.append(f"{len(lignes_incompletes)} ligne(s) incomplète(s) avant 'Total dépenses éligibles'. Détails :")
        for index, ligne in lignes_incompletes.iterrows():
            reference = ligne['Référence administrative - Demande']
            colonnes_vides = [col for col in colonnes_avant_total if pd.isnull(ligne[col])]
            rapport.append(f"  - Référence {reference} : colonnes vides -> {', '.join(colonnes_vides)}")
    else:
        rapport.append("Toutes les lignes sont complètes avant 'Total dépenses éligibles'.")

    instructions_invalides = df[~df['Instruction sur les dépenses de personnel - libelle'].isin(['retenue', 'non retenue']) | df['Instruction sur les dépenses de personnel - libelle'].isnull()]
    if not instructions_invalides.empty:
        rapport.append(f"{len(instructions_invalides)} ligne(s) avec une instruction non valide ou manquante.")
    else:
        rapport.append("Toutes les lignes ont une instruction valide ('retenue' ou 'non retenue').")

    non_retenues_non_conformes = df[
        (df['Instruction sur les dépenses de personnel - libelle'] == 'non retenue') &
        (df['Total dépenses éligibles'] != 0)
    ]
    if not non_retenues_non_conformes.empty:
        rapport.append(f"{len(non_retenues_non_conformes)} ligne(s) 'non retenue' avec des dépenses éligibles > 0.")
    else:
        rapport.append("Toutes les lignes 'non retenue' ont des dépenses éligibles à 0.")

    df['Somme Salaire + Charges'] = df['Salaire brut'] + df['Charges patronales']
    depenses_non_conformes = df[df['Total dépenses éligibles'] != df['Somme Salaire + Charges']]
    if not depenses_non_conformes.empty:
        rapport.append(f"{len(depenses_non_conformes)} ligne(s) avec un total des dépenses éligibles différent de la somme 'Salaire brut' + 'Charges patronales'.")
    else:
        rapport.append("Toutes les lignes ont un total des dépenses éligibles conforme à 'Salaire brut' + 'Charges patronales'.")

    return "\n".join(rapport)

def ouvrir_fichier():
    fichier = filedialog.askopenfilename(filetypes=[("Fichiers CSV et Excel", "*.csv;*.xls;*.xlsx")])
    if fichier:
        try:
            rapport = verifier_dossier(fichier)
            sauvegarder_rapport_pdf(rapport)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

def sauvegarder_rapport_pdf(rapport):
    dossier_sortie = filedialog.askdirectory()
    if dossier_sortie:
        save_path = f"{dossier_sortie}/rapport_controle.pdf"

        c = canvas.Canvas(save_path, pagesize=letter)
        c.setFont("Helvetica-Bold", 16)

        width, height = letter
        title = "Rapport de contrôle"
        title_width = c.stringWidth(title, "Helvetica-Bold", 16)
        c.drawString((width - title_width) / 2, height - 40, title)

        c.setFont("Helvetica", 10)

        y_position = height - 60  

        for line in rapport.split("\n"):
            c.drawString(40, y_position, line)
            y_position -= 12  
            if y_position < 40:  
                c.showPage()
                c.setFont("Helvetica", 10)
                y_position = height - 40
                c.setFont("Helvetica-Bold", 16)
                c.drawString((width - title_width) / 2, height - 40, title)
                c.setFont("Helvetica", 10)

        c.save()
        messagebox.showinfo("Succès", f"Le rapport a été généré et sauvegardé dans : {save_path}")

def main():
    root = tk.Tk()
    root.title("Application de Contrôle de Dossier")

    btn_ouvrir = tk.Button(root, text="Ouvrir un fichier à vérifier", command=ouvrir_fichier, padx=20, pady=10)
    btn_ouvrir.pack(padx=20, pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
