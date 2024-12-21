import os
import sys
import subprocess

# Fonction pour installer comtypes si ce n'est pas déjà installé
# comtypes, c'est une bibliothèque Python permettant d'interagir avec des applications Windows basées sur le modèle COM, 
# comme Microsoft Word, pour automatiser des tâches comme l'ouverture, la manipulation et la conversion de documents.
def install_comtypes():
    try:
        import comtypes # type: ignore
    except ImportError:
        print("comtypes n'est pas installé. Installation en cours...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "comtypes"])

# Fonction pour convertir les fichiers Word en PDF
def convert_word_to_pdf(doc_path, pdf_path):
    import comtypes.client # type: ignore
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 correspond au format PDF
    doc.Close()
    word.Quit()

# Fonction pour parcourir récursivement les dossiers et convertir les fichiers Word en PDF
def convert_all_word_in_directory(root_dir, output_dir):
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.endswith('.doc') or file.endswith('.docx'):
                doc_path = os.path.join(root, file)
                pdf_path = os.path.join(output_dir, os.path.splitext(file)[0] + '.pdf')
                
                print(f"Conversion de {doc_path} en {pdf_path}...")
                convert_word_to_pdf(doc_path, pdf_path)
                print(f"Conversion réussie : {pdf_path}")

# Fonction principale pour exécuter le programme
def main():
    # Vérifier si comtypes est installé et installer si nécessaire
    install_comtypes()

    # Récupérer le répertoire d'exécution du script
    current_directory = os.path.dirname(os.path.abspath(_file_))
    output_directory = os.path.join(current_directory, "PDF")

    # Créer le dossier PDF si nécessaire
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        print(f"Dossier 'PDF' créé à l'emplacement : {output_directory}")

    print(f"Début de la conversion des fichiers Word dans le répertoire : {current_directory}")

    # Lancer la conversion dans ce répertoire
    convert_all_word_in_directory(current_directory, output_directory)

    print("Conversion terminée.")

# Appel de la fonction principale si le script est exécuté directement
if _name_ == "_main_":
    main()
