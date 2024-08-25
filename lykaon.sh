#!/bin/bash

# Fonction pour nettoyer les fichiers temporaires
cleanup() {
    rm -f "$SCRIPT_FILE"
    rm -f "$DIR/all_authors.txt"
    rm -rf "$TMP_DIR"
    echo "Fichiers temporaires nettoyés."
}

# Configurer le nettoyage automatique en cas de sortie du script
trap cleanup EXIT

# Vérifier que toutes les commandes nécessaires sont disponibles
for cmd in cp unzip zip sed osascript python3; do
    if ! command -v $cmd &> /dev/null; then
        echo "Erreur : La commande $cmd est requise mais n'est pas installée. Veuillez l'installer et réessayer."
        exit 1
    fi
done

# Demander à l'utilisateur de sélectionner le fichier à modifier (.docx ou .doc)
FILE=$(osascript -e 'set docFile to choose file with prompt "Sélectionnez le fichier .docx ou .doc à modifier"' -e 'POSIX path of docFile' 2>/dev/null)

# Vérifier si l'utilisateur a annulé la sélection
if [[ -z "$FILE" ]]; then
    echo "Aucun fichier sélectionné. Script annulé."
    exit 1
fi

# Extraire le répertoire et le nom du fichier pour sauvegarder le fichier modifié au même endroit
DIR="$(dirname "$FILE")"
BASENAME=$(basename "$FILE" | sed 's/\.[^.]*$//') # Enlever l'extension du fichier
EXT="${FILE##*.}"

# Nom du fichier modifié
FINAL_FILE="$DIR/${BASENAME}_modified.$EXT"
ZIP_FILE="$DIR/${BASENAME}_modified.zip"
TMP_DIR=$(mktemp -d "${DIR}/tmp.XXXXXXXXXX")

# Étape 1: Copier et renommer le fichier en .zip si c'est un .docx
if [[ "$EXT" == "docx" ]]; then
    cp "$FILE" "$ZIP_FILE"
else
    echo "Le format $EXT n'est pas pris en charge pour la modification. Veuillez sélectionner un fichier .docx."
    exit 1
fi

# Étape 2: Décompresser le fichier .zip dans un répertoire temporaire
unzip -o "$ZIP_FILE" -d "$TMP_DIR"

# Extraire les auteurs de comments.xml et document.xml avec Python
COMMENTS_XML="$TMP_DIR/word/comments.xml"
DOCUMENT_XML="$TMP_DIR/word/document.xml"

# Exécuter le script Python pour extraire les auteurs
python3 extract_authors.py "$COMMENTS_XML" "$DOCUMENT_XML" > "$DIR/all_authors.txt"

# Vérifier si des auteurs ont été trouvés
if [[ ! -s "$DIR/all_authors.txt" ]]; then
    echo "Aucun auteur trouvé dans les fichiers XML. Script annulé."
    exit 1
fi

# Lire les noms du fichier et formater pour AppleScript
AUTHOR_LIST=$(awk '{print $0}' "$DIR/all_authors.txt" | sed 's/^/"/;s/$/"/' | paste -sd, -)

# Enregistrer le script AppleScript dans un fichier temporaire
SCRIPT_FILE=$(mktemp /tmp/select_author.XXXXXXXXXX.scpt)

# Créer le script AppleScript pour la sélection multiple
cat <<EOF > "$SCRIPT_FILE"
set authors to {$AUTHOR_LIST}
set selectedAuthors to choose from list authors with prompt "Sélectionnez un ou plusieurs auteurs" with multiple selections allowed
if selectedAuthors is not false then
    set AppleScript's text item delimiters to ","
    set selectedAuthors to selectedAuthors as string
    return selectedAuthors
else
    return "Aucun auteur sélectionné"
end if
EOF

# Exécuter le script AppleScript et récupérer les sélections
OLD_NAMES=$(osascript "$SCRIPT_FILE")

# Vérifier si l'utilisateur a annulé la sélection
if [[ "$OLD_NAMES" == "Aucun auteur sélectionné" ]]; then
    echo "Aucun auteur sélectionné. Script annulé."
    exit 1
fi

# Demander le nouveau nom via une boîte de dialogue
NEW_NAME=$(osascript -e 'Tell application "System Events" to display dialog "Entrez le nouveau nom :" default answer ""' -e 'text returned of result' 2>/dev/null)

# Vérifier si la variable NEW_NAME est non vide
if [[ -z "$NEW_NAME" ]]; then
    echo "Le nouveau nom ne peut pas être vide. Veuillez réessayer."
    exit 1
fi

# Afficher une boîte de dialogue pour choisir l'action à effectuer
CHOICE=$(osascript -e "Tell application \"System Events\" to display dialog \"Où souhaitez-vous modifier l'auteur ?\" buttons {\"Modifier tout\", \"Modifier l'auteur des commentaires\", \"Modifier l'auteur des révisions\"} default button 1" -e "button returned of result" 2>/dev/null)

# Vérifier le choix de l'utilisateur et définir la variable ACTION en conséquence
case "$CHOICE" in
    "Modifier tout")
        ACTION="all"
        ;;
    "Modifier l'auteur des commentaires")
        ACTION="comments"
        ;;
    "Modifier l'auteur des révisions")
        ACTION="document"
        ;;
    *)
        echo "Choix invalide. Script annulé."
        exit 1
        ;;
esac

# Convertir OLD_NAMES en un tableau pour traitement
IFS=',' read -r -a AUTHORS_ARRAY <<< "$OLD_NAMES"

# Initialiser une variable pour le message de confirmation
MESSAGE="*** La modification a été effectuée avec succès ***\n\n"

# Remplacer chaque auteur sélectionné par le nouveau nom
for OLD_NAME in "${AUTHORS_ARRAY[@]}"; do
    # Supprimer les espaces en début et fin de chaîne
    OLD_NAME=$(echo "$OLD_NAME" | xargs)
    
    # Échapper les caractères spéciaux pour sed
    OLD_NAME_ESCAPED=$(printf '%s\n' "$OLD_NAME" | sed 's/[&/\]/\\&/g')
    NEW_NAME_ESCAPED=$(printf '%s\n' "$NEW_NAME" | sed 's/[&/\]/\\&/g')

    # Remplacement conditionnel en fonction du choix de l'utilisateur
    case "$ACTION" in
        "comments")
            if [[ -f "$COMMENTS_XML" ]]; then
                # Remplacement uniquement dans comments.xml si le nom est précédé de w:author="
                sed -i '' "s/w:author=\"$OLD_NAME_ESCAPED\"/w:author=\"$NEW_NAME_ESCAPED\"/g" "$COMMENTS_XML"
                MESSAGE+="$OLD_NAME a été remplacé par $NEW_NAME dans les commentaires.\n"
            else
                echo "Le fichier comments.xml n'a pas été trouvé."
            fi
            ;;
        "document")
            if [[ -f "$DOCUMENT_XML" ]]; then
                # Remplacement uniquement dans document.xml si le nom est précédé de w:author="
                sed -i '' "s/w:author=\"$OLD_NAME_ESCAPED\"/w:author=\"$NEW_NAME_ESCAPED\"/g" "$DOCUMENT_XML"
                MESSAGE+="$OLD_NAME a été remplacé par $NEW_NAME dans les révisions.\n"
            else
                echo "Le fichier document.xml n'a pas été trouvé."
            fi
            ;;
        "all")
            if [[ -f "$COMMENTS_XML" ]]; then
                # Remplacement dans comments.xml si le nom est précédé de w:author="
                sed -i '' "s/w:author=\"$OLD_NAME_ESCAPED\"/w:author=\"$NEW_NAME_ESCAPED\"/g" "$COMMENTS_XML"
            fi
            if [[ -f "$DOCUMENT_XML" ]]; then
                # Remplacement dans document.xml si le nom est précédé de w:author="
                sed -i '' "s/w:author=\"$OLD_NAME_ESCAPED\"/w:author=\"$NEW_NAME_ESCAPED\"/g" "$DOCUMENT_XML"
            fi
            MESSAGE+="$OLD_NAME a été remplacé par $NEW_NAME dans les commentaires et les révisions.\n"
            ;;
        *)
            echo "Choix invalide. Script annulé."
            exit 1
            ;;
    esac
done

# Étape 4: Se positionner dans le répertoire temporaire pour recomprimer
cd "$TMP_DIR" || exit

# Étape 5: Recompresser le contenu en un fichier .zip, en respectant la structure originale
zip -r "$ZIP_FILE" . -x "*.DS_Store"

# Étape 6: Renommer le fichier .zip final en .docx
mv "$ZIP_FILE" "$FINAL_FILE"

# Afficher le message de confirmation dans une boîte de dialogue
osascript -e "Tell application \"System Events\" to display dialog \"$MESSAGE\nLe fichier modifié est disponible dans le même dossier.\" buttons {\"OK\"} default button \"OK\""
