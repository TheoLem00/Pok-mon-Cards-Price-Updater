# Pokemon-Cards-Price-Updater
Automatically update the price of pokemon cards using the cardmarket.com link


EN :

How to Use the Pokémon Card Price Updater

1 - Download the Required Files
First, download both the Excel file and the Python (.py) script. Save them in the same folder so you can easily locate them later, as their file paths will be needed.

2 - Set Up the Excel File
Open the Excel file. You will need to enable the Developer tab. To do this:

	-Click on File in the top-left corner.

	-Select Options at the bottom.

	-In the new window, go to Customize Ribbon.

	-On the right-hand side, scroll down and check the box next to Developer.

	-Click OK.

3 - Configure the VBA Script
Go to the Developer tab and click on Visual Basic.
In the new window, if it’s not already open, double-click on Module1 in the left panel to view the code.
Replace the placeholder path "C:\Users\..." with the actual path to your Python file.
Save and close the editor.

4 - Edit the Python Script
Open the .py file using a code editor like Visual Studio Code (recommended for its ease of use).
On line 28, replace the placeholder "C:\Users\..." with the full path to your Excel file.
Save the file by clicking File > Save.

5 - Using the Updater
Go back to the Excel file.

In the relevant columns, enter the name of the Pokémon card and paste the corresponding CardMarket.com link.

I recommende to apply filters on the CardMarket link (e.g., language, minimum card condition) for more accurate pricing.

6 - Run the Update
Once everything is set up, simply click on the Pokéball button in Excel.
The script will run for a few seconds, and the card price will automatically update in your file!


FR : 

Comment utiliser l’outil d’actualisation automatique des prix de cartes Pokémon

1 - Télécharger les fichiers nécessaires
Commencez par télécharger les deux fichiers : le fichier Excel et le script Python (.py).
Placez-les dans un même dossier pour les retrouver facilement, car vous aurez besoin de leur chemin d’accès plus tard.

2 - Configurer le fichier Excel
Ouvrez le fichier Excel. Vous devez activer l’onglet Développeur. Pour cela :

	-Cliquez sur Fichier en haut à gauche.

	-Sélectionnez Options en bas du menu.

	-Dans la fenêtre qui s’ouvre, cliquez sur Personnaliser le ruban.

	-Dans la liste de droite, cochez la case Développeur tout en bas.

	-Cliquez sur OK.

3 - Modifier le code VBA.
Rendez-vous dans l’onglet Développeur, puis cliquez sur Visual Basic à gauche.
Une nouvelle fenêtre s’ouvre. Si le code ne s’affiche pas directement, double-cliquez sur Module1 dans le panneau de gauche.
Remplacez le chemin "C:\Users\..." par le chemin complet de votre fichier .py.
Enregistrez les modifications et fermez l’éditeur.

4 - Modifier le script Python.
Ouvrez le fichier .py avec un éditeur de code.
Je vous recommande Visual Studio Code, simple et efficace.
Une fois le logiciel installé, ouvrez votre fichier .py.
À la ligne 28, remplacez "C:\Users\..." par le chemin de votre fichier Excel.
Enregistrez les modifications via Fichier > Enregistrer.

5 - Utiliser l’outil
Revenez dans le fichier Excel.

Saisissez le nom du Pokémon dans la colonne prévue.

Collez le lien correspondant sur CardMarket.com.

Pour plus de précision, ajoutez des filtres sur CardMarket (langue, état minimum de la carte, etc.).

6 - Lancer la mise à jour
Cliquez simplement sur la Pokéball dans Excel.
Le script se lance, travaille pendant quelques secondes, puis met à jour automatiquement le prix de la carte !
