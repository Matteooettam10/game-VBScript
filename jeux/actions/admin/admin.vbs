' Ouvre le fichier texte contenant les noms d'utilisateur et les mots de passe
Set fso = CreateObject("Scripting.FileSystemObject")
Dim file
Set file = fso.OpenTextFile("users.txt", 1)

' Lit le contenu du fichier texte dans une variable
Dim contents
contents = file.ReadAll()

' Ferme le fichier texte
file.Close()

' Divise le contenu du fichier texte en lignes
Dim lines
lines = Split(contents, vbCrLf)

' Demande à l'utilisateur de sélectionner un nom d'utilisateur
Dim selected_username
selected_username = InputBox("Selectionnez un nom d'utilisateur :")

' Parcourt les lignes du fichier texte et ajoute "|admins" à la fin de la ligne correspondante
Dim new_contents
Dim line
For Each line In lines
    If InStr(line, selected_username & ",") = 1 Then
        line = line & "|admins"
    End If
    new_contents = new_contents & line & vbCrLf
Next

' Ouvre le fichier texte en mode écriture et écrit le nouveau contenu
Set file = fso.OpenTextFile("users.txt", 2)
file.Write(new_contents)
file.Close()

' Affiche un message de confirmation
MsgBox "Le nom d'utilisateur " & selected_username & " a ete mis à jour avec le role d'administrateur."