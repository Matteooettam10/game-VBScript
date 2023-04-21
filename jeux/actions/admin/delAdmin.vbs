' Demande à l'utilisateur de sélectionner un nom d'utilisateur à modifier
Dim selected_username
selected_username = InputBox("Selectionnez un nom d'utilisateur a modifier :")

' Vérifie si le nom d'utilisateur existe dans le fichier texte
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("users.txt", 1)
Dim new_file_content
new_file_content = ""
Do Until file.AtEndOfStream
    line = file.ReadLine
    If InStr(line, selected_username & ",") = 1 Then
        ' Retire la chaîne "|admins" si elle est présente
        If InStr(line, "|admins") > 0 Then
            line = Replace(line, "|admins", "")
        End If
    End If
    new_file_content = new_file_content & line & vbCrLf
Loop
file.Close

' Écrit le contenu modifié dans le fichier texte
Set file = fso.OpenTextFile("users.txt", 2)
file.Write new_file_content
file.Close

MsgBox "Le nom d'utilisateur a ete modifie avec succes !"