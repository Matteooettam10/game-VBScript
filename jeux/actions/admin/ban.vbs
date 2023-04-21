' Demande à l'utilisateur le pseudo de l'utilisateur à bannir
Dim user_to_ban
user_to_ban = InputBox("Entrez le pseudo de l'utilisateur a bannir :")

' Ouvre le fichier "users.txt" en lecture
Set fso = CreateObject("Scripting.FileSystemObject")
Set users_file = fso.OpenTextFile("users.txt", 1)

' Ouvre le fichier "ban.txt" en écriture, ou en crée un nouveau s'il n'existe pas
Set ban_file = fso.OpenTextFile("actions\admin\ban.txt", 8, True)

' Lit le contenu de "users.txt" ligne par ligne, copie chaque ligne dans "ban.txt"
' si le pseudo correspond à celui de l'utilisateur à bannir, et écrit toutes les autres
' lignes dans un fichier temporaire
Set temp_file = fso.CreateTextFile("actions\admin\temp.txt")
Do Until users_file.AtEndOfStream
    line = users_file.ReadLine
    If InStr(line, user_to_ban & ",") = 1 Then
        ban_file.WriteLine line
    Else
        temp_file.WriteLine line
    End If
Loop

' Ferme les fichiers "users.txt", "ban.txt" et "temp.txt"
users_file.Close
ban_file.Close
temp_file.Close

' Supprime le fichier "users.txt" et renomme "temp.txt" en "users.txt" pour le remplacer
fso.DeleteFile("users.txt")
fso.MoveFile "actions\admin\temp.txt", "users.txt"

MsgBox "L'utilisateur " & user_to_ban & " a ete banni."