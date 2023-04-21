' Demande à l'utilisateur le nom d'utilisateur qu'il veut débannir
Dim username
username = InputBox("Nom d'utilisateur a debannir:")

' Ouvre le fichier "ban.txt" et lit son contenu
Dim fso, ban_file, users_file, line, new_users_contents
Set fso = CreateObject("Scripting.FileSystemObject")
Set ban_file = fso.OpenTextFile("actions\admin\ban.txt", 1)
new_users_contents = ""

' Vérifie si le nom d'utilisateur est dans le fichier "ban.txt"
Dim is_banned
is_banned = False
Do Until ban_file.AtEndOfStream
    line = ban_file.ReadLine
    If InStr(line, username & ",") = 1 Then
        is_banned = True
    Else
        new_users_contents = new_users_contents & line & vbCrLf
    End If
Loop
ban_file.Close

If Not is_banned Then
    MsgBox "Cet utilisateur n'est pas banni."
Else
    ' Ajoute l'utilisateur débanni à la fin du fichier "users.txt"
    Set users_file = fso.OpenTextFile("users.txt", 8, True)
    users_file.WriteLine line
    users_file.Close
    
    ' Écrit les utilisateurs restants dans le fichier "ban.txt"
    Set ban_file = fso.OpenTextFile("actions\admin\ban.txt", 2)
    ban_file.Write new_users_contents
    ban_file.Close
    
    MsgBox username & "debanni avec succès !"
End If