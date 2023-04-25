Dim choice
Dim shell
choice = MsgBox("oui = creer des questions de quiz" & vbCrLf & "non = sois mettre une personne admins  ou lui enleve ou bannir une personne ou la debannir" & vbCrLf & "annuler = pour parlez dans de chat", vbYesNoCancel, "Pannel admin")

If choice = vbYes Then
    ' Créer un objet FileSystemObject pour travailler avec des fichiers
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Demander à l'utilisateur la question et la réponse du quiz à ajouter
    Dim question_quiz, reponse_quiz
    question_quiz = InputBox("Saisissez la question du quiz a ajouter:")
    reponse_quiz = InputBox("Saisissez la réponse du quiz a ajouter:")

    ' Ouvrir le fichier texte en mode ajout
    Set file = fso.OpenTextFile("quiz.txt", 8, True)

    ' Ajouter la question et la réponse à la fin du fichier, sur une nouvelle ligne
    file.WriteLine vbCrLf & question_quiz & "|" & reponse_quiz

    ' Fermer le fichier et afficher un message de confirmation
    file.Close
    MsgBox "La question a ete ajoutee au quiz !"
ElseIf choice = vbNo Then
    
    choice = MsgBox("oui = mettre admins une personne ou lui enleve" & vbCrLf & "non = bannire ou debannir une personne" & vbCrLf & "annuler = pour annule", vbYesNoCancel)

    If choice = vbYes Then
        
        choice = MsgBox("oui = pour mettre admins une personnes" & vbCrLf & "non = enleve la perms admins", vbYesNoCancel, "panel choix pour admins")

        If choice = vbYes Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run "actions\admin\admin.vbs"
        ElseIf choice = vbNo Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run "actions\admin\delAdmin.vbs"
        ElseIf choice = vbCancel Then
            'annule son choix
        End If

    ElseIf choice = vbNo Then
        choice = MsgBox("oui = ban un utilisateur" & vbCrLf & "non = deban un utilisateur", vbYesNoCancel)

        If choice = vbYes Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run "actions\admin\ban.vbs"
        ElseIf choice = vbNo Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run "actions\admin\unban.vbs"
        ElseIf choice = vbCancel Then
            ' annule le choix
    End If
    ElseIf choice = vbCancel Then
    ' Code exécuté si l'utilisateur choisit "Annuler"
    End If

ElseIf choice = vbCancel Then
    Set shell = CreateObject("WScript.Shell")
    shell.Run "actions\admin\chat.vbs"
End If
