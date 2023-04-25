' Ouvre le fichier texte contenant le chat
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("chat.txt", 1)

If Err.Number <> 0 Then
    ' Gère l'erreur si le fichier n'existe pas
    MsgBox "Le fichier chat.txt est introuvable. Veuillez le créer avant de continuer.", vbCritical, "Erreur"
    WScript.Quit
End If

' Lit le contenu du fichier texte dans une variable
contents = file.ReadAll()

' Ferme le fichier texte
file.Close()

' Boucle jusqu'à ce que l'utilisateur quitte le programme
Do While True

    ' Affiche un menu pour l'utilisateur
    choice = MsgBox("Que souhaitez-vous faire ?" & vbCrLf & vbCrLf & "1 - Voir le chat" & vbCrLf & "2 - Envoyer un message" & vbCrLf & "3 - Supprimer un message" & vbCrLf & "4 - Quitter", vbQuestion + vbYesNoCancel, "Chat")

    ' Vérifie le choix de l'utilisateur
    Select Case choice
        Case vbYes ' Voir le chat
            ' Affiche le contenu du fichier chat.txt
            MsgBox contents, vbInformation, "Chat"
        Case vbNo ' Envoyer un message
            ' Demande le nom d'utilisateur à l'utilisateur
            username = InputBox("Entrez votre nom d'utilisateur :")

            ' Demande à l'utilisateur d'entrer un message
            message = InputBox("Entrez votre message :")

            ' Vérifie si l'utilisateur a entré un message
            If message <> "" Then
                ' Ouvre le fichier "chat.txt" en mode ajout
                Set file = fso.OpenTextFile("chat.txt", 8, True)

                If Err.Number <> 0 Then
                    ' Gère l'erreur si le fichier ne peut pas être ouvert en mode ajout
                    MsgBox "Impossible d'ouvrir le fichier chat.txt en mode ajout. Veuillez vérifier que le fichier n'est pas ouvert dans une autre application.", vbCritical, "Erreur"
                    WScript.Quit
                End If

                ' Vérifie si le fichier est vide
                If file.Size > 0 Then
                    ' Ajoute un retour à la ligne avant d'écrire le message
                    file.WriteLine ""
                End If

                ' Récupère le temps en secondes depuis minuit
                seconds = Timer()

                ' Génère un code aléatoire en fonction du temps en secondes depuis minuit
                Randomize seconds ' Initialise la séquence de nombres aléatoires avec le temps en secondes depuis minuit
                randomCode = Int(Rnd() * 1000000) ' Génère un nombre aléatoire compris entre 0 et 999999
                randomCode = Right("000000" & randomCode, 6) ' Formate le nombre aléatoire pour qu'il contienne 6 chiffres
                
                ' Demande à l'utilisateur s'il veut afficher "admin" à côté de son pseudo
                adminPseudo = MsgBox("Voulez-vous afficher admin à côté de votre pseudo ?", vbYesNo)

                If adminPseudo = vbYes Then
            ' Code exécuté si l'utilisateur clique sur le bouton "Oui"
            ' Ajoute le message à la fin du fichier et ferme le fichier
            file.Write "admin | " & username & " : " & message & " #" & randomCode & vbCrLf
            file.Close()

        ElseIf adminPseudo = vbNo Then
                ' Code exécuté si l'utilisateur clique sur le bouton "Non"
                file.Write username & " : " & message & " #" & randomCode
                file.Close()

            End If

            ' Affiche un message de confirmation
            MsgBox "Votre message a été envoyer"
        Else
            ' Affiche un message d'erreur si aucun message n'a été entré
            MsgBox "Aucun message entré"
        End Select
        End If
    Else
        ' Quitte la boucle
    End If

Loop