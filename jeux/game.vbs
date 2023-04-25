Option Explicit
Dim choice
Dim shell

' Demande à l'utilisateur s'il souhaite s'inscrire ou se connecter
choice = MsgBox("Voulez-vous vous inscrire ? Si non, appuyez sur 'Non' pour vous connecter.", vbYesNo)

If choice = vbYes Then ' Si l'utilisateur choisit de s'inscrire

    ' Demande à l'utilisateur de choisir un nom d'utilisateur et un mot de passe
    Dim new_username
    new_username = ""
    Do While new_username = "" ' Tant que le mot de passe est vide ou contient une virgule ou une barre verticale
    new_username = InputBox("Choisissez un nom d'utilisateur:")
    If InStr(new_username, ",") > 0 Or InStr(new_username, "|") > 0 Then ' Vérifie si le mot de passe contient une virgule ou une barre verticale
            MsgBox "Le nom d'utilisateur ne peut pas contenir de ',' ou de '|'. Veuillez en choisir un autre."
            new_username = ""
        End If
    Loop
    dim new_password
    new_password = ""

    Do While new_password = "" ' Tant que le mot de passe est vide ou contient une virgule ou une barre verticale
        new_password = InputBox("Choisissez un mot de passe:")
        If InStr(new_password, ",") > 0 Or InStr(new_password, "|") > 0 Then ' Vérifie si le mot de passe contient une virgule ou une barre verticale
            MsgBox "Le mot de passe ne peut pas contenir de virgule ou de barre verticale. Veuillez en choisir un autre."
            new_password = ""
        End If
    Loop

    ' Vérifie si le nom d'utilisateur est déjà utilisé
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile("users.txt", 1)
    dim username_exists
    username_exists = False
    Do Until file.AtEndOfStream
        line = file.ReadLine
        If InStr(line, new_username & ",") = 1 Then
            username_exists = True
            Exit Do
        End If
    Loop
    file.Close

    If username_exists Then
        MsgBox "Ce nom d'utilisateur est déjà utilisé. Veuillez en choisir un autre."
    Else ' Si le nom d'utilisateur n'est pas déjà utilisé, enregistre les informations d'inscription
        Set file = fso.OpenTextFile("users.txt", 8, True)
        file.WriteLine new_username & "," & new_password
        file.Close
        MsgBox "Inscription réussie !"
    End If
ElseIf choice = vbNo Then ' Si l'utilisateur choisit de se connecter

    ' Demande à l'utilisateur son nom d'utilisateur et son mot de passe
Dim username, password
username = InputBox("Nom d'utilisateur:")
password = InputBox("Mot de passe:")

' Ouvre le fichier texte contenant les noms d'utilisateur et les mots de passe
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim file
Set file = fso.OpenTextFile("users.txt", 1)

' Vérifie si le nom d'utilisateur et le mot de passe sont valides
Dim valid
Dim is_admin
valid = False
is_admin = False
Do Until file.AtEndOfStream
    dim line
    line = file.ReadLine
    If InStr(line, username & ",") = 1 Then
        dim stored_password
        stored_password = Mid(line, Len(username) + 2)
        If password = stored_password Then
            valid = True
            ' Vérifie si l'utilisateur est un administrateur
            If InStr(line, "|admins") > 0 Then
                is_admin = True
            End If
            Exit Do
        End If
    End If
Loop
file.Close

    ' Affiche un message différent si l'utilisateur est un administrateur ou non
    If valid Then
        If is_admin Then
            MsgBox "Connexion réussie en tant qu'administrateur !"
            Set shell = CreateObject("WScript.Shell")
            shell.Run "actions\admin_actions.vbs"
        Else
            choice = MsgBox("Que voulez-vous faire ?" & vbNewLine & "1. Jouer à pierre-papier-ciseaux" & vbNewLine & "2. Lancer le quiz" & vbNewLine & "3. Autre", vbExclamation + vbYesNoCancel, "Choix")

            If choice = vbYes Then
                Dim joueur1, joueur2
    Dim choix_joueur1, choix_joueur2

    joueur1 = username
    joueur2 = "Ordinateur"

    ' Fonction pour obtenir le choix du joueur 1
    Function obtenir_choix_joueur1()
        Dim choix
        choix = InputBox(username &", choisissez pierre, papier ou ciseaux:")
        obtenir_choix_joueur1 = choix
    End Function

    ' Fonction pour obtenir le choix du joueur 2
    Function obtenir_choix_joueur2()
        Dim choix
        
        If joueur2 = "Joueur 2" Then
            choix = InputBox("Joueur 2, choisissez pierre, papier ou ciseaux:")
        Else
            Randomize
            choix = Int((3 * Rnd) + 1)

            Select Case choix
                Case 1
                    choix = "pierre"
                Case 2
                    choix = "papier"
                Case 3
                    choix = "ciseaux"
            End Select
        End If
        
        obtenir_choix_joueur2 = choix
    End Function

    Function determiner_gagnant(joueur1, joueur2)
        If joueur1 = joueur2 Then
            determiner_gagnant = "Egalite"
        ElseIf (joueur1 = "pierre" And joueur2 = "ciseaux") Or (joueur1 = "papier" And joueur2 = "pierre") Or (joueur1 = "ciseaux" And joueur2 = "papier") Then
            determiner_gagnant = joueur1
        Else
            determiner_gagnant = joueur2
        End If
    End Function


    ' Demande si l'utilisateur veut jouer contre l'ordinateur ou contre un autre joueur
    Dim choix_mode
    choix_mode = InputBox("Voulez-vous jouer contre l'ordinateur (tapez 'ordi') ou contre un autre joueur (tapez 'joueur')?")

    If choix_mode = "ordi" Then
        joueur2 = "Ordinateur"
    ElseIf choix_mode = "joueur" Then
        joueur2 = "Joueur 2"
    Else
        MsgBox "Choix invalide. Le jeu va se fermer."
        WScript.Quit
    End If

    ' Obtient les choix des joueurs
    choix_joueur1 = obtenir_choix_joueur1()
    choix_joueur2 = obtenir_choix_joueur2()

    ' Affiche les choix des joueurs
    MsgBox username &" a choisi " & choix_joueur1 & vbCrLf & joueur2 & " a choisi " & choix_joueur2

    ' Détermine le gagnant et affiche le résultat
MsgBox "Le choix gagnant est: " & determiner_gagnant(choix_joueur1, choix_joueur2)

ElseIf choice = vbNo Then

' Ouvre le fichier texte contenant les questions et les réponses
Set fso = CreateObject("Scripting.FileSystemObject")
set file = fso.OpenTextFile("quiz.txt")

' Initialise le score à 0
dim score
score = 0

' Boucle à travers chaque ligne du fichier texte
Do While Not file.AtEndOfStream
' Lit la question et la réponse
line = file.ReadLine()
dim question
question = Split(line, "|")(0)
dim answer
answer = Split(line, "|")(1)
  ' Affiche la question et demande une réponse
  dim response
  response = InputBox(question)
  ' Vérifie si la réponse donnée correspond à la réponse correcte, avec ou sans majuscule
  If LCase(response) = LCase(answer) Then
      ' Si la réponse est correcte, ajoute 1 au score
      score = score + 1
      MsgBox "Bonne reponse !"
  Else
      ' Si la réponse est incorrecte, affiche la réponse correcte
      MsgBox "Mauvaise reponse. La reponse etait : " & answer
  End If
Loop
' Ferme le fichier texte
file.Close

' Affiche le score final
MsgBox "Votre score final est : " & score

' Stocke le score de l'utilisateur dans le fichier de résultats
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("results.txt", 8)
file.WriteLine(username & "," & score)
file.Close
' Affiche les résultats de tous les utilisateurs triés par score
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("results.txt", 1)
dim results
results = "Résultats :" & vbCrLf
Dim user_scores(100) ' Déclare un tableau pour stocker les scores des utilisateurs
dim i
i = 0 ' Initialise l'indice du tableau
Do Until file.AtEndOfStream
line = file.ReadLine
dim user_score
user_score = Split(line, ",")(1)
username = Split(line, ",")(0)
user_scores(i) = Array(username, user_score) ' Ajoute le nom d'utilisateur et le score au tableau
i = i + 1 ' Incrémente l'indice du tableau
Loop
file.Close

' Trie le tableau par score en ordre décroissant
dim j
For j = 0 To i - 2
    dim k
    For k = j + 1 To i - 1
        If user_scores(j)(1) < user_scores(k)(1) Then
            dim temp
            temp = user_scores(j)
            user_scores(j) = user_scores(k)
            user_scores(k) = temp
        End If
    Next
Next

' Ajoute chaque nom d'utilisateur et son score trié au message de résultats
For j = 0 To i - 1
    results = results & user_scores(j)(0) & " : " & user_scores(j)(1) & vbCrLf
Next

MsgBox results
ElseIf choice = vbCancel Then
    Dim choiceAutre
    choiceAutre = MsgBox("que voulez vous faire ?" & vbCrLf & "oui = generateur de mot de passe" & vbCrLf & "non = Le chat", vbYesNo)
    If choice = vbYes Then
        ' Code exécuté si l'utilisateur clique sur le bouton "Oui"
    ' Demande à l'utilisateur la longueur souhaitée pour le mot de passe
dim password_length
password_length = InputBox("Entrez la longueur souhaitee pour le mot de passe aleatoire")

' Initialise les caractères possibles pour le mot de passe
dim characters
characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=.£¤§"

' Génère le mot de passe aléatoire
password = ""
For i = 1 To password_length
    dim random_index
    random_index = Int(Len(characters) * Rnd + 1)
    password = password & Mid(characters, random_index, 1)
Next
' Affiche le mot de passe généré
MsgBox "Votre mot de passe aléatoire est : " & password

dim modifMotDePasse
modifMotDePasse = MsgBox("Voulez-vous changer votre mot de passe en " & password &" ?", vbYesNo)

If modifMotDePasse = vbYes Then
    ' Ouvre le fichier texte contenant les noms d'utilisateur et les mots de passe
    Set fso = CreateObject("Scripting.FileSystemObject")
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
    selected_username = username

    ' Demande à l'utilisateur de saisir le nouveau mot de passe
    new_password = password

    ' Parcourt les lignes du fichier texte et modifie le mot de passe correspondant à l'utilisateur sélectionné
    Dim new_contents
    For Each line In lines
        If InStr(line, selected_username & ",") = 1 Then
            line = selected_username & "," & new_password
        End If
        new_contents = new_contents & line & vbCrLf
    Next

    ' Ouvre le fichier texte en mode écriture et écrit le nouveau contenu
    Set file = fso.OpenTextFile("users.txt", 2)
    file.Write(new_contents)
    file.Close()

    ' Affiche un message de confirmation
    MsgBox "Votre mot de passe a été mis a jour avec succès en " & password & " !"

ElseIf modifMotDePasse = vbNo Then
    ' annuler
End If

elseIf choiceAutre = vbNo Then

' Boucle jusqu'à ce que l'utilisateur quitte le programme
Do While True

    ' Affiche un menu pour l'utilisateur
    choice = MsgBox("Que souhaitez-vous faire ?" & vbCrLf & vbCrLf & "1 - Voir le chat" & vbCrLf & "2 - Envoyer un message" & vbCrLf & "3 - Quitter", vbQuestion + vbYesNoCancel, "Chat")

    ' Vérifie le choix de l'utilisateur
    If choice = vbYes Then
        ' Vérifie si le fichier existe avant de l'ouvrir
Set fso = CreateObject("Scripting.FileSystemObject")
dim filePath
filePath = "chat.txt"

If fso.FileExists(filePath) Then
' Ouvre le fichier texte contenant le chat
Set file = fso.OpenTextFile(filePath, 1)
' Déclaration de la variable qui va contenir les messages
Dim messages
messages = ""

' Lit chaque ligne du fichier et extrait les messages
Do While Not file.AtEndOfStream
    ' Lit la ligne courante
    line = file.ReadLine()
    ' Trouve la position du premier "#" dans la ligne
    dim pos
    pos = InStr(line, "#")
    ' Vérifie si la ligne contient un "#"
    If pos > 0 Then
        ' Extrait le message de la ligne en supprimant tout ce qui se trouve après le "#"
        message = Left(line, pos - 1)
        ' Ajoute le message à la chaîne de caractères contenant tous les messages et ajoute une ligne vide
        messages = messages & message & vbCrLf & vbCrLf
    End If
Loop

' Ferme le fichier texte
file.Close()

' Affiche les messages dans une boîte de dialogue
If messages <> "" Then
    MsgBox messages
Else
    MsgBox "Aucun message trouvé dans le fichier chat.txt"
End If
Else
    ' Affiche un message d'erreur si le fichier n'est pas trouvé
    MsgBox "Le fichier chat.txt est introuvable"
    End If
    ElseIf choice = vbNo Then
        ' Demande à l'utilisateur d'entrer un message
        Dim message
        message = InputBox("Entrez votre message :")

        ' Vérifie si l'utilisateur a entré un message
        If message <> "" Then
            ' Vérifie si le message contient les caractères "|" ou ":"
            If InStr(message, "|") > 0 Or InStr(message, ":") > 0 Then
                ' Affiche un message d'erreur si le message contient les caractères "|" ou ":"
                MsgBox "Le message ne peut pas contenir les caractères '|' ou ':'"
            Else
                ' Ouvre le fichier "chat.txt" en mode ajout
                Set file = fso.OpenTextFile("chat.txt", 8, True)

                ' Récupère le temps en secondes depuis minuit
                Dim seconds
                seconds = Timer()

                ' Génère un code aléatoire en fonction du temps en secondes depuis minuit
                Dim randomCode
                Randomize seconds ' Initialise la séquence de nombres aléatoires avec le temps en secondes depuis minuit
                randomCode = Int(Rnd() * 1000000) ' Génère un nombre aléatoire compris entre 0 et 999999
                randomCode = Right("000000" & randomCode, 6) ' Formate le nombre aléatoire pour qu'il contienne 6 chiffres
                ' Ajoute le message à la fin du fichier et ferme le fichier
                file.WriteLine vbCrLf & username & " : " & message & " #" & randomCode
                file.Close()

                ' Affiche un message de confirmation
                MsgBox "Votre message a été enregistré dans le fichier chat.txt"
            End If
        Else
            ' Affiche un message d'erreur si aucun message n'a été entré
            MsgBox "Aucun message entré"
        End If
    Else
        ' Quitte la boucle
        Exit Do
    End If

Loop

End If
end if
end if
Else
    MsgBox "Nom d'utilisateur ou mot de passe invalide."
end if
end if
