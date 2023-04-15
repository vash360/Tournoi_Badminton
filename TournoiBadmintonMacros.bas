REM  *****  BASIC  *****
REM Fait par David Lanier Aout-Sept/ Nov 2021, 
REM Repris en Mars-Avril 2023 par David Lanier pour faire un algorithme plus perfectionné du choix des matches en ronde suisse pour que les joueurs les mieux
REM classés en pourcentage de matches gagnés jouent ensemble et ceux avec le moins de matches jouent ensemble, comme s'il y avait 2 tableaux.
REM On utilise le pourcentage de matches gagnés plutôt que le vrai nombre de matches gagnés car les joueus sont sensés pouvoir s'arrêter quelques tours ou arriver en cours de tournoi
REM donc si un joueur rentre et commence à tout gagner, son pourcentage de matches gagnés est elevé et on le fait jouer avec les meilleurs.

'Option Explicit

'Bug

REM Améliorations
REM	Quand on a plus dhommes ou de femmes, on devrait utiliser des complétants pour ne pas avoir des doubles hommes ou dames au lieu de mixtes
REM Pour le choix d'un partenaire ou adversaire en mode ron de suisse, on poourrait chercher quel est la personne la plus adaptée en fonction de son pourcentage lorsque il n'y en a plus du haut ou du bas du tableau
REM Ex : Il reste 1 hom et 1 dame dans le bas du tableau et d'autres joueurs dans le haut du tableau, on va prendre pour compléter des gens du haut du tableau mais on pourrait prendre des gens dans le bas du haut du tableau...

'Déclaration des variables
Dim Classeur As Object, Feuille As Object, Tour As Object

Dim IndexColumnName as Integer
Dim NumJoueurs as Integer
Dim NumMatches as Integer
Dim NumJoueursHomRestant as Integer
Dim NumJoueursFemRestant as Integer
Dim NumJoueursHomCompletantRestant as Integer
Dim NumJoueursFemCompletantRestant as Integer
Dim indexLigneSrcJoueurs as Integer
Dim indexLigneDstJoueurs as Integer
Dim TourName as String
Dim IndicesDesJoueursDeCeTour(100) as Integer 'Est l'indice relatif à la 1ere ligne de joueurs (= indexLigneSrcJoueurs)), cette indirection sert quand certains joueurs ne jouent pas ce tour

'Pour chaque joueur on indique les 2 numéros de matches dans lequel il peut joueur afin de ne pas
'le choisir en complétant dans un même match. Les indices sont, pour le joueur n :, 2n et (2n + 1) et on regarde leur valeur, -1 voulant dire : pas encore de match.
'Il n'y a que 2 matches ou un joueur peut jouer, un pour son tour et un autre en tant que complétant. Il n'y a que 3 complétants maximim par tour pour faire un match s'il reste un joueur seul qui doit jouer
Dim JoueurNumeroMatch(200) as Integer 'Le double des joueurs car on stocke 2 indices par joueur
Dim i as Integer

Sub CreeTourSuivant()
	
	Dim NumJoueursIndexes as Integer
	Dim actualIndex as Integer

	'La variable Classeur est le classeur en cours
	Classeur = ThisComponent
	
	'La variable Feuille est la feuille nommée Feuille2
	Feuille = Classeur.Sheets.GetByName("Inscriptions")
	
	IndexColumnName 		= 1
	indexLigneSrcJoueurs 	= 9 'Is the Number in the line -  1 (If it's line 10, it's 10-1 = 9)
	indexLigneDstJoueurs 	= 4
	
	NumJoueurs = Feuille.getCellByPosition(IndexColumnName+7,2).value
	'print Feuille.getCellByPosition(IndexColumnName+7,2).string
	
	'Recupere le numero du tour courant
	Dim TourNumeroAsText as String
	TourNumeroAsText = Feuille.getCellByPosition(IndexColumnName+7,6).String
	Dim TourNumeroAsInteger as Integer, NouveauTourNumeroAsInteger as Integer
	TourNumeroAsInteger = Feuille.getCellByPosition(IndexColumnName+7,6).Value
	NouveauTourNumeroAsInteger = TourNumeroAsInteger + 1
	Dim NombreDeMatcheJouesEnTout as Integer
	NombreDeMatcheJouesEnTout = Feuille.getCellByPosition(IndexColumnName+9,6).Value
	
	'Sauve le fichier avec le numéro du tour avant de créer le tour suivant
	'Utilisation de l'algorithme de Ronde Suisse ?
	Dim bSauvegardeautoAvantTourSuivant as Boolean
	Dim theString as String
	theString = Feuille.getCellByPosition(4,2).String ' Le choix de la sauvegarde auto est en cellule E3 soit (4,2)
	If (theString = "X" or theString = "x") Then
		bSauvegardeautoAvantTourSuivant = true
	Else
		bSauvegardeautoAvantTourSuivant = false
	END If
	
	If (bSauvegardeautoAvantTourSuivant) Then
		EnregistrerCopieduFichier TourNumeroAsText
	End If
	
	'Utilisation de l'algorithme de Ronde Suisse ?
	Dim bAlgoRondeSuisse as Boolean
	theString = Feuille.getCellByPosition(4,4).String ' Le choix de l'algorithme de ronde Suisse est en cellule E5 soit (4,4)
	If (theString = "X" or theString = "x") Then
		bAlgoRondeSuisse = true
	Else
		bAlgoRondeSuisse = false
	END If
	
	'Incremente le numéro du tour courant
	Feuille.getCellByPosition(IndexColumnName+7,6).value = NouveauTourNumeroAsInteger
	
	'Cree le nom du tour courant
	TourName = "Tour"+TourNumeroAsText
	
	If (NumJoueurs < 4) Then
		Print("Nombre de joueurs < 4")
		Exit Sub
	End If
	
	dim MoyennePourcentagesJoueurDuTour as Single
	MoyennePourcentagesJoueurDuTour = 0
	
	NumJoueursHomRestant = Feuille.getCellByPosition(IndexColumnName+7,4).value
	NumJoueursFemRestant = Feuille.getCellByPosition(IndexColumnName+7,3).value

	dim TabPourcentageHommes(NumJoueursHomRestant) as Integer
	dim TabPourcentageFemmes(NumJoueursFemRestant) as Integer
	dim CompteurJoueursHommes as Integer
	dim CompteurJoueursFemmes as Integer
	CompteurJoueursHommes = 0
	CompteurJoueursFemmes = 0
	dim PourcentageJoueurCourant as Integer
	
	'Met les indices des joueurs jouant ce tour dans un tableau
	actualIndex 		= 0
	NumJoueursIndexes 	= 0
	Do While (NumJoueursIndexes < NumJoueurs)
		If JoueCeTourSansIndirection(actualIndex) Then
			IndicesDesJoueursDeCeTour(NumJoueursIndexes) 	= actualIndex
			PourcentageJoueurCourant 						= RecuperePourcentageMatchsGagnesDuJoueur(actualIndex)
		
			If (EstUnHomme(NumJoueursIndexes)) Then 'EstUnHomme utilise la valeur remplie dans IndicesDesJoueursDeCeTour donc le faire avant.
				TabPourcentageHommes(CompteurJoueursHommes) = PourcentageJoueurCourant
				CompteurJoueursHommes 						= CompteurJoueursHommes + 1
			Else
				TabPourcentageFemmes(CompteurJoueursFemmes) = PourcentageJoueurCourant
				CompteurJoueursFemmes 						= CompteurJoueursFemmes + 1
			End If
			
			MoyennePourcentagesJoueurDuTour 				= MoyennePourcentagesJoueurDuTour + PourcentageJoueurCourant
			NumJoueursIndexes 								= NumJoueursIndexes + 1
		End If
		actualIndex = actualIndex + 1
	Loop
	MoyennePourcentagesJoueurDuTour = MoyennePourcentagesJoueurDuTour / NumJoueursIndexes
	
	'print CompteurJoueursFemmes
	'print CompteurJoueursHommes
	'print "MoyennePourcentagesJoueurDuTour " +CStr(MoyennePourcentagesJoueurDuTour)
	'print NumJoueursIndexes
	'print MoyennePourcentagesJoueurDuTour
	
	
	Dim NumJoueursHommeHautClassement as Integer
	Dim NumJoueursHommeBasClassement as Integer
	Dim NumJoueursFemmeHautClassement as Integer
	Dim NumJoueursFemmeBasClassement as Integer
	NumJoueursHommeHautClassement 	= 0
	NumJoueursHommeBasClassement 	= 0
	NumJoueursFemmeHautClassement	= 0
	NumJoueursFemmeBasClassement	= 0
	If (bAlgoRondeSuisse) Then
		For i = 0 to (NumJoueursHomRestant-1)
			If (TabPourcentageHommes(i) > MoyennePourcentagesJoueurDuTour) Then
				NumJoueursHommeHautClassement = NumJoueursHommeHautClassement + 1
			Else
				NumJoueursHommeBasClassement = NumJoueursHommeBasClassement + 1
			EndIf
		Next
		For i = 0 to (NumJoueursFemRestant-1)
			If (TabPourcentageFemmes(i) > MoyennePourcentagesJoueurDuTour) Then
				NumJoueursFemmeHautClassement = NumJoueursFemmeHautClassement + 1
			Else
				NumJoueursFemmeBasClassement = NumJoueursFemmeBasClassement + 1
			EndIf
		Next
	End If 
	'Print "NumJoueursHommeHautClassement = "+ CStr(NumJoueursHommeHautClassement)
	'Print "NumJoueursHommeBasClassement = "+ CStr(NumJoueursHommeBasClassement)
	'Print "NumJoueursFemmeHautClassement = "+ CStr(NumJoueursFemmeHautClassement)
	'Print "NumJoueursFemmeBasClassement = "+ CStr(NumJoueursFemmeBasClassement)
	
	'Prend les joueurs par 4 pour faire un match
	Dim NumCompletants as Integer
	NumCompletants 	= NumJoueurs Mod 4
	NumMatches 		= NumJoueurs \ 4
	If (NumCompletants  > 0) Then
		NumMatches = NumMatches + 1
	End If
	'Print NumMatches
	
	NumJoueursHomCompletantRestant  = NumJoueursHomRestant
	NumJoueursFemCompletantRestant	= NumJoueursFemRestant
	
	'Ces 2 tableaux n'utilisent pas les indices réels des joueurs mais leur indice dans la liste virtuelle des joueurs qui jouent ce tour
	for i = 0 to 99
		JoueurNumeroMatch(2*i) 		= -1 'None
		JoueurNumeroMatch(2*i+1) 	= -1 'None
	Next i
	
	'Créé une nouvelle page pour ce tour
	sheets 		= Classeur.Sheets
	position 	= sheets.count
	sheets.insertNewByName(TourName, position)
	'exists 	= sheets.hasByName(TourName)
	Tour 		= Classeur.Sheets.GetByName(TourName)
	
	'Cree le bouton pour calculer le score
	CreeBoutonScoreDuTour(Tour, TourName)
	Tour.getCellByPosition(5,1).setString("Matches")
	Tour.getCellByPosition(6,1).Value	= NumMatches
	
	Dim NumJoueursSet as Integer
	NumJoueursSet = 0
	
	Dim NumMatchesReady as Integer
	NumMatchesReady = 0
	Do While NumMatchesReady < NumMatches	
	    dim pairTeam1 as Object
	    dim pairTeam2 as Object
	    dim indexJoueur1Team1 as Integer, indexJoueur2Team1 as Integer
	    dim indexJoueur1Team2 as Integer, indexJoueur2Team2 as Integer
	    dim bIsJoueur1HommeTeam1 as Boolean, bIsJoueur2HommeTeam1 as Boolean
		dim bIsJoueur1HommeTeam2 as Boolean, bIsJoueur2HommeTeam2 as Boolean
		dim bIsJoueur1CompletantTeam1 as Boolean, bIsJoueur2CompletantTeam1 as Boolean
		dim bIsJoueur1CompletantTeam2 as Boolean, bIsJoueur2CompletantTeam2 as Boolean
		Dim textColor As Integer
			
		'L'algo ronde Suisse choisit un joueur puis essaie de prendre son partenaire et leurs adversaires dans le haut ou le bas du tableau afin que les niveaux soient plus ou moins similaires.
		'L'idée est que les joueurs du haut du tableau jouent avec ceux du haut du tableau et idem pour ceux du bas du tableau afin d'équilibrer les matches.
		If (bAlgoRondeSuisse) Then
			Dim bJoueursDuHautDuTableau As Boolean
			bJoueursDuHautDuTableau = False
			Dim bPremierePaire as Boolean
			
			'La fonction PaireMixteAlgoRondeSuisse peut modifier les valeurs de NumJoueursHommeHautClassement, NumJoueursHommeBasClassement, NumJoueursFemmeHautClassement, NumJoueursFemmeBasClassement et bJoueursDuHautDuTableau
			bPremierePaire			= True
			pairTeam1 				= PaireMixteAlgoRondeSuisse(NumMatchesReady, MoyennePourcentagesJoueurDuTour, NumJoueursHommeHautClassement, NumJoueursHommeBasClassement, NumJoueursFemmeHautClassement, NumJoueursFemmeBasClassement, bJoueursDuHautDuTableau, bPremierePaire)
			'bJoueursDuHautDuTableau a été mis à jour dans la fonction précédente
			bPremierePaire			= False
		    pairTeam2 				= PaireMixteAlgoRondeSuisse(NumMatchesReady, MoyennePourcentagesJoueurDuTour, NumJoueursHommeHautClassement, NumJoueursHommeBasClassement, NumJoueursFemmeHautClassement, NumJoueursFemmeBasClassement, bJoueursDuHautDuTableau, bPremierePaire)
		Else
		    pairTeam1 				= PaireMixteAuHasard(NumMatchesReady)
		    pairTeam2 				= PaireMixteAuHasard(NumMatchesReady)    
	    EndIf
	    
    	indexReelJoueur1Team1 		= pairTeam1(0)
	    bIsJoueur1HommeTeam1		= pairTeam1(1)
	    bIsJoueur1CompletantTeam1 	= pairTeam1(2)
	    indexReelJoueur2Team1 		= pairTeam1(3)
	    bIsJoueur2HommeTeam1		= pairTeam1(4)
	    bIsJoueur2CompletantTeam1 	= pairTeam1(5)
	    
	    indexReelJoueur1Team2 		= pairTeam2(0)
	    bIsJoueur1HommeTeam2		= pairTeam2(1)
	    bIsJoueur1CompletantTeam2 	= pairTeam2(2)
	    indexReelJoueur2Team2 		= pairTeam2(3)
	    bIsJoueur2HommeTeam2		= pairTeam2(4)
	    bIsJoueur2CompletantTeam2 	= pairTeam2(5)
	    
	    Dim NumeroMatchCell As Object

		NumeroMatchCell	 			= Tour.getCellByPosition(IndexColumnName - 1, indexLigneDstJoueurs+NumJoueursSet)
		NumeroMatchCell.String		= "Match "+CStr(NombreDeMatcheJouesEnTout + NumMatchesReady + 1)
		
	    'Copy first team
	    CopyData(indexLigneSrcJoueurs + indexReelJoueur1Team1, indexLigneDstJoueurs+NumJoueursSet, IndexColumnName, GetTextColor(bIsJoueur1HommeTeam1, bIsJoueur1CompletantTeam1), bIsJoueur1CompletantTeam1)
	    NumJoueursSet = NumJoueursSet +1
	    CopyData(indexLigneSrcJoueurs + indexReelJoueur2Team1, indexLigneDstJoueurs+NumJoueursSet, IndexColumnName, GetTextColor(bIsJoueur2HommeTeam1, bIsJoueur2CompletantTeam1), bIsJoueur2CompletantTeam1)
	   	NumJoueursSet = NumJoueursSet + 1
	   	'Copy second team
	   	CopyData(indexLigneSrcJoueurs + indexReelJoueur1Team2, indexLigneDstJoueurs+NumJoueursSet-2, IndexColumnName+5, GetTextColor(bIsJoueur1HommeTeam2, bIsJoueur1CompletantTeam2), bIsJoueur1CompletantTeam2)
	    NumJoueursSet = NumJoueursSet + 1 
	    CopyData(indexLigneSrcJoueurs + indexReelJoueur2Team2, indexLigneDstJoueurs+NumJoueursSet-2, IndexColumnName+5, GetTextColor(bIsJoueur2HommeTeam2, bIsJoueur2CompletantTeam2), bIsJoueur2CompletantTeam2)
	   	NumJoueursSet = NumJoueursSet + 1
	   	NumMatchesReady = NumMatchesReady + 1
	   	
	   	'Hide the columns with the real index
	   	oColumns = Tour.Columns
		oColumn = oColumns.GetByIndex(IndexColumnName+2)
		oColumn.IsVisible = False
		oColumn = oColumns.GetByIndex(IndexColumnName+7)
		oColumn.IsVisible = False
	Loop
	
	'Met à jour le nombre de matches joués par anticipation, ils ne sont pas encore joués.
	Feuille.getCellByPosition(IndexColumnName+9,6).Value = NombreDeMatcheJouesEnTout + NumMatchesReady
	
	'Selectionne la page créée
	Controller = ThisComponent.getcurrentController
	Controller.setActiveSheet(Tour)
End Sub

Sub EnregistrerCopieduFichier(NumeroTour as String)
    ' Déclaration des variables
    Dim cheminComplet As String
    Dim dossier As String
    Dim nomFichierAvecExtension As String
    Dim nomFichier As String
    Dim extension As String
    Dim positionDernierSeparateur As Integer
    
    ' Récupération du chemin complet du fichier actif
    cheminComplet = ThisComponent.getURL()
    
    ' Recherche de la position du dernier séparateur de dossier
    positionDernierSeparateur = LastOccurence (cheminComplet, "/")
    
    ' Extraction du dossier
    dossier = Left(cheminComplet, positionDernierSeparateur - 1)
    
    ' Extraction du nom de fichier avec extension
    nomFichierAvecExtension = Mid(cheminComplet, positionDernierSeparateur + 1)
    positionDernierSeparateur = LastOccurence(nomFichierAvecExtension, ".")-1 'Sur le Point
    nomFichier = Left(nomFichierAvecExtension, positionDernierSeparateur)
    extension = Right(nomFichierAvecExtension, 4)
    
    Dim NouveauNom as String
    NouveauNom = (dossier & "/" & nomFichier & "_AvantTour" & NumeroTour & extension )
    
    ' Enregistre une copie du fichier avec le nom du tour inclus
	dim document as object
	dim dispatcher as object
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	dim args1(1) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = NouveauNom
	args1(1).Name = "FilterName"
	args1(1).Value = "calc8"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	'Sauve sous le nom original pour retrouver le fichier d'origine
	args1(0).Value = cheminComplet
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
End Sub

Function LastOccurence(strString As String, strCharacter As String) As Integer

    Dim intPosition As Integer
    
    intPosition = 1
    
    While intPosition <= Len(strString) And strCharacter <> "" And InStr(intPosition, strString, strCharacter) <> 0
        intPosition = InStr(intPosition, strString, strCharacter)
        LastOccurence = intPosition
        intPosition = intPosition + 1
    Wend
End Function

Function RecuperePourcentageMatchsGagnesDuJoueur(ligneRelative as Integer) as Integer
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
  	ColonneNom 				= 1
  	RecuperePourcentageMatchsGagnesDuJoueur = oInscriptionPage.getCellByPosition(ColonneNom + 8, indexLigneSrcJoueurs + ligneRelative).Value	'indexLigneSrcJoueurs est le début des lignes des joueurs
End Function

Function RecupereNomDuJoueur(ligne as Integer) as String
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
  	ColonneNom 				= 1
  	Dim NomJoueurCourant as String
  	NomJoueurCourant  		= oInscriptionPage.getCellByPosition(ColonneNom  , indexLigneSrcJoueurs + ligne).String 'indexLigneSrcJoueurs est le début des lignes des joueurs
	PrenomJoueurCourant  	= oInscriptionPage.getCellByPosition(ColonneNom+1, indexLigneSrcJoueurs + ligne).String
	'print NomJoueurCourant
	RecupereNomDuJoueur		= PrenomJoueurCourant + " " + NomJoueurCourant
End Function

Function GetTextColor(bIshomme as Boolean, bIsCompletant As Boolean) as Long
	If bIsCompletant then
    	GetTextColor = 204 ' is x0000CC blue completant
    Else
    	If bIshomme Then
    		GetTextColor = 0 'Black color for men
    	Else
    		'GetTextColor =  15037813 ' Is #cbf5cb Pale red
    		GetTextColor = RGB(148, 0, 211) 'Purple for women
    	End If
    End If
End Function

Function UpdateNombreJoueursRestant(bIshomme as Boolean, bIsCompletant as Boolean)
	If bIsCompletant then
		If bIshomme then
	    	NumJoueursHomCompletantRestant = NumJoueursHomCompletantRestant - 1
	    Else
	    	NumJoueursFemCompletantRestant = NumJoueursFemCompletantRestant - 1
	    End If
	Else
		If bIshomme Then
	    	NumJoueursHomRestant = NumJoueursHomRestant - 1
	    Else
	    	NumJoueursFemRestant = NumJoueursFemRestant - 1
	    End If
	End If
End function

Function CopyData(theSrcLine As Integer, theDstLine As Integer, theDstColumn As Integer, theColor as Long, bJoueurCompletant as Boolean)
	Dim Nom_Src As Object, Prenom_Src As Object
	Dim Nom_Dst As Object, Prenom_Dst As Object

	Nom_Src 			= Feuille.getCellByPosition(IndexColumnName,theSrcLine)
 	Prenom_Src 			= Feuille.getCellByPosition(IndexColumnName+1,theSrcLine)
 	Nom_Dst 			= Tour.getCellByPosition(theDstColumn,theDstLine)
 	Prenom_Dst 			= Tour.getCellByPosition(theDstColumn+1,theDstLine)
 	IndexReel_Dst 		= Tour.getCellByPosition(theDstColumn+2,theDstLine)
 	
 	Nom_Dst.DataArray 		= Nom_Src.DataArray
	Prenom_Dst.DataArray 	= Prenom_Src.DataArray
	if (bJoueurCompletant) Then
		IndexReel_Dst.Value = -1 'Pour ne pas compter les points de son score
	Else
		IndexReel_Dst.Value = theSrcLine
	End If
	Nom_Dst.CharColor 		= theColor
	Prenom_Dst.CharColor 	= theColor

End Function

Function JoueCeTourSansIndirection(theIndex As Integer) As Boolean
	Dim theCell as Object
	
	'Ne pas utiliser IndicesDesJoueursDeCeTour puisque c'est pour le remplir qu'on appelle cette fonction
	'theCell = Feuille.getCellByPosition(IndexColumnName+1,theIndex + indexLigneSrcJoueurs)
	'print theIndex
	'print theCell.String
	theCell = Feuille.getCellByPosition(IndexColumnName+3, theIndex + indexLigneSrcJoueurs)
	
	Dim theString as String
	theString = theCell.String
	If (theString = "X" or theString = "x") Then
		JoueCeTourSansIndirection = true
	Else
		JoueCeTourSansIndirection = false
	END If
End Function

Function EstUnHomme(theIndex As Integer) As Boolean
	Dim theCell as Object
	
	Dim vraiIndex as Integer
	vraiIndex = IndicesDesJoueursDeCeTour(theIndex)
	
	'theCell = Feuille.getCellByPosition(IndexColumnName+1,vraiIndex + indexLigneSrcJoueurs)
	'print theIndex
	'print theCell.String
	theCell = Feuille.getCellByPosition(IndexColumnName+2, vraiIndex + indexLigneSrcJoueurs)
	
	Dim theString as String
	theString = theCell.String
	If (theString = "H" or theString = "h") Then
		EstUnHomme = true
	Else
		EstUnHomme = false
	END If
End Function

'Retourne un tableau avec 2 éléments {indexJoueur, bIsJoueurCompletant} 
Function ChoisisJoueur(bOnVeutUnHomme as Boolean, numeroMatchCourant as Integer) As Array
	Do While true
		Dim Hasard as Integer
		Dim bIsJoueurCompletant as Boolean
	    'Dim bJoue as Boolean
	    
		'Int((maxVal - minVal + 1) * Rnd + minVal)
	 	Hasard = Int((NumJoueurs) * Rnd) REM entre 0 et NumJoueurs-1
	    'Print Hasard
	     
	    'Est on en train de choisir un Completant car tous les joueurs sont déjà pris pour ce tour ?
	    If (JoueursEncoreDisponibles())Then
	    	If (bOnVeutUnHomme and (NumJoueursHomRestant = 0)) Then
	    		bIsJoueurCompletant = True 
	    	Else If ((bOnVeutUnHomme = False) and (NumJoueursFemRestant = 0))Then
	    			bIsJoueurCompletant = True 
	    		 End If
	    	End If
	    Else
	    	bIsJoueurCompletant = True  
	    End If
	    
	    If ((JoueurDisponible(Hasard, bIsJoueurCompletant, numeroMatchCourant))) Then
	    	Dim bIsMan as Boolean
	    	bIsMan = EstUnHomme(Hasard)
		 	If bOnVeutUnHomme and bIsMan Then
		 		'Is a man
		 		UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
		 		ChoisisJoueur = Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
		 		Exit Do
		 	Else 	If ((bOnVeutUnHomme = false) and (bIsMan = false))Then
		 				'Is a woman
		 				UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
		 				ChoisisJoueur =  Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
		 				Exit Do
		 			End If
		 	End If
	 	End If
	Loop
End Function

'Retourne un tableau avec 2 éléments {indexJoueur, bIsJoueurCompletant} 
Function ChoisisJoueurAlgoRondSuisse(bOnVeutUnHomme as Boolean, numeroMatchCourant as Integer, MoyennePourcentagesJoueurDuTour as Integer, Optional ByRef NumJoueursHommeHautClassement as Integer, Optional ByRef NumJoueursHommeBasClassement as Integer, Optional ByRef NumJoueursFemmeHautClassement as Integer, Optional ByRef NumJoueursFemmeBasClassement as Integer, ByVal bJoueursDuHautDuTableau as Boolean) As Array
	Do While true
		Dim Hasard as Integer
		Dim bIsJoueurCompletant as Boolean
	  
	 	Hasard = Int((NumJoueurs) * Rnd) REM entre 0 et NumJoueurs-1
	    'Print Hasard
	     
	    'Est on en train de choisir un Completant car tous les joueurs sont déjà pris pour ce tour ?
	    If (JoueursEncoreDisponibles())Then
	    	If (bOnVeutUnHomme and (NumJoueursHomRestant = 0)) Then
	    		bIsJoueurCompletant = True 
	    	Else If ((bOnVeutUnHomme = False) and (NumJoueursFemRestant = 0))Then
	    			bIsJoueurCompletant = True 
	    		 End If
	    	End If
	    Else
	    	bIsJoueurCompletant = True  
	    End If  
	    
	    If ((JoueurDisponible(Hasard, bIsJoueurCompletant, numeroMatchCourant))) Then
	    	Dim bEstUnHomme as Boolean
	    	Dim bJoueurcourantDansHautTableau as Boolean
	    	bEstUnHomme 					= EstUnHomme(Hasard)
	    	PourcentageMatchesGagnes 		= RecuperePourcentageMatchsGagnesDuJoueur(IndicesDesJoueursDeCeTour(Hasard))
	    	bJoueurcourantDansHautTableau	= (PourcentageMatchesGagnes > MoyennePourcentagesJoueurDuTour)
	    				
		 	If bOnVeutUnHomme and bEstUnHomme Then
		 		'C'est un homme
		 		If (bJoueursDuHautDuTableau)Then
		 			'On cherche un homme dans le haut du tableau 
		 			If (NumJoueursHommeHautClassement > 0) Then
		 				'Donc il en reste...
		 				If (bJoueurcourantDansHautTableau) Then
		 					'Cet homme fait partie du haut du tableau, c'est ce que l'on cherche
		 					UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
					 		ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
					 		NumJoueursHommeHautClassement 	= NumJoueursHommeHautClassement - 1
					 		Exit Do
		 				End If
		 			Else
		 				'Il n'y a plus d'hommes dans le haut du tableau, donc prend celui-ci qui est à priori dans le bas du tableau
		 				UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
					 	ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
					 	NumJoueursHommeBasClassement 	= NumJoueursHommeBasClassement - 1
					 	Exit Do
		 			End If
		 		Else
		 			'On cherche un homme dans le bas du tableau
		 			If (NumJoueursHommeBasClassement > 0) Then
		 				'Donc il en reste...
		 				If (bJoueurcourantDansHautTableau = False) Then
		 					'Cet homme fait partie du bas du tableau, c'est ce que l'on cherche
		 					UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
					 		ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
					 		NumJoueursHommeBasClassement 	= NumJoueursHommeBasClassement - 1
					 		Exit Do
		 				End If
		 			Else
		 				'Il n'y a plus d'hommes dans le bas du tableau, donc prend celui-ci qui est à priori dans le haut du tableau
		 				UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
					 	ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
					 	NumJoueursHommeHautClassement 	= NumJoueursHommeHautClassement - 1
					 	Exit Do
		 			End If
		 		End If
		 		
		 	
		 	Else 	If ((bOnVeutUnHomme  = false) and (bEstUnHomme = false))Then
		 				'C'est une femme
		 				If (bJoueursDuHautDuTableau)Then
				 			'On cherche une femme dans le haut du tableau 
				 			If (NumJoueursFemmeHautClassement > 0) Then
				 				'Donc il en reste...
				 				If (bJoueurcourantDansHautTableau) Then
				 					'Cette femme fait partie du haut du tableau, c'est ce que l'on cherche
				 					UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
							 		ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
							 		NumJoueursFemmeHautClassement 	= NumJoueursFemmeHautClassement - 1
							 		Exit Do
				 				End If
				 			Else
				 				'Il n'y a plus de femmes dans le haut du tableau, donc prend celle-ci qui est à priori dans le bas du tableau
				 				UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
							 	ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
							 	NumJoueursFemmeBasClassement 	= NumJoueursFemmeBasClassement - 1
							 	Exit Do
				 			End If
				 		Else
				 			'On cherche une femme dans le bas du tableau
				 			If (NumJoueursFemmeBasClassement > 0) Then
				 				'Donc il en reste...
				 				If (bJoueurcourantDansHautTableau = False) Then
				 					'Cette femme fait partie du bas du tableau, c'est ce que l'on cherche
				 					UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
							 		ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
							 		NumJoueursFemmeBasClassement 	= NumJoueursFemmeBasClassement - 1
							 		Exit Do
				 				End If
				 			Else
				 				'Il n'y a plus de femmes dans le bas du tableau, donc prend celle-ci qui est à priori dans le haut du tableau
				 				UpdateJoueurChoisi(Hasard, bIsJoueurCompletant, numeroMatchCourant)
							 	ChoisisJoueurAlgoRondSuisse 	= Array(IndicesDesJoueursDeCeTour(Hasard), bIsJoueurCompletant)'On renvoit le vrai index du joueur
							 	NumJoueursFemmeHautClassement 	= NumJoueursFemmeHautClassement - 1
							 	Exit Do
				 			End If
				 		End If
		 			End If
		 		End If
		 	End If
	Loop
End Function

Function UpdateJoueurChoisi(index as Integer, bIsCompletant as Boolean, numeroMatchCourant as Integer)
	if (JoueurNumeroMatch(2 * index) = -1) Then
		JoueurNumeroMatch(2 * index) = numeroMatchCourant 'C'est son premier match dans ce tour
	Else
		JoueurNumeroMatch(2*index + 1) = numeroMatchCourant 'C'est son 2e match dans ce tour dont 1 pour compléter
	End If
	
End Function

Function JoueurDisponible(indexJoueur As Integer, bCompletant as Boolean, numeroMatchCourant as Integer) As Boolean
	
	if ((JoueurNumeroMatch(2*indexJoueur) = numeroMatchCourant) Or (JoueurNumeroMatch(2*indexJoueur+1) = numeroMatchCourant))Then
		JoueurDisponible = false'Le joueur joue déjà dans ce match
	Else
		If bCompletant then
			JoueurDisponible = (JoueurNumeroMatch(2*indexJoueur + 1) = -1)
		Else
			JoueurDisponible = (JoueurNumeroMatch(2*indexJoueur + 0) = -1)
		End If
	End If
End Function

Function JoueursEncoreDisponibles() As Boolean
	JoueursEncoreDisponibles = (NumJoueursHomRestant + NumJoueursFemRestant) > 0
End Function

'Retourne un tableau avec 6 éléments {indexJoueur1, bIsJoueur1Homme, bIsJoueur1Completant, indexJoueur2, bIsJoueur2Homme, bIsJoueur2Completant} 
Public Function PaireMixteAuHasard(numeroMatchCourant as Integer) As Array
	Dim indexJoueur1 as Integer, indexJoueur2 as Integer
	dim bIsJoueur1Homme as Boolean, bIsJoueur2Homme as Boolean
	dim bIsJoueur1Completant as Boolean, bIsJoueur2Completant as Boolean
	Dim tabIndexAndCompletant1 As Object, tabIndexAndCompletant2 As Object
	
	'On essaye de choisir un homme ici
	If (JoueursEncoreDisponibles())Then
		bIsJoueur1Homme		= (NumJoueursHomRestant > 0) or (NumJoueursHomCompletantRestant > 0)
	Else
		bIsJoueur1Homme		= (NumJoueursHomCompletantRestant > 0)
	End If
	
	'ChoisisJoueur retourne un tableau avec 2 éléments {indexJoueur, bIsJoueurCompletant}
	tabIndexAndCompletant1 	= ChoisisJoueur(bIsJoueur1Homme, numeroMatchCourant)
	indexJoueur1			= tabIndexAndCompletant1(0)
	bIsJoueur1Completant	= tabIndexAndCompletant1(1)
	
	UpdateNombreJoueursRestant(bIsJoueur1Homme, bIsJoueur1Completant)
	
	'On essaye de choisir une femme pour compléter la paire
	If (JoueursEncoreDisponibles())Then
		If (NumJoueursFemRestant > 0) Then
			bIsJoueur2Homme	= false 'On choisit une femme
		Else
			bIsJoueur2Homme	= true 'On choisit un homme
		End If
	Else
		'Completants
		If bIsJoueur1Homme then
			'Le premier joueur de la paire est un homme
			If (NumJoueursFemCompletantRestant > 0) Then
				bIsJoueur2Homme	= false 'Choose a woman
			Else
				bIsJoueur2Homme	= true 'Choose a man
			End If
		Else
			'Le premier joueur de la paire est une femme
			If (NumJoueursHomCompletantRestant > 0) Then
				bIsJoueur2Homme	= true 'Choose a man
			Else
				bIsJoueur2Homme	= false 'Choose a woman
			End If
		End If
			End If
	
	tabIndexAndCompletant2	= ChoisisJoueur(bIsJoueur2Homme, numeroMatchCourant)
	indexJoueur2			= tabIndexAndCompletant2(0)
	bIsJoueur2Completant 	= tabIndexAndCompletant2(1)
	
	UpdateNombreJoueursRestant(bIsJoueur2Homme, bIsJoueur2Completant)
	
	PaireMixteAuHasard 	= Array(indexJoueur1, bIsJoueur1Homme, bIsJoueur1Completant, indexJoueur2, bIsJoueur2Homme, bIsJoueur2Completant)
End Function


'Retourne un tableau avec 6 éléments {indexJoueur1, bIsJoueur1Homme, bIsJoueur1Completant, indexJoueur2, bIsJoueur2Homme, bIsJoueur2Completant} 
'L'algo ronde Suisse choisit un joueur puis essaie de prendre son partenaire et leurs adversaires dans le haut ou le bas du tableau afin que les niveaux soient plus ou moins similaires.
'L'idée est que les joueurs du haut du tableau jouent avec ceux du haut du tableau et idem pour ceux du bas du tableau afin d'équilibrer les matches.
Public Function PaireMixteAlgoRondeSuisse(numeroMatchCourant as Integer, MoyennePourcentagesJoueurDuTour as Integer, Optional ByRef NumJoueursHommeHautClassement as Integer, Optional ByRef NumJoueursHommeBasClassement as Integer, Optional ByRef NumJoueursFemmeHautClassement, Optional ByRef NumJoueursFemmeBasClassement as Integer, Optional ByRef bJoueursDuHautDuTableau as Boolean, ByVal bPremierePaire as Boolean) As Array
	Dim indexJoueur1 as Integer, indexJoueur2 as Integer
	dim bIsJoueur1Homme as Boolean, bIsJoueur2Homme as Boolean
	dim bIsJoueur1Completant as Boolean, bIsJoueur2Completant as Boolean
	Dim tabIndexAndCompletant1 As Object, tabIndexAndCompletant2 As Object
	
	'On essaye de choisir un homme ici
	If (JoueursEncoreDisponibles())Then
		bIsJoueur1Homme		= (NumJoueursHomRestant > 0) or (NumJoueursHomCompletantRestant > 0)
	Else
		bIsJoueur1Homme		= (NumJoueursHomCompletantRestant > 0)
	End If
	
	'ChoisisJoueur retourne un tableau avec 2 éléments {indexJoueur, bIsJoueurCompletant}
	If (bPremierePaire)Then
		'Premier joueur de la première paire, n'utilise pas l'algo rond Suisse, prend au hasard
		tabIndexAndCompletant1 					= ChoisisJoueur(bIsJoueur1Homme, numeroMatchCourant) 'Au hasard, ne tient pas compte du classement
		indexJoueur1							= tabIndexAndCompletant1(0)
		bIsJoueur1Completant					= tabIndexAndCompletant1(1)
		'print "indexJoueur1 " + CStr(indexJoueur1)
		'print "Nom " + RecupereNomDuJoueur(indexJoueur1)
		
		PourcentageMatchesGagnesJoueurCourant 	= RecuperePourcentageMatchsGagnesDuJoueur(indexJoueur1)
		'print "PourcentageMatchesGagnesJoueurCourant "+ CStr(PourcentageMatchesGagnesJoueurCourant)
		bJoueursDuHautDuTableau					= PourcentageMatchesGagnesJoueurCourant > MoyennePourcentagesJoueurDuTour 'Met à jour cette variable si on prend des joueurs du haut ou du bas du tableau
		If (bJoueursDuHautDuTableau) Then
			'On est dans le haut du tableau
			If(bIsJoueur1Homme) Then
				NumJoueursHommeHautClassement = NumJoueursHommeHautClassement - 1
			Else
				NumJoueursFemmeHautClassement = NumJoueursFemmeHautClassement - 1
			End If
		Else
			'On est dans le bas du tableau
			If(bIsJoueur1Homme) Then
				NumJoueursHommeBasClassement = NumJoueursHommeBasClassement - 1
			Else
				NumJoueursFemmeBasClassement = NumJoueursFemmeBasClassement - 1
			End If
		End If
	Else
		'Algo ronde Suisse, tient compte du pourcentage de matches gagnés pour le choix du joueur
		tabIndexAndCompletant1 	= ChoisisJoueurAlgoRondSuisse(bIsJoueur1Homme, numeroMatchCourant, MoyennePourcentagesJoueurDuTour, NumJoueursHommeHautClassement, NumJoueursHommeBasClassement, NumJoueursFemmeHautClassement, NumJoueursFemmeBasClassement, bJoueursDuHautDuTableau) 
		indexJoueur1			= tabIndexAndCompletant1(0)
		bIsJoueur1Completant	= tabIndexAndCompletant1(1)
	End If	
	
	UpdateNombreJoueursRestant(bIsJoueur1Homme, bIsJoueur1Completant)
	
	'On essaye de choisir une femme pour compléter la paire
	If (JoueursEncoreDisponibles())Then
		If (NumJoueursFemRestant > 0) Then
			bIsJoueur2Homme	= false 'On choisit une femme
		Else
			bIsJoueur2Homme	= true 'On choisit un homme
		End If
	Else
		'Completants
		If bIsJoueur1Homme then
			'Le premier joueur de la paire est un homme
			If (NumJoueursFemCompletantRestant > 0) Then
				bIsJoueur2Homme	= false 'Choose a woman
			Else
				bIsJoueur2Homme	= true 'Choose a man
			End If
		Else
			'Le premier joueur de la paire est une femme
			If (NumJoueursHomCompletantRestant > 0) Then
				bIsJoueur2Homme	= true 'Choose a man
			Else
				bIsJoueur2Homme	= false 'Choose a woman
			End If
		End If
			End If
	
	tabIndexAndCompletant2	= ChoisisJoueurAlgoRondSuisse(bIsJoueur2Homme, numeroMatchCourant, MoyennePourcentagesJoueurDuTour, NumJoueursHommeHautClassement, NumJoueursHommeBasClassement, NumJoueursFemmeHautClassement, NumJoueursFemmeBasClassement, bJoueursDuHautDuTableau)
	indexJoueur2			= tabIndexAndCompletant2(0)
	bIsJoueur2Completant 	= tabIndexAndCompletant2(1)
	
	UpdateNombreJoueursRestant(bIsJoueur2Homme, bIsJoueur2Completant)
	
	PaireMixteAlgoRondeSuisse 	= Array(indexJoueur1, bIsJoueur1Homme, bIsJoueur1Completant, indexJoueur2, bIsJoueur2Homme, bIsJoueur2Completant)
End Function

Function CreeBoutonScoreDuTour(PageDuTour as Object, TourName as String)
  oDrawPage = PageDuTour.DrawPage
  sScriptURL = "vnd.sun.star.script:Standard.Modulerondesuisse.ButtonPushEvent?language=Basic&location=document"
  oButtonModel = AddNewButton(("Score du "+TourName), ("Score du "+TourName), ThisComponent, oDrawPage)
  oForm = oDrawPage.getForms().getByIndex(0)
  'Ffind index inside the form container
  nIndex = GetIndex(oButtonModel, oForm)
  AssignAction(nIndex, sScriptURL, oForm)
	
End Function

' assign sScriptURL event as css.awt.XActionListener::actionPerformed.
' event is assigned to the control described by the nIndex in the oForm container
Sub AssignAction(nIndex As Integer, sScriptURL As String, oForm As Object)
  aEvent = CreateUnoStruct("com.sun.star.script.ScriptEventDescriptor")
  With aEvent
    .AddListenerParam 	= ""
    .EventMethod 		= "actionPerformed"
    .ListenerType 		= "XActionListener"
    .ScriptCode 		= sScriptURL
    .ScriptType 		= "Script"
  End With
  oForm.registerScriptEvent(nIndex, aEvent)
End Sub


Function AddNewButton(sName As String, sLabel As String, oDoc As Object, oDrawPage As Object) As Object
  oControlShape = oDoc.createInstance("com.sun.star.drawing.ControlShape")
  aPoint 		= CreateUnoStruct("com.sun.star.awt.Point")
  aSize 		= CreateUnoStruct("com.sun.star.awt.Size")
  aPoint.X 		= 20000
  aPoint.Y 		= 1000
  aSize.Width 	= 3000
  aSize.Height 	= 1000
  oControlShape.setPosition(aPoint)
  oControlShape.setSize(aSize)
  oButtonModel 			= CreateUnoService("com.sun.star.form.component.CommandButton")
  oButtonModel.Name 	= sName
  oButtonModel.Label 	= sLabel

  oControlShape.setControl(oButtonModel)
  oDrawPage.add(oControlShape)

  AddNewButton = oButtonModel
End Function

Function GetIndex(oControl As Object, oForm As Object) As Integer
  Dim nIndex As Integer
  nIndex = -1
  For i = 0 To oForm.getCount() - 1 step 1
    If EqualUnoObjects(oControl, oForm.getByIndex(i)) Then
      nIndex = i
      Exit For
    End If
  Next
  GetIndex = nIndex
End Function

Sub ButtonPushEvent(ev as com.sun.star.awt.ActionEvent)
  oPage 					= ThisComponent.getCurrentController().getActiveSheet() 'Get active sheet
  oTourName 				= oPage.getName() 'Get its name
  oNumeroTour 				= CInt(Right(oTourName, Len(oTourName) - 4))
  'Print (oNumeroTour)
  oInscriptionPage 			= ThisComponent.Sheets.GetByName("Inscriptions")
  StartColumnIndexOfTours 	= 10
  StartLineIndexOfTours 	= 8
  ColonneDuTourDansPageInscriptions	= StartColumnIndexOfTours+oNumeroTour-1
  oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, StartLineIndexOfTours).String = oTourName
  
  'Get Number of matches
  oNumMatches = oPage.getCellByPosition(6, 1).Value
  'Print (oNumMatches)
  CalculeScoresDuTour(oPage, oNumMatches, ColonneDuTourDansPageInscriptions) 
End Sub

Sub CalculeScoresDuTour(oTourPage as Object, NumMatches as Integer, ColonneDuTourDansPageInscriptions as Integer)
  'Calcule les scores
   For i = 0 To (NumMatches-1) step 1
   	 CalculeScoreDuMatch(oTourPage, i, ColonneDuTourDansPageInscriptions)
   Next
   ''Cree la page de classement
   CreePageClassement() 'Classement par nombre de matches gagnés, puis sets et points. Ne marche que pour un tournoi ou tout le monde joue tout le temps. Si qqun fait une pause o arrive en cours de route, il est désavantagé.
   CreePageClassementParPourcentage() 'Classement par pourcentage de matches gagnés
End Sub

Sub CreePageClassement()'Classement par nombre de matches gagnés, puis sets et points
  	indexLigneSrcJoueurs 	= 9
  	Classeur 				= ThisComponent
  	sheets 					= Classeur.Sheets
  	'Efface la page classement si elle existe
	If (sheets.hasByName("Classement"))Then
	  sheets.removeByName("Classement")
	End If
	sheets.insertNewByName("Classement", sheets.count)'insert à la fin
	FeuilleClassement 	= Classeur.Sheets.GetByName("Classement")
	
	'Créé la légende Nom/Prénom/Score/Sets/Points
	FeuilleClassement.getCellByPosition(0, indexLigneSrcJoueurs-1).String = "Classement"
	FeuilleClassement.getCellByPosition(1, indexLigneSrcJoueurs-1).String = "Nom"
	FeuilleClassement.getCellByPosition(2, indexLigneSrcJoueurs-1).String = "Prénom"
	FeuilleClassement.getCellByPosition(3, indexLigneSrcJoueurs-1).String = "Score"
	FeuilleClassement.getCellByPosition(4, indexLigneSrcJoueurs-1).String = "Sets"
	FeuilleClassement.getCellByPosition(5, indexLigneSrcJoueurs-1).String = "Delta Points"
	FeuilleClassement.getCellByPosition(6, indexLigneSrcJoueurs-1).String = "Nbre matches"
	FeuilleClassement.getCellByPosition(7, indexLigneSrcJoueurs-1).String = "%tage matches gagnés"

	oInscriptionPage 	= ThisComponent.Sheets.GetByName("Inscriptions")
  	NumTotalJoueurs		= CalculeNombreTotalDeJoueurs()
  	Dim LignesExclues(NumTotalJoueurs) as Integer 'Array of integers
  	dim i As Integer
  	For i = 0 To (NumTotalJoueurs-1)
    	LignesExclues(i) = -1
	Next i
  	'print NumTotalJoueurs
  	JoueurCourant 		= 0
  
  	ADejaJoueUnMatchMeilleurJoueur = False
  	Do While (JoueurCourant <> NumTotalJoueurs)
  		LigneMeilleurJoueur				= PrendJoueurAvecLeMeilleurScore(NumTotalJoueurs, LignesExclues)
  		if (LigneMeilleurJoueur >= 0)Then
  			LignesExclues(JoueurCourant)= LigneMeilleurJoueur
  			RecopieClassement(JoueurCourant, LigneMeilleurJoueur)
		Else 
			Exit Do
		End If
		JoueurCourant 					= JoueurCourant + 1 
		'print LigneMeilleurJoueur
	Loop
	
	'Selectionne la page créée
	Controller = ThisComponent.getcurrentController
	Controller.setActiveSheet(FeuilleClassement)
End Sub

Sub CreePageClassementParPourcentage()'Classement par pourcentage de matches gagnés
  	indexLigneSrcJoueurs 	= 9
  	Classeur 				= ThisComponent
  	sheets 					= Classeur.Sheets
  	'Efface la page classement si elle existe
	If (sheets.hasByName("Classement %"))Then
	  sheets.removeByName("Classement %")
	End If
	sheets.insertNewByName("Classement %", sheets.count)'insert à la fin
	FeuilleClassement 	= Classeur.Sheets.GetByName("Classement %")
	
	'Créé la légende Nom/Prénom/Score/Sets/Points
	FeuilleClassement.getCellByPosition(0, indexLigneSrcJoueurs-1).String = "Classement par pourcentage"
	FeuilleClassement.getCellByPosition(1, indexLigneSrcJoueurs-1).String = "Nom"
	FeuilleClassement.getCellByPosition(2, indexLigneSrcJoueurs-1).String = "Prénom"
	FeuilleClassement.getCellByPosition(3, indexLigneSrcJoueurs-1).String = "Score"
	FeuilleClassement.getCellByPosition(4, indexLigneSrcJoueurs-1).String = "Sets"
	FeuilleClassement.getCellByPosition(5, indexLigneSrcJoueurs-1).String = "Delta Points"
	FeuilleClassement.getCellByPosition(6, indexLigneSrcJoueurs-1).String = "Nbre matches"
	FeuilleClassement.getCellByPosition(7, indexLigneSrcJoueurs-1).String = "%tage matches gagnés"

	oInscriptionPage 	= ThisComponent.Sheets.GetByName("Inscriptions")
  	NumTotalJoueurs		= CalculeNombreTotalDeJoueurs()
  	Dim LignesExclues(NumTotalJoueurs) as Integer 'Array of integers
  	dim i As Integer
  	For i = 0 To (NumTotalJoueurs-1)
    	LignesExclues(i) = -1
	Next i
  	'print NumTotalJoueurs
  	JoueurCourant 		= 0
  	Do While (JoueurCourant <> NumTotalJoueurs)
  		LigneMeilleurJoueur				= PrendJoueurAvecLeMeilleurPourcentage(NumTotalJoueurs, LignesExclues)
  		If (LigneMeilleurJoueur >= 0)Then
  			LignesExclues(JoueurCourant)= LigneMeilleurJoueur
			RecopieClassementPourcentage(JoueurCourant, LigneMeilleurJoueur)
		Else 
			Exit Do
		End If
		JoueurCourant 					= JoueurCourant + 1 
		'print LigneMeilleurJoueur
	Loop
End Sub

Sub RecopieClassement(JoueurCourant as Integer, LigneMeilleurJoueur as Integer)
	'print ("JoueurCourant :"+CStr(JoueurCourant)+" LigneMeilleurJoueur :"+CStr(LigneMeilleurJoueur))
	indexLigneSrcJoueurs 	= 9
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
  	FeuilleClassement 		= ThisComponent.Sheets.GetByName("Classement")
  	ColonneNom 				= 1
  	
  	NomJoueurCourant  							= oInscriptionPage.getCellByPosition(ColonneNom  , LigneMeilleurJoueur).String
	PrenomJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonneNom+1, LigneMeilleurJoueur).String
	NombreMatchesGagnesJoueurCourant  			= oInscriptionPage.getCellByPosition(ColonneNom+4, LigneMeilleurJoueur).Value
	SetsJoueurCourant  							= oInscriptionPage.getCellByPosition(ColonneNom+5, LigneMeilleurJoueur).Value
	PointsJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonneNom+6, LigneMeilleurJoueur).Value
	NombreMatchesJouesJoueurCourant  			= oInscriptionPage.getCellByPosition(ColonneNom+7, LigneMeilleurJoueur).Value
	PourcentageMatchesGagnesJoueurCourant 		= oInscriptionPage.getCellByPosition(ColonneNom+8, LigneMeilleurJoueur).Value	
	
	'Classement du joueur
	LigneClassement	= indexLigneSrcJoueurs + JoueurCourant 'JoueurCourant démarre à 0
	FeuilleClassement.getCellByPosition(0, LigneClassement).String = CStr(JoueurCourant+1)
	FeuilleClassement.getCellByPosition(1, LigneClassement).String = NomJoueurCourant
	FeuilleClassement.getCellByPosition(2, LigneClassement).String = PrenomJoueurCourant
	FeuilleClassement.getCellByPosition(3, LigneClassement).Value  = NombreMatchesGagnesJoueurCourant
	FeuilleClassement.getCellByPosition(4, LigneClassement).Value  = SetsJoueurCourant
	FeuilleClassement.getCellByPosition(5, LigneClassement).Value  = PointsJoueurCourant
	FeuilleClassement.getCellByPosition(6, LigneClassement).Value  = NombreMatchesJouesJoueurCourant
	FeuilleClassement.getCellByPosition(7, LigneClassement).Value  = PourcentageMatchesGagnesJoueurCourant
End Sub

Sub RecopieClassementPourcentage(JoueurCourant as Integer, LigneMeilleurJoueur as Integer)
	'print ("JoueurCourant :"+CStr(JoueurCourant)+" LigneMeilleurJoueur :"+CStr(LigneMeilleurJoueur))
	indexLigneSrcJoueurs 	= 9
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
  	FeuilleClassement 		= ThisComponent.Sheets.GetByName("Classement %")
  	ColonneNom 				= 1
  	
  	NomJoueurCourant  							= oInscriptionPage.getCellByPosition(ColonneNom  , LigneMeilleurJoueur).String
	PrenomJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonneNom+1, LigneMeilleurJoueur).String
	NombreMatchesGagnesJoueurCourant  			= oInscriptionPage.getCellByPosition(ColonneNom+4, LigneMeilleurJoueur).Value
	SetsJoueurCourant  							= oInscriptionPage.getCellByPosition(ColonneNom+5, LigneMeilleurJoueur).Value
	PointsJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonneNom+6, LigneMeilleurJoueur).Value
	NombreMatchesJouesJoueurCourant  			= oInscriptionPage.getCellByPosition(ColonneNom+7, LigneMeilleurJoueur).Value
	PourcentageMatchesGagnesJoueurCourant 		= oInscriptionPage.getCellByPosition(ColonneNom+8, LigneMeilleurJoueur).Value	
	
	'Classement du joueur
	LigneClassement	= indexLigneSrcJoueurs + JoueurCourant 'JoueurCourant démarre à 0
	FeuilleClassement.getCellByPosition(0, LigneClassement).String = CStr(JoueurCourant+1)
	FeuilleClassement.getCellByPosition(1, LigneClassement).String = NomJoueurCourant
	FeuilleClassement.getCellByPosition(2, LigneClassement).String = PrenomJoueurCourant
	FeuilleClassement.getCellByPosition(3, LigneClassement).Value  = NombreMatchesGagnesJoueurCourant
	FeuilleClassement.getCellByPosition(4, LigneClassement).Value  = SetsJoueurCourant
	FeuilleClassement.getCellByPosition(5, LigneClassement).Value  = PointsJoueurCourant
	FeuilleClassement.getCellByPosition(6, LigneClassement).Value  = NombreMatchesJouesJoueurCourant
	FeuilleClassement.getCellByPosition(7, LigneClassement).Value  = PourcentageMatchesGagnesJoueurCourant
End Sub

'Renvoit le numéro de ligne du meilleur joueur en excluant les lignes du tableau LignesExclues
'Renvoit -1 si le joueur avec le meilleur score n'a pas joué de matches
Sub PrendJoueurAvecLeMeilleurScore(NumTotalJoueurs As Integer, LignesExclues as Array) As Integer
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
	indexLigneSrcJoueurs 	= 9 'Is the Number in the line -  1 (If it's line 10, it's 10-1 = 9)
	
	JoueurCourant 						= 0
	LigneMeilleurJoueur 				= indexLigneSrcJoueurs
	NombreMatchesGagnesMeilleurJoueur 	= -1
	SetsMeilleurJoueur					= -1
	PointsMeilleurJoueur				= -1
	NombreMatchesJouesMeilleurJoueur 	= -1
	ADejaJoueUnMatchMeilleurJoueur		= False 					
		
	ColonneScore				= 5
	ColonneSets					= 6
	ColonnePoints				= 7
	ColonneNombreMatchesJoues	= 8
	'print (CStr(NombreMatchesGagnesJoueurCourant) +" "+CStr(SetsJoueurCourant) + " "+CStr(PointsJoueurCourant))
	
	Do While (JoueurCourant <> NumTotalJoueurs)
	  LigneCourante				= indexLigneSrcJoueurs + JoueurCourant
	  idx 						= ChercheValeur(LignesExclues, NumTotalJoueurs, LigneCourante)
      If idx < 0 Then
		  NombreMatchesGagnesJoueurCourant 	= oInscriptionPage.getCellByPosition(ColonneScore , LigneCourante).Value
		  SetsJoueurCourant  				= oInscriptionPage.getCellByPosition(ColonneSets  , LigneCourante).Value
		  PointsJoueurCourant  				= oInscriptionPage.getCellByPosition(ColonnePoints, LigneCourante).Value
		  NombreMatchesJouesJoueurCourant	= oInscriptionPage.getCellByPosition(ColonneNombreMatchesJoues, LigneCourante).Value
		  ADejaJoueUnMatch 					= NombreMatchesJouesJoueurCourant > 0
		  
		  If (ADejaJoueUnMatch) Then  	
			  If (NombreMatchesGagnesJoueurCourant > NombreMatchesGagnesMeilleurJoueur) Then
			  	'Score du joueur courant supérieur
			  	NombreMatchesGagnesMeilleurJoueur	= NombreMatchesGagnesJoueurCourant
				SetsMeilleurJoueur					= SetsJoueurCourant
				PointsMeilleurJoueur				= PointsJoueurCourant
				LigneMeilleurJoueur					= LigneCourante
				NombreMatchesJouesMeilleurJoueur 	= NombreMatchesJouesJoueurCourant
				ADejaJoueUnMatchMeilleurJoueur		= True
			  Else 
			  	  If (NombreMatchesGagnesJoueurCourant = NombreMatchesGagnesMeilleurJoueur) Then 'Le meilleur joueur a fait potentiellement moins de sets qu'un autre joueur
			         If (SetsJoueurCourant > SetsMeilleurJoueur) Then
			           'Scores égaux mais nombre de sets supérieur
			            NombreMatchesGagnesMeilleurJoueur 	= NombreMatchesGagnesJoueurCourant
						SetsMeilleurJoueur					= SetsJoueurCourant
						PointsMeilleurJoueur				= PointsJoueurCourant
						LigneMeilleurJoueur					= LigneCourante
						NombreMatchesJouesMeilleurJoueur 	= NombreMatchesJouesJoueurCourant
						ADejaJoueUnMatchMeilleurJoueur		= True
					 Else
					   If (SetsJoueurCourant = SetsMeilleurJoueur) Then
			           	 If (PointsJoueurCourant > PointsMeilleurJoueur) Then
			           	    'Scores et sets égaux mais nombre de points supérieur
			               	NombreMatchesGagnesMeilleurJoueur 	= NombreMatchesGagnesJoueurCourant
							SetsMeilleurJoueur					= SetsJoueurCourant
							PointsMeilleurJoueur				= PointsJoueurCourant
							LigneMeilleurJoueur					= LigneCourante
							NombreMatchesJouesMeilleurJoueur 	= NombreMatchesJouesJoueurCourant
							ADejaJoueUnMatchMeilleurJoueur		= True
			             Else
				             If (NombreMatchesJouesJoueurCourant > NombreMatchesJouesMeilleurJoueur) Then
				           	 	NombreMatchesGagnesMeilleurJoueur 	= NombreMatchesGagnesJoueurCourant
								SetsMeilleurJoueur					= SetsJoueurCourant
								PointsMeilleurJoueur				= PointsJoueurCourant
								LigneMeilleurJoueur					= LigneCourante
								NombreMatchesJouesMeilleurJoueur 	= NombreMatchesJouesJoueurCourant
								ADejaJoueUnMatchMeilleurJoueur		= True
				             End If
				         End If
			           End If
			         End If
			      End If
			  End If
	  	End If
	  End If
	  JoueurCourant = JoueurCourant + 1
	Loop
	If (ADejaJoueUnMatchMeilleurJoueur) Then
		PrendJoueurAvecLeMeilleurScore = LigneMeilleurJoueur
	Else
		PrendJoueurAvecLeMeilleurScore = -1
	End If
End Sub

'Renvoit le numéro de ligne du meilleur joueur en utilisant le pourcentage de réussite, en excluant les lignes du tableau LignesExclues
Sub PrendJoueurAvecLeMeilleurPourcentage(NumTotalJoueurs As Integer, LignesExclues as Array) As Integer
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
	indexLigneSrcJoueurs 	= 9 'Is the Number in the line -  1 (If it's line 10, it's 10-1 = 9)
	
	JoueurCourant 							= 0
	LigneMeilleurJoueur 					= indexLigneSrcJoueurs
	NombreMatchesGagnesMeilleurJoueur 		= -1
	SetsMeilleurJoueur						= -1
	PointsMeilleurJoueur					= -1
	NombreMatchesJouesMeilleurJoueur 		= -1
	PourcentageMatchesGagnesMeilleurJoueur 	= -1
	ADejaJoueUnMatchMeilleurJoueur			= False 	
	
	ColonneScore				= 5
	ColonneSets					= 6
	ColonnePoints				= 7
	ColonneNombreMatchesJoues	= 8
	ColonnePourcentage			= 9
	
	Dim ADejaJoueUnMatch as Boolean
	
	'print (CStr(NombreMatchesGagnesJoueurCourant) +" "+CStr(SetsJoueurCourant) + " "+CStr(PointsJoueurCourant))
	
	Do While (JoueurCourant <> NumTotalJoueurs)
	  LigneCourante				= indexLigneSrcJoueurs + JoueurCourant
	  idx 						= ChercheValeur(LignesExclues, NumTotalJoueurs, LigneCourante)
      If idx < 0 Then
		  NombreMatchesGagnesJoueurCourant  		= oInscriptionPage.getCellByPosition(ColonneScore , LigneCourante).Value
		  SetsJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonneSets  , LigneCourante).Value
		  PointsJoueurCourant  						= oInscriptionPage.getCellByPosition(ColonnePoints, LigneCourante).Value
		  NombreMatchesJouesJoueurCourant			= oInscriptionPage.getCellByPosition(ColonneNombreMatchesJoues, LigneCourante).Value
		  PourcentageMatchesGagnesJoueurCourant 	= oInscriptionPage.getCellByPosition(ColonnePourcentage, LigneCourante).Value
		  ADejaJoueUnMatch 							= NombreMatchesJouesJoueurCourant > 0
		  
		  If (ADejaJoueUnMatch) Then
			  If (PourcentageMatchesGagnesJoueurCourant > PourcentageMatchesGagnesMeilleurJoueur) Then
			  	'Pourcentage du joueur courant supérieur
			  	NombreMatchesGagnesMeilleurJoueur 		= NombreMatchesGagnesJoueurCourant
				SetsMeilleurJoueur						= SetsJoueurCourant
				PointsMeilleurJoueur					= PointsJoueurCourant
				LigneMeilleurJoueur						= LigneCourante
				NombreMatchesJouesMeilleurJoueur 		= NombreMatchesJouesJoueurCourant
				PourcentageMatchesGagnesMeilleurJoueur 	= PourcentageMatchesGagnesJoueurCourant
				ADejaJoueUnMatchMeilleurJoueur			= True
			  Else 
			  	  If (PourcentageMatchesGagnesJoueurCourant = PourcentageMatchesGagnesMeilleurJoueur) Then
			         If (NombreMatchesJouesJoueurCourant > NombreMatchesJouesMeilleurJoueur) Then
			           	'Pourcentages égaux mais nombre de matches gagnés supérieur
			           	NombreMatchesGagnesMeilleurJoueur 		= NombreMatchesGagnesJoueurCourant
						SetsMeilleurJoueur						= SetsJoueurCourant
						PointsMeilleurJoueur					= PointsJoueurCourant
						LigneMeilleurJoueur						= LigneCourante
						NombreMatchesJouesMeilleurJoueur 		= NombreMatchesJouesJoueurCourant
						PourcentageMatchesGagnesMeilleurJoueur 	= PourcentageMatchesGagnesJoueurCourant
						ADejaJoueUnMatchMeilleurJoueur			= True
					 Else
					   If (NombreMatchesJouesJoueurCourant = NombreMatchesJouesMeilleurJoueur) Then
			           	  If (NombreMatchesGagnesJoueurCourant > NombreMatchesGagnesMeilleurJoueur) Then
						  	'Score du joueur courant supérieur
							NombreMatchesGagnesMeilleurJoueur 		= NombreMatchesGagnesJoueurCourant
							SetsMeilleurJoueur						= SetsJoueurCourant
							PointsMeilleurJoueur					= PointsJoueurCourant
							LigneMeilleurJoueur						= LigneCourante
							NombreMatchesJouesMeilleurJoueur 		= NombreMatchesJouesJoueurCourant
							PourcentageMatchesGagnesMeilleurJoueur 	= PourcentageMatchesGagnesJoueurCourant
							ADejaJoueUnMatchMeilleurJoueur			= True
						  Else 
						  	  If (NombreMatchesGagnesJoueurCourant = NombreMatchesGagnesMeilleurJoueur) Then
						         If (SetsJoueurCourant > SetsMeilleurJoueur) Then
						           'Scores égaux mais nombre de sets supérieur
									NombreMatchesGagnesMeilleurJoueur 		= NombreMatchesGagnesJoueurCourant
									SetsMeilleurJoueur						= SetsJoueurCourant
									PointsMeilleurJoueur					= PointsJoueurCourant
									LigneMeilleurJoueur						= LigneCourante
									NombreMatchesJouesMeilleurJoueur 		= NombreMatchesJouesJoueurCourant
									PourcentageMatchesGagnesMeilleurJoueur 	= PourcentageMatchesGagnesJoueurCourant
									ADejaJoueUnMatchMeilleurJoueur			= True
								 Else
								   If (SetsJoueurCourant = SetsMeilleurJoueur) Then
						           	 If (PointsJoueurCourant > PointsMeilleurJoueur) Then
						           	    'Scores et sets égaux mais nombre de points supérieur
						               	NombreMatchesGagnesMeilleurJoueur 		= NombreMatchesGagnesJoueurCourant
										SetsMeilleurJoueur						= SetsJoueurCourant
										PointsMeilleurJoueur					= PointsJoueurCourant
										LigneMeilleurJoueur						= LigneCourante
										NombreMatchesJouesMeilleurJoueur 		= NombreMatchesJouesJoueurCourant
										PourcentageMatchesGagnesMeilleurJoueur 	= PourcentageMatchesGagnesJoueurCourant
										ADejaJoueUnMatchMeilleurJoueur			= True
						             End If ' If (PointsJoueurCourant > PointsMeilleurJoueur)
						           End If 'If (SetsJoueurCourant = SetsMeilleurJoueur
						         End If ' If (SetsJoueurCourant > SetsMeilleurJoueur)
						      End If ' If (NombreMatchesGagnesJoueurCourant = NombreMatchesGagnesMeilleurJoueur)
						  End If 'If (NombreMatchesGagnesJoueurCourant > NombreMatchesGagnesMeilleurJoueur)
			         	End If ' If (NombreMatchesJouesJoueurCourant = NombreMatchesJouesMeilleurJoueur)
			      End If '  If (NombreMatchesJouesJoueurCourant > NombreMatchesJouesMeilleurJoueur)
			  End If 'If (PourcentageMatchesGagnesJoueurCourant = PourcentageMatchesGagnesMeilleurJoueur)
		  End If 'If (PourcentageMatchesGagnesJoueurCourant > PourcentageMatchesGagnesMeilleurJoueur)
	  	End If 'If (ADejaJoueUnMatch)
	  End If 'If idx < 0
	  JoueurCourant = JoueurCourant + 1
	Loop 'Do While
	
	If (ADejaJoueUnMatchMeilleurJoueur) Then
		PrendJoueurAvecLeMeilleurPourcentage = LigneMeilleurJoueur
	Else
		PrendJoueurAvecLeMeilleurPourcentage = -1
	End If
End Sub

'Renvoit l'index de la valeur si elle est trouvée dans le tableau
Sub ChercheValeur(Tableau as Array, Total as Integer, Val as Integer) as Integer
	dim i As Integer
	Result = -1
	For i = 0 To (Total-1)
    	if (Tableau(i) = Val) Then
    		Result = i
    		Exit For
    	End If
	Next i
	
	ChercheValeur = Result
End Sub

Sub CalculeNombreTotalDeJoueurs() as Integer
	oInscriptionPage 		= ThisComponent.Sheets.GetByName("Inscriptions")
	indexLigneSrcJoueurs 	= 9 'Is the Number in the line -  1 (If it's line 10, it's 10-1 = 9)
	
	NumJoueurs 		= 0
	NomJoueur  		= oInscriptionPage.getCellByPosition(1, indexLigneSrcJoueurs + NumJoueurs).String
	PrenomJoueur  	= oInscriptionPage.getCellByPosition(2, indexLigneSrcJoueurs + NumJoueurs).String
	Do While (NomJoueur <> "" or PrenomJoueur <> "")
	  NumJoueurs 	= NumJoueurs + 1
	  NomJoueur 	= oInscriptionPage.getCellByPosition(1, indexLigneSrcJoueurs + NumJoueurs).String
	  PrenomJoueur 	= oInscriptionPage.getCellByPosition(2, indexLigneSrcJoueurs + NumJoueurs).String	
	Loop
	CalculeNombreTotalDeJoueurs = NumJoueurs
End Sub

Sub CalculeScoreDuMatch(oTourPage as Object, currentMatchIndex as Integer, ColonneDuTourDansPageInscriptions as Integer)
  oInscriptionPage 			= ThisComponent.Sheets.GetByName("Inscriptions")
  
  SetsEquipeGauche			= 0
  SetsEquipeDroite 			= 0
  DeltaPointsEquipeGauche	= 0
  DeltaPointsEquipeDroite	= 0
  indexLigneDstJoueurs 		= 4
  ColumnIndexOfSets 		= 6
  ColumnIndexOfPoints 		= 7
  ColumnIndexOfMatchesJoues = 8  
  ColumnIndexOfPourcentageMatchesGagnes = 9
   
  'Set 1
  PointsEquipeGaucheSet1 		= oTourPage.getCellByPosition(4, indexLigneDstJoueurs + (4*currentMatchIndex) + 0).Value '4 est l'index de la colonne du score equipe gauche
  PointsEquipeDroiteSet1 		= oTourPage.getCellByPosition(5, indexLigneDstJoueurs + (4*currentMatchIndex) + 0).Value '5 est l'index de la colonne du score equipe gauche
  DeltaPointsEquipeGaucheSet1 	= PointsEquipeGaucheSet1 - PointsEquipeDroiteSet1
  DeltaPointsEquipeDroiteSet1 	= PointsEquipeDroiteSet1 - PointsEquipeGaucheSet1 
  If (PointsEquipeGaucheSet1 <> 0 Or PointsEquipeDroiteSet1 <> 0) Then
	  If (PointsEquipeGaucheSet1 > PointsEquipeDroiteSet1) Then
	  	SetsEquipeGauche = SetsEquipeGauche + 1
	  Else
	  	SetsEquipeDroite = SetsEquipeDroite + 1
	  End If
  End If
  
  'Set 2
  PointsEquipeGaucheSet2 		= oTourPage.getCellByPosition(4, indexLigneDstJoueurs + (4*currentMatchIndex) + 1).Value
  PointsEquipeDroiteSet2 		= oTourPage.getCellByPosition(5, indexLigneDstJoueurs + (4*currentMatchIndex) + 1).Value
  DeltaPointsEquipeGaucheSet2 	= PointsEquipeGaucheSet2 - PointsEquipeDroiteSet2
  DeltaPointsEquipeDroiteSet2 	= PointsEquipeDroiteSet2 - PointsEquipeGaucheSet2 
  
  If (PointsEquipeGaucheSet2 <> 0 Or PointsEquipeDroiteSet2 <> 0) Then
	  If (PointsEquipeGaucheSet2 > PointsEquipeDroiteSet2) Then
	  	SetsEquipeGauche = SetsEquipeGauche + 1
	  Else
	  	SetsEquipeDroite = SetsEquipeDroite + 1
	  End If
  End If
  
  'Set 3
  PointsEquipeGaucheSet3 = oTourPage.getCellByPosition(4, indexLigneDstJoueurs + (4*currentMatchIndex) + 2).Value
  PointsEquipeDroiteSet3 = oTourPage.getCellByPosition(5, indexLigneDstJoueurs + (4*currentMatchIndex) + 2).Value
  DeltaPointsEquipeGaucheSet3 	= PointsEquipeGaucheSet3 - PointsEquipeDroiteSet3
  DeltaPointsEquipeDroiteSet3 	= PointsEquipeDroiteSet3 - PointsEquipeGaucheSet3 
  If (PointsEquipeGaucheSet3 <> 0 Or PointsEquipeDroiteSet3 <> 0) Then
	  If (PointsEquipeGaucheSet3 > PointsEquipeDroiteSet3) Then
	  	SetsEquipeGauche = SetsEquipeGauche + 1
	  Else
	  	SetsEquipeDroite = SetsEquipeDroite + 1
	  End If
  End If
  
  DeltaPointsEquipeGauche	= DeltaPointsEquipeGaucheSet1 + DeltaPointsEquipeGaucheSet2 + DeltaPointsEquipeGaucheSet3
  DeltaPointsEquipeDroite	= DeltaPointsEquipeDroiteSet1 + DeltaPointsEquipeDroiteSet2 + DeltaPointsEquipeDroiteSet3
  
  'print ("Match "+CStr(currentMatchIndex)+": Sets Equipe G : "+CStr(SetsEquipeGauche)+" Sets Equipe D : "+CStr(SetsEquipeDroite))
  
  'Points ajoutés pour gagnants et perdants
  PointsPourGagnants = 1
  PointsPourPerdants = 0
  
  'On recupere les index des lignes des joueurs pour leur mettre les points dans la colonne du joueur
  IndexJoueur1EquipeGauche = oTourPage.getCellByPosition(3, indexLigneDstJoueurs + (4*currentMatchIndex) + 0).Value
  IndexJoueur2EquipeGauche = oTourPage.getCellByPosition(3, indexLigneDstJoueurs + (4*currentMatchIndex) + 1).Value
 
  IndexJoueur1EquipeDroite = oTourPage.getCellByPosition(8, indexLigneDstJoueurs + (4*currentMatchIndex) + 0).Value
  IndexJoueur2EquipeDroite = oTourPage.getCellByPosition(8, indexLigneDstJoueurs + (4*currentMatchIndex) + 1).Value
 
  oJoueur1EquipeGaucheNom 		= oTourPage.getCellByPosition(1, indexLigneDstJoueurs + (4*currentMatchIndex) + 0) '1 est l'index de colonne du nom du joueur 1 equipe gauche
  oJoueur1EquipeGauchePrenom 	= oTourPage.getCellByPosition(2, indexLigneDstJoueurs + (4*currentMatchIndex) + 0) '2 est l'index de colonne du prenom du joueur 1 equipe gauche
  oJoueur2EquipeGaucheNom 		= oTourPage.getCellByPosition(1, indexLigneDstJoueurs + (4*currentMatchIndex) + 1) 'nom joueur 2 equipe gauche
  oJoueur2EquipeGauchePrenom 	= oTourPage.getCellByPosition(2, indexLigneDstJoueurs + (4*currentMatchIndex) + 1) 'prenom du joueur 2 equipe gauche

  oJoueur1EquipeDroiteNom 		= oTourPage.getCellByPosition(6, indexLigneDstJoueurs + (4*currentMatchIndex) + 0) '6 est l'index de colonne du nom du joueur 1 equipe droite
  oJoueur1EquipeDroitePrenom 	= oTourPage.getCellByPosition(7, indexLigneDstJoueurs + (4*currentMatchIndex) + 0) '7 est l'index de colonne du prenom du joueur 1 equipe droite
  oJoueur2EquipeDroiteNom 		= oTourPage.getCellByPosition(6, indexLigneDstJoueurs + (4*currentMatchIndex) + 1) 'joueur 2 equipe droite
  oJoueur2EquipeDroitePrenom 	= oTourPage.getCellByPosition(7, indexLigneDstJoueurs + (4*currentMatchIndex) + 1) 'joueur 2 equipe droite
  
  CouleurGagnant				= RGB(0, 200, 0)
  CouleurPerdant				= RGB(200, 50, 0)
  
  'Nombre de sets
  RemplitScore(ColumnIndexOfSets, IndexJoueur1EquipeGauche, SetsEquipeGauche, True)
  RemplitScore(ColumnIndexOfSets, IndexJoueur2EquipeGauche, SetsEquipeGauche, True)
  RemplitScore(ColumnIndexOfSets, IndexJoueur1EquipeDroite, SetsEquipeDroite, True)
  RemplitScore(ColumnIndexOfSets, IndexJoueur2EquipeDroite, SetsEquipeDroite, True)
  
  'Nombre des points
  RemplitScore(ColumnIndexOfPoints, IndexJoueur1EquipeGauche, DeltaPointsEquipeGauche, True)
  RemplitScore(ColumnIndexOfPoints, IndexJoueur2EquipeGauche, DeltaPointsEquipeGauche, True)
  RemplitScore(ColumnIndexOfPoints, IndexJoueur1EquipeDroite, DeltaPointsEquipeDroite, True)
  RemplitScore(ColumnIndexOfPoints, IndexJoueur2EquipeDroite, DeltaPointsEquipeDroite, True)
  
  'Nombre de matches joués
  RemplitScore(ColumnIndexOfMatchesJoues, IndexJoueur1EquipeGauche, 1, True)
  RemplitScore(ColumnIndexOfMatchesJoues, IndexJoueur2EquipeGauche, 1, True)
  RemplitScore(ColumnIndexOfMatchesJoues, IndexJoueur1EquipeDroite, 1, True)
  RemplitScore(ColumnIndexOfMatchesJoues, IndexJoueur2EquipeDroite, 1, True)
   	   
  If (SetsEquipeGauche <> 0 or SetsEquipeDroite <> 0) Then
	  If (SetsEquipeGauche > SetsEquipeDroite) Then
	  	
	  	'Equipe gauche gagnante
	  	'On ajoute les points sur le tour dans la page inscriptions
	  	
	  	'Nombre de points pour match	
	  	RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur1EquipeGauche, PointsPourGagnants, False)
		RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur2EquipeGauche, PointsPourGagnants, False)
		oJoueur1EquipeGaucheNom.CellBackColor 		= CouleurGagnant
		oJoueur1EquipeGauchePrenom.CellBackColor 	= CouleurGagnant
		oJoueur2EquipeGaucheNom.CellBackColor 		= CouleurGagnant
		oJoueur2EquipeGauchePrenom.CellBackColor 	= CouleurGagnant

		'Nombre de points pour match	
	    RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur1EquipeDroite, PointsPourPerdants, False)
	    RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur2EquipeDroite, PointsPourPerdants, False) 
	    oJoueur1EquipeDroiteNom.CellBackColor 		= CouleurPerdant
	    oJoueur1EquipeDroitePrenom.CellBackColor 	= CouleurPerdant
	    oJoueur2EquipeDroiteNom.CellBackColor 		= CouleurPerdant
	    oJoueur2EquipeDroitePrenom.CellBackColor 	= CouleurPerdant  
	  Else
	  	'Equipe droite gagnante
	    'On ajoute les points sur le tour dans la page inscriptions
	  	'Nombre de points pour match	
	  	RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur1EquipeGauche, PointsPourPerdants, False)
		RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur2EquipeGauche, PointsPourPerdants, False)	
		oJoueur1EquipeGaucheNom.CellBackColor 		= CouleurPerdant
		oJoueur1EquipeGauchePrenom.CellBackColor 	= CouleurPerdant
		oJoueur2EquipeGaucheNom.CellBackColor 		= CouleurPerdant
		oJoueur2EquipeGauchePrenom.CellBackColor 	= CouleurPerdant
		
	    RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur1EquipeDroite, PointsPourGagnants, False)
	    RemplitScore(ColonneDuTourDansPageInscriptions, IndexJoueur2EquipeDroite, PointsPourGagnants, False)
	    oJoueur1EquipeDroiteNom.CellBackColor 		= CouleurGagnant
	    oJoueur1EquipeDroitePrenom.CellBackColor 	= CouleurGagnant
	    oJoueur2EquipeDroiteNom.CellBackColor 		= CouleurGagnant
	    oJoueur2EquipeDroitePrenom.CellBackColor 	= CouleurGagnant	  
	  End If
  End If
  
  'Pourcentage de matches gagnés
  RemplitPourcentageMatchesGagnes(ColumnIndexOfPourcentageMatchesGagnes, IndexJoueur1EquipeGauche, PointsPourGagnants)
  RemplitPourcentageMatchesGagnes(ColumnIndexOfPourcentageMatchesGagnes, IndexJoueur2EquipeGauche, PointsPourGagnants)
  RemplitPourcentageMatchesGagnes(ColumnIndexOfPourcentageMatchesGagnes, IndexJoueur1EquipeDroite, PointsPourGagnants)
  RemplitPourcentageMatchesGagnes(ColumnIndexOfPourcentageMatchesGagnes, IndexJoueur2EquipeDroite, PointsPourGagnants)
End Sub

Sub RemplitPourcentageMatchesGagnes(ColumnIndexOfPourcentageMatchesGagnes as Integer, IndexLigneJoueurDansPageInscriptions as Integer, PointsPourMatchGagne as Integer)
	Dim Pourcentage as Double
	Dim NbreMatchesJoues as Integer
	
	ColumnIndexOfScore			= 5
	ColumnIndexOfMatchesJoues 	= 8 
	oInscriptionPage 			= ThisComponent.Sheets.GetByName("Inscriptions")

	If (IndexLigneJoueurDansPageInscriptions >= 0) Then
  	'IndexLigneJoueurDansPageInscriptions est negatif si le joueur est remplacant/completant donc on ignore ses points car il a déjà joué ce tour
		Score 				= oInscriptionPage.getCellByPosition(ColumnIndexOfScore, IndexLigneJoueurDansPageInscriptions).Value
		'print score
		NbreMatchesJoues 	= oInscriptionPage.getCellByPosition(ColumnIndexOfMatchesJoues, IndexLigneJoueurDansPageInscriptions).Value 
		'print NbreMatchesJoues
	  	Pourcentage 		= 100.0 * CDbl(Score) / CDbl(NbreMatchesJoues * PointsPourMatchGagne)
		'print Left(CStr(Pourcentage), 4)
		oInscriptionPage.getCellByPosition(ColumnIndexOfPourcentageMatchesGagnes, IndexLigneJoueurDansPageInscriptions).Value = CInt(Pourcentage) 'Convertit en entier
	End If
End Sub

Sub RemplitScore(ColonneDuTourDansPageInscriptions as Integer, IndexLigneJoueurDansPageInscriptions as Integer, Valeur as Integer, bAjoute as Boolean)
  oInscriptionPage = ThisComponent.Sheets.GetByName("Inscriptions")
  If (IndexLigneJoueurDansPageInscriptions >= 0) Then
  	'IndexLigneJoueurDansPageInscriptions est negatif si le joueur est remplacant/completant donc on ignore ses points car il a déjà joué ce tour
  	if (bAjoute) Then
  		if oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, IndexLigneJoueurDansPageInscriptions).String = "" Then
  		  oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, IndexLigneJoueurDansPageInscriptions).setFormula("="+CStr(Valeur))
  		Else
  		  oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, IndexLigneJoueurDansPageInscriptions).setFormula(oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, IndexLigneJoueurDansPageInscriptions).getFormula()+ "+" + CStr(Valeur))
  		End If
  	Else
  		oInscriptionPage.getCellByPosition(ColonneDuTourDansPageInscriptions, IndexLigneJoueurDansPageInscriptions).Value = Valeur
    End If
  End If
End Sub



