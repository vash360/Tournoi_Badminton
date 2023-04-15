# Tournoi de Badminton

Gestion de tournoi amical de badminton en double mixte
Créé par David Lanier
(c)Copyright 2021-2023, All rights reserved

Le logiciel gère le choix du partenaire et des adversaires en créant (le plus possible) 
suivant le nombre de joueurs des paires en mixte.
Le tournoi se déroule suivant des tours (pas de poules ou de phases finales). 
Un classement en fonction des matches est établi à la fin de chaque tour. 
Et aussi un classement par pourcentage de matches gagnés (qui peut être différent du classement global,
si un joueur n'a joué que 2 matches et un autre 9 matches.
Chacun est libre d'utiliser l'un ou l'autre ou définir le classement.

Il y a 2 modes :
1) le choix des paires qui s'affrontent dans un tour est fait au hasard, tout en essayant de faire 
des paires mixtes le plus possible.
2) le choix des paires qui s'affrontent dans un tour est fait en fonction du classement par pourcentage
de matches gagnés pour définir si un joueur/euse se trouve dans le haut ou le bas du tableau et le logiciel
essaie de choisir un partenaire et des adversaires dans le même endroit du tableau (haut ou bas).

Il est possible d'enlever des joueurs/euses à chaque tour ou d'en ajouter de nouveaux. 
Les joueurs/euses peuvent aussi faire une pause pendant un tour ou plus.

Pour les matches, le logiciel gère jusqu'à 3 sets par matches.
Le nombre de points est libre, il peut être différent de 21. 
Il est aussi possible de ne faire qu'un seul set.
Le logiciel regarde le premier set, il considère que l'équipe qui a le plus grande nombre de points 
a gagné ce set.
Puis il regarde le 2e set s'il y en a un et le 3e. 
L'équipe gagnante du match est l'équipe qui a gagné le plus de sets.

Un match gagné ramène 1 point, un match perdu 0 points. 
Le logiciel ne gère pas les égalités ou matches nuls, il faut obligatoirement une équipe gagnante.
