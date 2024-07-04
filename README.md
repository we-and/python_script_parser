<img src="https://github.com/we-and/python_script_parser/blob/main/screenshot.png?raw=true" width="500">

## Nouveautés de cette version
 - choix PDF multi colonnes
 - choix XLSX multi colonnes
 - Nouveaux formats
 - Nouveaux filtres
 
## Télécharger
Choisir en bas de cette page un .dmg en fonction de votre machine Mac

| Modèle de Mac | DMG |
| ------------- | ------------- |
| Apple Silicon (processeur M1, M2, M3)  | scripti_applesilicon_<version>.dmg |
| Intel (processeur Intel)  | scripti_intel_<version>.dmg |

## Démarrage
- Ouvrir le .dmg
- Double clic   Scripti
- Au premier lancement, attendre 30 secondes: Une console s'ouvre pour l'installation et la fenêtre s'affiche ensuite.
Version vidéo: https://youtu.be/51F_gElQ1aI


### Documentation
| Sujet | Vidéo |
| ------------- | ------------- |
| Installation | https://youtu.be/51F_gElQ1aI |
| Fusionner des personnages | https://youtu.be/9MqWrRHDRPk |
| Extraction PDF dialogue centré | https://youtu.be/1F8yu14x6u8 |
| Tableaux Word | https://youtu.be/XxZBVQKvihI |
  

### Formats de fichier 
| Fichier  | Format | Exemples | Extrait |
| ------------- | ------------- |------------- |------------- |
| TXT  | PERSONNAGE Tab Dialogue |  3 Secondes, La Esclava blanca | Isabel 2	Non, je viens avec vous. | 
| TXT | PERSONNAGE Espaces Dialogue |  You can't run forever |BURLY MAN        Shut up. | 
| TXT  | PERSONAGE: Dialogue | What you wish for | IMOGENE:<br>Of course. | 
| TXT  | PERSONAGE<br>Dialogue | The Clean Up Crew | GABRIEL<br>I know this one. | 
| TXT  | PERSONAGE:<br>Dialogue | Country Crush | CODY TO KATHERINE: <br>Baby, hey, come here. |
| TXT  | [Timecode] <br> [PERSONNAGE] Dialogue [Timecode] | M3 |[00:00:44.02]<br>[Courtney] Is he gone?<br>[00:00:46.16]| 
| TXT  | Timecode --> Timecode <br> [Personnage] Dialogue <br> Dialogue suite | DESP_LIST |00:01:06,691 --> 00:01:08,943<br>[Ana] A place where dreams <br>invade daytime| 
| TXT  | Timecode --> Timecode <br> PERSONNAGE: Dialogue <br> Dialogue suite | Young Hearts | 6<br>00:01:43,583 --> 00:01:48,083<br>LUK: <i>je vindt de vrijheid bij elkaar</i>| 
| TXT  | Timecode - Timecode <br> [Personnage]: <br> Dialogue | Badland Doves |00:01:17:20 - 00:01:19:18<br>Regina:<br>That was my idea.| 
| TXT  | PERSONNAGE <br> Dialogue | Ventino |MANOLO<br>We have to go, let's go. | 
| TXT  | MIN'SEC-PERSONNAGE : <br> Dialogue | DURAG DEF |00’24-CAUE :<br>La couleur préférée des frères, c’est le rose bébé.| 
| TXT  | Timecode - Timecode<br>Personnage <br> Dialogue | Transcript | 00:01:30:07 - 00:01:37:08<br>Speaker 2<br>Move in here because I'm retired.
| TXT  | IDX TIMECODE --> TIMECODE<br>Dialogue<br>Dialogue | Mascarpone 24fps |10 00:04:19,875 --> 00:04:21,608<br>Qu’est-ce qui s’est passé ?| 
| TXT  | PERSONNAGE Dialogue | GodLess |RON Things were subtle at the start.| 
| TXT  | TIMECODE TIMECODE_OR_--	Dialog | ENG Big Kill  | 01:02:52:15	01:02:53:15	Don't shoot me.| 
| TXT  | NUM:       TIMECODE TIMECODE TIME <br> Dialogue | Kokon 24fps Final | 3:       10:00:40:22 10:00:42:19 01:21<br>Trop drôle avec les poils !| 
| DOCX | Colonne Character / Dialog| Blackwater Lane, Gods of the Deep, Take Back | | 
| DOCX | Colonne Combined Continuity | Latency  | | 
| DOCX | Colonne Scene Description | Catch the bullet || 
| DOCX | Colonne Dialogue with Speaker Id | Kenya || 
| DOCX  | [Timecode] --><br>[Timecode] PERSONNAGE  Dialogue | Kokon || 
| DOCX | Format script de film avec styles | John C Walsh || 
| XLSX | Colonne Character / English | Day Zero |  | 
| PDF  | Dialogue en bloc centré | Radical | | 
| PDF  | Tables, découper la colonne | Unplugging || 

### Fonctions
Ouvrir un dossier de travail
Ouvrir un fichier individuel 
Changer la taille des répliques
Export dialogues
Export comptage
Export dialogue par personnage
Ouvrir le dossier de résultats 
Ouvrir un groupe de scripts

### Règles de filtrage
```
(O.S.) (OS) (O.S) (OS/ON)
(V.O) (VO) 
(CONT’D) (CONT’D) (CONT.) (CON'T.)  (CONT'D) (CONT’D)
SONG
NOTE D'AUTEUR
END CREDITS 
NARRATIVE TITLE
ON-SCREEN TEXT
MAIN TITLE 
OPENING CREDITS
(didascalies)
(Ambiance)
[notes]
“-“
“- “
Dialogue monoligne
Dialogue multiligne
Dialogue alterné
Détection encodage du fichier
Séparateur de scène
Personnage qui interrompt une chanson
Chanson chantée par un personnage/ chanson de fond
Personnage A AND B compté pour A et B. 
```
### A faire
- FILTRAGE: Dialogue alterné sans mention des personnages
- Support .doc et .rtf sous macOS 
- Étudier pour macOS 10.13

### Requis
MacOS 11.1 ou supérieur 


### Testés
| Modèle  | OS | 
| ------------- | ------------- |
| MacBook Pro M1  | MacOS 14 |
| MacBook Air M1  | MacOS 13 |
| MacBook Air Intel  | MacOS 11.7 |
