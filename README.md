# SportRadarTest
**Andrea De Filippo - SportRadar coding test**

### INTRODUCTION

Sample VB6 program implementing a football World Cup score board simple library, as per specification document "Coding Exercise v1.6.pdf".

The program is comprised of two executables:

"SportRadarScoreBoardLib.dll" which is the library containing the requested functionalities.

"prjScoreBoardTest.exe" a GUI program demonstrating such functionalities.

The GUI program requires the VB6 SP6 runtimes as it makes use the standard VB6 controls and a ListView from the Microsoft Windows Common Controls 6.0 (SP6) ActiveX.

### DESIGN

The overall design was built over the assumption that in a real-world scenario, during a World Cup tournament, multiple games can be played at once. I added error checking to ensure that a team in an active game cannot be entered in another new game.

The **clsGame** class is responsible for holding information about a game, but it's not directly instantiable, it can only be created through the StartNewGame method of the collection class **clsGames**. This partially makes up for the lack of constructors in VB6.

Each clsGame instance auto-generates a GUID on creation which also serves as Key in the collection.

Every time the collection is modified, because of a new game or a score change, it is re-sorted according to specifications, therefore all FOR...EACH cycles on the collection class instance always return a properly sorted list of games. For sorting, a BubbleSort algorithm was chosen, which is not the fastest approach but considering the limited number of teams in a World Cup tournament, it serves the purpose well.

### CODE

All code written by me, Andrea De Filippo, except where specified (Sorting and GUID generation). I followed what's considered to be the standard for VB6 code: generally CamelCase for variable names, prefixed with a letter indicating the type. All methods and modules have some kind of header showing information.

### RUN
Run the executable file for a quick try. It has no digital signature, so it might cause the Windows SmartScreen to pop up. In case you decide to use the Exe file, you will need to register the ActiveX dll file with RegSvr32. Alternatively, open the project group file **FootballScoreBoardTest.vbg** in the VB6 IDE.

### GUI

Everything in the GUI should be self-explanatory. As suggested, I did not spend much time on it, just a basic demo. A quick note on end game: it was requested that the score board be cleared upon ending a game, so when a closed game is selected in the listview, no information is shown. Alternatively, one could decide to still show the game as ended with start and end times.


Thank you for considering me for the position.
Cheers.
