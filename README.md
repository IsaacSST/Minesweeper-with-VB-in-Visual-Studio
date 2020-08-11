# Minesweeper-with-VB-in-Visual-Studio
Minesweeper reverse engineered from scratch with login and highscore features, built in Visual Studio with windows forms. Made as part of a project for A level Computer Science.
## Downloading and launching the program:
To download the game just download the MinesweeperProject.exe file in the base directory. This file appears in the /MinesweeperProject/MinesweeperProject/bin/Debug/ directory upon compiling, but I copied it to the base directory for ease of access.
Most likely the browser will attempt to block the download and operating system will warn you against opening as it is an untrusted exe file. If you really want to play you will have to go into advanced options and bypass these warnings.
The entire game is encompassed in the .exe file; no installation is necessary. When run the .exe will create Help.txt, Highscores.txt, and UserLoginRecord.txt files in whichever directory it is located.
## Playing the game:
### Navigation:

To play the game, you must first log in. You can do this by entering your username and password in the main menu and pressing the Log in button
if the account already exists, or the create account button if the account does not yet exist (Note account passwords are stored in plaintext in the UserLoginRecord.txt file, so it's advisable to use a throwaway password if you want to be completely safe; the account feature is really just there to tick a box on the A level mark scheme).
This will bring you to a Landing page where you should press the new game button.
Then you will be asked to state the dimensions of the game board. As the board is always a square, only one number is needed.
This should be a whole number between 10 and 40 but if you enter a decimal, the program will round it down.
After this the board will load and you can start playing.
The menu will be on the right of the screen and can be still be used when the game is running

### How to play:

The game starts with a simple board of squares, each containing either a mine,
 a number indicating number of surrounding mines, or nothing at all.
If a player clicks a square with a mine under, the game will end; if player clicks a square with a number on, that square will be revealed;
 if a player clicks a blank square, all connected blank and numbered squares will be revealed.
Also, the player may right click on a square to place a flag on it if they think a mine is there.
If a revealed numbered square has its number satisfied by adjacent flagged squares, 
 then the player may double click, or middle click on this number to clear all adjacent squares other than the flagged squares.
If there are too many or too few flags adjacent, nothing will happen, however if there are the right amount but they are incorrect,
 the actual mines will be revealed and the player will lose the game.

### Interface: 

When the game is being played, the number of mines left, a restart button and a game timer will be just above the board.
The restart button contains a yellow face, and when clicked will reset the game.

### Scoring:

The scoring of the game depends on how long it takes for the user to beat the game, therefore the lower the score, the better.
If they lose by triggering a mine, they have no score.
