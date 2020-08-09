Public Class Master
    'Defining class variables and instantiating important objects such as the Game board and the navigation menu
    Public n As Integer
    Public Mines As Integer
    Public Board As New Game
    Public User As String
    Public DynamicMenu As New MainControls
    Public BoardExists As Boolean



    Private Sub Master_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Text in the top bar of the form window set to "Minesweeper"
        Me.Text = "Minesweeper"
        'Border style of form set
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        'UserLoginRecord.txt file created if it does not already exist in the same directory as program executable
        If Not System.IO.File.Exists("UserLoginRecord.txt") Then
            System.IO.File.Create("UserLoginRecord.txt").Dispose()
        End If
        'Highscores.txt file created if it does not already exist in the same directory as program executable
        If Not System.IO.File.Exists("Highscores.txt") Then
            System.IO.File.Create("Highscores.txt").Dispose()
        End If
        'Help.txt file created if it does not already exist in the same directory as program executable
        If Not System.IO.File.Exists("Help.txt") Then
            System.IO.File.Create("Help.txt").Dispose()
        End If
        'Help text written to Help.txt file. Necessary with every launch of program incase the file did not exist before or had been edited
        Using Writer As New System.IO.StreamWriter("Help.txt")
            Writer.WriteLine("Navigation--------------------------------")
            Writer.WriteLine("")
            Writer.WriteLine("To play the game, you must first log in. You can do this by entering your username and password in the main menu and pressing the Log in button")
            Writer.WriteLine("if the account already exists, or the create account button if the account does not yet exist.")
            Writer.WriteLine("This will bring you to a Landing page where you should press the new game button.")
            Writer.WriteLine("Then you will be asked to state the dimensions of the game board. As the board is always a square, only one number is needed.")
            Writer.WriteLine("This should be a whole number between 10 and 40 but if you enter a decimal, the program will round it down.")
            Writer.WriteLine("After this the board will load and you can start playing.")
            Writer.WriteLine("The menu will be on the right of the screen and can be still be used when the game is running")
            Writer.WriteLine("")
            Writer.WriteLine("How to play-------------------------------")
            Writer.WriteLine("")
            Writer.WriteLine("The game starts with a simple board of squares, each containing either a mine,")
            Writer.WriteLine(" a number indicating number of surrounding mines, or nothing at all.")
            Writer.WriteLine("If a player clicks a square with a mine under, the game will end; if player clicks a square with a number on, that square will be revealed;")
            Writer.WriteLine(" if a player clicks a blank square, all connected blank and numbered squares will be revealed.")
            Writer.WriteLine("Also, the player may right click on a square to place a flag on it if they think a mine is there.")
            Writer.WriteLine("If a revealed numbered square has its number satisfied by adjacent flagged squares, ")
            Writer.WriteLine(" then the player may double click, or middle click on this number to clear all adjacent squares other than the flagged squares.")
            Writer.WriteLine("If there are too many or too few flags adjacent, nothing will happen, however if there are the right amount but they are incorrect,")
            Writer.WriteLine(" the actual mines will be revealed and the player will lose the game.")
            Writer.WriteLine("")
            Writer.WriteLine("Interface---------------------------------")
            Writer.WriteLine("")
            Writer.WriteLine("When the game is being played, the number of mines left, a restart button and a game timer will be just above the board.")
            Writer.WriteLine("The restart button contains a yellow face, and when clicked will reset the game.")
            Writer.WriteLine("")
            Writer.WriteLine("Scoring-----------------------------------")
            Writer.WriteLine("")
            Writer.WriteLine("The scoring of the game depends on how long it takes for the user to beat the game, therefore the lower the score, the better.")
            Writer.WriteLine("If they lose by triggering a mine, they have no score.")
        End Using
        'Navigation menu added to form that contains all navigational buttons
        Me.Controls.Add(DynamicMenu)



    End Sub
    Public Sub AddBoard()
        'Board removed and re-instantiated as a Game object if it already exists
        If BoardExists Then
            Me.Controls.Remove(Board)

            Board = New Game
        End If
        'Key properties of board are assigned: n is the side length in tiles of the board, Mines are the number of mines that should be present in the board
        Board.n = Me.n
        Board.Mines = Me.Mines
        'Menu object shifted to the right to make way for Game object. New position is based on the side length of the Board
        DynamicMenu.Left = 16 * n + 48
        DynamicMenu.Top = 31
        'Dimensions of form window changed based on side length of Board and the width of the Menu object
        Me.Width = 16 * n + DynamicMenu.Width + 48
        Me.Height = 16 * n + 95
        'Height of form window set to 390 if the side length of the Board was small enough that it was set to be less than 390 in the previous calculation. This is done because the Highscores page in the Menu object requires the Form height to be 390
        If Me.Height < 390 Then
            Me.Height = 390
        End If
        'Board added to Form
        Me.Controls.Add(Board)
        'BoardExists class variable set to true to show the board exists
        BoardExists = True

    End Sub
    Sub RecordTime()
        'ExistingScores string defined. This is to contain the string that exists in the Highscores.txt file before a new score is added
        Dim ExistingScores As String
        'Entirety of Highscores.txt file is read to ExistingScores string using StreamReader object
        Using Reader As New System.IO.StreamReader("Highscores.txt")
            ExistingScores = Reader.ReadToEnd()
        End Using
        'Score written to Highscores.txt file using StreamWriter object
        Using Writer As New System.IO.StreamWriter("Highscores.txt")
            'If there are no existing scores, the record of the current user and score is written straight to the file in the format !User|Score. A 'T' character is also placed at the start of the file to show that the scores present are sorted in ascending order. The score is calculated from the current time stored in the clock and the size of the board. The lower the score the better so it is calculated by dividing the user's time by the number of tiles in the board then multiplying this by 1000. This means a user has equal chance of getting a good score for any board size. It is multuplied by 1000 at the end to make it a nicer number.
            If ExistingScores = "" Then
                Writer.Write("T!" & User & "|" & CStr(Math.Floor((Board.Clock.Time / (n ^ 2)) * 1000)))
            Else
                'If there are existing scores and the existing scores are sorted as denoted by the character 'T' at the start of the string, then an 'F' is written to the start of the file followed be the existing scores string starting at the second character followed by the new score record
                If ExistingScores.First = "T" Then
                    Writer.Write("F" & Mid(ExistingScores, 2, Len(ExistingScores)) & "!" & User & "|" & CStr(Math.Floor((Board.Clock.Time / (n ^ 2)) * 1000)))
                    'If there are existing scores and they are not sorted then the starting character does not need to be changed so the just ExistingScores string followed by the new score record are written to the file.
                Else
                    Writer.Write(ExistingScores & "!" & User & "|" & CStr(Math.Floor((Board.Clock.Time / (n ^ 2)) * 1000)))

                End If
            End If

        End Using
    End Sub

End Class