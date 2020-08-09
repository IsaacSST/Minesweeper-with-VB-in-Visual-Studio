Public Class MainControls
    'All user controls required in menu instantiated here. Those that require event handling by this class are declared WithEvents to enable this.
    Private WithEvents Login As New Button
    Private WithEvents CreateAccount As New Button
    Private WithEvents ViewHighscores As New Button
    Private WithEvents NewGame As New Button
    Private WithEvents LogOut As New Button
    Private WithEvents Start As New Button
    Private WithEvents Back As New Button
    Private WithEvents Help As New Button
    Private WithEvents Dimensions As New TextBox
    Private WithEvents Username As New TextBox
    Private WithEvents Password As New TextBox
    Private lblDimensions As New Label
    Private lblUsername As New Label
    Private lblPassword As New Label
    Private Warning As New Label
    Private Scores As New RankedList
    'FromPage Variable holds the menu page most recently visited before the current one so the back button knows where to go
    Private FromPage As String


    Private Sub MainControls_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Master form and menu object set to be the same upon the loading of the menu object. Up to this point the master form had no dimensions
        Me.Width = 300
        Master.Width = 300
        Me.Height = 350
        Master.Height = 350
        'Dimension and default property defining for all objects that will appear in the menu
        Warning.BackColor = Color.Transparent
        Warning.ForeColor = Color.Red
        Warning.Font = New Font("Arial", 6, FontStyle.Bold)
        Warning.Height = 10
        Warning.Width = 160
        Warning.TextAlign = ContentAlignment.MiddleCenter
        'Width of greater Menu object used to define the horizontal position of component objects so the component objects are always centered
        Warning.Left = Math.Floor(Me.Width / 2) - 90

        lblUsername.BackColor = Color.Transparent
        lblUsername.ForeColor = Color.Black
        lblUsername.Text = "Username:"
        lblUsername.Font = New Font("Arial", 7, FontStyle.Bold)
        lblUsername.Height = 20
        lblUsername.Width = 60
        lblUsername.Top = 10
        lblUsername.Left = Math.Floor(Me.Width / 2) - 40

        Username.BackColor = Color.White
        Username.ForeColor = Color.Black
        Username.Font = New Font("Arial", 7, FontStyle.Regular)
        Username.BorderStyle = BorderStyle.FixedSingle
        Username.Width = 60
        Username.Height = 20
        Username.Top = 30
        Username.Left = Math.Floor(Me.Width / 2) - 40

        lblPassword.BackColor = Color.Transparent
        lblPassword.ForeColor = Color.Black
        lblPassword.Text = "Password:"
        lblPassword.Font = New Font("Arial", 7, FontStyle.Bold)
        lblPassword.Height = 20
        lblPassword.Width = 60
        lblPassword.Top = 50
        lblPassword.Left = Math.Floor(Me.Width / 2) - 40

        Password.BackColor = Color.White
        Password.ForeColor = Color.Black
        Password.Font = New Font("Arial", 7, FontStyle.Regular)
        Password.BorderStyle = BorderStyle.FixedSingle
        Password.Width = 60
        Password.Height = 20
        Password.Top = 70
        Password.Left = Math.Floor(Me.Width / 2) - 40
        Password.PasswordChar = "*"

        lblDimensions.BackColor = Color.Transparent
        lblDimensions.ForeColor = Color.Black
        lblDimensions.Text = "Board size (10-50):"
        lblDimensions.Font = New Font("Arial", 7, FontStyle.Bold)
        lblDimensions.Height = 20
        lblDimensions.Width = 60
        lblDimensions.Top = 10
        lblDimensions.Left = Math.Floor(Me.Width / 2) - 40

        Dimensions.BackColor = Color.White
        Dimensions.ForeColor = Color.Black
        Dimensions.Font = New Font("Arial", 7, FontStyle.Regular)
        Dimensions.BorderStyle = BorderStyle.FixedSingle
        Dimensions.Width = 60
        Dimensions.Height = 20
        Dimensions.Top = 30
        Dimensions.Left = Math.Floor(Me.Width / 2) - 40

        Login.BackColor = Color.Gray
        Login.ForeColor = Color.Black
        Login.Font = New Font("Arial", 8, FontStyle.Regular)
        Login.Text = "Login"
        Login.TextAlign = ContentAlignment.MiddleCenter
        Login.Width = 70
        Login.Height = 40
        Login.Top = 100
        Login.Left = Math.Floor(Me.Width / 2) - 45

        NewGame.BackColor = Color.Gray
        NewGame.ForeColor = Color.Black
        NewGame.Font = New Font("Arial", 8, FontStyle.Regular)
        NewGame.Text = "New Game"
        NewGame.TextAlign = ContentAlignment.MiddleCenter
        NewGame.Width = 70
        NewGame.Height = 40
        NewGame.Top = 10
        NewGame.Left = Math.Floor(Me.Width / 2) - 45

        LogOut.BackColor = Color.Gray
        LogOut.ForeColor = Color.Black
        LogOut.Font = New Font("Arial", 8, FontStyle.Regular)
        LogOut.Text = "Log Out"
        LogOut.TextAlign = ContentAlignment.MiddleCenter
        LogOut.Width = 70
        LogOut.Height = 40
        LogOut.Top = 60
        LogOut.Left = Math.Floor(Me.Width / 2) - 45

        CreateAccount.BackColor = Color.Gray
        CreateAccount.ForeColor = Color.Black
        CreateAccount.Font = New Font("Arial", 8, FontStyle.Regular)
        CreateAccount.Text = "Create Account"
        CreateAccount.TextAlign = ContentAlignment.MiddleCenter
        CreateAccount.Width = 70
        CreateAccount.Height = 40
        CreateAccount.Top = 150
        CreateAccount.Left = Math.Floor(Me.Width / 2) - 45

        Start.BackColor = Color.Gray
        Start.ForeColor = Color.Black
        Start.Font = New Font("Arial", 8, FontStyle.Regular)
        Start.Text = "Start Game"
        Start.TextAlign = ContentAlignment.MiddleCenter
        Start.Width = 70
        Start.Height = 40
        Start.Top = 60
        Start.Left = Math.Floor(Me.Width / 2) - 45

        Help.BackColor = Color.Gray
        Help.ForeColor = Color.Black
        Help.Font = New Font("Arial", 8, FontStyle.Regular)
        Help.Text = "Help"
        Help.TextAlign = ContentAlignment.MiddleCenter
        Help.Width = 70
        Help.Height = 40
        Help.Top = 250
        Help.Left = Math.Floor(Me.Width / 2) - 45

        Back.BackColor = Color.Gray
        Back.ForeColor = Color.Black
        Back.Font = New Font("Arial", 8, FontStyle.Regular)
        Back.Text = "Back"
        Back.TextAlign = ContentAlignment.MiddleCenter
        Back.Width = 70
        Back.Height = 40
        Back.Top = 110
        Back.Left = Math.Floor(Me.Width / 2) - 45

        ViewHighscores.BackColor = Color.Gray
        ViewHighscores.ForeColor = Color.Black
        ViewHighscores.Font = New Font("Arial", 8, FontStyle.Regular)
        ViewHighscores.Text = "View best scores"
        ViewHighscores.TextAlign = ContentAlignment.MiddleCenter
        ViewHighscores.Width = 70
        ViewHighscores.Height = 40
        ViewHighscores.Top = 200
        ViewHighscores.Left = Math.Floor(Me.Width / 2) - 45
        'AddInitialForm procedure called to load the first page of the menu object
        AddInitialForm()
    End Sub
    Public Sub AddInitialForm()
        'FromPage set to main as this is the main page
        FromPage = "Main"
        'Help button re-positioned for other pages so must be moved back here
        Help.Top = 250
        'All controls needed for the main page are added and those that might be present but not wanted are removed
        Me.Controls.Add(lblUsername)
        Me.Controls.Add(Username)
        Me.Controls.Add(lblPassword)
        Me.Controls.Add(Password)
        Me.Controls.Add(Login)
        Me.Controls.Add(CreateAccount)
        Me.Controls.Add(ViewHighscores)
        Me.Controls.Add(Warning)
        Me.Controls.Add(Help)
        Me.Controls.Remove(Scores)
        Me.Controls.Remove(Back)


    End Sub
    Sub LogOutPress() Handles LogOut.MouseUp
        'ViewHighscores button repositioned to its default positioning on the main page
        ViewHighscores.Top = 200
        'Menu will be returned to main page so buttons not wanted are removed
        Me.Controls.Remove(NewGame)
        Me.Controls.Remove(LogOut)
        'Board removed from form
        Master.Controls.Remove(Master.Board)
        'User logged out by clearing User variable
        Master.User = ""
        'Menu object moved into gap left by removing the Board object and the new excess space removed
        Master.Width = Me.Width
        Master.Height = Me.Height
        Me.Left = 0
        Me.Top = 0
        'Menu object returned to main page by calling AddInitialForm
        AddInitialForm()
    End Sub
    Sub CreateNewAccount() Handles CreateAccount.MouseUp
        'Two different validation booleans for two stages of validation. One does not occur unless the first has been passed
        Dim GoAhead1 As Boolean = True
        Dim GoAhead2 As Boolean = True
        'ExistingRecords string defined. It is to contain the contents of the UserLoginRecord.txt file before the file is edited by this procedure
        Dim ExistingRecords As String

        'Creation of user login record is halted if the Username or Password boxes are left unfilled. The user is told why through the Warning label
        If Username.Text = "" Or Password.Text = "" Then
            Warning.Text = "A box is blank"
            GoAhead1 = False

        Else
            'Loop that goes through entire length of the contents of Username text box and checks each character for spaces, '|'s or '!'s. These characters cannot be allowed in the username so if it contains any then the account creation is not allowed to go ahead.
            For i = 1 To Len(Username.Text)
                Select Case Mid(Username.Text, i, 1)
                    Case " ", "|", "!"
                        Warning.Text = "Cannot use spaces ,'!' or '|' In username"
                        GoAhead1 = False

                End Select

            Next
            'Loop that goes through entire length of the contents of Password text box and checks each character for spaces, '|'s or '!'s. These characters cannot be allowed in the password so if it contains any then the account creation is not allowed to go ahead.
            For i = 1 To Len(Password.Text)
                Select Case Mid(Password.Text, i, 1)
                    Case " ", "|", "!"
                        Warning.Text = "Cannot use spaces ,'!' or '|' In password"
                        GoAhead1 = False

                End Select

            Next
        End If
        'If the contents of the Username and Password textboxes are valid as a potential login, then GoAhead1 will not have been changed to false, so the contents of this selection statement will go ahead.
        If GoAhead1 = True Then
            'Proposed username is checked against the existing records to see if the account already exists. If it does then GoAhead2 is set to false so account creation is not allowed to go ahead.
            Using Reader As New System.IO.StreamReader("UserLoginRecord.txt")
                'Assigning entire file string to ExistingRecords variable
                ExistingRecords = Reader.ReadToEnd()
                'Loop goes through entire length of ExistingRecords and checks each character. If it finds an '!' then a username will follow, so the section of string immediately following the '!' with the same length as the new proposed Username is checked to see if it is the same as the proposed Username. If it is, then the user is informed that their desired account already exists through a message displayed in the Warning label.
                For i = 1 To Len(ExistingRecords)
                    If Mid(ExistingRecords, i, 1) = "!" Then
                        If Username.Text = Mid(ExistingRecords, i + 1, Len(Username.Text)) Then
                            Warning.Text = "User already exists"
                            GoAhead2 = False
                        End If
                    End If
                Next
            End Using
            'If the proposed Username does not already exist then account creation goes ahead
            If GoAhead2 = True Then
                'The existing records are written into the UserLoginRecord.txt file along with the new account record at the end
                Using Writer As New System.IO.StreamWriter("UserLoginRecord.txt")
                    Dim RecordString = ExistingRecords & ("!" & Username.Text & "|" & Password.Text)
                    Writer.Write(RecordString)
                End Using
                'It is presumed the user will want to login immediately after creating an account so this is automatically done for them here
                Master.User = Username.Text
                LoginLanding()


            End If
        End If
        'Username and Password text boxes cleared ready for next use
        Username.Clear()
        Password.Clear()

    End Sub
    Sub UserLogin() Handles Login.MouseUp
        'UserFound boolean defined, this holds whether or not the account the user has attempted to log in with has been found in the UserLoginRecord.txt file
        Dim UserFound As Boolean = False
        'Login procedure not allowed to go ahead if one of the boxes is blank as this would definitely not be a valid login. User is informed through Warning label
        If Username.Text = "" Or Password.Text = "" Then
            Warning.Text = "A box is blank"

        Else
            'Code block checks if account exists and if combination of username and password is correct
            Using Reader As New System.IO.StreamReader("UserLoginRecord.txt")
                'Existing Records string defined as being entirety of contents of UserLoginRecord.txt file
                Dim ExistingRecords As String = Reader.ReadToEnd()
                'Loop goes through entire length of the Existing Records string
                For i = 1 To Len(ExistingRecords)
                    'If current character in string is an '!' then a user record will follow, so program proceeds to check the contents of Username and Password boxes against the record
                    If Mid(ExistingRecords, i, 1) = "!" Then
                        'The section of string immediately following the '!' with the same length as the attempted login username is checked to see if it is the same as the attempted login username. If it is then the account exists so UserFound is set to true and the program proceeds to check the password
                        If Mid(ExistingRecords, i + 1, Len(Username.Text)) = Username.Text Then
                            UserFound = True
                            'If the section of string in the existing records immediately following the Username with length one more than the length of the entered Password is  a '|' followed by the attempted password, then the password is correct and the user is logged in and sent to the Login landing page
                            If Mid(ExistingRecords, i + Len(Username.Text) + 1, Len(Password.Text) + 1) = ("|" & Password.Text) Then
                                Master.User = Username.Text
                                LoginLanding()

                                'Otherwise, they are informed through the warning label that the password is incorrect
                            Else
                                Warning.Text = "Password Incorrect"
                            End If
                        End If
                    End If
                Next
                'If the username was not matched at any point when scanning through the record file then there is no record of the attempted login account so the login is rejected and the user is informed their account is not recognised through the Warning label
                If UserFound = False Then

                    Warning.Text = "User not recognised"
                End If
            End Using
        End If
        'Username and Password text boxes cleared ready for next use
        Username.Clear()
        Password.Clear()
    End Sub
    Sub LoginLanding()
        'FromPage set to "Login" to tell the Back button to return to the Login landing page from any subsequent pages
        FromPage = "Login"
        'Warning label is cleared in case it was displaying a warning left over from the Main page
        Warning.Text = ""
        'All potentially present user controls that are not wanted on this page are removed from the form
        Me.Controls.Remove(lblUsername)
        Me.Controls.Remove(Username)
        Me.Controls.Remove(lblPassword)
        Me.Controls.Remove(Password)
        Me.Controls.Remove(Login)
        Me.Controls.Remove(CreateAccount)
        Me.Controls.Remove(lblDimensions)
        Me.Controls.Remove(Dimensions)
        Me.Controls.Remove(Start)
        Me.Controls.Remove(Back)
        Me.Controls.Remove(Scores)


        'ViewHighscores and Help buttons repositioned in accordance with the layout of this page
        ViewHighscores.Top = 110
        Help.Top = 160
        'All controls that are potentially not already on the form that are needed on this page are added
        Me.Controls.Add(NewGame)
        Me.Controls.Add(LogOut)
        Me.Controls.Add(ViewHighscores)

    End Sub
    Sub NewGameLanding() Handles NewGame.MouseUp
        'All potentially present user controls that are not wanted on this page are removed from the form
        Me.Controls.Remove(NewGame)
        Me.Controls.Remove(LogOut)
        Me.Controls.Remove(ViewHighscores)
        'All controls that are potentially not already on the form that are needed on this page are added
        Me.Controls.Add(lblDimensions)
        Me.Controls.Add(Dimensions)
        Me.Controls.Add(Start)
        Me.Controls.Add(Back)
        'The Back button is repositioned in accordance with the layout of this page
        Back.Top = 110

    End Sub

    Sub InitiateGame() Handles Start.MouseUp
        'If contents of Dimensions text box is a number between 10 and 40 inclusive then this procedure will create a game board with side length of that number
        If IsNumeric(Dimensions.Text) = True Then
            If Dimensions.Text <= 40 And Dimensions.Text >= 10 Then
                Master.n = CInt(Dimensions.Text)
                'Number of mines set depending on the size of the board. Benchmark is that there should be 50 mines in a 16x16 board and the calculation reflects this
                Master.Mines = Math.Floor(50 * ((Master.n) ^ 2 / 256))
                Master.AddBoard()
                'Menu object returned to the login landing page
                LoginLanding()
            Else
                'User informed through Warning label if their input is invalid
                Warning.Text = "Must be number 10-40"
            End If
        Else
            Warning.Text = "Must be number 10-40"
        End If

    End Sub
    Sub HelpIsAtHand() Handles Help.MouseUp
        'Help.txt file in same directory as program executable is opened in default text file viewer program
        Process.Start("Help.txt")
    End Sub
    Sub HighScores() Handles ViewHighscores.MouseUp
        'All potentially present user controls that are not wanted on this page are removed from the form
        Me.Controls.Remove(lblUsername)
        Me.Controls.Remove(Username)
        Me.Controls.Remove(lblPassword)
        Me.Controls.Remove(Password)
        Me.Controls.Remove(Login)
        Me.Controls.Remove(CreateAccount)
        Me.Controls.Remove(LogOut)
        Me.Controls.Remove(NewGame)
        Me.Controls.Remove(ViewHighscores)
        'All controls that are potentially not already on the form that are needed on this page are added
        Me.Controls.Add(Scores)
        Me.Controls.Add(Back)
        'Back and help buttons repositioned in accordance with the layout of this page
        Back.Top = 220
        Help.Top = 270
        'Variables defined
        Dim ScoreCounter As Integer
        Dim RecordingCounter As Integer
        Dim CurrentScores As String
        Dim AlreadySorted As Boolean
        Dim RecordUsername As Boolean = False
        Dim RecordScore As Boolean = False
        Dim CurrentChar As Char
        'Entire contents of Highscores.txt file read as a string to CurrentScores variable. This means CurrentScores now contains all HighScore records
        Using Reader As New System.IO.StreamReader("Highscores.txt")
            CurrentScores = Reader.ReadToEnd()
        End Using

        'Selection statement to avoid wasting time, if there are no highscores then the user will just see a blank highscores list with the Warning label showing "No Highscores" and a Back button at the bottom of the form
        If CurrentScores = "" Then
            Warning.Text = "No Highscores"
        Else
            'If the first character in the file is a T then it is 'True' that the scores present are already sorted in ascending order so time does not need to be wasted on a sorting algorithm later on
            If Mid(CurrentScores, 1, 1) = "T" Then
                AlreadySorted = True
            Else
                AlreadySorted = False
            End If
            'An '!' comes before every username in every score record, so iterating through the length of the score string and counting how many '!' characters there are will give the number of highscores there are
            For i = 1 To Len(CurrentScores)
                If Mid(CurrentScores, i, 1) = "!" Then
                    ScoreCounter = ScoreCounter + 1
                End If
            Next
            'A 2 dimensional array is defined that is to contain all of the Highscores in the highscores record. One dimension for Username and the other for Score.
            Dim ScoreArray(ScoreCounter, 2)
            'ScoreCounter will be needed again later so another counter is needed for use in recording the Highscores into the ScoreArray
            RecordingCounter = ScoreCounter
            'For loop goes through entire length of the Highscores record and fills the ScoreArray character by character
            For i = 1 To Len(CurrentScores)
                'CurrentChar set to be character at position in CurrentScores string corresponding to the iteration number of the loop
                CurrentChar = Mid(CurrentScores, i, 1)
                'If the CurrentChar in the string is an '!' then the characters following it will spell out a Username, so the CurrentChars in the following iterations will need to be recorded in the ScoreArray as a Username, so RecordUsername is set to true and RecordScore is set to false so the program knows it is recording a username. The loop will continue recording characters as part of a particular username until a '|' character is reached. The '|' character always comes before a score so RecordScore is set to True and RecordUsername is set to false and the loop will record the CurrentChars into the ScoreArray as a Score until another '!' is reached. At which point a new highscore record will be started in a new row of the ScoreArray. The RecordingCounter variable is decremented every time an '!' is reached and this is how the program knows when to start recording into a new row in the ScoreArray
                If CurrentChar = "!" Then
                    RecordUsername = True
                    RecordScore = False
                    RecordingCounter = RecordingCounter - 1
                ElseIf CurrentChar = "|" Then
                    RecordUsername = False
                    RecordScore = True
                    'Recording code is inside selection statement so program only attempts to record when the CurrentChar is not one of the pointer characters: '!' or '|'
                Else
                    'Two booleans used so first character which will be 'T' or 'F' is not recorded into the array
                    If RecordUsername = True Then
                        'All Usernames recorded into column 1
                        'Row number defined by ScoreCounter - RecordingCounter. This means Highscores are recorded starting in row 1 going up to row ScoreCounter, as RecordingCounter is inititally set as equal to ScoreCounter but is decremented before each Highscore record is recorded, meaning for the first Highscore record, RecordingCounter will be equal to ScoreCounter-1, and for the last Highscore record, RecordingCounter will be equal to 0
                        'Usernames and passwords recorded by concatenating CurrentChar to previous contents of ScoreArray position
                        ScoreArray(ScoreCounter - RecordingCounter, 1) = (ScoreArray(ScoreCounter - RecordingCounter, 1) & CurrentChar)
                    ElseIf RecordScore = True Then
                        'All Scores recorded into column 2
                        ScoreArray(ScoreCounter - RecordingCounter, 2) = (ScoreArray(ScoreCounter - RecordingCounter, 2) & CurrentChar)

                    End If
                End If
            Next
            'If Highscores are not already sorted then they are sorted in ascending order by calling Quicksort function
            If AlreadySorted = False Then
                'All of the scores converted to integer form from string form so they will work in the sorting algorithm
                For i = 1 To ScoreCounter
                    ScoreArray(i, 2) = CInt(ScoreArray(i, 2))
                Next
                'Quicksort function takes the ScoreArray and number of Scores in the array and returns the sorted array
                ScoreArray = QuickSort(ScoreArray, ScoreCounter)
            End If
            'All labels in the Highscore display object set to default background colour of transparent
            For i = 1 To 20
                Scores.DisplayArray(i, 1).BackColor = Color.Transparent
                Scores.DisplayArray(i, 2).BackColor = Color.Transparent
            Next
            'The highscore display array can fit 20 scores so if the number of highscores recorded is greater than or equal to 20 then scores 1 to 20 are taken from the ScoreArray and put into the highscores display array. However if the number of scores is less than 20 then doing this will cause an out of bounds error as there will not be 20 indexes in the ScoreArray, so a seperate loop is needed that only goes up to the size of the ScoreArray and leaves the rest of the Highscore spaces in the display array blank.
            If ScoreCounter >= 20 Then

                For i = 1 To 20

                    'Leftmost labels in display array contain the Username of the scorers have the position from 1 to 20 written before the usernames
                    Scores.DisplayArray(i, 1).Text = i & ": " & ScoreArray(i, 1)
                    'Rightmost labels in display array contain the Scores of the top scorers
                    Scores.DisplayArray(i, 2).Text = ScoreArray(i, 2)
                    'If the user is logged in then their highscores are highlighed in yellow on the display array
                    If ScoreArray(i, 1) = Master.User And Master.User <> "" Then
                        Scores.DisplayArray(i, 1).BackColor = Color.Yellow
                        Scores.DisplayArray(i, 2).BackColor = Color.Yellow
                    End If
                Next
            Else
                For i = 1 To ScoreCounter

                    Scores.DisplayArray(i, 1).Text = i & ": " & ScoreArray(i, 1)
                    Scores.DisplayArray(i, 2).Text = ScoreArray(i, 2)
                    If ScoreArray(i, 1) = Master.User And Master.User <> "" Then
                        Scores.DisplayArray(i, 1).BackColor = Color.Yellow
                        Scores.DisplayArray(i, 2).BackColor = Color.Yellow
                    End If
                Next
            End If
        End If

    End Sub
    Sub GoBack() Handles Back.MouseUp
        'Warning text cleared as it may contain "No Highscores" from the Highscores viewing page
        Warning.Text = ""
        'If the previous page is the Login Landing page then when the back button is clicked the user is returned to the Login Landing page. The only other page they could be returning to is the Initial page so a simple Else is used that calls the AddInitialForm sub which returns the user to this page
        If FromPage = "Login" Then
            LoginLanding()

        Else
            AddInitialForm()
        End If
    End Sub
    Sub ClickUsername() Handles Username.MouseUp
        'When the user clicks inside the Username text box, this means they are going to type something so the Warning label is cleared to avoid them thinking the thing they are typing is invalid when it might not be
        Warning.Text = ""
    End Sub
    Sub ClickPassword() Handles Password.MouseUp
        'When the user clicks inside the Password text box, this means they are going to type something so the Warning label is cleared to avoid them thinking the thing they are typing is invalid when it might not be
        Warning.Text = ""
    End Sub
    Sub ClickDimensions() Handles Dimensions.MouseUp
        'When the user clicks inside the Dimensions text box, this means they are going to type something so the Warning label is cleared to avoid them thinking the thing they are typing is invalid when it might not be
        Warning.Text = ""
    End Sub
    Function QuickSort(Arr(,), Length)
        'If there is only one score then the array can't not be in order, so the array that was passed in is returned.
        If Length = 1 Then
            Return Arr
        Else
            'The pivot position is defined as being the score exactly halfway along the array if the number of scores is odd. If the number of scores is even then there is no score that is exactly at halfway so the one just past halfway is used.
            Dim Pivot As Double = Math.Ceiling((Length + 1) / 2)
            Dim LessThanPivot As Integer = 0
            Dim GreaterThanPivot As Integer = 0
            Dim LessThanCounter As Integer
            Dim GreaterThanCounter As Integer
            'NewArr shall contain a version of the array that was passed in sorted in ascending order by the Score
            Dim NewArr(Length, 2)
            'Number of scores in the array that are less than or equal to pivot score counted. LessThanPivot used as counter name as LessThanOrEqualToPivot seems too long as a variable name
            For i = 1 To Length
                If Arr(i, 2) <= Arr(Pivot, 2) Then
                    LessThanPivot = LessThanPivot + 1
                End If
            Next
            'LessThanPivot reduced by 1 as the pivot itself will have been counted and it should not be
            LessThanPivot = LessThanPivot - 1
            'The number of scores strictly greater than the pivot score will be equal to the total number of scores (Length) that aren't the pivot (-1) and aren't otherwise less than or equal to the pivot (-LessThanPivot)
            GreaterThanPivot = Length - LessThanPivot - 1

            'Provided there are some scores that are less than or equal to the pivot score, these scores are placed into their own array and then sorted using quicksort, and then placed into the NewArr starting in the first row. This means this function is recursive as it calls itself. 
            If LessThanPivot <> 0 Then
                'A new array is defined with space for all the score records of scores less than the pivot score
                Dim arrLessThanPivot(LessThanPivot, 2)
                'Loop that goes through all score records in the array Arr that was passed into the function and copies all will scores less than or equal to the pivot score into the new array arrLessThanPivot
                For i = 1 To Length
                    If Arr(i, 2) < Arr(Pivot, 2) Then
                        LessThanCounter = LessThanCounter + 1
                        arrLessThanPivot(LessThanCounter, 1) = Arr(i, 1)
                        arrLessThanPivot(LessThanCounter, 2) = Arr(i, 2)
                    End If
                    If i <> Pivot And Arr(i, 2) = Arr(Pivot, 2) Then
                        LessThanCounter = LessThanCounter + 1
                        arrLessThanPivot(LessThanCounter, 1) = Arr(i, 1)
                        arrLessThanPivot(LessThanCounter, 2) = Arr(i, 2)
                    End If
                Next
                'Another array SortedLessThan is defined that shall contain the arrLessThanPivot array in sorted form
                Dim SortedLessThan(LessThanPivot, 2)
                'SortedLessThan set to be equal to the returned array from sorting the arrLessThanPivot array using QuickSort
                SortedLessThan = QuickSort(arrLessThanPivot, LessThanPivot)
                'Scores that are Less than or equal to the pivot, now in sorted in ascending order are put into the first rows of the NewArr
                For i = 1 To LessThanPivot
                    NewArr(i, 1) = SortedLessThan(i, 1)
                    NewArr(i, 2) = SortedLessThan(i, 2)
                Next
            End If
            'Provided there are some scores that are greater than or equal to the pivot score, these scores are placed into their own array and sorted using quicksort. These sorted scores are then added to the NewArr after the scores that are Less than or equal to the pivot, leaving space for the pivot inbetween the two sets of scores
            If GreaterThanPivot <> 0 Then
                'A new array is defined with space for all the score records of scores greater than the pivot score
                Dim arrGreaterThanPivot(GreaterThanPivot, 2)
                'Loop that goes through all score records in the array Arr that was passed into the function and copies all will scores greater than the pivot score into the new array arrGreaterThanPivot
                For i = 1 To Length

                    If Arr(i, 2) > Arr(Pivot, 2) Then
                        GreaterThanCounter = GreaterThanCounter + 1
                        arrGreaterThanPivot(GreaterThanCounter, 1) = Arr(i, 1)
                        arrGreaterThanPivot(GreaterThanCounter, 2) = Arr(i, 2)
                    End If
                Next
                'Another array SortedGreaterThan is defined that shall contain the arrGreaterThanPivot array in sorted form
                Dim SortedGreaterThan(GreaterThanPivot, 2)
                'SortedGreaterThan set to be equal to the returned array from sorting the arrGreaterThanPivot array using QuickSort
                SortedGreaterThan = QuickSort(arrGreaterThanPivot, GreaterThanPivot)
                'The Sorted scores greater than the pivot are placed in the NewArr starting in the Row after where the pivot is to go
                For i = LessThanPivot + 2 To Length
                    'Row Indexes defined such that the NewArr index always starts just after where the pivot should be and the SortedGreaterThan index always starts at 1, the first score that is greater than the pivot
                    NewArr(i, 1) = SortedGreaterThan(i - LessThanPivot - 1, 1)
                    NewArr(i, 2) = SortedGreaterThan(i - LessThanPivot - 1, 2)
                Next
            End If
            'Pivot score record copied into the NewArr in the space left for it
            NewArr(LessThanPivot + 1, 1) = Arr(Pivot, 1)
            NewArr(LessThanPivot + 1, 2) = Arr(Pivot, 2)
            'NewArr is now a sorted version of the array Arr which was passed in, so the function is complete and NewArr is returned
            Return NewArr
        End If
    End Function
End Class
