Public Class Game

    'Defining of variables and instantiating objects
    'setting default values of some variables also
    Public n As Integer
    Public Mines As Integer
    Public GameStarted As Boolean = False
    Public arrBoardPieces(n, n) As MinesweeperButton
    Public NumberRevealed As Integer = 0
    Public Clock As New TimerBox
    Public MineCounter As New Label
    Public ReactionFaces As New RestartGameButton
    Public BlockGame As Boolean
    Private Sub Game_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Redim of arrBoardPieces required as the value of n changed to desired value after the Game has been instantiated,
        'meaning without Redim the array would have dimensions 0,0
        ReDim arrBoardPieces(n, n)
        'Assigning dimensions for the Game object based on the number of tiles in the array. These dynamic dimensions mean that the window is never unnecessarily large or too small to fit all the objects
        Me.Height = 16 * n + 95
        Me.Width = 16 * n + 48

        'Positioning and properties set for ReactionFaces, which is a RestartGameButton instantiation
        'Horizontal positioning set based on the width of the Game object so that it is in the middle of it.

        ReactionFaces.Top = 5
        ReactionFaces.Left = Math.Floor(Me.Width / 2) - 12
        ReactionFaces.BackColor = Color.White
        ReactionFaces.BringToFront()

        'Positioning set for clock object
        'Horizontal position set so it is always 2/3 of the way along the Board from left

        Clock.Top = 5
        Clock.Left = 2 * Math.Floor(Me.Width / 3)

        'Position and properties set for MineCounter, which is a Label.
        'Horizontal position set so that it is 1/3 the way along the board from left. Set based on the position of the clock as this has already been set at 2/3 the way along.
        MineCounter.Width = 40
        MineCounter.Height = 20
        MineCounter.Top = 5
        MineCounter.Left = Math.Floor(Clock.Left / 2) - 40
        MineCounter.BackColor = Color.Black
        MineCounter.ForeColor = Color.Red
        MineCounter.Font = New Font("Arial", 10, FontStyle.Bold)
        MineCounter.Text = Mines
        MineCounter.BorderStyle = BorderStyle.Fixed3D

        'All three objects added to board so that they will be visible
        Me.Controls.Add(ReactionFaces)
        Me.Controls.Add(MineCounter)
        Me.Controls.Add(Clock)
        'Loop that iterates through every index in the arrBoardPieces array. It instantiates every position as a new MinesweeperButton and sets its dimensions based on the x and y (i and j) of the index. Sets the Tilei and Tilej properties of each MinesweeperButton to be equal to its position in the array which is the current i and j of the loop. Adds each object so it is visible on the Game object. This loop creates an interesting cascade effect when loading in the buttons which is particularly noticable with larger boards.
        For j = 1 To n
            For i = 1 To n
                arrBoardPieces(i, j) = New MinesweeperButton
                arrBoardPieces(i, j).Left = 16 + ((i - 1) * 16)
                arrBoardPieces(i, j).Top = ((n - j) * 16) + 41
                arrBoardPieces(i, j).BringToFront()
                arrBoardPieces(i, j).Tilei = i
                arrBoardPieces(i, j).Tilej = j
                Me.Controls.Add(arrBoardPieces(i, j))
            Next i
        Next j
    End Sub

    Public Sub RecursiveReveal(ByRef i As Integer, ByRef j As Integer)
        'Assigning default values for counter bound variables.
        'For the following counters , left to right and top to bottom are the directions of index increasing.
        'l stands for left and takes the value of the relative horizontal index to the point passed in of the leftmost tile that will be revealed by this procedure
        'r stands for right and takes the value of the relative horizontal index to the point passed in of the rightmost tile that will be revealed by this procedure
        't stands for top and takes the value of the relative vertical index to the point passed in of the topmost tile that will be revealed by this procedure
        'b stands for botton and takes the value of the relative vertical index to the point passed in of the bottommost tile that will be revealed by this procedure
        Dim l As Integer = -1
        Dim r As Integer = 1
        Dim t As Integer = -1
        Dim b As Integer = 1

        'Selection statements to avoid exceptions where the program tries to access indexes outside of the board. They check if the tile the procedure is being called on is at the edge of the board and if it is the relevent counter is set to 0
        Select Case i
            Case 1
                l = 0
            Case n
                r = 0
        End Select

        Select Case j
            Case 1
                t = 0
            Case n
                b = 0
        End Select
        'For Loop that goes through every tile surrounding the tile with index (i,j) in the array and calls the 'Reveal' procedure of the tile if it is not flagged and not already revealed.
        'This creates recursion as the 'Reveal' procedure will call this procedure on the new tile index if it is a blank tile
        'The end result after all the calls have completed is a large revealed area that will contain all tiles connected by blank tiles to the first tile called.
        For k = l To r
            For m = t To b
                If arrBoardPieces(i + k, j + m).Revealed = False And arrBoardPieces(i + k, j + m).Flagged = False Then

                    arrBoardPieces(i + k, j + m).Reveal()

                End If
            Next m
        Next k
    End Sub
    Public Function FlagCheck(ByVal TileType As Integer, ByVal i As Integer, ByVal j As Integer)
        'Assigning default values for counter bound variables.
        'For the following variables , left to right and top to bottom are the directions of index increasing.
        'l stands for left and takes the value of the relative horizontal index to the point passed in of the leftmost tile that will be checked by this procedure
        'r stands for right and takes the value of the relative horizontal index to the point passed in of the rightmost tile that will be checked by this procedure
        't stands for top and takes the value of the relative vertical index to the point passed in of the topmost tile that will be checked by this procedure
        'b stands for botton and takes the value of the relative vertical index to the point passed in of the bottommost tile that will be checked by this procedure
        Dim l As Integer = -1
        Dim r As Integer = 1
        Dim t As Integer = -1
        Dim b As Integer = 1
        'FlagCount is a counter and takes the number of flags currently counted that surround the tile with index (i,j)
        Dim FlagCount As Integer = 0
        'Selection statements to avoid exceptions where the program tries to access indexes outside of the board. They check if the tile the procedure is being called on is at the edge of the board and if it is the relevent counter bound is set to 0
        Select Case i
            Case 1
                l = 0
            Case n
                r = 0
        End Select

        Select Case j
            Case 1
                t = 0
            Case n
                b = 0
        End Select
        'For loop that goes through every tile surrounding the tile with index (i,j) using counter bound variables, and counts how many are flagged by incrementing the FlagCount counter if the current tile is flagged.
        For k = l To r
            For m = t To b
                If arrBoardPieces(i + k, j + m).Flagged = True Then
                    FlagCount = FlagCount + 1
                End If
            Next m
        Next k
        'Returns true if the number of flags surrounding the tile satisfies the TileType of the tile, otherwise returns false.
        If TileType = FlagCount Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub ResetGame()
        'Makes it so the game object cannot be interacted with. All interaction event handlers check the BlockGame variable before they do anything
        BlockGame = False
        'Resets the board to its default state if it is not already in its default state
        If GameStarted = True Then
            GameStarted = False
            NumberRevealed = 0
            Clock.Time = 1
            Clock.Text = 0
            Clock.Timer.Stop()
            MineCounter.Text = Mines
            For i = 1 To n
                For j = 1 To n
                    arrBoardPieces(i, j).TileType = 0
                    arrBoardPieces(i, j).Revealed = False
                    arrBoardPieces(i, j).Flagged = False
                    arrBoardPieces(i, j).BackgroundImage = My.Resources.button
                Next
            Next
        End If
    End Sub
    Public Sub LargePress(ByVal i As Integer, ByVal j As Integer, ByVal Up As Boolean)
        'Assigning default values for counter bound variables.
        'For the following variables , left to right and top to bottom are the directions of index increasing.
        'l stands for left and takes the value of the relative horizontal index to the point passed in of the leftmost tile that will be pressed or unpressed by this procedure
        'r stands for right and takes the value of the relative horizontal index to the point passed in of the rightmost tile that will be pressed or unpressed by this procedure
        't stands for top and takes the value of the relative vertical index to the point passed in of the topmost tile that will be pressed or unpressed by this procedure
        'b stands for botton and takes the value of the relative vertical index to the point passed in of the bottommost tile that will be pressed or unpressed by this procedure
        Dim l As Integer = -1
        Dim r As Integer = 1
        Dim t As Integer = -1
        Dim b As Integer = 1
        'Selection statements to avoid exceptions where the program tries to access indexes outside of the board. They check if the tile the procedure is being called on is at the edge of the board and if it is the relevent counter bound is set to 0
        Select Case i
            Case 1
                l = 0
            Case n
                r = 0
        End Select

        Select Case j
            Case 1
                t = 0
            Case n
                b = 0
        End Select
        'Loops through a 3x3 square of tiles with center at the tile with index (i,j) that was passed into the procedure. For each tile, provided it is not flagged, its background image is set according to the value of the 'Up' boolean that is passed into the procedure. If Up is false, the buttons are being pressed so their background image is set to the 'pressed' resource. If Up is true then the LargePress is being released so the background image should be set to the default 'button' resource.
        For k = l To r
            For m = t To b
                If Up = False Then
                    If arrBoardPieces(i + k, j + m).Flagged = False And arrBoardPieces(i + k, j + m).Revealed = False Then
                        arrBoardPieces(i + k, j + m).BackgroundImage = My.Resources.pressed

                    End If
                Else
                    If arrBoardPieces(i + k, j + m).Flagged = False And arrBoardPieces(i + k, j + m).Revealed = False Then
                        arrBoardPieces(i + k, j + m).BackgroundImage = My.Resources.button

                    End If
                End If
            Next m
        Next k
    End Sub
    Public Sub GenerateMines(ByVal x As Integer, ByVal y As Integer)
        'Stores the number of mines left to be placed. Initial value is the total number of desired mines as is stored in the 'Mines' variable.
        Dim MineCounter As Integer = Mines
        'Stores the total number of tiles in the array left to be processed. Initial value is the total number of tiles in the array
        Dim TileTotal = n ^ 2
        'An array that contains a number of 1s equal to the percentage chance that the tile being processed should contain a mine
        Dim arrProb(100) As Integer
        'The index at which the arrProb will be accessed to see if a tile should be a mine
        Dim Index As Integer
        'These two booleans store whether or not the current tile being processed is outside of the bounds of the 3x3 square with center (x,y) that should be left clear of mines
        Dim AllCleari As Boolean = True
        Dim AllClearj As Boolean = True
        'Stores whether or not the algorithm should go ahead based on if the tile is outside of the 3x3 square around the tile at position (x,y)
        Dim GoAhead As Boolean = True
        'Stores the calculated percentage chance that a mine will be placed on the current tile
        Dim Percentage As Integer
        'Loop goes through all x co-ordinates in the array arrBoardPieces
        For i = 1 To n
            'Sets default value of Allcleari boolean to True, otherwise it will be permanently false for all iterations after the first time it is set to false
            AllCleari = True
            'Sets Allcleari to false if current x coordinate is within the x coordinates of the 3x3 mine free square around the clicked tile (x,y)
            Select Case i
                Case x - 1, x, x + 1
                    AllCleari = False

            End Select
            'Loop goes through all y co-ordinates in the array arrBoardPieces. Combined with the outer x loop, this means every tile in the array is visited
            For j = 1 To n
                'Sets default values of AllClearj and GoAhead booleans to true
                AllClearj = True
                GoAhead = True
                'Sets Allclearj to false if current y coordinate is within the y coordinates of the 3x3 mine free square around the clicked tile (x,y)
                Select Case j
                    Case y - 1, y, y + 1
                        AllClearj = False
                End Select
                'Sets GoAhead boolean to false if the current tile is within the 3x3 mine free square around the called on tile (x,y)
                If AllCleari = False And AllClearj = False Then
                    GoAhead = False
                End If
                'Mine placement algorithm proceeds if the current tile is outside of the 3x3 mine free square
                If GoAhead Then
                    'Algorithm proceeds if there are still mines left to place, otherwise nothing is done
                    If MineCounter > 0 Then
                        'Sets the Percentage variable to be equal to the percentage of tiles left to process that need to contain a mine
                        Select Case (MineCounter / TileTotal) * 100
                            Case Is >= 100
                                Percentage = 100
                            Case Is >= 0, Is < 100
                                'For the algorithm to work, Percentage must be an integer so the actual percentage is rounded up
                                Percentage = Math.Ceiling(((MineCounter / TileTotal) * 100))
                        End Select
                        'Sets a number, equal to the variable Percentage, of indexes in the arrProb array to 1
                        For k = 1 To Percentage

                            arrProb(k) = 1
                        Next
                        'Index randomly chosen from the arrProb array. This is essentially the moment when it is decided whether or not the current tile should be a mine based on the percentage chance that it should be
                        Index = RandomIndex(1, 100)
                        'If the index in the arrProb array corresponding to the Index variable is equal to 1, then the current tile should be a mine, so its TileType is set to 9, which means it is a mine, and the MineCounter is decremented.

                        If arrProb(Index) = 1 Then
                            arrBoardPieces(i, j).TileType = 9
                            MineCounter = MineCounter - 1
                        End If
                    End If
                End If
                'Sets all array indexes to 0 to avoid incorrect percentage for the next tile if it should have a lower percentage
                For clear = 1 To 100
                    arrProb(clear) = 0
                Next
                'The TileTotal, ie the number of tiles left to process is decremented and the loop moves on to the next tile
                TileTotal = TileTotal - 1
            Next j
        Next i
    End Sub
    Public Sub GameEnd(ByVal x As Integer, ByVal y As Integer, Win As Boolean)
        'BlockGame variable set to True, meaning user's interactions with the game board no longer do anything
        BlockGame = True
        'Selects what to do based on whether the user has lost or not
        If Win = False Then
            'Setting face in the RestartGameButton at top of board to be a dead face
            ReactionFaces.Faces.Top = -48
            'Loop to go through all x co-ordinates in the arrBoardPieces array
            For i = 1 To n
                'Loop to go through all y co-ordinates in the arrBoardPieces array. Together with the outer loop, this creates a loop that goes through every position in the arrBoardPieces array.
                For j = 1 To n
                    'Selects what to do based on if the current loop position is the tile (x,y) which is passed into this procedure. This is the tile the user revealed to lose the game, so nothing is done if the loop position is at this tile as its background image was already updated when the user revealed it
                    If Not (i = x And j = y) Then
                        'Shows all non flagged mines remaining on the board
                        If arrBoardPieces(i, j).Flagged = False And arrBoardPieces(i, j).TileType = 9 Then
                            arrBoardPieces(i, j).BackgroundImage = My.Resources.mine
                        End If
                        'Shows all incorrectly placed flags to the user
                        If arrBoardPieces(i, j).Flagged = True And arrBoardPieces(i, j).TileType <> 9 Then
                            arrBoardPieces(i, j).BackgroundImage = My.Resources.wrong
                        End If

                    End If

                Next
            Next
            'Freezes the counter clock in the top right of the game
            Clock.Timer.Stop()
            'Displays a dialog box informing the user they have lost
            MsgBox("You LOSE")

        Else
            'Setting face in the RestartGameButton at top of board to be a cool face with sunglasses
            ReactionFaces.Faces.Top = -24
            'Loop to go through all x co-ordinates in the arrBoardPieces array
            For i = 1 To n
                'Loop to go through all y co-ordinates in the arrBoardPieces array. Together with the outer loop, this creates a loop that goes through every position in the arrBoardPieces array.
                For j = 1 To n
                    'Flags all remaining unflagged mines
                    If arrBoardPieces(i, j).Flagged = False And arrBoardPieces(i, j).TileType = 9 Then
                        arrBoardPieces(i, j).BackgroundImage = My.Resources.flag
                    End If

                Next
            Next

            'Freezes counter clock in top right of Game object
            Clock.Timer.Stop()
            'Displays a dialog box informing the user they have won and showing them the time they did it in
            MsgBox("You Win with time: " & Clock.Time)
            'Calls RecordTime sub which saves the user's score in a text file
            Master.RecordTime()
        End If

    End Sub

    Public Sub Adjacency()
        'Assigning default values for counter bound variables.
        'For the following variables , left to right and top to bottom are the directions of index increasing.
        'l stands for left and takes the value of the relative horizontal index to the current point in the for loop of the leftmost tile that will be checked in the current iteration of the loop
        'r stands for right and takes the value of the relative horizontal index to the current point in the for loop of the rightmost tile that will be checked in the current iteration of the loop
        't stands for top and takes the value of the relative vertical index to the current point in the for loop of the topmost tile that will be checked in the current iteration of the loop
        'b stands for botton and takes the value of the relative vertical index to the current point in the for loop of the bottommost tile that will be checked in the current iteration of the loop
        'MineCount is a counter that takes the value of the number of mines counted that are touching the current tile of the loop
        Dim MineCount As Integer
        Dim l As Integer = -1
        Dim r As Integer = 1
        Dim t As Integer = -1
        Dim b As Integer = 1
        'Loop to go through all x co-ordinates in the arrBoardPieces array
        For i = 1 To n
            'Selection statements to determine the values of horizontal counter bounds to ensure the algorithm doesn't attempt to access indexes outside of the bounds of the array. Both counters assigned a value for each case to ensure the values are correct for every tile, otherwise potentially incorrect values could carry over from the previous iteration of the loop
            If i = 1 Then
                l = 0
                r = 1
            ElseIf i = n Then
                l = -1
                r = 0
            Else
                l = -1
                r = 1
            End If
            'Loop to go through all y co-ordinates in the arrBoardPieces array. Together with the outer loop, this creates a loop that goes through every position in the arrBoardPieces array.
            For j = 1 To n
                'Minecounter reset to 0 for every tile
                MineCount = 0
                'Selection statements to determine the values of vertical counter bounds to ensure the algorithm doesn't attempt to access indexes outside of the bounds of the array. Both counters assigned a value for each case to ensure the values are correct for every tile, otherwise potentially incorrect values could carry over from the previous iteration of the loop
                If j = 1 Then
                    t = 0
                    b = 1
                ElseIf j = n Then
                    t = -1
                    b = 0
                Else
                    t = -1
                    b = 1
                End If
                'Loop that goes through every tile surrounding and including the current tile in the whole array loop and checks if it is a mine. If it is then the MineCount counter is incremented
                For k = l To r
                    For m = t To b

                        If arrBoardPieces(i + k, j + m).TileType = 9 Then
                            MineCount = MineCount + 1
                        End If

                    Next m

                Next k
                'Assigns TileType of current tile to be equal to the number of mines counted in the previous loop. Not done for tiles that have already been assigned mine status as this would break the game
                If arrBoardPieces(i, j).TileType <> 9 Then
                    arrBoardPieces(i, j).TileType = MineCount
                End If

            Next j
        Next i
    End Sub



    Private Function RandomIndex(Min As Integer, Max As Integer) As Integer
        'Random number Generator instantiated
        Static Generator As New System.Random
        'New random number generated between Min and Max inclusive.Min and Max +1 passed in as lower and upper bound as in this procedure the lower bound is inclusive but the upper bound is inexplicably non inclusive
        Return Generator.Next(Min, Max + 1)
    End Function

End Class

