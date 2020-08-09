Public Class MinesweeperButton

    'Object property definitions
    Public Property Revealed As Boolean = False
    Public Property Tilei As Integer
    Public Property Tilej As Integer
    Public Property TileType As Integer
    Public Property Flagged As Boolean = False
    Private Sub MinesweeperButton_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Only the button needs to be loaded into the form is its dimensions and background image set. 16x16 pixels is the dimensions of the background image so this is set as the dimensions of the button so there is no excess space
        Me.BackgroundImage = My.Resources.button
        Me.Width = 16
        Me.Height = 16
    End Sub
    Private Sub RevealProcedure()
        'Tile's status formally set to revealed so program knows that interactions with it must change
        Me.Revealed = True
        'Another tile has been revealed so the NumberRevealed tracker is incremented
        Master.Board.NumberRevealed = Master.Board.NumberRevealed + 1
        'After this tile has been revealed, if the number of revealed tiles in the board is equal to the number of tiles that aren't mines, and because the user hasn't lost the game by revealing this tile, it stands to reason that the player has won the game, therefore the GameEnd procedure is called on this tile with the Win parameter set to true
        If Master.Board.NumberRevealed = (Master.Board.n) ^ 2 - Master.Board.Mines Then
            Master.Board.GameEnd(Tilei, Tilej, True)
        End If
    End Sub
    Public Sub Reveal()
        'If the BlockGame boolean is True then the user has lost the game and they must click the RestartGameButton to play again, therefore there should not be any changes to the game board until they do this. This block is necessary in this procedure as otherwise if the user loses the game by indirectly revealing a mine through the RecursiveReveal procedure, then the RecursiveReveal will continue revealing tiles after they have closed the game loss dialog box. With this selection statement and by setting BlockGame to True in the GameEnd procedure, the RecursiveReveal will blocked from revealing any tiles and from continuing to call itself
        If Master.Board.BlockGame = False Then
            'The TileType of the tile tells this procedure what to do when the tile is revealed
            Select Case Me.TileType
                Case 0
                    'The revealed image for a blank tile is the same as that of a pressed tile
                    Me.BackgroundImage = My.Resources.pressed
                    'Setting Revealed to true is required in this case before the RecursiveReveal is called as otherwise the recursion will be infinite as the algorithm will keep attempting to reveal tiles that have already been visited
                    Me.Revealed = True
                    'Calling the RecursiveReveal process with center at this tile
                    Master.Board.RecursiveReveal(Tilei, Tilej)
                Case 1
                    'If there is 1 mine adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 1 overlaid on it
                    Me.BackgroundImage = My.Resources.one
                Case 2
                    'If there are 2 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 2 overlaid on it
                    Me.BackgroundImage = My.Resources.two
                Case 3
                    'If there are 3 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 3 overlaid on it
                    Me.BackgroundImage = My.Resources.three
                Case 4
                    'If there are 4 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 4 overlaid on it
                    Me.BackgroundImage = My.Resources.four
                Case 5
                    'If there are 5 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 5 overlaid on it
                    Me.BackgroundImage = My.Resources.five
                Case 6
                    'If there are 6 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 6 overlaid on it
                    Me.BackgroundImage = My.Resources.six
                Case 7
                    'If there are 7 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 7 overlaid on it
                    Me.BackgroundImage = My.Resources.seven
                Case 8
                    'If there are 8 mines adjacent to this tile then when it is revealed it will show the same backround image as with the blank tile but with a 8 overlaid on it
                    Me.BackgroundImage = My.Resources.eight
                Case 9
                    'If the tiletype of this tile is 9 then the user has revealed a mine and therefore lost the game, so the background image of this tile is set to an exploded mine and the GameEnd procedure is called with the Win parameter set to False
                    Me.BackgroundImage = My.Resources.blown
                    Master.Board.GameEnd(Tilei, Tilej, False)
            End Select
            'Reveal procedure called which handles trivial reveal related things
            RevealProcedure()
        End If
    End Sub
    Private Sub MinesweeperButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        'If BlockGame is true then any interaction with the board tiles should yield no results, therefore things are only done when the user Mouses down on a tile if BlockGame is False
        If Master.Board.BlockGame = False Then
            'Changing the positioning of the image inside the RestartGameButton so that a nervous face is shown
            Master.Board.ReactionFaces.Faces.Top = -72
            'Selection statement that changes what is done depending on which mouse button is being pressed
            If e.Button = MouseButtons.Right Then
                'If a tile is revealed then right clicking on it should do nothing, therefore things are only done if this tile has not been revealed
                If Me.Revealed = False Then
                    'If the tile is flagged then it should be unflagged when the user right clicks on it. If the tile is unflagged then it should be flagged when the user right clicks on it
                    If Me.Flagged = True Then
                        'Flag physically removed from tile
                        Me.BackgroundImage = My.Resources.button
                        'Flag formally declared as having been removed
                        Me.Flagged = False
                        'Minecounter increased by 1 as this counter is a user aid for telling them how many mines there are left to flag, so if a tile is unflagged then there is one more flag left to place
                        Master.Board.MineCounter.Text = Master.Board.MineCounter.Text + 1
                    Else
                        'Flag physically placed in tile
                        Me.BackgroundImage = My.Resources.flag
                        'Flag formally declared as having been placed
                        Me.Flagged = True
                        'Minecounter decreased by 1, as now there is 1 less mine to flag
                        Master.Board.MineCounter.Text = Master.Board.MineCounter.Text - 1
                    End If
                End If
            ElseIf e.Button = MouseButtons.Left Then
                'Pressing the left mouse button should have no effect if the tile has already been revealed
                If Me.Revealed = False Then
                    'Flags are where the user thinks a mine is, so they are blocked from pressing flagged tiles
                    If Me.Flagged = False Then
                        'Tile is shown to be pressed. Things that happen when a user left clicks a tile happen when the user releases the button, not when they press it
                        Me.BackgroundImage = My.Resources.pressed
                    End If
                End If
            ElseIf e.Button = MouseButtons.Middle Then
                'If the user middle clicks a tile then a 3x3 square of tiles around the pressed tile is also pressed as well
                Master.Board.LargePress(Tilei, Tilej, False)
            End If
        End If
    End Sub
    Private Sub TheBigReveal()
        'If this tile is revealed and showing a number and if the number of flags on touching tiles around this tile satisfies the number on the tile, then the RecursiveReveal procedure is called with center at this tile when this procedure is called
        If Me.Revealed = True And Me.TileType <> 0 Then
            Dim FlagConditionsmet As Boolean = Master.Board.FlagCheck(Me.TileType, Me.Tilei, Me.Tilej)
            If FlagConditionsmet Then
                Master.Board.RecursiveReveal(Tilei, Tilej)
            End If
        End If
    End Sub
    Private Sub MinesweeperButton_DoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.DoubleClick
        'Upon double clicking on a tile, provided the game is not blocked, TheBigReveal procedure is called. This starts a RecursiveReveal process if the tile is a revealed numbered tile whose number is satisfied by adjacent flags
        If Master.Board.BlockGame = False Then
            TheBigReveal()
        End If
    End Sub

    Private Sub MinesweeperButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        'If BlockGame is true then any interaction with the board tiles should yield no results, therefore things are only done when the user Mouses up on a tile if BlockGame is False
        If Master.Board.BlockGame = False Then
            'Changing the positioning of the image inside the RestartGameButton so that the default face is shown
            Master.Board.ReactionFaces.Faces.Top = -96
            'Selection statement that changes what is done depending on which mouse button is being released
            If e.Button = MouseButtons.Left Then
                'Left clicking does nothing on flagged tiles as these tiles are where the user thinks there is a mine
                If Me.Flagged = False Then
                    'Before the user left clicks the first tile, all tiles in the game board are actually blank tiles. The status of tiles is generated after the first click so that the user can be given a fair chance at the start by leaving a 3x3 area around the clicked tile free of mines
                    If Master.Board.GameStarted = False Then
                        'GenerateMines procedure called with this tile's co-ordinates passed in as the center of the 3x3 area that should be left free of mines
                        Master.Board.GenerateMines(Tilei, Tilej)
                        'Adjacency procedure called which determines the TileType of all non-mine tiles
                        Master.Board.Adjacency()
                        'Timer clock in top right of game started at 1
                        Master.Board.Clock.Timer.Start()
                        Master.Board.Clock.Text = 1
                        'Game formally declared as being started
                        Master.Board.GameStarted = True

                    End If
                    'If tile is not already revealed then it is revealed
                    If Me.Revealed = False Then
                        Me.Reveal()
                    End If
                End If

            ElseIf e.Button = MouseButtons.Middle Then
                'Large press is released upon release of middle mouse button
                Master.Board.LargePress(Tilei, Tilej, True)
                'TheBigReveal procedure is called, which starts a RecursiveReveal process if the tile is a revealed numbered tile whose number is satisfied by adjacent flags.
                TheBigReveal()
            End If
        End If
    End Sub
End Class
