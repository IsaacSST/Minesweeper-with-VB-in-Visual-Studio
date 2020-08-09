Public Class RestartGameButton
    'Faces label declared WithEvents as mouse events on the label will need to be handled. This is because when a user clicks on an object, component objects take priority as the event handling object
    Public WithEvents Faces As New Label

    Private Sub RestartGameButton_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Faces label added to object with dimensions and background image set
        '24 pixels is the height and width of a single face image
        Me.Faces.Width = 24
        '120 pixels is the height of the entire faces image that contains all the different faces needed
        Me.Faces.Height = 120
        'The default face is at the bottom of the faces image so the Top of the faces label is set to be 96 pixels above the top of the RestartGameButton object. This has the effect of showing the default face in the button whilst all the other potential faces are hidden
        Me.Faces.Top = -96
        'Width and height of the RestartGameButton set to be 24 by 24 as these are the dimensions of a single face
        Me.Height = 24
        Me.Width = 24

        Me.Faces.BackgroundImage = My.Resources.Minesweeper_faces
        Me.Controls.Add(Faces)
    End Sub
    Public Sub RestartGameButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Faces.MouseDown
        'When the user mouses down with their left mouse button the image shown is changed to be a pressed version of the default face. This face is positioned at the very top of the faces image so the Top of the faces label is set to be in line with the top of the RestartGameButton object
        If e.Button = MouseButtons.Left Then
            Me.Faces.Top = 0
        End If
    End Sub
    Public Sub RestartGameButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Faces.MouseUp
        'When the user releases the left mouse button, the face is reset to the default face and the ResetGame procedure is called, which resets the game board ready for a new game. Hence the name of this object
        If e.Button = MouseButtons.Left Then
            Me.Faces.Top = -96
            Master.Board.ResetGame()
        End If

    End Sub

End Class
