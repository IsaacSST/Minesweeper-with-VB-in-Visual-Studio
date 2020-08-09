Public Class RankedList
    'Defining the array of labels that will display the highscores
    Public DisplayArray(20, 2) As Label
    Private Sub RankedList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dimensions and positioning of object as a whole defined
        Me.Width = 100
        Me.Height = 200
        Me.Left = 90
        Me.Top = 10
        'Loop that goes through all rows of the DisplayArray
        For i = 1 To 20
            'Loop the goes through both columns of the DisplayArray. Together with the outer loop this means every element of the DisplayArray is visited
            For j = 1 To 2
                'Each element instantiated as a Label and default properties and dimensions set
                DisplayArray(i, j) = New Label
                DisplayArray(i, j).BackColor = Color.Transparent
                DisplayArray(i, j).Height = 10
                DisplayArray(i, j).Font = New Font("Arial", 6, FontStyle.Bold)
                DisplayArray(i, j).ForeColor = Color.Black
                'Vertical position of the label depends on the row number. This means the DisplayArray is shown as a stack of its rows
                DisplayArray(i, j).Top = (i - 1) * 10
                DisplayArray(i, j).BorderStyle = BorderStyle.None
                'If the element in question is in the first column then it is to display a username so is placed at the leftmost edge of the RankedList object, and given a width of 60. Otherwise, it must be in the second column which means it is to display a score so it is placed 60 pixels away from the leftmost edge of the RankedList object, meaning it is placed flush to the right of the first column, and it is given a width of 40 as scores do not need as much space as usernames so the first column gets a greater share of the 100 width available
                If j = 1 Then
                    DisplayArray(i, j).Width = 60
                Else
                    DisplayArray(i, j).Width = 40
                    DisplayArray(i, j).Left = 60
                End If
                'Label element added to the Form and brought to the front to avoid bugs where it may accidently be hidden
                DisplayArray(i, j).BringToFront()
                Me.Controls.Add(DisplayArray(i, j))
            Next
        Next
    End Sub
End Class
