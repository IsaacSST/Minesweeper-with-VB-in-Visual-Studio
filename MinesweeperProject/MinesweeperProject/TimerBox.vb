Public Class TimerBox
    Inherits Label
    'Time is to start at 1 so the user cannot have a time of 0
    Public Time As Integer = 1
    'Timer declared withevents so this class can handle the event of the timer ticking
    Public WithEvents Timer As New System.Windows.Forms.Timer
    Sub New()
        'Dimensions and properties set upon instantiation of the timerbox
        Me.Width = 40
        Me.Height = 20
        Me.ForeColor = Color.Red
        Me.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
        Me.Font = New Font("Arial", 10, FontStyle.Bold)
        Me.Text = 0
        'Timer ticking interval set to 1 second
        Me.Timer.Interval = 1000
        Me.BackColor = Color.Black
    End Sub
    Private Sub Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        'When the timer ticks the time is incremented and so is the number displayed as the text in the TimerBox
        Me.Time = Time + 1
        Me.Text = Time
    End Sub
End Class
