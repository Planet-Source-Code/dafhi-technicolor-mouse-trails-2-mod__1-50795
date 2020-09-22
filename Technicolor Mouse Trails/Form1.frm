VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Mouse Trails by Peter Wilson (http://dev.midar.com/)"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_FPS As Long = 60 'setting this too high can cause system to hang

Private Type mdrBallType
    ParentIndex As Integer   ' Link to parent. This is a really cool feature as you can create complex hierarchies!
    Size As Single
    Mass As Single
    Colour As OLE_COLOR
    
    DesiredPositionX As Single
    DesiredPositionY As Single
    
    CurrentPositionX As Single
    CurrentPositionY As Single
    
    OffSetX As Single   ' Just for fun!
    OffSetY As Single   ' Just for fun!
End Type

Private m_Ball() As mdrBallType

Dim Frame&
Dim Tick&
Dim NextTick&

Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Private Sub Form_Load()
    
    ' Set some basic properties
    AutoRedraw = True
    FillStyle = vbFSSolid
    
    ' Create the ball hierarchy.
    Call InitBalls

    Show
    
    Do While DoEvents
     Tick = timeGetTime
     If Tick >= NextTick Then
      NextTick = Tick + Int(1000 / MAX_FPS)
      Frame = (Frame + 1) Mod 360
      BackColor = HSV(Frame, 1, 1)
      Call DrawCrossHairs
      Call DrawBalls(-1)
      ' Note: Once the root ball is drawn,
      '       it will then draw it's children,
      '       and they in turn will draw their children, etc.
     End If
    Loop

End Sub



Private Sub InitBalls()

    Dim intN As Integer
    ReDim m_Ball(60)    ' How many balls do we want? Don't forget 0 to 50 is actually 51 balls!
    
    ' Create the "root" ball.
    With m_Ball(0)
        .ParentIndex = -1
        .Size = 60                                     ' <<< Change for fun!
        .Mass = 3                                      ' <<< Change for fun!
        .Colour = HSV(0, 1, 1)                         ' <<< Change for fun!
    End With
    
    For intN = 1 To 60
        With m_Ball(intN)
            .ParentIndex = intN - 1                     ' <<<  Make each ball link to the previous ball (or any other way you like!)
            .Size = Abs(30 - intN) * 2                  ' <<< Change for fun!
            .Mass = 4                                   ' <<< Change for fun!
            .Colour = HSV((intN / 50) * 360, 1, 1)      ' <<< Change for fun!
            .OffSetX = 0
            .OffSetY = 0
        End With
    Next intN
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intN As Integer
    Static intCounter As Integer
    
    intCounter = intCounter + 1
    If intCounter > 8 Then intCounter = 1
    
    For intN = 1 To 60
        With m_Ball(intN)
            If intCounter = 1 Then
                .OffSetX = ((Rnd * 300) - 150)
                .OffSetY = ((Rnd * 300) - 150)
                .Size = Abs(30 - intN) * 2
                m_Ball(0).Size = 60 ' << This shouldn't really be in the loop, but what the heck, it's simple.
            ElseIf intCounter = 2 Then
                .OffSetX = Cos((60 / intN) * 20) * 100
                .OffSetY = Sin((60 / intN) * 20) * 100
                .Size = 20
                m_Ball(0).Size = 20
            ElseIf intCounter = 3 Then
                .OffSetX = Cos(intN) * 100
                .OffSetY = Sin(intN) * 100
                .Size = 20
                m_Ball(0).Size = 20
            ElseIf intCounter = 4 Then
                .OffSetX = -10
                .OffSetY = 0
            ElseIf intCounter = 5 Then
                .OffSetX = 10
                .OffSetY = 0
            ElseIf intCounter = 6 Then
                .OffSetX = 0
                .OffSetY = -10
            ElseIf intCounter = 7 Then
                .OffSetX = 0
                .OffSetY = 10
            ElseIf intCounter = 8 Then
                .OffSetX = 0
                .OffSetY = 0
            End If
        End With
    Next intN
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Set where you would like the "root ball" to be.
    m_Ball(0).DesiredPositionX = x
    m_Ball(0).DesiredPositionY = y
        
End Sub

Private Sub Form_Resize()

    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier.
    Dim sngAspectRatio As Single
    
    If WindowState <> vbMinimized And ScaleHeight > 0 Then
     sngAspectRatio = Width / Height
     ScaleLeft = -1000
     ScaleWidth = 2000
     ScaleHeight = 2000 / sngAspectRatio
     ScaleTop = -ScaleHeight / 2
    End If
    
End Sub


Private Sub DrawCrossHairs()

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Me.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    Me.ForeColor = RGB(0, 64, 64)
    Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
    
    ' Draw Vertical line
    Me.ForeColor = RGB(0, 96, 96)
    Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
    
End Sub

Private Sub DrawBalls(ParentIndex As Long)

    ' ====================================================================
    ' This is a recursive procedure, this means it calls itself!
    ' If you are a slacker, and put in the wrong parent id's you might get
    ' stuck in an infinite loop and your comptuer will run out of memory.
    ' ====================================================================
    'On Local Error Resume Next ' Ignore errors (which can occur if you use really small masses)
    
    Dim lngIndex As Long
    Dim lngNewParent As Long
    
    Dim sngDeltaX As Single
    Dim sngDeltaY As Single
    
    Dim sngBallX As Single
    Dim sngBallY As Single
    
    ' Loop through the balls from the Lower Boundry to the Upper Boundry of the array.
    For lngIndex = LBound(m_Ball) To UBound(m_Ball)
        If m_Ball(lngIndex).ParentIndex = ParentIndex Then
            
            With m_Ball(lngIndex)
                
                Me.ForeColor = .Colour
                Me.FillColor = .Colour
                
                If ParentIndex = -1 Then ' "root ball"
                    
                    ' Calculate the difference between where the ball currently is, and where we would like it to be.
                    sngDeltaX = (.CurrentPositionX - .DesiredPositionX)
                    sngDeltaY = (.CurrentPositionY - .DesiredPositionY)
                    
                    ' Then move the ball closer to where it should be, depending on it's mass.
                    .CurrentPositionX = .CurrentPositionX - (sngDeltaX / .Mass)
                    .CurrentPositionY = .CurrentPositionY - (sngDeltaY / .Mass)
                
                Else
                
                    ' Calculate the difference between where the ball currently is, and where we would like it to be.
                    ' Note: Each child ball, seeks it's parents current location.
                    sngDeltaX = (.CurrentPositionX - m_Ball(ParentIndex).CurrentPositionX + .OffSetX)
                    sngDeltaY = (.CurrentPositionY - m_Ball(ParentIndex).CurrentPositionY + .OffSetY)   ' <<< Change this -20 for fun!
                    
                    ' Then move the ball closer to where it should be, depending on it's mass.
                    .CurrentPositionX = .CurrentPositionX - (sngDeltaX / .Mass)
                    .CurrentPositionY = .CurrentPositionY - (sngDeltaY / .Mass)
                    
                    ' Draw a line to the parent (optional)
                    Me.Line (.CurrentPositionX, .CurrentPositionY)-(m_Ball(ParentIndex).CurrentPositionX, m_Ball(ParentIndex).CurrentPositionY)
                    
                End If
                
                ' Draw a pretty circle on Me (ie. the form)
                Me.Circle (.CurrentPositionX, .CurrentPositionY), .Size
                
                ' Now go a draw my children's balls! (Gee - I never thought I would ever type that sentance! ha ha)
                Call DrawBalls(lngIndex)
                ' Do not place code after this point (since this is a recursive routine) and any code placed
                ' here could potentially get called many times (which you probably don't want!)
                
            End With
        End If
    Next lngIndex
    
End Sub


