VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "CheckBox"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdCombo 
      Caption         =   "Combo-Box"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Text Box"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Line Line4 
      Index           =   6
      X1              =   328
      X2              =   72
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Line Line4 
      Index           =   5
      X1              =   328
      X2              =   72
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line4 
      Index           =   4
      X1              =   328
      X2              =   72
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   328
      X2              =   72
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   328
      X2              =   72
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   328
      X2              =   72
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   328
      X2              =   328
      Y1              =   304
      Y2              =   136
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   264
      X2              =   264
      Y1              =   136
      Y2              =   304
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   136
      X2              =   136
      Y1              =   136
      Y2              =   304
   End
   Begin VB.Line Line4 
      Index           =   7
      X1              =   72
      X2              =   328
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   72
      X2              =   328
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   200
      X2              =   200
      Y1              =   136
      Y2              =   304
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   72
      X2              =   72
      Y1              =   136
      Y2              =   304
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This little program was a experiment for a huge project that I am currently writing
'This demonstrates how to use the drag and drop routines easily.
'It also demonstrates how to add controls to a project dynamically in run time
'It shows you how to move controls around the screen in run time
'
'I hope this helps somebody,
'Any comments, help required etc. email me:  c.woor@beaumontsoftware.com
'
'
Public WithEvents txtBox As TextBox
Attribute txtBox.VB_VarHelpID = -1
Dim TextNum As Long, CombNum As Long, ChkNum As Long
Dim Lines(100)              'Array to store all line data ie. x1,y1, index num etc.
Dim NumControls As Long     'Var to keep number of controls
Dim XFactor As Long, YFactor As Long    'Stores the twips / pixel calculation
Dim mousepos As POINTAPI

Private Sub Check1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    
    'Has something been draged over a checkbox
    'If so then change the mouse pointer, when it leaves put the mouse pointer back
    
Select Case State
    Case vbEnter
        Source.MousePointer = 12
    Case vbLeave
        Source.MousePointer = 0
End Select
End Sub

Private Sub cmdCheck_Click()

'We want to create a new copy of the check box that is hidden

Load Check1(ChkNum)     'Create new copy of Check1 with an index of ChkNum
SetCursorPos (cmdCheck.Left / 2), (cmdCheck.Top / 2)    'Move the mouse to the centre
                                                        'Of the checkbox
Check1(ChkNum).Left = cmdCheck.Left     'Place the new check box
Check1(ChkNum).Top = cmdCheck.Top       'on top of cmdCheck button

Check1(ChkNum).Visible = True           'Display it
Check1(ChkNum).Drag                     'start the drag process

ChkNum = ChkNum + 1                     'Increase the array index

End Sub

Private Sub cmdCombo_Click()

    'We want to create a new copy of the combo box that is hidden

Load Combo1(CombNum)    'Create new copy of Combobox with an index of Combnum
SetCursorPos (cmdCombo.Left / 2), (cmdCombo.Top / 2)    'Move the mouse to the centre
                                                        'Of the combobox
Combo1(CombNum).Left = cmdCombo.Left    'Place the new Combo box
Combo1(CombNum).Top = cmdCombo.Top      'on top of cmdCombo button

Combo1(CombNum).Visible = True          'Display it
Combo1(CombNum).Drag                    'start the drag process

CombNum = CombNum + 1                   'Increase the array index


End Sub

Private Sub cmdText_Click()

    'We want to create a new copy of the Text box that is hidden

Load Text1(TextNum)     'Create new copy of Text box with an index of Textnum
SetCursorPos (cmdText.Left / 2), (cmdText.Top / 2)  'Move the mouse to the centre
                                                    'Of the Text Box
Text1(TextNum).Left = cmdText.Left                  'Place the new Text box
Text1(TextNum).Top = cmdText.Top                    'on top of cmdText button

Text1(TextNum).Visible = True                       'Display it
Text1(TextNum).Drag                                 'start the drag process

TextNum = TextNum + 1                                'Increase the array index

End Sub

Private Sub Combo1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    
    'Has something been draged over a combobox
    'If so then change the mouse pointer, when it leaves put the mouse pointer back

Select Case State
    Case vbEnter
        Source.MousePointer = 12
    Case vbLeave
        Source.MousePointer = 0
End Select
End Sub

'This event gets fired if an object has been droped on the form

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
Dim Result As Boolean       'Does the control get unloaded in the MoveClosestX Function

'If the control has been droped and a line cannot be found with
'the MoveClosetX function, then Result is set to false with basically means that
'the drag-drop process must be cancelled

Result = False      'Init Result

Source.Move (x - Source.Width / 2), (y - Source.Height / 2)
    
    
GetCursorPos mousepos   'Get the mouses position

'Call MoveClosestX function to find the closet line along the X axis
'I am calling it with:  (mouses x -(forms x / number of pixels/twip)
'this gives us the exact position of the mouse.
'We need to do this as the API GetCursorPos returns the mouse co-ord's in pixels
'from the edge of the screen, we need the position of it within our form.
'Were also passing across the control that has been droped i.e. Source
Result = MoveClosestX(mousepos.x - (frmMain.Left / XFactor), Source)

'If the control hasn't been droped somewhere else the call MoveClosetY.
'This does exactly the same as above except it does it with the vertical axis
If Result = False Then MoveClosestY mousepos.y - (frmMain.Top / YFactor), Source

End Sub

Private Sub Form_Load()

'Reset all array indexes
TextNum = 1
CombNum = 1
ChkNum = 1

NumControls = frmMain.Controls.Count    'Count the controls on the form
For i = 0 To NumControls - 1        'For every control
    If TypeOf frmMain.Controls(i) Is Line Then  'Check if it's a line
        'If it is a line then get the lines properties and put
        'them into the array at position (i)
        'We only realy need it's name, index, x & y.
        Lines(i) = frmMain.Controls(i).Name & "," & frmMain.Controls(i).Index & "," & frmMain.Controls(i).X1 & "," & frmMain.Controls(i).Y1 & "," & frmMain.Controls(i).Y2
    End If
Next i

'Work out how many twips/pixel their are for the current
'resolution

XFactor = Screen.TwipsPerPixelX
YFactor = Screen.TwipsPerPixelY

End Sub


Private Sub Text1_Click(Index As Integer)
'If the control gets clicked after it's been positioned it can
'be redragged.

SetCursorPos (Text1(Index).Left / 2), (Text1(Index).Top / 2) 'Put the mouse in the middle
                                                            'of the control
Text1(Index).Drag   'Start dragging

End Sub

Private Sub Text1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)

    'Has something been draged over a textbox
    'If so then change the mouse pointer, when it leaves put the mouse pointer back


Select Case State
    Case vbEnter
        Source.MousePointer = 12
    Case vbLeave
        Source.MousePointer = 0
End Select


End Sub
'This is the funtion that looks along the horizontal axis for a
'line, it first looks left and stores how far it looked in pixels.
'It then looks right and stores how far it moved.
'It then compares the two figures and if either are below a certain limit
'and is lower than the other then the lowest is where we move the control
'too !    ---Bit confusing I know, walk through it


Function MoveClosestX(xPos As Long, Cntrl As Control) As Boolean
Dim XTemp As Long, CountLeft As Long, CountRight As Long    'Our counting var's
Dim Left As Long, Right As Long
Dim Strg As String  'String containing line data
Dim Nme As String, Idx As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long  'Line's info
Dim Fin As Long     'End position in string
Dim Moved As Boolean

For i = 0 To NumControls - 1
    Strg = Lines(i)
    If Strg > "" Then
        'This bit just go's through the string and stores each bit
        'in a variable
        If Mid$(Strg, 1, InStr(1, Strg, ",") - 1) = "Line1" Then
            Fin = InStr(1, Strg, ",")
            Nme = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            Idx = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            X1 = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            Y1 = Mid$(Strg, 1, Fin - 1)
            
                'First move left and count how many pixels we move
                'until we find a line.
            CountLeft = 0
            XTemp = Cntrl.Left  'Set XTemp to controls x
            For z = 1 To 30     'Walk left a maximum of 30 pixels
                If XTemp = X1 Then  'If we found a line
                    Left = XTemp    'Set the left var to the lines x1 or xtemp, x1==xtemp
                    Exit For        'Quit this little for loop
                End If
                XTemp = XTemp - 1   'Didn't find a line, move 1 pixel left
                CountLeft = CountLeft + 1   'Increase the pixel count
                DoEvents
            Next z
                        
                'Now we'll go right and see if thats closer
                
            CountRight = 0                      '
            XTemp = Cntrl.Left                  '
            For z = 1 To 30                     '
                If XTemp = X1 Then              '
                    Right = XTemp               'Same as above only this
                    Exit For                    'time we're moving right
                End If                          '
                CountRight = CountRight + 1     '
                XTemp = XTemp + 1               '
                DoEvents                        '
            Next z                              '
            
            'Now see which was closer and move the control accordinly
            'Make sure it's not too far
            If CountRight < 30 Or CountLeft < 30 Then   'We don't want to move more than 30 pixels
                If CountLeft > CountRight Then
                    Cntrl.Left = Right  'Move the control to the closet line to the right
                    Moved = True
                    Exit For
                Else
                    Cntrl.Left = Left   'Move the control to the closest line to the left
                    Moved = True
                    Exit For
                End If
            End If
        End If
    End If
Next i
'If we did'nt move the control then we must lose the control being dragged
'To do this we need to first unload it then we have to decrease the
'array index count
If Moved = False Then
    Unload Cntrl
    If TypeOf Cntrl Is TextBox Then
        TextNum = TextNum - 1
    End If
    If TypeOf Cntrl Is ComboBox Then
        CombNum = CombNum - 1
    End If
    If TypeOf Cntrl Is CheckBox Then
        ChkNum = ChkNum - 1
    End If
    MoveClosestX = True
End If


End Function

'Exactly the same as MoveClosestY except that it checks along the vertical axis.

Function MoveClosestY(yPos As Long, Cntrl As Control)
Dim YTemp As Long, CountUp As Long, CountDown As Long
Dim Up As Long, Down As Long
Dim Strg As String
Dim Nme As String, Idx As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim Srt As Long, Fin As Long
Dim Moved As Boolean

For i = 0 To NumControls - 1
    Strg = Lines(i)
    If Strg > "" Then
        If Mid$(Strg, 1, InStr(1, Strg, ",") - 1) = "Line4" Then
            Fin = InStr(1, Strg, ",")
            Nme = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            Idx = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            X1 = Mid$(Strg, 1, Fin - 1)
            Strg = Mid$(Strg, Fin + 1)
            Fin = InStr(1, Strg, ",")
            Y1 = Mid$(Strg, 1, Fin - 1)
            
                'First move up and count how many pixels we move
                'until we find a line.
            CountUp = 0
            YTemp = Cntrl.Top
            For z = 1 To 20
                If YTemp = Y1 Then
                    Up = YTemp
                    Exit For
                End If
                YTemp = YTemp - 1
                CountUp = CountUp + 1
                DoEvents
            Next z
                        
                'Now we'll go down and see if thats closer
                
            CountDown = 0
            YTemp = Cntrl.Top
            For z = 1 To 20
                If YTemp = Y1 Then
                    Down = YTemp
                    Exit For
                End If
                YTemp = YTemp + 1
                CountDown = CountDown + 1
                DoEvents
            Next z
            
            'Now see which was closer and move the control accordinly
            'Make sure it's not too far
            If CountDown < 15 Or CountUp < 15 Then
                If CountUp > CountDown Then
                    Cntrl.Top = Down
                    Moved = True
                    Exit For
                Else
                    Cntrl.Top = Up
                    Moved = True
                    Exit For
                End If
            End If
        End If
    End If
Next i

If Moved = False Then
    Unload Cntrl
    If TypeOf Cntrl Is TextBox Then
        TextNum = TextNum - 1
    End If
    If TypeOf Cntrl Is ComboBox Then
        CombNum = CombNum - 1
    End If
    If TypeOf Cntrl Is CheckBox Then
        ChkNum = ChkNum - 1
    End If
End If

End Function

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'If the right button's pressed then the text box under the mouse say's bye bye.


If Button = 2 Then
    Unload Text1(Index)
End If

End Sub
