VERSION 5.00
Begin VB.Form frmOCR 
   AutoRedraw      =   -1  'True
   Caption         =   "OCR Motion Reader"
   ClientHeight    =   6285
   ClientLeft      =   1470
   ClientTop       =   2010
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   Begin VB.CommandButton cmdWorks 
      Caption         =   "About"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Timer tmrStatusPause 
      Left            =   3000
      Top             =   5880
   End
   Begin VB.ComboBox cmbLetList 
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Text            =   "Select Letter"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   1320
      MousePointer    =   4  'Icon
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   519
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Label Label2 
         Caption         =   "The speed at which you write should be the same as when you edited"
         Height          =   735
         Left            =   5880
         TabIndex        =   7
         Top             =   4920
         Width           =   1815
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Drawing"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   5880
      Width           =   615
   End
End
Attribute VB_Name = "frmOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LetterType
    Direc(100) As Integer
End Type

Dim Alphabet(25) As LetterType

Dim WriteLet As Boolean
Dim HoldX As Integer, HoldY As Integer
Dim LetterMovement(200) As Integer
Dim Letter(-1 To 25) As String
Dim NumLet As Integer
Dim WriteFile As Boolean

Private Function Direction(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
'x1 and y1 are the center points

Dim Slope As Single

If x2 - x1 = 0 Then
    Slope = 50
Else
    Slope = -(y2 - y1) / (x2 - x1)
End If

If Slope <= 0 And Slope > -0.5 Then
    Direction = 0
ElseIf Slope <= -0.5 And Slope > -1 Then
    Direction = 1
ElseIf Slope <= -1 And Slope > -2 Then
    Direction = 2
ElseIf Slope < -2 Then
    Direction = 3
ElseIf Slope > 2 Then
    Direction = 4
ElseIf Slope <= 2 And Slope > 1 Then
    Direction = 5
ElseIf Slope <= 1 And Slope > 0.5 Then
    Direction = 6
ElseIf Slope <= 0.5 And Slope > 0 Then
    Direction = 7
End If

If y2 > y1 Then
    Direction = Direction + 8
End If

End Function

Private Sub cmbLetList_Change()
lblStatus.Caption = "Editing: " & Letter(cmbLetList.ListIndex)
End Sub

Private Sub cmdClear_Click()
picDraw.Cls
End Sub

Private Sub cmdEdit_Click()
If cmbLetList.ListIndex <> -1 Then
    WriteFile = True
    picDraw.Cls
    lblStatus.Caption = "Editing: " & Letter(cmbLetList.ListIndex)
Else
    MsgBox "Please select a lettet to edit.", vbCritical, "OCR Motion Reader"
End If
End Sub

Private Sub cmdWorks_Click()
frmWorks.Show 1
End Sub

Private Sub Form_Activate()
MsgBox "Note if this is your first time using the program you should probably " & vbNewLine & "edit every letter because right now they are custom to my handwriting." & vbNewLine & "So you should edit it to yours to make it more accurate.  Also a letter " & vbNewLine & "is finished once you release the mouse so be careful on letter like 'i'."
End Sub

Private Sub Form_Load()
Dim strFileLine As String
Dim Count As Integer
Dim i As Integer
Dim Start As Integer

Letter(-1) = ""
Letter(0) = "A"
Letter(1) = "B"
Letter(2) = "C"
Letter(3) = "D"
Letter(4) = "E"
Letter(5) = "F"
Letter(6) = "G"
Letter(7) = "H"
Letter(8) = "I"
Letter(9) = "J"
Letter(10) = "K"
Letter(11) = "L"
Letter(12) = "M"
Letter(13) = "N"
Letter(14) = "O"
Letter(15) = "P"
Letter(16) = "Q"
Letter(17) = "R"
Letter(18) = "S"
Letter(19) = "T"
Letter(20) = "U"
Letter(21) = "V"
Letter(22) = "W"
Letter(23) = "X"
Letter(24) = "Y"
Letter(25) = "Z"

For i = 0 To 25
    cmbLetList.AddItem Letter(i), i
Next i

Start = 2
Open App.Path & "\LetterMov.txt" For Input As 1
While Not EOF(1)
    Line Input #1, strFileLine
    For i = Start To 200 Step 2
        Alphabet(Count).Direc(Int(i / 2)) = Val(Mid(strFileLine, i, 2))
    Next i
    Start = 1
    Count = Count + 1
Wend
Close 1

End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteLet = True
HoldX = X
HoldY = Y
picDraw.CurrentX = X
picDraw.CurrentY = Y
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Direc As Integer
Dim BuffX As Integer, BuffY As Integer
Static Count As Integer


If WriteLet = True Then
    Count = Count + 1
    If Count Mod 2 = 0 Then
        If NumLet < 200 Then
            BuffX = X
            BuffY = Y
            Direc = Direction(HoldX, HoldY, BuffX, BuffY)
            HoldX = X
            HoldY = Y
            
            picDraw.Line -(BuffX, BuffY)
            
    
            LetterMovement(NumLet) = Direc
            
            NumLet = NumLet + 1
        Else
            lblStatus.Caption = "Letter Limit"
        End If
    End If
End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BuffArray(100) As String
Dim strFile As String
Dim Distort As Single
Dim Difference As Integer
Dim Score(25) As Single
Dim i As Integer
Dim F As Integer
Dim Highest As Integer
Dim HighScore As Single

frmOCR.Cls

'stretch or compress array to fit 100 into the BuffArray
Distort = NumLet / 100
For i = 0 To 100
    BuffArray(i) = LetterMovement(Int(i * Distort))
    If WriteFile = True Then
        Alphabet(cmbLetList.ListIndex).Direc(i) = LetterMovement(Int(i * Distort))
    End If
Next i

If WriteFile = False Then
    'calculate the score for each letter
    For F = 0 To 25
        Dim Total As Integer
        For i = 0 To 100

            
            If BuffArray(i) > Alphabet(F).Direc(i) Then
                Difference = BuffArray(i) - Alphabet(F).Direc(i)
            Else
                Difference = Alphabet(F).Direc(i) - BuffArray(i)
            End If
            
            'exceptions because the where the circle ends
            If BuffArray(i) = 0 And Alphabet(F).Direc(i) = 15 Then
                Difference = 1
            ElseIf BuffArray(i) = 0 And Alphabet(F).Direc(i) = 14 Then
                Difference = 2
            ElseIf BuffArray(i) = 1 And Alphabet(F).Direc(i) = 15 Then
                Difference = 2
            ElseIf BuffArray(i) = 1 And Alphabet(F).Direc(i) = 14 Then
                Difference = 3
            End If
            
            Score(F) = Score(F) + (8 - Difference)
            Total = Total + 8
        Next i
        'put score into percent
        Score(F) = Score(F) / Total * 100
        Total = 0
        
        frmOCR.Print Letter(F) & ":  " & CInt(Score(F)) & "%"
    Next F
    
    Highest = 0
    HighScore = Score(0)
    
    For i = 1 To 25
        If Score(i) > HighScore Then
            Highest = i
            HighScore = Mid(Score(i), 1, 2)
        End If
    Next i
    frmOCR.Print ""
    If HighScore < 50 Then
        frmOCR.Print "?" & Letter(Highest) & "?  Percent: "; CInt(HighScore) & "%"
    Else
        frmOCR.Print Letter(Highest) & "  Percent: "; CInt(HighScore) & "%"
    End If
    
    
    lblStatus.Caption = "Drawing"


Else
    For F = 0 To 25
        For i = 0 To 100
            If Val(Alphabet(F).Direc(i)) < 10 Then
                strFile = strFile & "0"
            End If
            strFile = strFile & Alphabet(F).Direc(i)
        Next i
        If F <> 25 Then
            strFile = strFile & vbNewLine
        End If
    Next F
    
    F = FreeFile
    Open App.Path & "\LetterMov.txt" For Output As #F
    Write #F, strFile
    Close #F
    lblStatus.Caption = "Letter Saved"
    tmrStatusPause.Interval = 1000
End If

WriteFile = False
WriteLet = False
NumLet = 0
End Sub

Private Sub tmrStatusPause_Timer()
lblStatus.Caption = "Drawing"
tmrStatusPause.Interval = 0
End Sub
