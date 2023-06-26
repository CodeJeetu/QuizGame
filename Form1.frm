VERSION 5.00
Begin VB.Form Quiz 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Quiz"
   ClientHeight    =   8565
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      Caption         =   "Next Question"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox txtAnswer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0%"
      Height          =   765
      Left            =   5040
      TabIndex        =   9
      Top             =   6960
      Width           =   1485
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   5880
      Width           =   5895
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   5040
      Width           =   5895
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   5895
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   5895
   End
   Begin VB.Label lblComment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblHeadAnswer 
      BackStyle       =   0  'Transparent
      Caption         =   "Captial"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblGiven 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label lblHeadGiven 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "Bar"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptionsCapitals 
         Caption         =   "Name &Capitals"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsState 
         Caption         =   " Name &State"
      End
      Begin VB.Menu mnuOptionsBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsMC 
         Caption         =   "&Multiple Choice Answers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsType 
         Caption         =   "&Type In Answers"
      End
   End
End
Attribute VB_Name = "Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CorrectAnswer As Integer
Dim NumAns As Integer
Dim NumCorrect As Integer
Dim Wsound(26) As Integer
Dim State(29) As String
Dim Capital(29) As String
Private Function SoundEx(W As String, Wsound() As Integer) As String
Dim Wtemp As String, S As String
Dim L As Integer
Dim I As Integer
Dim Wprev As Integer, Wsnd As Integer, Cindex As Integer
Wtemp = UCase(W)
L = Len(W)
If L <> 0 Then
 S = Left(Wtemp, 1)
 Wprev = 0
 If L > 1 Then
 For I = 2 To L
 Cindex = Asc(Mid(Wtemp, I, 1)) - 64
 If Cindex >= 1 And Cindex <= 26 Then
 Wsnd = Wsound(Cindex) + 48
 If Wsnd <> 48 And Wsnd <> Wprev Then S = S + Chr(Wsnd)
Wprev = Wsnd
 End If
  Next I
  End If
Else
 S = ""
End If
SoundEx = S
End Function
Private Sub Update_Score(Iscorrect As Integer)
Dim I As Integer
                'Check if answer is correct
cmdNext.Enabled = True
cmdNext.SetFocus
If Iscorrect = 1 Then
 NumCorrect = NumCorrect + 1
 lblComment.Caption = "Correct!"
  Else
 lblComment.Caption = "Sorry ..."
End If
            'Display correct answer and update score
If mnuOptionsMC.Checked = True Then
For I = 0 To 3
 If mnuOptionsCapitals.Checked = True Then
 If lblAnswer(I).Caption <> Capital(CorrectAnswer) Then
 lblAnswer(I).Caption = ""
 End If
 Else
 If lblAnswer(I).Caption <> State(CorrectAnswer) Then
 lblAnswer(I).Caption = ""
 End If
 End If
 Next I
Else
 If mnuOptionsCapitals.Checked = True Then
 txtAnswer.Text = Capital(CorrectAnswer)
 Else
 txtAnswer.Text = State(CorrectAnswer)
 End If
End If
lblScore.Caption = Format(NumCorrect / NumAns, "##0%")
End Sub

Private Sub About_Click()
About1.Show
Quiz.Hide
End Sub

Private Sub cmdExit_Click()
            'Exit program
Call mnuFileExit_Click
End Sub

Private Sub cmdNext_Click()
            'Generate the next question
cmdNext.Enabled = False
Call Next_Question(CorrectAnswer)

End Sub

Private Sub Form_Activate()
Call mnufilenew_click
End Sub

Private Sub Form_Load()

Randomize Timer
                'Load soundex function array
Wsound(1) = 0: Wsound(2) = 1: Wsound(3) = 2: Wsound(4) = 3
Wsound(5) = 0: Wsound(6) = 1: Wsound(7) = 2: Wsound(8) = 0
Wsound(9) = 0: Wsound(10) = 2: Wsound(11) = 2: Wsound(12) = 4
Wsound(13) = 5: Wsound(14) = 5: Wsound(15) = 0: Wsound(16) = 1
Wsound(17) = 2: Wsound(18) = 6: Wsound(19) = 2: Wsound(20) = 3
Wsound(21) = 0: Wsound(22) = 1: Wsound(23) = 0: Wsound(24) = 2
Wsound(25) = 0: Wsound(26) = 2
                'Load state/capital arrays
State(1) = "Utter Pradesh": Capital(1) = "Lucknow"
State(2) = "Bihar": Capital(2) = "Patana"
State(3) = "ArunChanalPradesh": Capital(3) = "ItaNagar"
State(4) = "AndraPradesh": Capital(4) = "HaidraBad"
State(5) = "Asam": Capital(5) = "DisPur"
State(6) = "Jharkhand": Capital(6) = "Ranchi"
State(7) = "Gujrat": Capital(7) = "Gandhi Nagar"
State(8) = "Hariyana": Capital(8) = "Chandigadh"
State(9) = "Haimanchal Pradesh": Capital(9) = "Tallahassee"
State(10) = "Karnatka": Capital(10) = "Bainglore"
State(11) = "Keral": Capital(11) = "Triwendram"
State(12) = "MadhyaPradesh": Capital(12) = "Bhopal"
State(13) = "Maharashta": Capital(13) = "Mumbai"
State(14) = "Meghalya": Capital(14) = "Shilong"
State(15) = "Nagaland": Capital(15) = "Kohima"
State(16) = "Udisa": Capital(16) = "Bhuneswar"
State(17) = "Punjab": Capital(17) = "Chandigadh"
State(18) = "Rajsthan": Capital(18) = "Jaipur"
State(19) = "Taminlandu": Capital(19) = "Chennai"
State(20) = "PashimBengal": Capital(20) = "Kolkata"
State(21) = "Manipur": Capital(21) = "Imfal"
State(22) = "Sikkim": Capital(22) = "GangTauk"
State(23) = "Tripura": Capital(23) = "AgarLatta"
State(24) = "Mijoram": Capital(24) = "Ejal"
State(25) = "Goa": Capital(25) = "Pangi"
State(26) = "JambhuKasmir": Capital(26) = "Shrinagar"
State(27) = "Utaranchal": Capital(27) = "Deharadun"
State(28) = "ChhAtisgarh": Capital(28) = "Raipur"
State(29) = "Ladakh": Capital(29) = "Leh"
End Sub
Private Sub lblAnswer_Click(Index As Integer)
'Check multiple choice answers
Dim Iscorrect As Integer
'If already answered, exit
If cmdNext.Enabled = True Then Exit Sub
Iscorrect = 0
If mnuOptionsCapitals.Checked = True Then
 If lblAnswer(Index).Caption = Capital(CorrectAnswer) Then Iscorrect = 1
Else
 If lblAnswer(Index).Caption = State(CorrectAnswer) Then
Iscorrect = 1
End If
End If
Call Update_Score(Iscorrect)
End Sub


Private Sub mnuFileExit_Click()
'End the application
End
End Sub


Private Sub mnufilenew_click()
'Reset the score and start again
NumAns = 0
NumCorrect = 0
lblScore.Caption = "0%"
lblComment.Caption = ""
cmdNext.Enabled = False
Call Next_Question(CorrectAnswer)
End Sub
Private Sub mnuOptionsCapitals_Click()
'Set up for providing capital, given state
mnuOptionsState.Checked = False
mnuOptionsCapitals.Checked = True
lblHeadGiven.Caption = "State:"
lblHeadAnswer.Caption = "Capital:"
Call mnufilenew_click
End Sub
Private Sub mnuOptionsMC_Click()
'Set up for multiple choice answers
Dim I As Integer
mnuOptionsMC.Checked = True
mnuOptionsType.Checked = False
For I = 0 To 3
 lblAnswer(I).Visible = True
Next I
txtAnswer.Visible = False
Call mnufilenew_click
End Sub
Private Sub mnuOptionsState_Click()
'Set up for providing state, given capital
mnuOptionsState.Checked = True
mnuOptionsCapitals.Checked = False
lblHeadGiven.Caption = "Capital:"
lblHeadAnswer.Caption = "State:"
Call mnufilenew_click
End Sub
Private Sub mnuOptionsType_Click()
'Set up for type in answers
Dim I As Integer
mnuOptionsMC.Checked = False
mnuOptionsType.Checked = True
For I = 0 To 3
 lblAnswer(I).Visible = False
Next I
txtAnswer.Visible = True
Call mnufilenew_click
End Sub
Private Sub Next_Question(Answer As Integer)
Dim VUsed(29) As Integer, I As Integer, J As Integer
Dim Index(3)
lblComment.Caption = ""
NumAns = NumAns + 1
'Generate the next question based on selected options
Answer = Int(Rnd * 29) + 1
If mnuOptionsCapitals.Checked = True Then
 lblGiven.Caption = State(Answer)
Else
 lblGiven.Caption = Capital(Answer)
End If
If mnuOptionsMC.Checked = True Then
'Multiple choice answers
'Vused array is used to see which states have
'been selected as possible answers
 For I = 1 To 29
 VUsed(I) = 0
 Next I
'Pick four different state indices (J) at random
'These are used to set up multiple choice answers
'Stored in the Index array
 I = 0
 Do
 Do
 J = Int(Rnd * 29) + 1
 Loop Until VUsed(J) = 0 And J <> Answer
VUsed(J) = 1
 Index(I) = J
 I = I + 1
 Loop Until I = 4
'Now replace one index (at random) with correct answer
 Index(Int(Rnd * 4)) = Answer
'Display multiple choice answers in label boxes
 For I = 0 To 3
 If mnuOptionsCapitals.Checked = True Then
 lblAnswer(I).Caption = Capital(Index(I))
 Else
 lblAnswer(I).Caption = State(Index(I))
 End If
 Next I
Else
'Type-in answers
 txtAnswer.Locked = False
 txtAnswer.Text = ""
 txtAnswer.SetFocus
End If
End Sub
Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
'Check type in answer'
Dim Iscorrect As Integer
Dim YourAnswer As String, TheAnswer As String
'Exit if already answered
If cmdNext.Enabled = True Then Exit Sub
If (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= vbKeyA + 32 And KeyAscii <= vbKeyZ + 32) Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
'Acceptable keystroke
 If KeyAscii <> vbKeyReturn Then Exit Sub
'Lock text box once answer entered
 txtAnswer.Locked = True
 Iscorrect = 0
'Convert response and correct answers to all upper
'case for typing problems
 YourAnswer = UCase(txtAnswer.Text)
 If mnuOptionsCapitals.Checked = True Then
 TheAnswer = UCase(Capital(CorrectAnswer))
 Else
 TheAnswer = UCase(State(CorrectAnswer))
 End If
'Check for both exact and approximate spellings
 If YourAnswer = TheAnswer Or SoundEx(YourAnswer, Wsound()) = SoundEx(TheAnswer, Wsound()) Then Iscorrect = 1
 Call Update_Score(Iscorrect)
Else
'Unacceptable keystroke
 KeyAscii = 0
End If
End Sub
