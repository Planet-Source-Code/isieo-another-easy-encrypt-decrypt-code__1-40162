VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encrypt/Decrypt"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Line Line7 
      X1              =   2160
      X2              =   2160
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   1800
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   1920
      Y1              =   2760
      Y2              =   2640
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2520
      Y1              =   2520
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   2520
      Y1              =   2520
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   160
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter0 As Integer
Dim counter As Integer
Dim counter1 As Integer
Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
Text1 = Text1 + vbCrLf
counter = 0
counter1 = 0
counter0 = 0
Timer1.Enabled = True


End Sub

Private Sub Command2_Click()
counter = 0
counter1 = 0
counter0 = 0
Timer2.Enabled = True

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Len(Text1) > Len(Text3) Then
If counter0 = Len(Text1) Then Timer1.Enabled = False
Else
If counter0 = Len(Text3) Then Timer1.Enabled = False
End If
If counter = Len(Text1) Then counter = 1
If counter1 = Len(Text3) Then counter1 = 1
mixz = Asc(Mid(Text1, counter, 1)) + Asc(Mid(Text3, counter1, 1))
counter = counter + 1
counter1 = counter1 + 1
If mixz > 255 Then mixz = mixz - 225
DoEvents
Text2 = Text2 + Chr(mixz)
counter0 = counter0 + 1
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Len(Text2) > Len(Text3) Then
If counter0 = Len(Text2) Then
Text2 = Text2 + vbCrLf
Timer2.Enabled = False
End If
Else
If counter0 = Len(Text3) Then
Text2 = Text2 + vbCrLf
Timer2.Enabled = False
End If
End If
If counter = Len(Text2) Then counter = 1
If counter1 = Len(Text3) Then counter1 = 1
If Asc(Mid(Text2, counter, 1)) > Asc(Mid(Text3, counter1, 1)) Then
mixz = Asc(Mid(Text2, counter, 1)) - Asc(Mid(Text3, counter1, 1))
Else
mixz = Asc(Mid(Text3, counter, 1)) - Asc(Mid(Text2, counter1, 1))
End If
counter = counter + 1
counter1 = counter1 + 1
If mixz > 255 Then mixz = mixz - 225
Text1 = Text1 + Chr(mixz)
counter0 = counter0 + 1
End Sub
