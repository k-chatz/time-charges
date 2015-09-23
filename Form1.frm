VERSION 5.00
Begin VB.Form frmTOS 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2715
   ClientLeft      =   5685
   ClientTop       =   2070
   ClientWidth     =   3855
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Cooper Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option4 
      Caption         =   "Start"
      Height          =   330
      Left            =   45
      TabIndex        =   26
      Top             =   3465
      Width           =   1860
   End
   Begin VB.OptionButton Option3 
      Caption         =   " € / Hour"
      Height          =   330
      Left            =   1935
      TabIndex        =   25
      Top             =   3465
      Width           =   1860
   End
   Begin VB.OptionButton Option2 
      Caption         =   "End debit"
      Height          =   330
      Left            =   1935
      TabIndex        =   24
      Top             =   3105
      Width           =   1860
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Time complete"
      Height          =   330
      Left            =   45
      TabIndex        =   23
      Top             =   3105
      Value           =   -1  'True
      Width           =   1860
   End
   Begin VB.CommandButton cmdPause 
      Appearance      =   0  'Flat
      Caption         =   "Pause - Resume"
      Default         =   -1  'True
      Height          =   375
      Left            =   45
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2745
      Width           =   3795
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3285
      Top             =   3825
   End
   Begin VB.Label lblHours 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   10
      Top             =   540
      Width           =   450
   End
   Begin VB.Label lblSeconds2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3420
      TabIndex        =   19
      Top             =   2340
      Width           =   390
   End
   Begin VB.Label lblMinites2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   18
      Top             =   2340
      Width           =   405
   End
   Begin VB.Label lblHours2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   17
      Top             =   2340
      Width           =   450
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2745
      TabIndex        =   16
      Top             =   2340
      Width           =   165
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3285
      TabIndex        =   15
      Top             =   2340
      Width           =   165
   End
   Begin VB.Label lblEndAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2295
      TabIndex        =   13
      Top             =   1620
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3285
      TabIndex        =   12
      Top             =   540
      Width           =   165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2745
      TabIndex        =   11
      Top             =   540
      Width           =   165
   End
   Begin VB.Label lblMinites 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   9
      Top             =   540
      Width           =   405
   End
   Begin VB.Label lblSeconds 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3420
      TabIndex        =   8
      Top             =   540
      Width           =   390
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2295
      TabIndex        =   6
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label lblPoso 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   2295
      TabIndex        =   3
      Top             =   45
      Width           =   1560
   End
   Begin VB.Label lblEuroPerHour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2295
      TabIndex        =   1
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Time elapsed:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   20
      Top             =   2340
      Width           =   2265
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   " End debit:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   14
      Top             =   1620
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Time passed:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   540
      Width           =   2265
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Start:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   5
      Top             =   900
      Width           =   2265
   End
   Begin VB.Label lblCurrentTime 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   2325
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Complete:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   1980
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " € / Hour:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   1260
      Width           =   2265
   End
   Begin VB.Label lblEndTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2295
      TabIndex        =   21
      Top             =   1980
      Width           =   1515
   End
End
Attribute VB_Name = "frmTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iStart, ISTOP As Variant
Dim Second_1 As Single
Dim EuroPerHour, EndAmount As Currency

Sub BehaviorCtr()
lblSeconds2.Caption = ""
lblMinites2.Caption = ""
lblHours2.Caption = ""

Label6.BackColor = &HFF8080
Label6.ForeColor = vbBlack
lblStart.BackColor = &HFF8080
lblStart.ForeColor = vbBlack

Label2.BackColor = &HFF8080
Label2.ForeColor = vbBlack
lblEuroPerHour.BackColor = &HFF8080
lblEuroPerHour.ForeColor = vbBlack

Label5.BackColor = &HFF8080
Label5.ForeColor = vbBlack
lblEndTime.BackColor = &HFF8080
lblEndTime.ForeColor = vbBlack

Label9.BackColor = &HFF8080
Label9.ForeColor = vbBlack
lblEndAmount.BackColor = &HFF8080
lblEndAmount.ForeColor = vbBlack

If Option1.Value = True Then
Label9.BackColor = vbBlack
Label9.ForeColor = &H80FF&
lblEndAmount.BackColor = vbBlack
lblEndAmount.ForeColor = &H80FF&

ElseIf Option2.Value = True Then
Label5.BackColor = &H0&
Label5.ForeColor = &H80FF&
lblEndTime.BackColor = &H0&
lblEndTime.ForeColor = &H80FF&

ElseIf Option3.Value = True Then
Label2.BackColor = &H0&
Label2.ForeColor = &H80FF&
lblEuroPerHour.BackColor = &H0&
lblEuroPerHour.ForeColor = &H80FF&

ElseIf Option4.Value = True Then
Label6.BackColor = &H0&
Label6.ForeColor = &H80FF&
lblStart.BackColor = &H0&
lblStart.ForeColor = &H80FF&

End If

End Sub

Private Sub cmdPause_Click()
If Timer1.Enabled = True Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
Call Calculate
End If
End Sub

Private Sub Form_Load()
Call BehaviorCtr
iStart = Format(Time(), "hh:mm:ss")
lblStart.Caption = iStart
Call Calculate
End Sub

Private Sub Label2_Click()
lblEuroPerHour.Caption = ""
Option3.Value = True
Call Calculate
End Sub

Private Sub Label5_Click()
lblEndTime.Caption = ""
Option2.Value = True
Call Calculate
End Sub

Private Sub Label6_Click()
lblStart.Caption = ""
Option4.Value = True
Call Calculate
End Sub

Private Sub Label9_Click()
lblEndAmount.Caption = ""
Option1.Value = True
Call Calculate
End Sub


Private Sub lblEndAmount_Click()
On Error Resume Next

If Option2.Value = True Or Option3.Value = True Or Option4.Value = True Then
EndAmount = InputBox("", "", lblEndAmount.Caption)
lblEndAmount.Caption = Format(EndAmount, "Currency")

End If

End Sub

Private Sub lblEndTime_Click()
If Option1.Value = True Or Option3.Value = True Or Option4.Value = True Then
TimeComplete = InputBox("", "", lblEndTime.Caption)
lblEndTime.Caption = Format(TimeComplete, "hh:mm:ss")

End If
End Sub

Private Sub lblEuroPerHour_Click()
On Error Resume Next

If Option3.Value = False Then
EuroPerHour = InputBox("", "", lblEuroPerHour)
Call Calculate
lblEuroPerHour = Format(EuroPerHour, "Currency")
End If

End Sub

Private Sub lblStart_Click()
On Error Resume Next
If Option4.Value = False Then
iStart = InputBox("", "", lblStart.Caption)
Call Calculate
lblStart.Caption = Format(iStart, "hh:mm:ss")
End If
End Sub


Sub Calculate()
On Error Resume Next
lblCurrentTime.Caption = " " & Format(Time(), "hh:mm:ss")

If iStart <> "" Then
Second_1 = DateDiff("s", Format(iStart, "hh:mm:ss"), Format(lblCurrentTime, "hh:mm:ss"))
iPoso = (EuroPerHour * Second_1) / 60 / 60
iYpoloipo = lblEndAmount - iPoso
SecElapsed = ((iYpoloipo / EuroPerHour) * 60) * 60
SecElapsed = Format(SecElapsed, "0")

If Option2.Value = True Then
lblEndTime.Caption = Format(DateAdd("s", SecElapsed, Time()), "hh:mm:ss")
ElseIf Option1.Value = True Then
lblEndAmount = Format((((DateDiff("s", lblStart, lblEndTime)) / 60) / 60) * EuroPerHour, "Currency")
End If

If Option3.Value = True Then
lblEuroPerHour.Caption = Format(lblEndAmount.Caption / ((DateDiff("s", lblStart, lblEndTime) / 60) / 60), "Currency")
End If
End If

Second_Start = (((lblEndAmount / lblEuroPerHour.Caption) * 60) * 60) * (-1)
If Option4.Value = True Then
iStart = Format(DateAdd("s", Second_Start, lblEndTime), "hh:mm:ss")
lblStart.Caption = iStart
End If

If Second_1 > 0 Then
lblSeconds.Caption = Format(Second_1 Mod 60, "00")
lblMinites.Caption = Format((Second_1 \ 60) Mod 60, "00")
lblHours.Caption = Format(((Second_1 \ 60) \ 60), "00")
Else
lblSeconds.Caption = ""
lblMinites.Caption = ""
lblHours.Caption = ""
End If

If EuroPerHour <> "" Then
lblPoso = Format(iPoso, "Currency")
Else
lblPoso = ""
End If

If SecElapsed > 0 Then
lblSeconds2.Caption = Format(SecElapsed Mod 60, "00")
lblMinites2.Caption = Format((SecElapsed \ 60) Mod 60, "00")
lblHours2.Caption = Format(((SecElapsed \ 60) \ 60), "00")
Else
lblSeconds2.Caption = ""
lblMinites2.Caption = ""
lblHours2.Caption = ""

If Option2.Value = True Then
lblEndTime.Caption = ""
End If

End If
frmTOS.Caption = " " & Format(iPoso, "Currency") & " | " & lblHours2.Caption & ":" & lblMinites2.Caption & ":" & lblSeconds2.Caption & " | Time charges"
End Sub

Private Sub Option1_Click()
Call BehaviorCtr
End Sub

Private Sub Option2_Click()
Call BehaviorCtr
End Sub

Private Sub Option3_Click()
Call BehaviorCtr
End Sub

Private Sub Option4_Click()
Call BehaviorCtr
End Sub

Private Sub Timer1_Timer()
Call Calculate
End Sub
