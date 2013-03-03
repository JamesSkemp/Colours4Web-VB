VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "Colours4Web"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame FrameSaved 
      Caption         =   "Saved"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   4815
      Begin VB.TextBox TextSaved 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton CommandClear 
      Caption         =   "Restart/Clear"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CommandRandom 
      Caption         =   "Randomize"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame FrameAbout 
      Caption         =   "About"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   4815
      Begin VB.CommandButton CommandEmail 
         Caption         =   "Email the creator"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CommandSite 
         Caption         =   "Visit StrivingLife.net"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TextAbout 
         Enabled         =   0   'False
         Height          =   765
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.TextBox TextHexFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox TextHexBlue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox TextHexGreen 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox TextHexRed 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.HScrollBar HScrollBlue 
      Height          =   255
      Left            =   2520
      Max             =   5
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TextBlue 
      Enabled         =   0   'False
      Height          =   975
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   1110
   End
   Begin VB.TextBox TextGreen 
      Enabled         =   0   'False
      Height          =   975
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1110
   End
   Begin VB.TextBox TextRed 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1110
   End
   Begin VB.HScrollBar HScrollGreen 
      Height          =   255
      Left            =   1320
      Max             =   5
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.HScrollBar HScrollRed 
      Height          =   255
      Left            =   120
      Max             =   5
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TextFinal 
      Enabled         =   0   'False
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label LabelBlue 
      Alignment       =   2  'Center
      Caption         =   "Blue"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LabelGreen 
      Alignment       =   2  'Center
      Caption         =   "Green"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LabelRed 
      Alignment       =   2  'Center
      Caption         =   "Red"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeToHex As Integer
Dim MadeToHex As String
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Sub MakeHexCode(MakeToHex)
 If MakeToHex = 0 Then MadeToHex = "00"
 If MakeToHex = 1 Then MadeToHex = "33"
 If MakeToHex = 2 Then MadeToHex = "66"
 If MakeToHex = 3 Then MadeToHex = "99"
 If MakeToHex = 4 Then MadeToHex = "CC"
 If MakeToHex = 5 Then MadeToHex = "FF"
End Sub

Sub DoStart()
 HScrollRed.Value = 0
 HScrollGreen.Value = 0
 HScrollBlue.Value = 0
 TextFinal.BackColor = RGB(HScrollRed.Value * 51, HScrollGreen.Value * 51, HScrollBlue.Value * 51)
 TextRed.BackColor = RGB(HScrollRed.Value * 51, 0, 0)
 TextGreen.BackColor = RGB(0, HScrollGreen.Value * 51, 0)
 TextBlue.BackColor = RGB(0, 0, HScrollBlue.Value * 51)

End Sub

Sub DoTextHex()
 MakeHexCode (HScrollRed.Value)
 TextHexRed.Text = MadeToHex
 TextHexFinal.Text = "#" + MadeToHex
 MakeHexCode (HScrollGreen.Value)
 TextHexGreen.Text = MadeToHex
 TextHexFinal.Text = TextHexFinal.Text + "" + MadeToHex
 MakeHexCode (HScrollBlue.Value)
 TextHexBlue.Text = MadeToHex
 TextHexFinal.Text = TextHexFinal.Text + "" + MadeToHex
End Sub

Private Sub CommandClear_Click()
 DoStart
End Sub

Private Sub CommandEmail_Click()
ShellExecute hwnd, "open", "mailto:homeofjrs@eml.cc", vbNullString, vbNullString, SW_SHOW
End Sub

Sub RandomRed()
 Randomize
 HScrollRed.Value = Int((6 * Rnd))
End Sub

Sub RandomGreen()
 Randomize
 HScrollGreen.Value = Int((6 * Rnd))
End Sub

Sub RandomBlue()
 Randomize
 HScrollBlue.Value = Int((6 * Rnd))
End Sub

Private Sub CommandRandom_Click()
 RandomRed
 RandomGreen
 RandomBlue
End Sub

Private Sub CommandSave_Click()
 TextSaved.Text = "" + TextSaved.Text + TextHexFinal.Text + ", "
End Sub

Private Sub CommandSite_Click()
ShellExecute hwnd, "open", "http://strivinglife.net/programs.htm", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  CommandRandom_Click
 ElseIf KeyCode = 65 Then
  RandomRed
 ElseIf KeyCode = 83 Then
  RandomGreen
 ElseIf KeyCode = 68 Then
  RandomBlue
 ElseIf KeyCode = 70 Then
  CommandRandom_Click
 ElseIf KeyCode = 82 Then
  CommandClear_Click
 ElseIf KeyCode = 69 Then
  CommandSave_Click
 End If

End Sub

Private Sub Form_Load()
 TextAbout.Text = "Colours4Web - James Skemp" & vbCrLf & "a - Random Red | s - Random Green | d - Random Blue" & vbCrLf & "f - Random all | r - Restart/Clear | e - Save"
 DoStart
 DoTextHex
End Sub

Private Sub Form_Resize()
 FormMain.Height = 5130
 FormMain.Width = 5160
End Sub

Private Sub HScrollBlue_Change()
 TextFinal.BackColor = RGB(HScrollRed.Value * 51, HScrollGreen.Value * 51, HScrollBlue.Value * 51)
 TextBlue.BackColor = RGB(0, 0, HScrollBlue.Value * 51)
 DoTextHex
End Sub

Private Sub HScrollGreen_Change()
 TextFinal.BackColor = RGB(HScrollRed.Value * 51, HScrollGreen.Value * 51, HScrollBlue.Value * 51)
 TextGreen.BackColor = RGB(0, HScrollGreen.Value * 51, 0)
 DoTextHex
End Sub

Private Sub HScrollRed_Change()
 TextFinal.BackColor = RGB(HScrollRed.Value * 51, HScrollGreen.Value * 51, HScrollBlue.Value * 51)
 TextRed.BackColor = RGB(HScrollRed.Value * 51, 0, 0)
 DoTextHex
End Sub
