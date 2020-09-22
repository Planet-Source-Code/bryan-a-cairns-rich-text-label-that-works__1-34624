VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4995
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Edit Mode"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display Mode"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin Project1.RTFLabel RTFLabel1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5953
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextRTF         =   $"Form1.frx":1FB26
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RTFLabel1.DisplayMode
End Sub

Private Sub Command2_Click()
RTFLabel1.EditMode
End Sub

Private Sub Form_Load()
Me.Show
DoEvents

'set the ID so the control can tell itself from others
RTFLabel1.SetID "1"

'load a rtf document
RTFLabel1.LoadFile App.Path & "\Document.rtf"
'alternatly if you are an RTF wizard you can put the raw data in there
'RTFLabel1.TextRTF = "Some RTF formated text"

'allow dynamic movement
RTFLabel1.AllowMove = True
End Sub

Private Sub RTFLabel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'force the control to update when moved
    If Button = 1 Then RTFLabel1.UpdateMoving
End Sub



