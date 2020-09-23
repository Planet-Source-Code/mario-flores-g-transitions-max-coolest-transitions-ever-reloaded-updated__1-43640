VERSION 5.00
Begin VB.Form ToolBox 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image CloseDN 
      Height          =   270
      Left            =   2040
      Picture         =   "ToolBox.frx":0000
      Top             =   720
      Width           =   270
   End
   Begin VB.Image CloseUP 
      Height          =   270
      Left            =   1680
      Picture         =   "ToolBox.frx":00B2
      Top             =   720
      Width           =   270
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   1080
      Picture         =   "ToolBox.frx":015E
      Top             =   120
      Width           =   270
   End
   Begin VB.Image PauseDN 
      Height          =   270
      Left            =   1320
      Picture         =   "ToolBox.frx":020A
      Top             =   720
      Width           =   270
   End
   Begin VB.Image PauseUP 
      Height          =   270
      Left            =   960
      Picture         =   "ToolBox.frx":0317
      Top             =   720
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   600
      Picture         =   "ToolBox.frx":0423
      Top             =   120
      Width           =   270
   End
   Begin VB.Image PlayDN 
      Height          =   270
      Left            =   480
      Picture         =   "ToolBox.frx":052F
      Top             =   720
      Width           =   270
   End
   Begin VB.Image PlayUp 
      Height          =   270
      Left            =   120
      Picture         =   "ToolBox.frx":079B
      Top             =   720
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   120
      Picture         =   "ToolBox.frx":09FA
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "ToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagOver As Boolean
Private Sub Form_Activate()
Image1 = PlayUp
Image2 = PauseUP
Image3 = CloseUP
Me.Move 0, 0
AlwaysonTop Me, True
End Sub

Private Sub Form_Load()
FlagOver = False
End Sub

Private Sub Image1_Click()
Image2 = PauseUP
Paused = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FlagOver = True Then Exit Sub
FlagOver = True
Image1 = PlayDN
End Sub

Private Sub Image2_Click()
Paused = True
Image2 = PauseDN
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Paused = True Then Exit Sub
If FlagOver = True Then Exit Sub
FlagOver = True
Image2 = PauseDN
End Sub

Private Sub Image3_Click()
FullSize = False
ENDSHOW
Paused = False
Unload Me
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FlagOver = True Then Exit Sub
FlagOver = True
Image3 = CloseDN
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FlagOver = False Then Exit Sub
FlagOver = False
If Image1 = PlayDN Then Image1 = PlayUp
If Image2 = PauseDN And Paused = False Then Image2 = PauseUP
If Image3 = CloseDN Then Image3 = CloseUP
End Sub

