VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Transitions MAX using DirectDraw"
   ClientHeight    =   9765
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":2372
   ScaleHeight     =   9765
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8940
      Left            =   3720
      Picture         =   "frmMain.frx":8091
      ScaleHeight     =   8910
      ScaleWidth      =   7830
      TabIndex        =   23
      Top             =   120
      Width           =   7860
      Begin VB.CommandButton Command1 
         Caption         =   "Let Me Try It"
         Height          =   495
         Left            =   3480
         TabIndex        =   26
         Top             =   8160
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":D313
         ForeColor       =   &H00800080&
         Height          =   795
         Left            =   1320
         TabIndex        =   35
         Top             =   6960
         Width           =   5205
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":D3C6
         ForeColor       =   &H00800080&
         Height          =   795
         Left            =   1320
         TabIndex        =   34
         Top             =   6240
         Width           =   5205
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   720
         TabIndex        =   33
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":D470
         ForeColor       =   &H00800080&
         Height          =   795
         Left            =   1320
         TabIndex        =   32
         Top             =   5280
         Width           =   5205
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hints:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   960
         TabIndex        =   31
         Top             =   4920
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>   3:Step Click the Play Button"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   30
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":D534
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   960
         TabIndex        =   29
         Top             =   3240
         Width           =   7320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>  1 Step : Browse For Pictures in Browse icon (Top LEFT)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   28
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<>  Please in You dont Understand Read The README FILE!!...."
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   27
         Top             =   1560
         Width           =   4620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Updated 04.03.03"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3240
         TabIndex        =   25
         Top             =   720
         Width           =   1290
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   6840
         Picture         =   "frmMain.frx":D5F5
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transitions MAX Using Direct Show by MArio Flores G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BE6B47&
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   480
         Width           =   4590
      End
   End
   Begin VB.FileListBox files 
      Height          =   480
      Left            =   120
      Pattern         =   "*bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.mpg;*.mpg2;*.avi"
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Slide 
      BackColor       =   &H00BE6B47&
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   2
      Left            =   2400
      Picture         =   "frmMain.frx":DA37
      ScaleHeight     =   660
      ScaleWidth      =   1935
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ComboBox cmbTransitions 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Default Transition"
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.PictureBox Slide 
      BackColor       =   &H00BE6B47&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   660
      Index           =   3
      Left            =   2400
      Picture         =   "frmMain.frx":DD48
      ScaleHeight     =   660
      ScaleWidth      =   1935
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ComboBox cmbEffects 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Default Transition"
         Top             =   120
         Width           =   1515
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   9390
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   13996
            Text            =   "Please Read The ReadMe File...................And Remember to give FeedBack on Code..Thanks! MArio Flores G"
            TextSave        =   "Please Read The ReadMe File...................And Remember to give FeedBack on Code..Thanks! MArio Flores G"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Espere 
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.PictureBox Frame 
      BorderStyle     =   0  'None
      Height          =   6450
      Left            =   2520
      ScaleHeight     =   6450
      ScaleWidth      =   8940
      TabIndex        =   9
      Top             =   710
      Width           =   8940
      Begin MSComctlLib.Slider SlideTimer 
         Height          =   630
         Left            =   0
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1111
         _Version        =   393216
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4470
         TabIndex        =   20
         Top             =   6000
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transitions MAX Using Direct Show by MArio Flores G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BE6B47&
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Top             =   2520
         Width           =   4590
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4080
         Picture         =   "frmMain.frx":E059
         Top             =   2760
         Width           =   480
      End
   End
   Begin VB.Image StopDN 
      Height          =   645
      Left            =   7800
      Picture         =   "frmMain.frx":E49B
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image StopUP 
      Height          =   645
      Left            =   7080
      Picture         =   "frmMain.frx":EDCE
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image cmdStop 
      Height          =   645
      Left            =   5160
      Picture         =   "frmMain.frx":F755
      Top             =   7365
      Width           =   645
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   7440
      TabIndex        =   22
      Top             =   240
      Width           =   195
   End
   Begin VB.Label LabelFull 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Left            =   720
      MouseIcon       =   "frmMain.frx":100DC
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPlaymode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transition"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2800
      TabIndex        =   18
      Top             =   330
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   4
      Left            =   1920
      Picture         =   "frmMain.frx":1022E
      Top             =   3000
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   3
      Left            =   720
      MouseIcon       =   "frmMain.frx":107B8
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   1800
      Picture         =   "frmMain.frx":1090A
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReadMe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   4
      Left            =   720
      MouseIcon       =   "frmMain.frx":114CC
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4920
      Width           =   720
   End
   Begin VB.Image PlayDN 
      Height          =   645
      Left            =   10560
      Picture         =   "frmMain.frx":1161E
      Top             =   4560
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image PlayUP 
      Height          =   645
      Left            =   10560
      Picture         =   "frmMain.frx":11F71
      Top             =   3840
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image PauseDN 
      Height          =   645
      Left            =   10560
      Picture         =   "frmMain.frx":12922
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image PauseUP 
      Height          =   645
      Left            =   10560
      Picture         =   "frmMain.frx":13271
      Top             =   2160
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image cmdPause 
      Height          =   645
      Left            =   4320
      Picture         =   "frmMain.frx":13C0F
      Top             =   7365
      Width           =   645
   End
   Begin VB.Image cmdPlay 
      Height          =   645
      Left            =   3480
      Picture         =   "frmMain.frx":145AD
      Top             =   7365
      Width           =   645
   End
   Begin VB.Image CloseUP 
      Height          =   315
      Left            =   10920
      Picture         =   "frmMain.frx":14F5E
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image CloseDN 
      Height          =   315
      Left            =   10920
      Picture         =   "frmMain.frx":15237
      Top             =   1200
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   11040
      Picture         =   "frmMain.frx":1553B
      Top             =   80
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  MArio Flores G"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   99
      Left            =   8640
      TabIndex        =   5
      Top             =   7440
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   1920
      Picture         =   "frmMain.frx":15814
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transitions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   2
      Left            =   720
      MouseIcon       =   "frmMain.frx":15D9E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmMain.frx":15EF0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "frmMain.frx":16042
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   5
      Left            =   690
      MouseIcon       =   "frmMain.frx":1B824
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5550
      Width           =   1140
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   1680
      Picture         =   "frmMain.frx":1B976
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   6
      Left            =   840
      MouseIcon       =   "frmMain.frx":21C00
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6240
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEC7B1&
      Height          =   195
      Index           =   0
      Left            =   720
      MouseIcon       =   "frmMain.frx":21D52
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1080
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ButtonFlag As Boolean


Private Sub cmbEffects_Click()
cmbTransitions.Text = vbNullString
lblPlaymode.Caption = "Effect"
Label1_Click 3
End Sub

Private Sub cmbTransitions_Click()
cmbEffects.Text = vbNullString
lblPlaymode.Caption = "Transition"
Label1_Click 2
End Sub



Private Sub cmdStop_Click()
ENDSHOW
Paused = False
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If objVideoWindow.Visible = False Then Exit Sub
If ButtonFlag = True Then Exit Sub
cmdStop = StopDN
ButtonFlag = True
End Sub

Private Sub Command1_Click()
Picture1.Visible = False

End Sub

Private Sub Form_Initialize()

On Local Error GoTo ErrLine

           Set gbl_objTimeline = New AMTimeline
           Exit Sub
             
ErrLine:
            Err.Clear
            Exit Sub
End Sub
         
Private Sub Form_Load()

On Local Error GoTo ErrLine
            
            SLIDETIME = 4
            SlideTimer.Value = SLIDETIME
'*********************************************************************
'PUT THE TRANSITIONS & EFFECTS ON COMBOBOX
'*********************************************************************
            
            cmbTransitions.Text = vbNullString
            cmbEffects.Text = vbNullString
            
            Call ViewTransitionFriendlyNamesDirect(cmbTransitions)
            Call ViewEffectFriendlyNamesDirect(cmbEffects)
            
            'assign the default transition
            If TransitionCLSIDToFriendlyName(gbl_objTimeline.GetDefaultTransitionB) <> vbNullString Then
               
               cmbTransitions.Text = TransitionCLSIDToFriendlyName(gbl_objTimeline.GetDefaultTransitionB)
            End If
            
           
            
            'assign the default effect
            If TransitionCLSIDToFriendlyName(gbl_objTimeline.GetDefaultTransitionB) <> vbNullString Then
               cmbEffects.Text = EffectCLSIDToFriendlyName(gbl_objTimeline.GetDefaultEffectB)
            End If
            
            Exit Sub
ErrLine:
            Err.Clear
            Exit Sub
End Sub
            
            


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ButtonFlag = False Then Exit Sub
    
    For i = 0 To 6
    Label1(i).ForeColor = &HFEC7B1
    Next i
    If cmdPlay = PlayDN Then cmdPlay = PlayUp
    If Image3 = CloseDN Then Image3 = CloseUP
    If cmdPause = PauseDN And Paused = False Then cmdPause = PauseUP
    If cmdStop = StopDN Then cmdStop = StopUP
    ButtonFlag = False
End Sub




Private Sub Form_Unload(Cancel As Integer)
            On Local Error GoTo ErrLine
                                  
            If Not gbl_objTimeline Is Nothing Then
                Call ENDSHOW
                Set gbl_objTimeline = Nothing
            End If
            
            Unload Me
            End
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub

Private Sub cmdExit_Click()

On Local Error GoTo ErrLine
                      
            Form_Unload 0
            Exit Sub
ErrLine:
            Err.Clear
            Exit Sub
End Sub
              
Public Sub cmdPlay_Click()
On Error Resume Next
Dim TempCLSID As String

    
    If files.ListCount = 0 Then Exit Sub 'Case Not Browsed
    
    If Running = True And Paused = False Then Exit Sub  'Do Nothing While is Running
    
    
    If Paused = True Then 'In case Slide Is Active
        Paused = False
        cmdPause = PauseUP
        Exit Sub
    End If
            
    
    
    ENDSHOW ' Yust in case Is active
    
    If cmbTransitions.Text = vbNullString Then
       Style = XEffect
       TempCLSID = EffectFriendlyNameToCLSID(cmbEffects)
    Else
       Style = XTransition
       TempCLSID = TransitionFriendlyNameToCLSID(cmbTransitions)
    End If
            
    Set Pfiles = New Collection
    
    For i = 1 To files.ListCount
        files.ListIndex = i - 1
        Pfiles.add Array(SourcePath & "\" & files.FileName)
    Next i
            
    Set gbl_objTimeline = CreateSLIDESHOW(TempCLSID, Pfiles)
          
           
    For i = 0 To 3
        Label1(i).Enabled = False
    Next i
    
    lblNumber.Caption = vbNullString
    SlideTimer.Visible = False
    LabelFull.Visible = True 'Bugs
    Slide(2).Visible = False
    Slide(3).Visible = False
'****************************************************************************************
        'READY>>>>>> SLIDE PRESENTATION(STARTS HERE)
         Call SHOWALL(gbl_objTimeline)
'****************************************************************************************
         LabelFull.Visible = False
         FullSize = False
    
         For i = 0 To 3
            Label1(i).Enabled = True
         Next i
        
         Exit Sub
            
End Sub
  
            
' ******************************************************************************************************************************
' * Maps Transitions friendly names to a Combobox
' *
' ******************************************************************************************************************************
            Private Sub ViewTransitionFriendlyNamesDirect(cmbComboBox As Control)
            On Local Error GoTo ErrLine
            
            If Not cmbComboBox Is Nothing Then
               If TypeName(cmbComboBox) = "ComboBox" Then
                  With cmbComboBox
                    
                     .AddItem "Barn" '
                     .AddItem "Bars" '
                     .AddItem "Basics"
                     .AddItem "Blinds" '
                     .AddItem "Burn Film"
                     .AddItem "CenterCurls"
                     .AddItem "Checkerboard"
                     .AddItem "ColorFade"
                     .AddItem "Compositor"
                     .AddItem "Curls"
                     .AddItem "Curtains"
                     .AddItem "Disolve"
                     .AddItem "Fade"
                     .AddItem "FadeWhite"
                     .AddItem "FlowMotion"
                     .AddItem "GlassBlock"
                     .AddItem "Grid"
                     .AddItem "Inset"
                     .AddItem "Iris"
                     .AddItem "Jaws"
                     .AddItem "Lens"
                     .AddItem "LightWipe"
                     .AddItem "Liquid"
                     .AddItem "PageCurl"
                     .AddItem "PeelABCD"
                     .AddItem "Pixelate"
                     .AddItem "RadialWipe"
                     .AddItem "Ripple"
                     .AddItem "RollDown"
                     .AddItem "Slide"
                     .AddItem "SMPTE Wipe"
                     .AddItem "Spiral"
                     .AddItem "Strips"
                     .AddItem "Stretch"
                     .AddItem "Threshold"
                     .AddItem "Twister"
                     .AddItem "Vacuum"
                     .AddItem "Water"
                     .AddItem "Wheel"
                     .AddItem "Wipe"
                     .AddItem "WormHole"
                     .AddItem "Zigzag"
             
                  End With
               End If
            End If
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub
' ******************************************************************************************************************************
' * Maps Effects friendly names to a Combobox
' *
' ******************************************************************************************************************************
            Private Sub ViewEffectFriendlyNamesDirect(cmbComboBox As Control)
            On Local Error GoTo ErrLine
            
            If Not cmbComboBox Is Nothing Then
               If TypeName(cmbComboBox) = "ComboBox" Then
                  With cmbComboBox
                     .AddItem "Additive"
                     .AddItem "BasicImage"
                     .AddItem "Blur"
                     .AddItem "Brightness"
                     .AddItem "Chroma"
                     .AddItem "DropShadow"
                     .AddItem "Emboss"
                     .AddItem "Engrave"
                     .AddItem "Fade"
                     .AddItem "Glow"
                     .AddItem "MaskFilter"
                     .AddItem "Pixelate"
                     .AddItem "Shadow"
                     .AddItem "Wave"

                  End With
               End If
            End If
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub
            
Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonFlag = True Then Exit Sub
cmdPlay = PlayDN
ButtonFlag = True
End Sub

Private Sub Image3_Click()
  Dim frm As Form
            On Local Error GoTo ErrLine
                 
            Unload frmMain
     
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonFlag = True Then Exit Sub
Image3 = CloseDN
ButtonFlag = True
End Sub

Public Sub cmdPause_Click()
On Error Resume Next
If objVideoWindow.Visible = False Then Exit Sub
Paused = True
cmdPause = PauseDN
End Sub

Private Sub cmdPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If objVideoWindow.Visible = False Then Exit Sub
If Paused = True Then Exit Sub
If ButtonFlag = True Then Exit Sub
cmdPause = PauseDN
ButtonFlag = True
End Sub

Private Sub Label1_Click(Index As Integer)

'*******************************BROWSE*******************************
If Index = 0 Then
    'FrmBrowse.Show 0
    
ShowTree

'files.Path = SourcePath
End If

'*******************************TIMER*******************************
If Index = 1 Then

If SlideTimer.Visible = True Then
  SlideTimer.Visible = False
Else
  SlideTimer.Visible = True
  SlideTimer.SetFocus
End If


End If

'*******************************SLIDER*******************************
If Index = 2 Or Index = 3 Then

    If Slide(Index).Visible = True Then
        Slide(Index).Visible = False
        Exit Sub
    End If

    If Slide(Index).Visible = False Then
        If Slide(2).Visible = True Then Slide(2).Visible = False
        If Slide(3).Visible = True Then Slide(3).Visible = False
        Slide(Index).Visible = True
        Exit Sub
    End If

End If

If Index = 4 Then Shell ("notepad.exe " & App.Path & "\readme.txt "), vbMaximizedFocus
'************************************************************************
If Index = 5 Then
  
  If WindowFlag = False Then
    Espere.Visible = True
    Me.Refresh
    Call SetAutoRgn(Me)
    Espere.Visible = False
    WindowFlag = True
    Exit Sub
  End If

If WindowFlag = True Then
    Call NoTranparency(Me, vbBlack)
    WindowFlag = False
    Exit Sub
  End If



End If
If Index = 6 Then Image3_Click
End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonFlag = True Then Exit Sub
Label1(Index).ForeColor = vbWhite
ButtonFlag = True
End Sub

Private Sub LabelFull_Click()
Paused = False
FullSize = True
ToolBox.Show
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub SlideTimer_Change()
SLIDETIME = SlideTimer.Value
End Sub

Private Sub SlideTimer_LostFocus()
SlideTimer.Visible = False
End Sub
