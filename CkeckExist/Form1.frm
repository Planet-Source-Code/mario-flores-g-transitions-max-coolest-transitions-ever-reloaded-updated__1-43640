VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TRANSITIONS CHECKER"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Enable to check for 3rd Party Transitions"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   83
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3rd &Party"
      Height          =   495
      Left            =   5040
      TabIndex        =   82
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&heck"
      Height          =   495
      Left            =   1680
      TabIndex        =   76
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HINT:"
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
      Left            =   600
      TabIndex        =   81
      Top             =   5640
      Width           =   525
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H00F57049&
      Height          =   1095
      Left            =   600
      TabIndex        =   80
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Bars)"
      Height          =   195
      Left            =   1800
      TabIndex        =   79
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Disolve)"
      Height          =   195
      Left            =   1800
      TabIndex        =   78
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Basics)"
      Height          =   195
      Left            =   1800
      TabIndex        =   77
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   37
      Left            =   6480
      TabIndex        =   75
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   36
      Left            =   6480
      TabIndex        =   74
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   35
      Left            =   6480
      TabIndex        =   73
      Top             =   4440
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   34
      Left            =   6480
      TabIndex        =   72
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   33
      Left            =   6480
      TabIndex        =   71
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   32
      Left            =   6480
      TabIndex        =   70
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   31
      Left            =   6480
      TabIndex        =   69
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   30
      Left            =   6480
      TabIndex        =   68
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   29
      Left            =   6480
      TabIndex        =   67
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   28
      Left            =   6480
      TabIndex        =   66
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   27
      Left            =   6480
      TabIndex        =   65
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   26
      Left            =   6480
      TabIndex        =   64
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   25
      Left            =   6480
      TabIndex        =   63
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   24
      Left            =   6480
      TabIndex        =   62
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   23
      Left            =   6480
      TabIndex        =   61
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   22
      Left            =   6480
      TabIndex        =   60
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   6480
      TabIndex        =   59
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   6480
      TabIndex        =   58
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   6480
      TabIndex        =   57
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   3240
      TabIndex        =   56
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   17
      Left            =   3240
      TabIndex        =   55
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   16
      Left            =   3240
      TabIndex        =   54
      Top             =   4440
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   15
      Left            =   3240
      TabIndex        =   53
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   14
      Left            =   3240
      TabIndex        =   52
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   3240
      TabIndex        =   51
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   3240
      TabIndex        =   50
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   3240
      TabIndex        =   49
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   3240
      TabIndex        =   48
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   3240
      TabIndex        =   47
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   3240
      TabIndex        =   46
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   3240
      TabIndex        =   45
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   3240
      TabIndex        =   44
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   43
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3240
      TabIndex        =   42
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3240
      TabIndex        =   41
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   40
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   39
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   38
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strips"
      Height          =   195
      Index           =   29
      Left            =   5400
      TabIndex        =   37
      Top             =   3000
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stretch"
      Height          =   195
      Index           =   30
      Left            =   5400
      TabIndex        =   36
      Top             =   3240
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      Height          =   195
      Index           =   31
      Left            =   5400
      TabIndex        =   35
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Twister"
      Height          =   195
      Index           =   32
      Left            =   5400
      TabIndex        =   34
      Top             =   3720
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacuum"
      Height          =   195
      Index           =   33
      Left            =   5400
      TabIndex        =   33
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Water"
      Height          =   195
      Index           =   34
      Left            =   5400
      TabIndex        =   32
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel"
      Height          =   195
      Index           =   35
      Left            =   5400
      TabIndex        =   31
      Top             =   4440
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wipe"
      Height          =   195
      Index           =   36
      Left            =   5400
      TabIndex        =   30
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WormHole"
      Height          =   195
      Index           =   37
      Left            =   5400
      TabIndex        =   29
      Top             =   4920
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spiral"
      Height          =   195
      Index           =   28
      Left            =   5400
      TabIndex        =   28
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slide"
      Height          =   195
      Index           =   27
      Left            =   5400
      TabIndex        =   27
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RollDown"
      Height          =   195
      Index           =   26
      Left            =   5400
      TabIndex        =   26
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ripple"
      Height          =   195
      Index           =   25
      Left            =   5400
      TabIndex        =   25
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RadialWipe"
      Height          =   195
      Index           =   24
      Left            =   5400
      TabIndex        =   24
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixelate"
      Height          =   195
      Index           =   23
      Left            =   5400
      TabIndex        =   23
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PeelABCD"
      Height          =   195
      Index           =   22
      Left            =   5400
      TabIndex        =   22
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PageCurl"
      Height          =   195
      Index           =   21
      Left            =   5400
      TabIndex        =   21
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liquid"
      Height          =   195
      Index           =   20
      Left            =   5400
      TabIndex        =   20
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LightWipe"
      Height          =   195
      Index           =   19
      Left            =   5400
      TabIndex        =   19
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lens"
      Height          =   195
      Index           =   18
      Left            =   480
      TabIndex        =   18
      Top             =   4920
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jaws"
      Height          =   195
      Index           =   17
      Left            =   480
      TabIndex        =   17
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iris"
      Height          =   195
      Index           =   16
      Left            =   480
      TabIndex        =   16
      Top             =   4440
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inset"
      Height          =   195
      Index           =   15
      Left            =   480
      TabIndex        =   15
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GlassBlock"
      Height          =   195
      Index           =   14
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FlowMotion"
      Height          =   195
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Top             =   3720
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FadeWhite"
      Height          =   195
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   3480
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fade"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RandomDissolve"
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curtains"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curls"
      Height          =   195
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compositor"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColorFade"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckerBoard"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CenterCurls"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blinds"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RevealTrans"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RandomBars"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barn"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LocalizarTransitions As String
Dim i

Private Sub Check1_Click()
If Check1.Value = 1 Then Check1.Visible = False
End Sub

Private Sub Command1_Click()
On Error Resume Next
'*****CHECK MICROSOFT SLIDES """
For i = 0 To 37
    
    LocalizarTransitions = ViewStringRegistryValue(HKEY_CLASSES_ROOT, "DXImageTransform.Microsoft." & Label1(i).Caption)
    If LocalizarTransitions = "YES" Then
        Label2(i).ForeColor = vbBlue
        Label2(i).Caption = "TRUE"
      End If
    If LocalizarTransitions = "NO" Then
        Label2(i).Caption = "FALSE"
    End If
  
Next i
Command1.Enabled = False
Check1.Enabled = True
End Sub



Private Sub Command2_Click()
On Error Resume Next
'*****CHECK METAFILECREATIONS SLIDES ""DXTMETA2.DLL""
For i = 0 To 37
    LocalizarTransitions = ViewStringRegistryValue(HKEY_CLASSES_ROOT, "DXImageTransform.Metacreations." & Label1(i).Caption)
    If LocalizarTransitions = "YES" Then
        Label2(i).ForeColor = &H8000&
        Label2(i).Caption = "TRUE"
    End If
           
    If LocalizarTransitions = "NO" Then
        If Label2(i).Caption <> "TRUE" Then
            Label2(i).ForeColor = &H8000&
            Label2(i).Caption = "FALSE"
        End If
    End If
Next i
End Sub
