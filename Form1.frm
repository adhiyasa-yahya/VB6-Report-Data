VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1560
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bumi Putra 1912"
      BeginProperty Font 
         Name            =   "Swis721 Blk BT"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
    Timer1.Interval = Rnd * 300 + 10
    ProgressBar1.Value = ProgressBar1.Value + 5
    Label1.Caption = ProgressBar1.Value & "%"
    If Label1.Caption = 100 & "%" Then
        Form_Homescreen.Show
        Unload Me
        
    End If
End Sub
