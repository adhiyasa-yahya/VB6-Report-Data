VERSION 5.00
Begin VB.Form Form_tampilan 
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text8"
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   12
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Nasabah"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Label Label3 
         Caption         =   "Uang pertanggungan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Lama Asuransi"
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
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Premi"
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
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor BK "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Mulai "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tanggal Selesai "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Polis "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Nama Pemegang Polis "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form_tampilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db  As New ADODB.Connection
Dim WithEvents Rs  As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Dim querry As String

Private Sub Koneksi()
    Dim Constr As String
    
    Constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "/si-arsip.accdb"
    Db.Open Constr

End Sub

Private Sub Command1_Click()
    Dim Form As New DataReport1
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    
    querry = "SELECT * FROM data_nasabah WHERE no_bk = '" & Text1.Text & "'"
    Rs.Open querry, Db, adOpenDynamic, adLockOptimistic
 
    With Form.Sections("Section2")
        .Controls("Label7").Caption = Text1.Text
        .Controls("Label8").Caption = Text2.Text
        .Controls("Label9").Caption = Text3.Text
        .Controls("Label10").Caption = Text4.Text
        .Controls("Label11").Caption = Text5.Text
        .Controls("Label24").Caption = Text5.Text
        .Controls("Label25").Caption = Text7.Text
        .Controls("Label26").Caption = Text8.Text
    End With
 
    Set Form.DataSource = Rs
    Form.Show vbModal
    
End Sub

Private Sub Form_Load()
    Call Koneksi
End Sub
