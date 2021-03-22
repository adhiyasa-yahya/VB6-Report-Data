VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Pegawai 
   Caption         =   "Form_pegawai"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   10404.77
   ScaleMode       =   0  'User
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox asuransi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   22
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   483
      Left            =   8520
      TabIndex        =   17
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Cari 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   840
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "PEGAWAI"
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.TextBox uang 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   4920
         Width           =   6855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   4080
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tampilkan"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Batal"
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
         Left            =   7560
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cari"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox namePolis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   6855
      End
      Begin VB.TextBox noPolis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   6855
      End
      Begin VB.TextBox selesai 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox mulai 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox noBk 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   360
         TabIndex        =   16
         Top             =   5520
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Uang Pertanggungan"
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
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Lama Asuransi   ( Bulan )"
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
         Index           =   1
         Left            =   4080
         TabIndex        =   21
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label Labels 
         Caption         =   "Premi (Rp)"
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
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Cari"
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
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nama Pemegang Polis  "
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
         Left            =   480
         TabIndex        =   5
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Polis  "
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
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tanggal Selesai  "
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
         Left            =   4080
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Mulai  "
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
         Left            =   480
         TabIndex        =   2
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor BK  "
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
         Left            =   3960
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form_Pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db  As New ADODB.Connection
Dim WithEvents Rs  As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Dim querry As String
Private lngFormWidth As Long
Private lngFormHeight As Long

Private Sub Koneksi()
    Dim Constr As String
    
    Constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "/si-arsip.accdb"
    Db.Open Constr

End Sub

  
Private Sub Form_Load()
    Call Koneksi
    
    Dim Ctl As Control
    lngFormWidth = ScaleWidth
    lngFormHeight = ScaleHeight
    On Error Resume Next
    For Each Ctl In Me
        Ctl.Tag = Ctl.Left & " " & Ctl.Top & " " & _
            Ctl.Width & " " & Ctl.Height & " "
            Ctl.Tag = Ctl.Tag & Ctl.FontSize & " "
    Next Ctl
    On Error GoTo 0
    
End Sub


Private Sub Form_Resize()
    Dim D(4) As Double
    Dim i As Long
    Dim TempPoz As Long
    Dim StartPoz As Long
    Dim Ctl As Control
    Dim TempVisible As Boolean
    Dim ScaleX As Double
    Dim ScaleY As Double
    ScaleX = ScaleWidth / lngFormWidth
    ScaleY = ScaleHeight / lngFormHeight
    On Error Resume Next
    For Each Ctl In Me
        TempVisible = Ctl.Visible
        Ctl.Visible = False
        StartPoz = 1
        For i = 0 To 4
            TempPoz = InStr(StartPoz, Ctl.Tag, " ", _
                vbTextCompare)
            If TempPoz > 0 Then
                D(i) = Mid(Ctl.Tag, StartPoz, _
                    TempPoz - StartPoz)
                StartPoz = TempPoz + 1
            Else
                D(i) = 0
            End If
            Ctl.Move D(0) * ScaleX, D(1) * ScaleY, _
                D(2) * ScaleX, D(3) * ScaleY
            Ctl.Width = D(2) * ScaleX
            Ctl.Height = D(3) * ScaleY
            If ScaleX < ScaleY Then
                   'Ctl.FontSize = D(4) * ScaleX
            Else
                   'Ctl.FontSize = D(4) * ScaleY
            End If
        Next i
        Ctl.Visible = TempVisible
    Next Ctl
    On Error GoTo 0
End Sub

Sub Bersih()
    Cari.Text = ""
    noBk.Text = ""
    mulai.Text = ""
    selesai.Text = ""
    noPolis.Text = ""
    namePolis.Text = ""
    Text1.Text = ""
    uang.Text = ""
    asuransi.Text = ""
End Sub

Private Sub Command1_Click()
 Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
      
    If Cari.Text = "" Then
        MsgBox "Please Enter Cari Value", vbInformation, "Need Value"
        Cari.SetFocus
    Else
        querry = "SELECT * FROM data_nasabah WHERE no_bk = '" & Cari.Text & "' OR nama_pemegang LIKE '%" & Cari.Text & "%'"
        Rs.Open querry, Db, adOpenDynamic, adLockOptimistic
        
        If Rs.EOF Then
            MsgBox "Data tidak ditemukan!", vbInformation
            Rs.Close
        Else
            Set DataGrid1.DataSource = Rs
            DataGrid1.Columns(0).Visible = False
 
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim Form As New Form_tampilan

    Form.Text1.Text = noBk.Text
    Form.Text2.Text = mulai.Text
    Form.Text3.Text = selesai.Text
    Form.Text4.Text = noPolis.Text
    Form.Text5.Text = namePolis.Text
    Form.Text6.Text = Text1.Text
    Form.Text7.Text = asuransi.Text
    Form.Text8.Text = uang.Text

    Form.Show vbModal
End Sub

Private Sub Command5_Click()
   
    helper.Show vbModal
End Sub

Private Sub Rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Rs.RecordCount > 0 Then 'jika recordnya lebih dari sama dengan nol maka
        noBk.Text = Rs!no_bk & ""
        mulai.Text = Rs!tgl_mulai & ""
        selesai.Text = Rs!tgl_selesai & ""
        noPolis.Text = Rs!no_polis & ""
        namePolis.Text = Rs!nama_pemegang & ""
        Text1.Text = Rs!premi
        asuransi.Text = Rs!jangka_waktu
        uang.Text = Rs!uang
    End If
End Sub

Private Sub Command2_Click()
    Call Bersih
    Set DataGrid1.DataSource = Nothing
End Sub
 

Private Sub Command4_Click()
    konfirmasi = MsgBox("Yakin ingin keluar?", vbYesNo + vbInformation, "Keluar")
    If konfirmasi = vbYes Then
        Unload Me
        Form_Homescreen.Show
        Db.Close
    Else
    End If
End Sub
 

 
