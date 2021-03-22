VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Admin 
   Caption         =   "Form_Admin"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      TabIndex        =   29
      Top             =   4320
      Width           =   3495
   End
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
      TabIndex        =   27
      Top             =   5160
      Width           =   7215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      TabIndex        =   26
      Text            =   "Bulan"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton help 
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
      Index           =   1
      Left            =   10680
      TabIndex        =   23
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command7 
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
      Height          =   495
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   22
      Top             =   5040
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMM yyyy"
      Format          =   127729667
      CurrentDate     =   43928
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMM yyyy"
      Format          =   127729667
      CurrentDate     =   43928
   End
   Begin VB.TextBox noBk 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADMIN"
      ClipControls    =   0   'False
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   360
         TabIndex        =   21
         Top             =   5880
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4683
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
      Begin VB.CommandButton Command6 
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Hapus"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "Tambah Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Edit Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3360
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
         Left            =   8160
         TabIndex        =   13
         Top             =   4920
         Width           =   855
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
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
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   7095
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
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   7095
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
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   3495
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
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   4800
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
         TabIndex        =   25
         Top             =   3960
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
         Left            =   360
         TabIndex        =   24
         Top             =   3960
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
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nama Pemegang Polis :"
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
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Polis :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tanggal Selesai :"
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
         TabIndex        =   9
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Mulai :"
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
         Left            =   360
         TabIndex        =   8
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor BK :"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.OLE OLE1 
      Height          =   1095
      Left            =   4200
      TabIndex        =   20
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Form_Admin"
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
 


Private Sub Combo1_Click()
    uang.Text = Combo1.Text * Text1.Text
End Sub
 

Private Sub Command5_Click()
    If noBk.Text = "" Then
        MsgBox "data belum lengkap, pilih data yang akan dihapus. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        noBk.SetFocus
    Else
        konfirmasi = MsgBox("Yakin ingin dihapus?", vbYesNo + vbInformation, "Keluar")
        If konfirmasi = vbYes Then
            Rs.Delete
            Call Bersih
        Else
        End If
    End If
End Sub

Private Sub Form_Load()
    Call Koneksi
    
    Text1.Text = 0
    
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
    
    Dim i As Long
    With Me.Combo1
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
  
    
End Sub
 
 
 
 

Private Sub Form_Resize()
    Dim D(4) As Double
    Dim E(4) As Double
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
            
            If ScaleWidth > 15295 Then
                Combo1.Left = 4000 + (ScaleWidth - 4200) / 4
                
            Else
                Combo1.Left = 4200
                
            End If
            
            If ScaleHeight > 9945 Then
                Combo1.Top = 4320 + D(2) * ScaleY
            Else
                Combo1.Top = 4320
            End If
 
             Debug.Print ScaleHeight
            
         
            If ScaleX < ScaleY Then
                   'Ctl.FontSize = D(4) * ScaleX
                   Combo1.FontSize = Int((Screen.Width * Combo1.FontSize) / ScaleX)
            Else
                   'Ctl.FontSize = D(4) * ScaleY
                   Combo1.FontSize = Int((Screen.Width * Combo1.FontSize) / ScaleY)
            End If
        Next i
        Ctl.Visible = TempVisible
    Next Ctl
    On Error GoTo 0
   
    
End Sub
 

Private Sub Koneksi()
    Dim Constr As String
    
    Constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "/si-arsip.accdb"
    Db.Open Constr

End Sub

Sub Bersih()
    Cari.Text = ""
    noBk.Text = ""
    noPolis.Text = ""
    namePolis.Text = ""
    Text1.Text = ""
    uang.Text = ""
    Combo1.Clear
End Sub
 
Private Sub Command1_Click()
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
      
    If Cari.Text = "" Then
        MsgBox "Isi data Nama Nasabah atau No BK pada fill cari. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Need Value"
        Cari.SetFocus
    Else
        querry = "SELECT * FROM data_nasabah WHERE no_bk = '" & Cari.Text & "' OR nama_pemegang LIKE '%" & Cari.Text & "%'"
        Rs.Open querry, Db, adOpenDynamic, adLockOptimistic
        
        If Rs.EOF Then
            MsgBox "Data tidak ditemukan!. Klik menu bantuan untuk informasi aplikasi", vbInformation
            Rs.Close
        Else
            Set DataGrid1.DataSource = Rs
            DataGrid1.Columns(0).Visible = False
            
        End If
         
    End If
     
End Sub



Private Sub help_Click(Index As Integer)
 
    helper.Show vbModal
End Sub

  

Private Sub Rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Rs.RecordCount > 0 Then 'jika recordnya lebih dari sama dengan nol maka
        noBk.Text = Rs!no_bk & ""
        DTPicker1.Value = Rs!tgl_mulai & ""
        DTPicker2.Value = Rs!tgl_selesai & ""
        noPolis.Text = Rs!no_polis & ""
        namePolis.Text = Rs!nama_pemegang & ""
        Text1.Text = Rs!premi
        Combo1.Text = Rs!jangka_waktu
        uang.Text = Rs!uang
    End If
End Sub
    
Private Sub Command2_Click()
    Call Bersih
    Set DataGrid1.DataSource = Nothing
End Sub

Sub Cekdata()
    If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" Then
        MsgBox ("Data belum lengkap!")
    End If
End Sub
 
Private Sub Command3_Click()
    If noBk.Text = "" Then
        MsgBox "data belum lengkap, pilih data yang akan diedit. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        noBk.SetFocus
    Else
        konfirmasi = MsgBox("Yakin ingin mengedit data?", vbYesNo + vbInformation, "Keluar")
        If konfirmasi = vbYes Then
            Rs.Update
            Rs!no_bk = noBk.Text
            Rs!tgl_mulai = DTPicker1.Value
            Rs!tgl_selesai = DTPicker2.Value
            Rs!no_polis = noPolis.Text
            Rs!nama_pemegang = namePolis.Text
            Rs!premi = Text1.Text
            Rs!jangka_waktu = Combo1.Text
            Rs!uang = uang.Text
            Rs.Update
            notifikasi = MsgBox("Data berhasil diupdate", vbInformation, "Notifikasi")
        End If
    End If
End Sub
 
Private Sub Command4_Click()
    Set Rs = New ADODB.Recordset
    querry = "SELECT * FROM data_nasabah"
    Rs.Open querry, Db, adOpenDynamic, adLockOptimistic
    
    If noBk.Text = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        noBk.SetFocus
    ElseIf DTPicker1.Value = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        DTPicker1.SetFocus
    ElseIf DTPicker2.Value = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        DTPicker2.SetFocus
    ElseIf noPolis.Text = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        noPolis.SetFocus
    ElseIf namePolis.Text = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        namePolis.SetFocus
    ElseIf Text1.Text = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        premi.SetFocus
    ElseIf Combo1.Text = "" Then
        MsgBox "data belum lengkap. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        premi.SetFocus
    Else
        konfirmasi = MsgBox("Yakin ingin menyimpan data?", vbYesNo + vbInformation, "Keluar")
        If konfirmasi = vbYes Then
            Rs.AddNew
            Rs!no_bk = noBk.Text
            Rs!tgl_mulai = DTPicker1.Value
            Rs!tgl_selesai = DTPicker2.Value
            Rs!no_polis = noPolis.Text
            Rs!nama_pemegang = namePolis.Text
            Rs!premi = Text1.Text
            Rs!jangka_waktu = Combo1.Text
            Rs!uang = uang.Text
            Rs.Update
            notifikasi = MsgBox("Data berhasil diupdate", vbInformation, "Notifikasi")
            Call Bersih
            
        End If
    End If
End Sub

Private Sub Command6_Click()
     If noBk.Text = "" Then
        MsgBox "Pilih data yang akan ditampilkan. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Peringatan!!"
        noBk.SetFocus
    Else
        Dim Form As New Form_tampilan
    
        Form.Text1.Text = noBk.Text
        Form.Text2.Text = DTPicker1.Value
        Form.Text3.Text = DTPicker2.Value
        Form.Text4.Text = noPolis.Text
        Form.Text5.Text = namePolis.Text
        Form.Text6.Text = Text1.Text
        Form.Text7.Text = Combo1.Text
        Form.Text8.Text = uang.Text
        
        Form.Show vbModal
    End If
End Sub

Private Sub Command7_Click()
    konfirmasi = MsgBox("Yakin ingin keluar?", vbYesNo + vbInformation, "Keluar")
    If konfirmasi = vbYes Then
        Unload Me
        Form_Homescreen.Show
        Db.Close
    Else
    End If
End Sub
  




 

