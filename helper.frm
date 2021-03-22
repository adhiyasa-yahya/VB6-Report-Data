VERSION 5.00
Begin VB.Form helper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Helper"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Mengerti"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   " Login Pegawai "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   6495
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
         Index           =   2
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
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
         Index           =   2
         Left            =   6840
         TabIndex        =   30
         Top             =   3720
         Width           =   855
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
         Index           =   2
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2520
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
         Index           =   2
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1320
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
         Index           =   2
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3120
         Width           =   1695
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
         Index           =   2
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "4. Untuk menambah data harap kosongkan form dengan cara klik batal"
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
         TabIndex        =   37
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Pencarian nasabah dapat diketikan pada inputan cari "
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
         TabIndex        =   36
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"helper.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   35
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label13 
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
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         TabIndex        =   33
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Dalam pencarian nasabah dapat menggunakan No. BK atau nama nasabah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   " Login Pegawai "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   6495
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
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1920
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
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3120
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
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
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
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
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
         Index           =   1
         Left            =   6840
         TabIndex        =   15
         Top             =   3720
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
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Dalam pencarian nasabah dapat menggunakan No. BK atau nama nasabah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label12 
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
         TabIndex        =   23
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label11 
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
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Untuk mencetak data nasabah dapat mengklik tombol tampilkan lalu cetak."
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
         TabIndex        =   21
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Pencarian nasabah dapat diketikan pada inputan cari "
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
         TabIndex        =   20
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   " Login "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
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
         Index           =   0
         Left            =   6840
         TabIndex        =   6
         Top             =   3720
         Width           =   855
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1695
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Harap masukan Username dan Password dengan benar"
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
         TabIndex        =   12
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Setelah selesai harap logout atau keluar dari akun anda"
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
         TabIndex        =   11
         Top             =   960
         Width           =   5655
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
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   1335
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
         TabIndex        =   9
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Pastikan anda login sesuai user level yang anda miliki"
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
         TabIndex        =   8
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pusat Bantuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command7_Click()
    Unload Me
End Sub


