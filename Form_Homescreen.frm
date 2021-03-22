VERSION 5.00
Begin VB.Form Form_Homescreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Masuk"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Login Pegawai"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Login Admin"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Masuk"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
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
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1800
         Width           =   3375
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
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Password  :"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Username  :"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form_Homescreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db  As New ADODB.Connection
Dim Rs  As New ADODB.Recordset
Dim querry As String

Private Sub Koneksi()
    Dim Constr As String
    
    Constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "/si-arsip.accdb"
    Db.Open Constr
End Sub

Private Sub Command2_Click(Index As Integer)
    Frame2.Caption = "Login Admin"
    Command1.Visible = True
    Command4.Visible = False
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command3_Click(Index As Integer)
    Frame2.Caption = "Login Pegawai"
    Command1.Visible = False
    Command4.Visible = True
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command4_Click()
        If Text1.Text = "" Then
        MsgBox "Username tidak boleh kosong. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Need Username"
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        MsgBox "Password tidak boleh kosong. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Need Password"
        Text2.SetFocus
    Else
        
        querry = "SELECT * FROM users WHERE username = '" & Text1.Text & "' And isAdmin = False "
        Rs.Open querry, Db, adOpenDynamic, adLockOptimistic

        If Rs.EOF Then
            MsgBox "User Id tidak terdaftar!", vbInformation
            Rs.Close
        Else
            If Trim(Text2.Text) = Trim(Rs.Fields("password")) Then
                Unload Me
                Form_Pegawai.Show
                
                Db.Close
            Else
                MsgBox "Password Salah!", vbInformation
                Rs.Close
            End If
            
        End If
    End If
    
End Sub

Private Sub Command5_Click()
    Db.Close
    helper.Show vbModal
End Sub

Private Sub Form_Activate()
    Command1.Default = True
    Call Koneksi
    Frame2.Caption = "Login Admin"
    Command1.Visible = True
    Command4.Visible = False
End Sub


Private Sub Command1_Click()
     
    If Text1.Text = "" Then
        MsgBox "Username tidak boleh kosong. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Need Username"
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        MsgBox "Password tidak boleh kosong. Klik menu bantuan untuk informasi aplikasi", vbInformation, "Need Password"
        Text2.SetFocus
    Else
        
        querry = "SELECT * FROM users WHERE username = '" & Text1.Text & "' And isAdmin = True "
        Rs.Open querry, Db, adOpenDynamic, adLockOptimistic

        If Rs.EOF Then
            MsgBox "User Id tidak terdaftar!", vbInformation
            Rs.Close
        Else
            If Trim(Text2.Text) = Trim(Rs.Fields("password")) Then
                Unload Me
                Form_Admin.Show
                
                Db.Close
            Else
                MsgBox "Password Salah!", vbInformation
                Rs.Close
            End If
            
        End If
    End If
 
End Sub
  
 
