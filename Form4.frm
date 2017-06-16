VERSION 5.00
Begin VB.Form Form5_login 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5085
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Out"
      Height          =   855
      Left            =   4320
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&IN"
      Height          =   855
      Left            =   3600
      Picture         =   "Form4.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox tpass 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox tuser 
      DataSource      =   "Adodc1_login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "--] WELCOME TO LOGIN [--"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name Password"
         BeginProperty Font 
            Name            =   "Snap ITC"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1680
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form5_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If tuser.Text = "" Then
    MsgBox "Username Masih Kosong!", vbCritical + vbOKOnly, "Error"
    tuser.SetFocus
ElseIf tpass.Text = "" Then
    MsgBox "Password masih Kosong", vbCritical + vbOKOnly, "Error"
    tpass.SetFocus
Else
    koneksi
    rslogin.Open " select * from login " & " where username = '" & tuser & "' " & " and pass = '" & tpass & " ' ", conn
    If rslogin.EOF Then
        MsgBox "Login Salah!", vbCritical + vbOKOnly, "Error"
        tuser.Text = ""
        tpass.Text = ""
        tuser.SetFocus
    ElseIf rslogin!UserName = "ADMIN" Then
        Form1_home.Show
        Unload Me
    Else
        With Form1_home
            .Show
            .cmd_pengaturan.Enabled = False
            Unload Me
        End With
    End If
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Anda Akan Keluar Dari Program?", vbYesNo + vbInformation, "Konfirmasi") = vbYes Then End
End Sub

Private Sub Form_Load()
koneksi
End Sub

Private Sub tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub tuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    tpass.SetFocus
End If
End Sub
