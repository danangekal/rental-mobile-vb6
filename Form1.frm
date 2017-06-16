VERSION 5.00
Begin VB.Form Form1_home 
   BackColor       =   &H00404040&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   9585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   10245
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tugas Mata Kuliah Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   9720
      Width           =   3780
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ".:Rental Mobil:."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   705
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3060
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRE APPLE"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   80.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Danang Eko Alfianto"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   855
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4080
      Picture         =   "Form1.frx":1194
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ".:Rental Mobil:."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Index           =   5
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3060
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   10680
      Left            =   0
      Picture         =   "Form1.frx":149E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5400
   End
   Begin VB.Menu cmd_mobil 
      Caption         =   "MOBIL"
      Begin VB.Menu cmd_merk 
         Caption         =   "Merk Mobil"
      End
      Begin VB.Menu cmd_daftarmobil 
         Caption         =   "Daftar Mobil"
      End
   End
   Begin VB.Menu cmd_anggota 
      Caption         =   "ANGGOTA"
   End
   Begin VB.Menu cmd_transaksi 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu cmd_penyewaan 
         Caption         =   "PENYEWAAN"
         Shortcut        =   ^A
      End
      Begin VB.Menu cmd_Pengembalian 
         Caption         =   "PENGEMBALIAN"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu cmd_pengaturan 
      Caption         =   "PENGATURAN"
      Begin VB.Menu cmd_pengguna 
         Caption         =   "Pengguna"
      End
   End
   Begin VB.Menu cmd_keluar 
      Caption         =   "KELUAR"
      Begin VB.Menu cmd_program 
         Caption         =   "Program"
      End
      Begin VB.Menu cmd_ganti 
         Caption         =   "Ganti Pengguna"
      End
   End
End
Attribute VB_Name = "Form1_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_anggota_Click()
Form4_anggota.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_daftarmobil_Click()
Form3_mobil.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_ganti_Click()
If MsgBox("Anda Akan Ganti Pengguna?", vbYesNo + vbInformation, "Informasi") = vbYes Then
    Form5_login.Show
    Form1_home.Hide
End If
End Sub

Private Sub cmd_merk_Click()
Form2_merk.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_Pengembalian_Click()
Form7_pengembalian.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_pengguna_Click()
Form0_user.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_penyewaan_Click()
Form6_penyewaan.Show
Form1_home.Enabled = False
End Sub

Private Sub cmd_program_Click()
If MsgBox("Anda Akan Keluar Program?", vbYesNo + vbInformation, "Informasi") = vbYes Then
    End
End If
End Sub

Private Sub Form_Load()
Label7.Caption = Date
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Time
End Sub

