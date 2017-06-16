VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form0_user 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   5475
   ClientLeft      =   5565
   ClientTop       =   780
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tpass 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox tuser 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmd_perbarui 
      Caption         =   "Perbarui"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmd_ubah 
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmd_home 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form0_user.frx":0000
      Height          =   2295
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "username"
         Caption         =   "Username"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "pass"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2700,284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   705
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   240
      X2              =   2640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   240
      X2              =   2640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   240
      X2              =   3960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   240
      X2              =   5040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk Ubah Dan Hapus Klik Panah Pada Kolom Tabel Yang Ditunjuk"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   4905
   End
End
Attribute VB_Name = "Form0_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim bm As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rslogin.Open "select * from login", conn
With rslogin
'rsmerk.Open "select * from merk", conn
'With rsmerk
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Masih Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rslogin.DataSource
'Set DataGrid1.DataSource = rsmerk.DataSource
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
koneksi
cmd_tambah.Caption = "Tambah"
cmd_simpan.Enabled = False
cmd_perbarui.Enabled = False
cmd_ubah.Enabled = False
cmd_hapus.Enabled = False
nonaktif
End Sub

Sub bersih()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Text = ""
    End If
Next
End Sub

Sub aktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Enabled = True
    End If
Next
End Sub

Sub nonaktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Enabled = False
    End If
Next
End Sub

Sub tampilkandata()
    tuser = DataGrid1.Columns(0)
    tpass = DataGrid1.Columns(1)
    cmd_perbarui.Enabled = False
    cmd_tambah.Enabled = True
    cmd_tambah.Caption = "Batal"
    cmd_ubah.Enabled = True
    cmd_ubah.Caption = "Ubah"
    cmd_hapus.Enabled = True
End Sub

Private Sub cmd_hapus_Click()
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" & vbCrLf & "" & "Username: " & tuser, vbYesNo + vbInformation, "INFORMASI")
If Tanya = vbYes Then
    SQLHapus = "delete from login where " & " username = '" & tuser & " '"
    conn.Execute SQLHapus, adCmdText
    rslogin.Requery
    form_activate
    bersih
    nonaktif
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_simpan.Enabled = False
    cmd_tambah.Caption = "Tambah"
Else
    bersih
    nonaktif
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_simpan.Enabled = False
    cmd_tambah.Enabled = True
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub cmd_ubah_Click()
If cmd_ubah.Caption = "Ubah" Then
    aktif
    tuser.Enabled = False
    tpass.SetFocus
    cmd_perbarui.Enabled = True
    cmd_simpan.Enabled = False
    cmd_tambah.Enabled = False
    cmd_ubah.Caption = "Batal"
    cmd_tambah.Caption = "Tambah"
Else
    cmd_simpan.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_ubah.Enabled = False
    cmd_tambah.Enabled = True
    nonaktif
    bersih
    cmd_ubah.Caption = "Ubah"
End If
End Sub

Private Sub cmd_home_Click()
Form_Load
bersih
nonaktif
Form1_home.Show
Form1_home.Enabled = True
Form0_user.Hide
End Sub

Private Sub cmd_perbarui_Click()
koneksi
SQLPerbarui = ""
SQLPerbarui = "update login set username= '" & tuser & "', pass= '" & tpass & "' where username='" & tuser & "'"
conn.Execute SQLPerbarui
MsgBox "Data Telah Diperbarui", vbOKOnly + vbInformation, "Sukses"
bersih
nonaktif
cmd_ubah.Enabled = False
cmd_tambah.Caption = "Tambah"
cmd_tambah.Enabled = True
cmd_perbarui.Enabled = False
form_activate
End Sub

Private Sub cmd_simpan_Click()
koneksi
If tuser.Text = "" Or tpass.Text = "" Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    SQLTambah = ""
    SQLTambah = "insert into login (username, pass) values ('" & tuser & "', '" & tpass & "')"
    conn.Execute SQLTambah
    MsgBox "Data Telah Disimpan", vbOKOnly + vbInformation, "Sukses"
    form_activate
    bersih
    nonaktif
    DataGrid1.Enabled = True
    cmd_simpan.Enabled = False
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub cmd_tambah_Click()
If cmd_tambah.Caption = "Tambah" Then
    cmd_simpan.Enabled = True
    aktif
    bersih
    tuser.SetFocus
    DataGrid1.Enabled = False
    cmd_tambah.Caption = "Batal"
Else
    cmd_simpan.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_ubah.Enabled = False
    bersih
    nonaktif
    DataGrid1.Enabled = True
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub DataGrid1_Click()
    tampilkandata
    nonaktif
End Sub

Private Sub tuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    tpass.SetFocus
End If
End Sub
