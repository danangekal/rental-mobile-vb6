VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2_merk 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   5280
   ClientLeft      =   5565
   ClientTop       =   780
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
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
      Left            =   3960
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
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
      Left            =   1680
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
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
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
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
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox tkode_merk 
      Height          =   375
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox tmerk 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   2160
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
         DataField       =   "kode_merk"
         Caption         =   "Kode Merk"
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
         DataField       =   "merk"
         Caption         =   "Merk"
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
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2700,284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk Ubah Klik Panah Pada Kolom Tabel Yang Ditunjuk"
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
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   4905
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   120
      X2              =   5040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   120
      X2              =   3840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   120
      X2              =   3000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   120
      X2              =   3000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merk"
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
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Merk"
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
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   870
   End
End
Attribute VB_Name = "Form2_merk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim bm As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rsmerk.Open "select * from merk", conn, adOpenDynamic, adLockOptimistic
With rsmerk
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Masih Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rsmerk.DataSource
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
koneksi
cmd_tambah.Caption = "Tambah"
cmd_simpan.Enabled = False
cmd_perbarui.Enabled = False
cmd_ubah.Enabled = False
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

Sub kode_otomatis()
Dim kode As String
koneksi
rsmerk.Open "select * from merk", conn, adOpenDynamic, adLockOptimistic
rsmerk.Requery

If rsmerk.EOF Then
    tkode_merk.Text = "ME" + "001"
Else
    rsmerk.MoveLast
    kode = Right(rsmerk!kode_merk, 3) + 1
    If Len(kode) = 1 Then
        tkode_merk.Text = "ME00" & kode
    ElseIf Len(kode) = 2 Then
        tkode_merk.Text = "ME0" & kode
    ElseIf Len(kode) = 3 Then
        tkode_merk.Text = "ME" & kode
    Else
        MsgBox "Silahkan Lakukan Backup Database Dan Kosongkan Tabel Merk Atau Hapus Transaksi Pengembalian...", vbCritical, "Informasi"
        nonaktif
        End
    End If
End If
rsmerk.Close
End Sub

Sub tampilkandata()
    tkode_merk = DataGrid1.Columns(0)
    tmerk = DataGrid1.Columns(1)
    cmd_perbarui.Enabled = False
    cmd_tambah.Enabled = True
    cmd_tambah.Caption = "Batal"
    cmd_ubah.Enabled = True
    cmd_ubah.Caption = "Ubah"
End Sub

Private Sub cmd_ubah_Click()
If cmd_ubah.Caption = "Ubah" Then
    aktif
    tkode_merk.Enabled = False
    tmerk.SetFocus
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
Form2_merk.Hide
End Sub

Private Sub cmd_perbarui_Click()
koneksi
SQLPerbarui = ""
SQLPerbarui = "update merk set kode_merk= '" & tkode_merk & "', merk= '" & tmerk & "' where kode_merk='" & tkode_merk & "'"
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
If tkode_merk.Text = "" Or tmerk.Text = "" Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    SQLTambah = ""
    SQLTambah = "insert into merk (kode_merk, merk) values ('" & tkode_merk & "', '" & tmerk & "')"
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
    kode_otomatis
    tkode_merk.Enabled = False
    tmerk.SetFocus
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

Private Sub tmerk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
