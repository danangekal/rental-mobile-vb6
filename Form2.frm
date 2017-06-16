VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3_mobil 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   7845
   ClientLeft      =   5565
   ClientTop       =   585
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   ScaleHeight     =   7531.2
   ScaleMode       =   0  'User
   ScaleWidth      =   14502.13
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tkode_merk 
      Height          =   360
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   33
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmd_home 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmd_ubah 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd_perbarui 
      Caption         =   "Perbarui"
      Height          =   495
      Left            =   8520
      TabIndex        =   26
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox ckode_merk 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2040
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   9375
      _ExtentX        =   16536
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
      Caption         =   "Untuk Mengubah Atau Menghapus Data Klik Panah Pada Kolom Tabel Yang Ditunjuk"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "nomor_mobil"
         Caption         =   "Nomor Mobil"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "nama_mobil"
         Caption         =   "Nama Mobil"
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
      BeginProperty Column04 
         DataField       =   "tgl_terdaftar"
         Caption         =   "Tgl Terdaftar"
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
      BeginProperty Column05 
         DataField       =   "harga_sewa"
         Caption         =   "Harga Sewa"
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
      BeginProperty Column06 
         DataField       =   "status"
         Caption         =   "Status"
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
            ColumnWidth     =   1767,661
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1042,269
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1835,319
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2266,101
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1653,757
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1813,052
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1427,661
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000D&
      Height          =   1095
      Index           =   1
      Left            =   4920
      TabIndex        =   19
      Top             =   2520
      Width           =   2175
      Begin VB.OptionButton o2 
         BackColor       =   &H00404040&
         Caption         =   "Tidak Ada"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton o1 
         BackColor       =   &H00404040&
         Caption         =   "Ada"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Ketersediaan"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox ttgl_terdaftar 
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Nomor Mobil"
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   1440
      TabIndex        =   14
      Top             =   3000
      Width           =   2175
      Begin VB.TextBox tnomor_mobil 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox tharga_sewa 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tmerk 
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox tnama_mobil 
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox tcari 
      Height          =   360
      Left            =   6960
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")** Pilih Kode Merk"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   4
      Left            =   6240
      TabIndex        =   35
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")**"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   3
      Left            =   4200
      TabIndex        =   34
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Merk"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   1
      Left            =   7320
      Picture         =   "Form2.frx":08CA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   7200
      Picture         =   "Form2.frx":0BD4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   6480
      Picture         =   "Form2.frx":149E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   585
      Index           =   0
      Left            =   5760
      Picture         =   "Form2.frx":1D68
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   5040
      Picture         =   "Form2.frx":2632
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   4320
      Picture         =   "Form2.frx":2EFC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   8640
      Picture         =   "Form2.frx":37C6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   3600
      Picture         =   "Form2.frx":4090
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   1
      Left            =   7920
      Picture         =   "Form2.frx":495A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   4
      X1              =   181.277
      X2              =   14320.85
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   181.277
      X2              =   14320.85
      Y1              =   3571.2
      Y2              =   3571.2
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   181.277
      X2              =   14320.85
      Y1              =   1382.4
      Y2              =   1382.4
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   181.277
      X2              =   14320.85
      Y1              =   7027.2
      Y2              =   7027.2
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   7251.065
      X2              =   7251.065
      Y1              =   1382.4
      Y2              =   3571.2
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ")* Setelah Isi Nomor Kendaraan, Tekan Enter Untuk Selanjutnya"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   5895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")*"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   1
      Left            =   3720
      TabIndex        =   15
      Top             =   3360
      Width           =   150
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/ Hari"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   8040
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   6360
      TabIndex        =   11
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Sewa @"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4920
      TabIndex        =   10
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Terdaftar"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Merk"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Mobil"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Berdasarkan Nama Mobil"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   181.277
      X2              =   14320.85
      Y1              =   7027.2
      Y2              =   7027.2
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Daftar Mobil"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terlengkap Dan Terbaru FIRE     APPLE"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   9375
   End
End
Attribute VB_Name = "Form3_mobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim status As String
Dim bm As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rsmobil.Open "select * from mobil", conn
With rsmobil
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Masih Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rsmobil.DataSource
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
tcari.Enabled = True
End Sub

Sub tampil_kode()
koneksi
rsmerk.Open "select * from merk", conn, adOpenDynamic, adLockOptimistic
rsmerk.Requery
With rsmerk
    If .EOF And .BOF Then
        MsgBox "Data Merk Tidak ada", vbOKOnly + vbCritical, "Perhatian"
        Exit Sub
    Else
        ckode_merk.Clear
        Do Until .EOF
            ckode_merk.AddItem ![kode_merk] + " - " + ![merk]
            .MoveNext
        Loop
            .MoveFirst
        End If
End With
End Sub

Sub bersih()
For Each a In Me.Controls
    If TypeOf a Is TextBox Or TypeOf a Is ComboBox Then
        a.Text = ""
    End If
    o1.Value = False
    o2.Value = False
Next
End Sub

Sub aktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Or TypeOf a Is ComboBox Then
        a.Enabled = True
    End If
    o1.Enabled = True
    o2.Enabled = True
Next
End Sub

Sub nonaktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Or TypeOf a Is ComboBox Then
        a.Enabled = False
    End If
    o1.Enabled = False
    o2.Enabled = False
Next
End Sub

Sub tampilkandata()
    tnomor_mobil = DataGrid1.Columns(0)
    tkode_merk = DataGrid1.Columns(1)
    tmerk = DataGrid1.Columns(2)
    tnama_mobil = DataGrid1.Columns(3)
    ttgl_terdaftar = DataGrid1.Columns(4)
    tharga_sewa = DataGrid1.Columns(5)
    If DataGrid1.Columns(6) = o1.Caption Then
        o1.Value = True
    Else
        o2.Value = True
    End If
    
    cmd_perbarui.Enabled = False
    cmd_tambah.Enabled = True
    cmd_tambah.Caption = "Batal"
    cmd_ubah.Enabled = True
    cmd_ubah.Caption = "Ubah"
    cmd_hapus.Enabled = True
End Sub

Private Sub DataGrid1_Click()
    tampilkandata
    cmd_cari.Enabled = False
    nonaktif
End Sub

Private Sub ckode_merk_Click()
tampil_kode
Dim x As String
Dim PanjangKanan As Integer

PanjangKanan = Len(ckode_merk.Text) - 8
x = Left(ckode_merk.Text, 5)
tkode_merk.Text = x
tmerk.Text = Right(ckode_merk.Text, PanjangKanan)
tnama_mobil.SetFocus
End Sub

Private Sub tcari_Change()
If Len(tcari) > 0 Then
    rsmobil.Filter = "nama_mobil like '" & tcari & "%" & "'"
Else
    form_activate
End If
End Sub

Private Sub cmd_cari_Click()
rsmobil.Find "nama_mobil ='" & tcari & "'"
If rsmobil.BOF Then
    MsgBox "Data Tidak Ditemukan!", vbCritical, "Informasi"
    tcari.Text = ""
    tcari.SetFocus
    Exit Sub
Else
    aktif
    bersih
    tampilkandata
End If
End Sub

Private Sub cmd_ubah_Click()
If cmd_ubah.Caption = "Ubah" Then
    aktif
    tampil_kode
    tnomor_mobil.Enabled = False
    tkode_merk.SetFocus
    cmd_perbarui.Enabled = True
    cmd_simpan.Enabled = False
    cmd_tambah.Enabled = False
    cmd_hapus.Enabled = False
    tcari.Enabled = False
    cmd_cari.Enabled = False
    cmd_ubah.Caption = "Batal"
    cmd_tambah.Caption = "Tambah"
Else
    cmd_simpan.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    cmd_tambah.Enabled = True
    nonaktif
    bersih
    cmd_ubah.Caption = "Ubah"
End If
End Sub

Private Sub cmd_hapus_Click()
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" & vbCrLf & "" & "Nomor Mobil : " & tnomor_mobil + vbCrLf & "" & "Merk : " & tmerk + vbCrLf & "" & "Nama Mobil: " & tnama_mobil, vbYesNo + vbInformation, "INFORMASI")
If Tanya = vbYes Then
    SQLHapus = "delete from mobil where " & " nomor_mobil = '" & tnomor_mobil.Text & " '"
    conn.Execute SQLHapus, , adCmdText
    rsmobil.Requery
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

Private Sub cmd_home_Click()
Form_Load
bersih
nonaktif
Form1_home.Show
Form1_home.Enabled = True
Form3_mobil.Hide
End Sub

Private Sub cmd_simpan_Click()
koneksi
If tnomor_mobil.Text = "" Or tkode_merk.Text = "" Or tmerk.Text = "" Or tnama_mobil.Text = "" Or ttgl_terdaftar.Text = "" Or tharga_sewa.Text = "" Or o1.Value = False And o2.Value = False Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    If o1.Value = True Then
        status = o1.Caption
    Else
        status = o2.Caption
    End If
    SQLTambah = ""
    SQLTambah = "insert into mobil (nomor_mobil, kode_merk, merk, nama_mobil, tgl_terdaftar, harga_sewa, status) values ('" & tnomor_mobil & "', '" & tkode_merk & "', '" & tmerk & "', '" & tnama_mobil & "', '" & ttgl_terdaftar & "', '" & tharga_sewa & "', '" & status & "')"
    conn.Execute SQLTambah
    MsgBox "Data Telah Disimpan", vbOKOnly + vbInformation, "Sukses"
    form_activate
    bersih
    nonaktif
    DataGrid1.Enabled = True
    cmd_tambah.Caption = "Tambah"
    cmd_simpan.Enabled = False
End If
End Sub

Private Sub cmd_tambah_Click()
If cmd_tambah.Caption = "Tambah" Then
    koneksi
    tampil_kode
    cmd_simpan.Enabled = True
    cmd_perbarui.Enabled = False
    aktif
    bersih
    ttgl_terdaftar.Text = Date
    tnomor_mobil.SetFocus
    DataGrid1.Enabled = False
    tcari.Enabled = False
    cmd_cari.Enabled = False
    form_activate
    cmd_tambah.Caption = "Batal"
Else
    form_activate
    cmd_simpan.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    bersih
    nonaktif
    tcari.Enabled = True
    cmd_cari.Enabled = True
    DataGrid1.Enabled = True
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub cmd_perbarui_Click()
koneksi
    If o1.Value = True Then
        status = o1.Caption
    Else
        status = o2.Caption
    End If
    SQLPerbarui = ""
    SQLPerbarui = "update mobil set nomor_mobil= '" & tnomor_mobil & "', kode_merk= '" & tkode_merk & "', merk='" & tmerk & "', nama_mobil='" & tnama_mobil & "', tgl_terdaftar='" & ttgl_terdaftar & "', harga_sewa='" & tharga_sewa & "', status='" & status & "' where nomor_mobil='" & tnomor_mobil & "'"
    conn.Execute SQLPerbarui
    MsgBox "Data Telah Diperbarui", vbOKOnly + vbInformation, "Sukses"
    bersih
    nonaktif
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    cmd_tambah.Caption = "Tambah"
    cmd_tambah.Enabled = True
    cmd_perbarui.Enabled = False
    form_activate
End Sub

Private Sub tkode_merk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tkode_mobil_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tmerk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tnama_mobil_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tnomor_mobil_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    ckode_merk.SetFocus
End If
End Sub
