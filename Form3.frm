VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4_anggota 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   7485
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   7485
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ttgl_terdaftar 
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker ttgl_lahir 
      Bindings        =   "Form3.frx":0000
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107806721
      CurrentDate     =   18264
   End
   Begin VB.CommandButton cmd_perbarui 
      Caption         =   "Perbarui"
      Height          =   375
      Left            =   8520
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2880
      Width           =   975
   End
   Begin VB.OptionButton o2 
      BackColor       =   &H00404040&
      Caption         =   "Perempuan"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton o1 
      BackColor       =   &H00404040&
      Caption         =   "Laki-laki"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox tphone 
      Height          =   375
      Left            =   6480
      MaxLength       =   13
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox talamat 
      Height          =   1455
      Left            =   6480
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox tnama_anggota 
      Height          =   375
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   15
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox tkode_anggota 
      Height          =   375
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_ubah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_home 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   8520
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox tcari 
      Height          =   360
      Left            =   5640
      TabIndex        =   1
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   8040
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   19
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
         DataField       =   "kode_anggota"
         Caption         =   "Kode Anggota"
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
         DataField       =   "nama_anggota"
         Caption         =   "Nama Anggota"
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
         DataField       =   "tgl_lahir"
         Caption         =   "Tgl Lahir"
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
         DataField       =   "kelamin"
         Caption         =   "Kelamin"
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
         DataField       =   "alamat"
         Caption         =   "Alamat"
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
         DataField       =   "phone"
         Caption         =   "Phone"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1830,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2025,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   3
      Left            =   5400
      Picture         =   "Form3.frx":000B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")* Klik Tombol Cari Untuk Menampilkan Hasil Pencarian"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3480
      TabIndex        =   27
      Top             =   6840
      Width           =   4470
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ")*"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   9120
      TabIndex        =   26
      Top             =   4200
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   4440
      Picture         =   "Form3.frx":0315
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   5160
      Picture         =   "Form3.frx":0BDF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   5880
      Picture         =   "Form3.frx":14A9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   6960
      Picture         =   "Form3.frx":1D73
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   7680
      Picture         =   "Form3.frx":263D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   8400
      Picture         =   "Form3.frx":2F07
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   120
      X2              =   9480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   120
      X2              =   9480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   120
      X2              =   9480
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   120
      X2              =   9480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Terdaftar"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4920
      TabIndex        =   21
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4920
      TabIndex        =   12
      Top             =   3120
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Lahir"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4920
      TabIndex        =   10
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Anggota"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Anggota"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   " Daftar Anggota"
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
      TabIndex        =   7
      Top             =   120
      Width           =   9375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   4680
      X2              =   4680
      Y1              =   1320
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9480
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Berdasarkan Nama Anggota"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rental Mobil FIRE         APPLE"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1440
      TabIndex        =   28
      Top             =   720
      Width           =   6150
   End
End
Attribute VB_Name = "Form4_anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim kelamin As String
Dim bm As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rsanggota.Open "select * from anggota", conn
With rsanggota
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Masih Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rsanggota.DataSource
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

Sub bersih()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Text = ""
    End If
    o1.Value = False
    o2.Value = False
Next
End Sub

Sub aktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Enabled = True
    End If
    o1.Enabled = True
    o2.Enabled = True
    ttgl_lahir.Enabled = True
Next
End Sub

Sub nonaktif()
For Each a In Me.Controls
    If TypeOf a Is TextBox Then
        a.Enabled = False
    End If
    o1.Enabled = False
    o2.Enabled = False
    ttgl_lahir.Enabled = False
Next
End Sub

Sub kode_anggotaOtomatis()
Dim kode As String
koneksi
rsanggota.Open "select * from anggota", conn, adOpenDynamic, adLockOptimistic
rsanggota.Requery

If rsanggota.EOF Then
    tkode_anggota.Text = "FA" + "001"
Else
    rsanggota.MoveLast
    kode = Right(rsanggota!kode_anggota, 3) + 1
    If Len(kode) = 1 Then
        tkode_anggota.Text = "FA00" & kode
    ElseIf Len(kode) = 2 Then
        tkode_anggota.Text = "FA0" & kode
    ElseIf Len(kode) = 3 Then
        tkode_anggota.Text = "FA" & kode
    
    Else
        MsgBox "Silahkan Lakukan Backup Database Dan Kosongkan Tabel Anggota Atau Hapus Transaksi Pengembalian...", vbCritical, "Informasi"
        nonaktif
        End
    End If
End If
rsanggota.Close
End Sub

Sub tampilkandata()
'If Not rsanggota.EOF Then
    tkode_anggota = rsanggota!kode_anggota
    tnama_anggota = rsanggota!nama_anggota
    ttgl_lahir.Value = rsanggota!tgl_lahir
    If rsanggota!kelamin = o1.Caption Then
        o1.Value = True
    Else
        o2.Value = True
    End If
    talamat = rsanggota!alamat
    tphone = rsanggota!phone
    ttgl_terdaftar = rsanggota!tgl_terdaftar
    cmd_perbarui.Enabled = False
    cmd_tambah.Enabled = True
    cmd_tambah.Caption = "Batal"
    cmd_ubah.Enabled = True
    cmd_ubah.Caption = "Ubah"
    cmd_hapus.Enabled = True
'End If
End Sub

Private Sub tcari_Change()
If Len(tcari) > 0 Then
    rsanggota.Filter = "nama_anggota like '" & tcari & "%" & "'"
Else
    form_activate
End If
End Sub

Private Sub cmd_cari_Click()
rsanggota.Find "nama_anggota ='" & tcari.Text & "'"
If rsanggota.EOF Then
    MsgBox "Data Tidak Ditemukan!", vbCritical, "Informasi"
    tcari.Text = ""
    tcari.SetFocus
    Exit Sub
Else
    tampilkandata
End If
End Sub

Private Sub cmd_ubah_Click()
If cmd_ubah.Caption = "Ubah" Then
    aktif
    tkode_anggota.Enabled = False
    tnama_anggota.SetFocus
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
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" & vbCrLf & "" & "Kode Anggota : " & tkode_anggota + vbCrLf & "" & "Nama Anggota : " & tnama_anggota + vbCrLf & "" & "Tanggal Lahir: " & ttgl_lahir, vbYesNo + vbInformation, "INFORMASI")
If Tanya = vbYes Then
    SQLHapus = "delete from anggota where " & " kode_anggota = '" & tkode_anggota & " '"
    conn.Execute SQLHapus, , adCmdText
    rsanggota.Requery
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
Form4_anggota.Hide
End Sub

Private Sub cmd_perbarui_Click()
koneksi
If o1.Value = True Then
    kelamin = o1.Caption
Else
    kelamin = o2.Caption
End If
SQLPerbarui = ""
SQLPerbarui = "update anggota set kode_anggota= '" & tkode_anggota & "', nama_anggota= '" & tnama_anggota & "', tgl_lahir='" & ttgl_lahir & "', kelamin='" & kelamin & "', alamat='" & talamat & "', phone='" & tphone & "', tgl_terdaftar='" & ttgl_terdaftar & "' where kode_anggota='" & tkode_anggota & "'"
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

Private Sub cmd_simpan_Click()
koneksi
If tkode_anggota.Text = "" Or tnama_anggota.Text = "" Or talamat.Text = "" Or tphone.Text = "" Or ttgl_terdaftar.Text = "" Or o1.Value = False And o2.Value = False Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    If o1.Value = True Then
        kelamin = o1.Caption
    Else
        kelamin = o2.Caption
    End If
    SQLTambah = ""
    SQLTambah = "insert into anggota (kode_anggota, nama_anggota, tgl_lahir, kelamin, alamat, phone, tgl_terdaftar) values ('" & tkode_anggota & "', '" & tnama_anggota & "', '" & ttgl_lahir & "', '" & kelamin & "', '" & talamat & "', '" & tphone & "', '" & ttgl_terdaftar & "')"
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
    cmd_perbarui.Enabled = False
    aktif
    bersih
    ttgl_terdaftar.Text = Date
    kode_anggotaOtomatis
    tkode_anggota.Enabled = False
    tnama_anggota.SetFocus
    tcari.Enabled = False
    cmd_cari.Enabled = False
    form_activate
    DataGrid1.Enabled = False
    cmd_tambah.Caption = "Batal"
Else
    form_activate
    cmd_simpan.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    tcari.Enabled = True
    cmd_cari.Enabled = True
    bersih
    nonaktif
    DataGrid1.Enabled = True
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub DataGrid1_Click()
    tampilkandata
    tcari.Enabled = False
    cmd_cari.Enabled = False
    nonaktif
End Sub



