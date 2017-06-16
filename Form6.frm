VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7_pengembalian 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   8355
   ClientLeft      =   5565
   ClientTop       =   585
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tnama_mobil 
      Height          =   360
      Left            =   3240
      MaxLength       =   15
      TabIndex        =   47
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox tno_penyewaan 
      DataSource      =   "Adodc1_penyewaan"
      Height          =   360
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   46
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cno_penyewaan 
      Height          =   360
      Left            =   3600
      TabIndex        =   44
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox ttotal_bayar 
      Height          =   360
      Left            =   7320
      TabIndex        =   43
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox tcari 
      Height          =   360
      Left            =   6840
      TabIndex        =   39
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox tterlambat 
      Height          =   360
      Left            =   7320
      TabIndex        =   34
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox tlama_sewa 
      Height          =   360
      Left            =   1920
      TabIndex        =   33
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox ttgl_kembali 
      Height          =   360
      Left            =   1920
      TabIndex        =   32
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox ttgl_sewa 
      Height          =   360
      Left            =   1920
      TabIndex        =   31
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox ttgl_transaksi 
      Height          =   360
      Left            =   1920
      TabIndex        =   30
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox tnomor_mobil 
      Height          =   360
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   29
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox tnama_anggota 
      Height          =   360
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   28
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox tkode_anggota 
      Height          =   360
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   27
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox tno_pengembalian 
      DataSource      =   "Adodc1_penyewaan"
      Height          =   360
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   26
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmd_home 
      Caption         =   "Keluar"
      Height          =   360
      Left            =   8400
      TabIndex        =   10
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "Hapus"
      Height          =   360
      Left            =   2520
      TabIndex        =   9
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_ubah 
      Caption         =   "Ubah"
      Height          =   360
      Left            =   1320
      TabIndex        =   8
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "Tambah"
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   1920
      Width           =   3855
      Begin VB.TextBox tdenda 
         Height          =   360
         Left            =   1680
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox tbayar_denda 
         Height          =   360
         Left            =   1680
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox tkembalian 
         Height          =   360
         Left            =   1680
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denda         Rp."
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar Denda  Rp."
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Kembalian     Rp."
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")*"
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   5
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   150
      End
   End
   Begin VB.CommandButton cmd_perbarui 
      Caption         =   "Perbarui"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000D&
      Height          =   1455
      Index           =   1
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
      Begin VB.OptionButton o2 
         BackColor       =   &H00404040&
         Caption         =   "Tidak Ada"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton o1 
         BackColor       =   &H00404040&
         Caption         =   "Ada"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Ketersediaan"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   45
      Top             =   4920
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3836
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "no_pengembalian"
         Caption         =   "No Pengembalian"
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
         DataField       =   "no_penyewaan"
         Caption         =   "No Penyewaan"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column07 
         DataField       =   "tgl_transaksi"
         Caption         =   "Tgl Transaksi"
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
      BeginProperty Column08 
         DataField       =   "tgl_sewa"
         Caption         =   "Tgl Sewa"
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
      BeginProperty Column09 
         DataField       =   "tgl_kembali"
         Caption         =   "Tgl Kembali"
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
      BeginProperty Column10 
         DataField       =   "lama_sewa"
         Caption         =   "Lama Sewa"
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
      BeginProperty Column11 
         DataField       =   "total_bayar"
         Caption         =   "Total Bayar"
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
      BeginProperty Column12 
         DataField       =   "terlambat"
         Caption         =   "Terlambat"
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
      BeginProperty Column13 
         DataField       =   "denda"
         Caption         =   "Denda"
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
      BeginProperty Column14 
         DataField       =   "bayar_denda"
         Caption         =   "Bayar Denda"
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
      BeginProperty Column15 
         DataField       =   "kembalian"
         Caption         =   "Kembalian"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2025,071
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1289,764
         EndProperty
      EndProperty
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bayar  Rp."
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   5640
      TabIndex        =   42
      Top             =   960
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   1
      Left            =   9120
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Transaksi"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Pengembalian"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Anggota"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Berdasarkan Kode Transaksi"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   9480
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor mobil"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Kembali"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Sewa"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Sewa"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   ")* Setelah diIsi, Tekan Enter Untuk Proses Selanjutnya"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   4575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Terlambat"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari )*"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   2
      Left            =   8040
      TabIndex        =   14
      Top             =   1560
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   5520
      X2              =   5520
      Y1              =   840
      Y2              =   4320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/Hari"
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   420
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   120
      X2              =   9480
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   120
      X2              =   9480
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   120
      X2              =   9480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   120
      X2              =   9480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   8640
      Picture         =   "Form6.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   7920
      Picture         =   "Form6.frx":0BD4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   5640
      Picture         =   "Form6.frx":149E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   6360
      Picture         =   "Form6.frx":1D68
      Stretch         =   -1  'True
      Top             =   240
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   7080
      Picture         =   "Form6.frx":2632
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   " Pengembalian Mobil"
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
      TabIndex        =   21
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "Form7_pengembalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim status, x, kode As String
Dim bm  As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rspengembalian.Open "select * from pengembalian", conn
With rspengembalian
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        DataGrid1.Refresh
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rspengembalian.DataSource
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

Sub tampil_penyewaan()
rspenyewaan.Open "select * from penyewaan", conn, adOpenDynamic, adLockOptimistic
rspenyewaan.Requery
With rspenyewaan
    If .EOF And .BOF Then
        MsgBox "Transaksi Penyewaan Tidak ada", vbOKOnly + vbCritical, "Perhatian"
        Exit Sub
    Else
        cno_penyewaan.Clear
        Do Until .EOF
            cno_penyewaan.AddItem ![no_penyewaan] + " - " + ![nama_anggota]
            .MoveNext
        Loop
            .MoveFirst
    End If
End With
End Sub

Private Sub cno_penyewaan_Click()
koneksi
x = Left(cno_penyewaan, 6)
With rspenyewaan
    .Open "select * from penyewaan " & " where no_penyewaan='" & x & "'", conn, adOpenDynamic, adLockOptimistic
    If Not .EOF Then
    tno_penyewaan = !no_penyewaan
    tkode_anggota = !kode_anggota
    tnama_anggota = !nama_anggota
    tnomor_mobil = !nomor_mobil
    tnama_mobil = !nama_mobil
    If (!status) = o1.Caption Then
        o1.Value = True
    ElseIf (!status) = o2.Caption Then
        o2.Value = True
    End If
    ttgl_transaksi = !tgl_transaksi
    ttgl_sewa = !tgl_sewa
    ttgl_kembali = !tgl_kembali
    tlama_sewa = !lama_sewa
    ttotal_bayar = !total_bayar
    tno_penyewaan.Enabled = False
    tkode_anggota.Enabled = False
    tnama_anggota.Enabled = False
    tnomor_mobil.Enabled = False
    tnama_mobil.Enabled = False
    o1.SetFocus
    o1.Enabled = True
    MsgBox "Status Berubah Menjadi Ada..", vbInformation, "Informasi"
    ttgl_transaksi.Enabled = False
    ttgl_sewa.Enabled = False
    ttgl_kembali.Enabled = False
    tlama_sewa.Enabled = False
    ttotal_bayar.Enabled = False
    tterlambat.SetFocus
    End If
End With
End Sub

Private Sub cnomor_mobil_Click()
'tampil_mobil
Dim x As String
p = Val(10 Or 9)
x = Left(cnomor_mobil.Text, 10)
rsmobil.Open "select * from mobil " & " where nomor_mobil='" & x & "'", conn, adOpenDynamic, adLockOptimistic

If Not rsmobil.EOF Then
tnomor_mobil = rsmobil!nomor_mobil
tnama_mobil = rsmobil!nama_mobil
tharga_sewa = rsmobil!harga_sewa
If (rsmobil!status) = o1.Caption Then
    o1.Value = True
ElseIf (rsmobil!status) = o2.Caption Then
    o1.Value = True
End If
ttgl_sewa.SetFocus
MsgBox "Silahkan Rubah Digit Kedua Untuk Tanggal Sewa Dan Tanggal Kembali..!!", vbInformation, "Informasi"
End If
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

Sub no_pengembalianOtomatis()
koneksi
rspengembalian.Open "select * from pengembalian", conn, adOpenDynamic, adLockOptimistic
rspengembalian.Requery

If rspengembalian.EOF Then
    tno_pengembalian.Text = "TRK" + "001"
Else
    rspengembalian.MoveLast
    kode = Right(rspengembalian!no_pengembalian, 3) + 1
    If Len(kode) = 1 Then
        tno_pengembalian.Text = "TRK00" & kode
    ElseIf Len(kode) = 2 Then
        tno_penyewaan.Text = "TRK0" & kode
    ElseIf Len(kode) = 3 Then
        tno_penyewaan.Text = "TRK" & kode
    Else
        MsgBox "Silahkan Lakukan Backup Database Dan Kosongkan Tabel Pengembalian Atau Hapus Transaksi Pengembalian...", vbCritical, "Informasi"
        nonaktif
        End
    End If
End If
rspengembalian.Close
End Sub

Sub tampilkandata()
With DataGrid1
    tno_pengembalian = .Columns(0)
    tno_penyewaan = .Columns(1)
    tkode_anggota = .Columns(2)
    tnama_anggota = .Columns(3)
    tnomor_mobil = .Columns(4)
    tnama_mobil = .Columns(5)
    If .Columns(6) = o1.Caption Then
        o1.Value = True
    ElseIf .Columns(6) = o2.Caption Then
        o2.Value = True
    End If
    ttgl_transaksi = .Columns(7)
    ttgl_sewa = .Columns(8)
    ttgl_kembali = .Columns(9)
    tlama_sewa = .Columns(10)
    ttotal_bayar = .Columns(11)
    tterlambat = .Columns(12)
    tdenda = .Columns(13)
    tbayar_denda = .Columns(14)
    tkembalian = .Columns(15)
End With
cmd_perbarui.Enabled = False
cmd_tambah.Enabled = True
cmd_tambah.Caption = "Batal"
cmd_ubah.Enabled = True
cmd_ubah.Caption = "Ubah"
cmd_hapus.Enabled = True
End Sub

Private Sub tcari_Change()
If Len(tcari) > 0 Then
    rspengembalian.Filter = "no_pengembalian like '" & tcari & "%" & "'"
Else
    form_activate
End If
End Sub

Private Sub cmd_cari_Click()
rspengembalian.Find "no_pengembalian ='" & tcari & "'"
If rspengembalian.EOF Then
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
    tampil_penyewaan
    tno_pengembalian.Enabled = False
    cno_penyewaan.SetFocus
    cmd_perbarui.Enabled = True
    cmd_simpan.Enabled = False
    cmd_tambah.Enabled = False
    cmd_hapus.Enabled = False
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
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" & vbCrLf & "" & "No Pengembalian : " & tno_pengembalian, vbYesNo + vbInformation, "INFORMASI")
If Tanya = vbYes Then
    SQLHapus = "delete from pengembalian where " & " no_pengembalian = '" & tno_pengembalian & " '"
    conn.Execute SQLHapus, , adCmdText
    rspengembalian.Requery
    bersih
    nonaktif
    cmd_ubah.Enabled = False
    cmd_hapus.Enabled = False
    cmd_perbarui.Enabled = False
    cmd_simpan.Enabled = False
    form_activate
    cmd_tambah.Caption = "Tambah"
Else
    form_activate
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
Form7_pengembalian.Hide
End Sub

Private Sub cmd_perbarui_Click()
koneksi
If o1.Value = True Then
    status = o1.Caption
Else
    status = o2.Caption
End If
SQLPerbarui1 = ""
SQLPerbarui2 = ""
SQLPerbarui1 = "update pengembalian set no_pengembalian = '" & tno_pengembalian & "', no_penyewaan='" & tno_penyewaan & "', kode_anggota= '" & tkode_anggota & "', nama_anggota= '" & tnama_anggota & "', nomor_mobil= '" & tnomor_mobil & "', nama_mobil= '" & tnama_mobil & "', status='" & status & "', tgl_transaksi='" & ttgl_transaksi & "', tgl_sewa='" & ttgl_sewa & "', tgl_kembali='" & ttgl_kembali & "', lama_sewa='" & tlama_sewa & "', total_bayar='" & ttotal_bayar & "', terlambat= '" & tterlambat & "', denda='" & tdenda & "', bayar_denda='" & tbayar_denda & "', kembalian='" & tkembalian & "' where no_pengembalian='" & tno_pengembalian & "'"
SQLPerbarui2 = "update mobil set status= '" & status & "' where nomor_mobil= '" & tnomor_mobil & "'"
SQLPerbarui3 = "update penyewaan set status= '" & status & "' where no_penyewaan= '" & tno_penyewaan & "'"
conn.Execute SQLPerbarui1
conn.Execute SQLPerbarui2
conn.Execute SQLPerbarui2
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
If tterlambat.Text = "" Or tdenda.Text = "" Or tbayar_denda.Text = "" Or tkembalian.Text = "" Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    If o1.Value = True Then
        status = o1.Caption
    Else
        status = o2.Caption
    End If
    SQLTambah = ""
    SQLPerbarui1 = ""
    SQLPerbarui2 = ""
    SQLTambah = "insert into pengembalian (no_pengembalian, no_penyewaan, kode_anggota, nama_anggota, nomor_mobil, nama_mobil, status, tgl_transaksi, tgl_sewa, tgl_kembali, lama_sewa, total_bayar, terlambat, denda, bayar_denda, kembalian) values ('" & tno_pengembalian & "', '" & tno_penyewaan & "', '" & tkode_anggota & "', '" & tnama_anggota & "', '" & tnomor_mobil & "', '" & tnama_mobil & "', '" & status & "', '" & ttgl_transaksi & "', '" & ttgl_sewa & "', '" & ttgl_kembali & "', '" & tlama_sewa & "', '" & ttotal_bayar & "', '" & tterlambat & "', '" & tdenda & "', '" & tbayar_denda & "', '" & tkembalian & "')"
    SQLPerbarui1 = "update penyewaan set status= '" & status & "' where no_penyewaan= '" & tno_penyewaan & "'"
    SQLPerbarui2 = "update mobil set status= '" & status & "' where nomor_mobil= '" & tnomor_mobil & "'"
    conn.Execute SQLTambah
    conn.Execute SQLPerbarui1
    conn.Execute SQLPerbarui2
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
    koneksi
    tampil_penyewaan
    cmd_simpan.Enabled = True
    cmd_perbarui.Enabled = False
    aktif
    bersih
    no_pengembalianOtomatis
    tno_pengembalian.Enabled = False
    cno_penyewaan.SetFocus
    DataGrid1.Enabled = False
    tcari.Enabled = False
    cmd_cari.Enabled = False
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

Private Sub DataGrid1_Click()
    tampilkandata
    cmd_cari.Enabled = False
    nonaktif
End Sub

Private Sub tterlambat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    tdenda.Text = Val(tterlambat) * (Val(ttotal_bayar) * 10 / 100)
    tbayar_denda.SetFocus
End If
End Sub

Private Sub tbayar_denda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(tbayar_denda) < Val(tdenda) Then
        MsgBox "Uang Bayar Kurang..", vbCritical, "Informasi"
        tbayar_denda.Text = ""
        tbayar_denda.SetFocus
        Exit Sub
    Else
        tkembalian.Text = Val(tbayar_denda) - Val(tdenda)
        cmd_simpan.SetFocus
    End If
End If
End Sub

