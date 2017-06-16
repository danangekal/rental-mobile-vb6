VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6_penyewaan 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Rental Mobil ""FIRE APPLE"""
   ClientHeight    =   8535
   ClientLeft      =   5565
   ClientTop       =   585
   ClientWidth     =   9600
   LinkTopic       =   "Form5"
   ScaleHeight     =   8535
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ttgl_kembali 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   47
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox ttgl_sewa 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   46
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox tnomor_mobil 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   44
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox tkode_anggota 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   42
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cnomor_mobil 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   40
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox ckode_anggota 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   39
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox tlama_sewa 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Index           =   1
      Left            =   3840
      TabIndex        =   33
      Top             =   3240
      Width           =   1815
      Begin VB.OptionButton o1 
         BackColor       =   &H00404040&
         Caption         =   "Ada"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton o2 
         BackColor       =   &H00404040&
         Caption         =   "Tidak Ada"
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Ketersediaan"
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
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1605
      End
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
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox tharga_sewa 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000D&
      Height          =   1695
      Index           =   0
      Left            =   5880
      TabIndex        =   23
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox ttotal_bayar 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox tbayar_sewa 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox tkembalian 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")*"
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
         Index           =   5
         Left            =   3240
         TabIndex        =   35
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Kembalian    Rp."
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
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar Sewa  Rp."
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
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bayar  Rp."
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
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.TextBox tnama_mobil 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   20
      Top             =   2760
      Width           =   2535
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
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   7560
      Width           =   1095
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
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   7560
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
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   7560
      Width           =   1095
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
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox tcari 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox tno_penyewaan 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox ttgl_transaksi 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
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
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox tnama_anggota 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   41
      Top             =   5400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   14
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
         DataField       =   "bayar_sewa"
         Caption         =   "Bayar Sewa"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2025,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1335,118
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
            ColumnWidth     =   1289,764
         EndProperty
      EndProperty
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " )*"
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
      Index           =   1
      Left            =   8760
      TabIndex        =   48
      Top             =   1320
      Width           =   225
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Mobil"
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
      Left            =   120
      TabIndex        =   45
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Anggota"
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
      Left            =   120
      TabIndex        =   43
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   6360
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   5640
      Picture         =   "Form5.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   4920
      Picture         =   "Form5.frx":1194
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   1
      Left            =   9120
      Picture         =   "Form5.frx":1A5E
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   7320
      Picture         =   "Form5.frx":1D68
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   8040
      Picture         =   "Form5.frx":2632
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   8760
      Picture         =   "Form5.frx":2EFC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   120
      X2              =   9480
      Y1              =   720
      Y2              =   720
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
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   120
      X2              =   9480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari )*"
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
      Index           =   3
      Left            =   8160
      TabIndex        =   32
      Top             =   1800
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   5760
      X2              =   5760
      Y1              =   720
      Y2              =   4800
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   ")* Setelah diIsi, Tekan Enter Untuk Proses Selanjutnya"
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
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   8160
      Width           =   4575
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/ Hari"
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
      Left            =   3120
      TabIndex        =   30
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Sewa @  Rp."
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
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Sewa"
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
      Left            =   5880
      TabIndex        =   24
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Sewa"
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
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Kembali"
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
      Height          =   495
      Index           =   0
      Left            =   5880
      TabIndex        =   21
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Mobil"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   " Penyewaan Mobil"
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
      TabIndex        =   18
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Berdasarkan No Penyewaan"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Anggota"
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
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "No Penyewaan"
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
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Transaksi"
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
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "Form6_penyewaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Control
Dim status, x As String
Dim bm  As Variant

Private Sub form_activate()
koneksi
conn.CursorLocation = adUseClient
rspenyewaan.Open "select * from penyewaan", conn
With rspenyewaan
    If Not (.BOF And .EOF) Then
      bm = .Bookmark
    Else
        MsgBox "Data Kosong...!", vbCritical, "Informasi"
        DataGrid1.Enabled = False
        DataGrid1.Refresh
        Exit Sub
    End If
End With
Set DataGrid1.DataSource = rspenyewaan.DataSource
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

Sub tampil_anggota()
rsanggota.Open "select * from anggota", conn, adOpenDynamic, adLockOptimistic
rsanggota.Requery
With rsanggota
    If .EOF And .BOF Then
        MsgBox "Data Anggota Tidak ada", vbOKOnly + vbCritical, "Perhatian"
        Exit Sub
    Else
        ckode_anggota.Clear
        Do Until .EOF
            ckode_anggota.AddItem ![kode_anggota] + " - " + ![nama_anggota]
            .MoveNext
        Loop
            .MoveFirst
    End If
End With
End Sub

Sub tampil_mobil()
rsmobil.Open "select * from mobil " & " where status='" & o1.Caption & "'", conn, adOpenDynamic, adLockBatchOptimistic
rsmobil.Requery
With rsmobil
    If .EOF And .BOF Then
        MsgBox "Stok Mobil Tidak ada", vbOKOnly + vbCritical, "Perhatian"
        cnomor_mobil.Clear
        Exit Sub
    Else
        cnomor_mobil.Clear
        Do Until .EOF
            cnomor_mobil.AddItem ![nomor_mobil]
            .MoveNext
        Loop
            .MoveFirst
        End If
End With
End Sub

Private Sub ckode_anggota_Click()
koneksi
x = Left(ckode_anggota, 5)
With rsanggota
    .Open "select kode_anggota, nama_anggota from anggota " & " where kode_anggota='" & x & "'", conn, adOpenDynamic, adLockOptimistic
    If Not .EOF Then
    tkode_anggota = rsanggota!kode_anggota
    tnama_anggota = rsanggota!nama_anggota
    cnomor_mobil.SetFocus
    End If
End With
End Sub

Private Sub cnomor_mobil_Click()
koneksi
rsmobil.Open "select * from mobil " & " where nomor_mobil='" & cnomor_mobil.Text & "'", conn, adOpenDynamic, adLockOptimistic

If Not rsmobil.EOF Then
tnomor_mobil = rsmobil!nomor_mobil
tnama_mobil = rsmobil!nama_mobil
tharga_sewa = rsmobil!harga_sewa
If (rsmobil!status) = o1.Caption Then
    o1.Value = True
ElseIf (rsmobil!status) = o2.Caption Then
    o1.Value = True
End If
tnomor_mobil.Enabled = False
tnama_mobil.Enabled = False
tharga_sewa.Enabled = False
ttgl_transaksi.Enabled = False
o2.SetFocus
o2.Enabled = True
MsgBox "Status Berubah Menjadi Tidak Ada.. " & vbCrLf & "" & "Silahkan Rubah Dua Digit Untuk " & vbCrLf & "" & "Tanggal Kembali..!!", vbInformation, "Informasi"
ttgl_kembali.SetFocus
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

Sub no_penyewaanOtomatis()
Dim kode As String
koneksi
rspenyewaan.Open "select * from penyewaan", conn, adOpenDynamic, adLockOptimistic
rspenyewaan.Requery

If rspenyewaan.EOF Then
    tno_penyewaan.Text = "TRS" + "001"
Else
    rspenyewaan.MoveLast
    kode = Right(rspenyewaan!no_penyewaan, 3) + 1
    If Len(kode) = 1 Then
        tno_penyewaan.Text = "TRS00" & kode
    ElseIf Len(kode) = 2 Then
        tno_penyewaan.Text = "TRS0" & kode
    ElseIf Len(kode) = 3 Then
        tno_penyewaan.Text = "TRS0" & kode
    Else
        MsgBox "Silahkan Lakukan Backup Database Dan Kosongkan Tabel Penyewaan Atau Hapus Transaksi Pengembalian...", vbCritical, "Informasi"
        nonaktif
        End
    End If
End If
rspenyewaan.Close
End Sub

Sub tampilkandata()
With DataGrid1
    tno_penyewaan = .Columns(0)
    tkode_anggota = .Columns(1)
    tnama_anggota = .Columns(2)
    tnomor_mobil = .Columns(3)
    tnama_mobil = .Columns(4)
    tharga_sewa = .Columns(5)
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
    tbayar_sewa = .Columns(12)
    tkembalian = .Columns(13)
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
    rspenyewaan.Filter = "no_penyewaan like '" & tcari & "%" & "'"
Else
    form_activate
End If
End Sub

Private Sub cmd_cari_Click()
rspenyewaan.Find "no_penyewaan ='" & tcari & "'"
If rspenyewaan.EOF Then
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
    tampil_anggota
    tampil_mobil
    tkode_anggota.Enabled = False
    tnama_anggota.SetFocus
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
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" & vbCrLf & "" & "No Penyewaan : " & tno_penyewaan, vbYesNo + vbInformation, "INFORMASI")
If Tanya = vbYes Then
    SQLHapus = "delete from penyewaan where " & " no_penyewaan = '" & tno_penyewaan & " '"
    conn.Execute SQLHapus, , adCmdText
    rspenyewaan.Requery
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
Form6_penyewaan.Hide
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
SQLPerbarui1 = "update penyewaan set kode_anggota= '" & tkode_anggota & "', nama_anggota= '" & tnama_anggota & "', nomor_mobil= '" & tnomor_mobil & "', nama_mobil= '" & tnama_mobil & "', harga_sewa= '" & tharga_sewa & "', status='" & status & "', tgl_transaksi='" & ttgl_transaksi & "', tgl_sewa='" & ttgl_sewa & "', tgl_kembali='" & ttgl_kembali & "', lama_sewa='" & tlama_sewa & "', total_bayar='" & ttotal_bayar & "', kembalian='" & tkembalian & "' where no_penyewaan='" & tno_penyewaan & "'"
SQLPerbarui2 = "update mobil set status= '" & status & "' where nomor_mobil= '" & tnomor_mobil & "'"
conn.Execute SQLPerbarui1
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
If tkode_anggota.Text = "" Or tnama_anggota.Text = "" Or tnomor_mobil.Text = "" Or tnama_mobil.Text = "" Or tharga_sewa.Text = "" Or ttgl_transaksi.Text = "" Or ttgl_sewa.Text = "" Or ttgl_kembali.Text = "" Or tlama_sewa.Text = "" Or ttgl_sewa.Text = "" Or ttotal_bayar.Text = "" Or tbayar_sewa.Text = "" Or tkembalian.Text = "" Or o1.Value = False And o2.Value = False Then
    MsgBox "Data Belum Terisi Lengkap..!!", vbOKOnly, "INFORMASI"
    Exit Sub
Else
    If o1.Value = True Then
        status = o1.Caption
    Else
        status = o2.Caption
    End If
    SQLTambah = ""
    SQLPerbarui = ""
    SQLTambah = "insert into penyewaan (no_penyewaan, kode_anggota, nama_anggota, nomor_mobil, nama_mobil, harga_sewa, status, tgl_transaksi, tgl_sewa, tgl_kembali, lama_sewa, total_bayar, bayar_sewa, kembalian) values ('" & tno_penyewaan & "', '" & tkode_anggota & "', '" & tnama_anggota & "', '" & tnomor_mobil & "', '" & tnama_mobil & "', '" & tharga_sewa & "', '" & status & "', '" & ttgl_transaksi & "', '" & ttgl_sewa & "', '" & ttgl_kembali & "', '" & tlama_sewa & "', '" & ttotal_bayar & "', '" & tbayar_sewa & "', '" & tkembalian & "')"
    SQLPerbarui = "update mobil set status= '" & status & "' where nomor_mobil= '" & tnomor_mobil & "'"
    conn.Execute SQLTambah
    conn.Execute SQLPerbarui
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
    tampil_anggota
    tampil_mobil
    cmd_simpan.Enabled = True
    cmd_perbarui.Enabled = False
    aktif
    bersih
    ttgl_transaksi.Text = Date
    no_penyewaanOtomatis
    tno_penyewaan.Enabled = False
    ckode_anggota.SetFocus
    ttgl_sewa.Text = (Left(Date, 1)) & (Left(Date, 2) + 1) & (Right(Date, 8))
    ttgl_kembali.Text = (Left(Date, 2) + 2) & (Right(Date, 8))
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
    DataGrid1.Refresh
    cmd_tambah.Caption = "Tambah"
End If
End Sub

Private Sub DataGrid1_Click()
    tampilkandata
    cmd_cari.Enabled = False
    nonaktif
End Sub

Private Sub tlama_sewa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ttotal_bayar.Text = Val(tharga_sewa.Text) * (tlama_sewa.Text)
    tbayar_sewa.SetFocus
End If
End Sub

Private Sub tbayar_sewa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(tbayar_sewa) < Val(ttotal_bayar) Then
        MsgBox "Uang Bayar Kurang..", vbCritical, "Informasi"
        tbayar_sewa.Text = ""
        tbayar_sewa.SetFocus
        Exit Sub
    Else
        tkembalian.Text = Val(tbayar_sewa.Text) - Val(ttotal_bayar.Text)
        cmd_simpan.SetFocus
    End If
End If
End Sub

Private Sub ttgl_kembali_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    tlama_sewa.Text = Val(Left(ttgl_kembali, 2) + 2) - Val(Left(ttgl_sewa, 2) + 2)
    tlama_sewa.SetFocus
End If
End Sub

