VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   480
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7515
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait Loading"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1590
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   3
         Left            =   6120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   5400
         Picture         =   "frmSplash.frx":08D6
         Stretch         =   -1  'True
         Top             =   600
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   600
         Index           =   2
         Left            =   4680
         Picture         =   "frmSplash.frx":11A0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   2
         Left            =   3960
         Picture         =   "frmSplash.frx":1A6A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3240
         Picture         =   "frmSplash.frx":2334
         Stretch         =   -1  'True
         Top             =   600
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   600
         Index           =   0
         Left            =   2520
         Picture         =   "frmSplash.frx":2BFE
         Stretch         =   -1  'True
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "AB"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   4320
         TabIndex        =   9
         Top             =   3120
         Width           =   225
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   3240
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   3240
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Image Image1 
         Height          =   600
         Index           =   4
         Left            =   360
         Picture         =   "frmSplash.frx":34C8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   7
         Left            =   1080
         Picture         =   "frmSplash.frx":3D92
         Stretch         =   -1  'True
         Top             =   600
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   0
         Left            =   1800
         Picture         =   "frmSplash.frx":465C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   465
         Index           =   1
         Left            =   5160
         Picture         =   "frmSplash.frx":4F26
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Danang Ekal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4680
         TabIndex        =   3
         Top             =   3300
         Width           =   2625
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company: Dekal Corp."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4680
         TabIndex        =   2
         Top             =   3600
         Width           =   2565
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   5850
         TabIndex        =   4
         Top             =   2460
         Width           =   1605
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Aplication"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   5265
         TabIndex        =   5
         Top             =   2100
         Width           =   2190
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Mobil Fire  Apple"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   825
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   7065
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Terlengkap, Termurah Dan Terbaru"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   4950
      End
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   3480
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

If Label2.Visible = True Then
Label2.Visible = False
Else
Label2.Visible = True
End If

a = a + 1
Label1.Caption = CStr(a) & "% " & "Completed"
ProgressBar1.Value = a
If a = 100 Then
Unload Me
Form5_login.Show
End If

End Sub
