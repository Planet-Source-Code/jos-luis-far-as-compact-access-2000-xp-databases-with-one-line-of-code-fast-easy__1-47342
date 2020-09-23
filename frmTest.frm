VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compact Access Database"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9000
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.vbalProgressBar pbOriginal 
      Height          =   1935
      Left            =   2280
      TabIndex        =   12
      Top             =   1380
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   3413
      Picture         =   "frmTest.frx":08CA
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":0C2E
      BarPictureMode  =   0
      BackPictureMode =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3300
      PasswordChar    =   "l"
      TabIndex        =   11
      Top             =   4260
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdCompactDatabase 
      Caption         =   "&Compact Database"
      Height          =   315
      Left            =   3660
      TabIndex        =   7
      Top             =   4620
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1020
      Width           =   6255
   End
   Begin VB.CommandButton cmdSelectFile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "...."
      Height          =   255
      Left            =   7500
      TabIndex        =   0
      Top             =   1020
      Width           =   315
   End
   Begin Proyecto1.vbalProgressBar pbFinal 
      Height          =   1935
      Left            =   5880
      TabIndex        =   13
      Top             =   1380
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   3413
      Picture         =   "frmTest.frx":0F92
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":12F6
      BarPictureMode  =   0
      BackPictureMode =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFinalSize 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5100
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblOriginalSize 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1500
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Size After Optimize (Compact)"
      Height          =   255
      Left            =   5100
      TabIndex        =   4
      Top             =   3420
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Original Size"
      Height          =   255
      Left            =   1500
      TabIndex        =   3
      Top             =   3420
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Access DataBase to Compact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   300
      Width           =   6135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1290
      TabIndex        =   8
      Top             =   330
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   5040
      TabIndex        =   9
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   1440
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSelectFile_Click()
    txtPassword.Text = ""
    txtPassword.Visible = False
    lblFinalSize = ""
    lblOriginalSize = ""
    pbFinal.Value = 0
    Text1 = OpenCommonDialog("Select Access DataBase to Optimize (Compact)", "Access Databases|*.mdb", "*.mdb")
    If Text1 <> "" Then
        lblOriginalSize = GetFileSize(Text1)
        pbFinal.Max = FileLen(Text1)
    End If
End Sub
Private Sub cmdCompactDatabase_Click()
    If Text1 = "" Then cmdSelectFile_Click
    
    CompactDB Text1, False
    pbFinal.Value = FileLen(Text1)
    lblFinalSize = GetFileSize(Text1)
    
    If Ok = False Then
        txtPassword.Visible = True
        txtPassword.SetFocus
        CompactDB Text1, False, txtPassword
        pbFinal.Value = FileLen(Text1)
        lblFinalSize = GetFileSize(Text1)
    End If

End Sub
