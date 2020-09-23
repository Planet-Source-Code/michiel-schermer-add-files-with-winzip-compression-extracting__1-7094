VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZipping 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Files To Zip"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmZipping.frx":0000
   LinkTopic       =   "frmZipping"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   2730
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDestFile 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "C:\Example.zip"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtSourceDir 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "C:\ZipTest"
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start to add files to zip"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WinZip for Visual Basic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Destination file (files will be packed in this file)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Source dir (this directory will be packed)"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmZipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program is made by a M. Schermer from the Netherlands
' I'm a prof. programmer for a lot of compagny's in the Netherlands and was in America
' a short time ago to program for a compagny that was making radars and that kind of stuff
' So you know from where you got this code

' Look in a couple of days on www.Planet-Source-Code.com for an update of this version.
' Sorry if there are some spelling faults in my notes or that there was left over some Dutch
' documentation

' Good luck with the code - Michiel.Schermer@Bit-ic.nl

Private Sub Command1_Click()
    With Options
        .ActionToDo = Add
        .Compression = eXtra
        .FilesToAdd = AddHiddenSystem
        .Options = Recurse_Directories
        .IfFileAlreadyExists = AlwaysOverwrite
    End With
    
    Call AddFilesToZip(txtSourceDir, txtDestFile)
End Sub

