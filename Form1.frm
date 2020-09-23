VERSION 5.00
Begin VB.Form frmZipping 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Files To Zip"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "frmZipping"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "C:\Example.zip"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "C:\ZipTest"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add files to zip"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Destination file (files will be packed in this file)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Source dir (this directory will be packed)"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmZipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
