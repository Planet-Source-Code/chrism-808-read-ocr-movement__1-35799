VERSION 5.00
Begin VB.Form frmWorks 
   Caption         =   "About OCR Motion Reader"
   ClientHeight    =   6735
   ClientLeft      =   3255
   ClientTop       =   3030
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3570
      Left            =   120
      Picture         =   "frmWorks.frx":0000
      ScaleHeight     =   3510
      ScaleWidth      =   3645
      TabIndex        =   4
      Top             =   3120
      Width           =   3705
   End
   Begin VB.Label Label4 
      Caption         =   $"frmWorks.frx":29D5A
      Height          =   3135
      Left            =   4260
      TabIndex        =   3
      Top             =   300
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Purpose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "How it Works?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWorks.frx":29FE7
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   3975
   End
End
Attribute VB_Name = "frmWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub
