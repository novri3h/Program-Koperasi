VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form COBA 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog DLG1 
      Left            =   4440
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   7440
      Width           =   8415
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   240
      Width           =   8775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6840
      Width           =   1695
   End
End
Attribute VB_Name = "COBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
DLG1.ShowOpen
Text2 = DLG1.FileName
Text1 = Right(Text2, 5)
Picture1.Picture = LoadPicture(Text2)
End Sub
