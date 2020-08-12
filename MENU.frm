VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENU 
   Caption         =   "Menu Utama Program Koperasi"
   ClientHeight    =   3090
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "MENU.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNANGGOTA 
         Caption         =   "ANGGOTA"
      End
   End
   Begin VB.Menu MNTRANSAKSI 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu MNSIMPAN 
         Caption         =   "SIMPAN DANA"
      End
      Begin VB.Menu MNPINJAM 
         Caption         =   "PINJAM DANA"
      End
   End
   Begin VB.Menu MNLAPORAN 
      Caption         =   "LAPORAN"
      Begin VB.Menu MNLAPANGGOTA 
         Caption         =   "DATA ANGGOTA"
      End
      Begin VB.Menu MNLAPSIMPAN 
         Caption         =   "PENYINPANAN DANA"
      End
      Begin VB.Menu MNLAPPINJAM 
         Caption         =   "PEMINJAMAN DANA"
      End
   End
   Begin VB.Menu MNKELUAR 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub MNANGGOTA_Click()
ANGGOTA.Show vbModal
End Sub

Private Sub MNKELUAR_Click()
End
End Sub

Private Sub MNLAPANGGOTA_Click()
Laporan.Show
End Sub

Private Sub MNLAPPINJAM_Click()
LAPPINJAM.Show vbModal
End Sub

Private Sub MNLAPSIMPAN_Click()
LAPSIMPAN.Show vbModal
End Sub

Private Sub MNPINJAM_Click()
PINJAMAN.Show vbModal
End Sub

Private Sub mnsimpan_Click()
SIMPANAN.Show
End Sub

Private Sub MNSQL_Click()
UjiSQL.Show vbModal
End Sub
