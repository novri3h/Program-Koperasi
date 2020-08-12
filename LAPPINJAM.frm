VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LAPPINJAM 
   Caption         =   "Laporan Peminjaman Dana"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
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
   ScaleHeight     =   3375
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      Height          =   345
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo6 
      Height          =   345
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox Combo7 
      Height          =   345
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Laporan Per Angota"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1750
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Laporan Harian"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1750
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Awal"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1750
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Akhir"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1750
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bulan"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1750
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tahun"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1750
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Laporan Tahun"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1750
   End
End
Attribute VB_Name = "LAPPINJAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
RSPinjam.Open "select distinct (no_anggota) from tblPinjam", Conn
Do While Not RSPinjam.EOF
    Combo1.AddItem RSPinjam!no_anggota
    RSPinjam.MoveNext
Loop

Conn.Close
Call BukaDB
RSPinjam.Open "select distinct tanggal from tblPinjam", Conn
Do While Not RSPinjam.EOF
    Combo2.AddItem Format(RSPinjam!Tanggal, "DD-MMM-YYYY")
    Combo3.AddItem Format(RSPinjam!Tanggal, "YYYY, MM, DD")
    Combo4.AddItem Format(RSPinjam!Tanggal, "YYYY, MM, DD")
    RSPinjam.MoveNext
Loop
Conn.Close


Call BukaDB
Dim RSBLN As New ADODB.Recordset
RSBLN.Open "select distinct month(tanggal) as Bulan from TBLPinjam", Conn
Do While Not RSBLN.EOF
    Combo5.AddItem RSBLN!BULAN & Space(5) & MonthName(RSBLN!BULAN)
    RSBLN.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(tanggal)  as Tahun from tblPinjam", Conn
Do While Not RSTHN.EOF
    Combo6.AddItem RSTHN!TAHUN
    Combo7.AddItem RSTHN!TAHUN
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'lap per anggota
Private Sub Combo1_Click()
    CR.SelectionFormula = "{tblPinjam.no_anggota}='" & Combo1 & "' "
    CR.ReportFileName = App.Path & "\lap Pinjaman per anggota.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

'lap harian
Private Sub Combo2_Click()
    CR.SelectionFormula = "totext({tblPinjam.tanggal})='" & CDate(Combo2) & "' "
    CR.ReportFileName = App.Path & "\lap Pinjaman per tanggal.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'lap mingguan
Private Sub Combo4_Click()
    If Combo3 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo3.SetFocus
        Exit Sub
    Else
        If Combo4 < Combo3 Or Combo3 > Combo4 Then
            MsgBox "Tanggal terbalik"
            Combo4.SetFocus
            Exit Sub
        ElseIf Combo4 = Combo3 Then
            MsgBox "pilih tanggal yang berbeda"
            Combo4.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{TBLPinjam.Tanggal} in date (" & Combo3 & ") to date (" & Combo4 & ")"
    CR.ReportFileName = App.Path & "\Lap Pinjaman per minggu.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

'lap bulanan
Private Sub Combo6_Click()
Call BukaDB
    RSPinjam.Open "select * from TBLPinjam where month(tanggal)='" & Val(Left(Combo5, 2)) & "' and year(tanggal)='" & (Combo6) & "'", Conn
    If RSPinjam.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({TBLPinjam.Tanggal})=" & Val(Left(Combo5, 2)) & " and Year({TBLPinjam.Tanggal})=" & Val(Combo6.Text)
    CR.ReportFileName = App.Path & "\Lap Pinjaman per bulan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'lap tahunan
Private Sub Combo7_Click()
    CR.SelectionFormula = "year({TBLPinjam.Tanggal})=" & Val(Combo7.Text)
    CR.ReportFileName = App.Path & "\Lap Pinjaman tahunan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub


