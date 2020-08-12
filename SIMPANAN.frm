VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SIMPANAN 
   Caption         =   "Data Tabungan"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox FolderFoto 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   720
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   8475
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SIMPANAN.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "No_Simpan"
         Caption         =   "No_Simpan"
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
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
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
         DataField       =   "No_Anggota"
         Caption         =   "No_Anggota"
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
         DataField       =   "JmlSimpan"
         Caption         =   "JmlSimpan"
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
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6480
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox JmlSimpan 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1500
   End
   Begin VB.ComboBox CBOAgt 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4755
      Left            =   5640
      ScaleHeight     =   4755
      ScaleWidth      =   3555
      TabIndex        =   13
      Top             =   120
      Width           =   3555
      Begin MSComDlg.CommonDialog Cdlg1 
         Left            =   1560
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Foto"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label Saldo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   14
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label Nama 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Simpanan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Anggota"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "SIMPANAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FolderFoto_Change()
Picture1.Picture = LoadPicture(FolderFoto)
End Sub

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBKoperasi.mdb"
Adodc1.RecordSource = "TBLSimpan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call NoSimpan
Tanggal = Format(Date, "DD-MMM-YYYY")
End Sub

'prosedur untuk membuat nomor pinjam otomatis dengan pola SPNYYMMDD999
Private Sub NoSimpan()
Call BukaDB
RSSimpan.Open "select * from TBLSIMPAN Where NO_SIMPAN In(Select Max(NO_SIMPAN)From TBLSIMPAN)Order By NO_SIMPAN Desc", Conn
RSSimpan.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSSimpan
        If .EOF Then
            Urutan = "SPN" + Format(Date, "yymmdd") + "001"
            Nomor = Urutan
        Else
            If Mid(!No_Simpan, 4, 6) <> Format(Date, "yymmdd") Then
                Urutan = "SPN" + Format(Date, "yymmdd") + "001"
            Else
                Hitung = Right(!No_Simpan, 9) + 1
                Urutan = "SPN" + Format(Date, "yymmdd") + Right("000" & Hitung, 3)
            End If
        End If
        Nomor = Urutan
    End With
End Sub

'pada saat form dipanggil, tampilkan nomor anggota di dalam combo
Private Sub Form_Load()
Call BukaDB
RSAnggota.Open "select * from tblanggota", Conn
CBOAgt.Clear
Do While Not RSAnggota.EOF
    CBOAgt.AddItem RSAnggota!no_anggota
    RSAnggota.MoveNext
Loop
CBOAgt.Enabled = False
JmlSimpan.Enabled = False
End Sub

Sub KondisiAwal()
CBOAgt = ""
Nama = ""
Saldo = ""
JmlSimpan = ""
Picture1.Picture = LoadPicture()
CBOAgt.Enabled = False
JmlSimpan.Enabled = False
CmdInput.Caption = "&Input"
CmdTutup.Caption = "&Tutup"
End Sub

'pada saat nomor anggota dipilih dalam combo, tampilkan nama anggota, saldo dan fotonya
Private Sub CBOAgt_Click()
Call BukaDB
RSAnggota.Open "select * from tblanggota where no_anggota='" & CBOAgt & "'", Conn
If RSAnggota.EOF Then
    MsgBox "Nomor anggota tidak terdaftar"
    CBOAgt.SetFocus
    Exit Sub
Else
    Nama = RSAnggota!Nama
    Saldo = Format(RSAnggota!Saldo, "##,###,###")
    FolderFoto = RSAnggota!lokasi
'    Picture1.Picture = LoadPicture(RSAnggota!foto)
End If
End Sub

'pada saat nomor anggota dipilih dalam combo, tampilkan nama anggota, saldo dan fotonya
Private Sub CBOAgt_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSAnggota.Open "select * from tblanggota where no_anggota='" & CBOAgt & "'", Conn
    If RSAnggota.EOF Then
        MsgBox "Nomor anggota tidak terdaftar"
        CBOAgt.SetFocus
        Exit Sub
    Else
        Nama = RSAnggota!Nama
        Saldo = Format(RSAnggota!Saldo, "##,###,###")
'        Picture1.Picture = LoadPicture(RSAnggota!foto)
        JmlSimpan.SetFocus
    End If
End If
End Sub

Private Sub CmdInput_Click()
If CmdInput.Caption = "&Input" Then
    CmdInput.Caption = "&Simpan"
    CmdTutup.Caption = "&Batal"
    CBOAgt.Enabled = True
    JmlSimpan.Enabled = True
    CBOAgt.SetFocus
    Exit Sub
Else
    If CBOAgt = "" Or JmlSimpan = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        'simpan data ke tabel pinjam
        Dim simpan As String
        simpan = "Insert into tblsimpan (no_simpan,tanggal,no_anggota,jmlsimpan,KODEKSR) values " & _
        "('" & Nomor & "','" & CDate(Tanggal) & "','" & CBOAgt & "','" & JmlSimpan & "','" & MENU.StatusBar1.Panels(1) & "')"
        Conn.Execute simpan
        
        'tambah saldo anggota dengan simpanan yang baru saja dilakukan
        Call BukaDB
        RSAnggota.Open "select * from tblanggota where no_anggota='" & CBOAgt & "'", Conn
        If Not RSAnggota.EOF Then
            Dim edit As String
            edit = "update tblanggota set saldo= '" & RSAnggota!Saldo + JmlSimpan & "' where no_anggota='" & CBOAgt & "'"
            Conn.Execute edit
            Call KondisiAwal
            Form_Activate
        End If
    End If
End If
End Sub

Private Sub CmdTutup_Click()
If CmdTutup.Caption = "&Tutup" Then
    Unload Me
ElseIf CmdTutup.Caption = "&Batal" Then
    CBOAgt = ""
    Call KondisiAwal
    Form_Activate
End If
End Sub

Private Sub JmlSimpan_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    JmlSimpan = Format(JmlSimpan, "##,###,###")
    CmdInput.SetFocus
End If
End Sub
