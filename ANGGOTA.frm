VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ANGGOTA 
   Caption         =   "DATA ANGGOTA"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
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
   ScaleHeight     =   4995
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6240
      ScaleHeight     =   4755
      ScaleWidth      =   3555
      TabIndex        =   17
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ANGGOTA.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3625
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
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Nama"
         Caption         =   "Nama"
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
         DataField       =   "Wajib"
         Caption         =   "Wajib"
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
         DataField       =   "Pokok"
         Caption         =   "Pokok"
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
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
         DataField       =   "Foto"
         Caption         =   "Foto"
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
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   2040
      Width           =   1845
      _ExtentX        =   3254
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
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   1850
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   1850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   1850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   400
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1850
   End
   Begin VB.TextBox FolderFoto 
      Height          =   350
      Left            =   720
      TabIndex        =   5
      Top             =   4560
      Width           =   5475
   End
   Begin VB.TextBox Saldo 
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   2000
   End
   Begin VB.TextBox SmpPokok 
      Height          =   350
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox SmpWajib 
      Height          =   350
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   2000
   End
   Begin VB.TextBox Nama 
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   2000
   End
   Begin VB.ComboBox NoAnggota 
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label Label7 
      Caption         =   " Klik Frame Foto untuk mengambil gambar"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Foto"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Simpanan Pokok"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1755
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Simpanan Wajib"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Anggota"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "ANGGOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'LblLokasi.Caption = (App.Path & "\FOTO\" & Trim(Text1.Text) & ".JPEG")

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBKoperasi.mdb"
Adodc1.RecordSource = "TBLAnggota"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Adodc1.Visible = False
End Sub

Sub Form_Load()
Nama.MaxLength = 30
SmpWajib.MaxLength = 8
SmpPokok.MaxLength = 8
Saldo.MaxLength = 8
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSAnggota.Open "Select * From TBLAnggota where No_Anggota='" & NoAnggota & "'", Conn
End Function

Private Sub KosongkanText()
    NoAnggota = ""
    Nama = ""
    SmpWajib = ""
    SmpPokok = ""
    Saldo = ""
    FolderFoto = ""
End Sub

Private Sub SiapIsi()
    'enabled = true menyebabkan objek dpt dimasuki kursor
    NoAnggota.Enabled = True
    Nama.Enabled = True
    SmpWajib.Enabled = True
    SmpPokok.Enabled = True
    Saldo.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    'enabled = false menyebabkan objek tdk dpt dimasuki kursor
    NoAnggota.Enabled = False
    Nama.Enabled = False
    SmpWajib.Enabled = False
    SmpPokok.Enabled = False
    Saldo.Enabled = False
    FolderFoto.Enabled = False
End Sub

Private Sub KondisiAwal()
    Form_Activate
    KosongkanText
    TidakSiapIsi
    Command1.Caption = "&Input"
    Command2.Caption = "&Edit"
    Command3.Caption = "&Hapus"
    Command4.Caption = "&Tutup"
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
 
End Sub

Private Sub TampilkanData()
With RSAnggota
    Nama = RSAnggota!Nama
    SmpWajib = RSAnggota!wajib
    SmpPokok = RSAnggota!Pokok
    Saldo = RSAnggota!Saldo
    FolderFoto = RSAnggota!lokasi
    'Picture1.Picture = LoadPicture(RSAnggota!foto)
End With
End Sub

Private Sub Command1_Click()
   
    If Command1.Caption = "&Input" Then
        Command1.Caption = "&Simpan"
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Caption = "&Batal"
        NoAnggota.Clear
        SiapIsi
        KosongkanText
        NoAnggota.SetFocus
    Else
        If NoAnggota = "" Or Nama = "" Or SmpWajib = "" Or SmpPokok = "" Or Saldo = "" Then
            MsgBox "Data Belum Lengkap...!"
        ElseIf FolderFoto = "" Then
            MsgBox "Belum ada foto"
            Picture1_Click
            Exit Sub
        Else
'            FolderFoto = (App.Path & "\FOTO\" & Trim(Text1.Text) & ".JPEG")
            Dim SQLTambah As String
            SQLTambah = "Insert Into TBLAnggota (No_Anggota,Nama,wajib,Pokok,Saldo,lokasi,foto) values " & _
            "('" & NoAnggota & "','" & Nama & "','" & SmpWajib & "','" & SmpPokok & "','" & Saldo & "','" & FolderFoto & "','" & Picture1 & "')"
            Conn.Execute SQLTambah
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub command2_Click()

    If Command2.Caption = "&Edit" Then
        Command1.Enabled = False
        Command2.Caption = "&Simpan"
        Command3.Enabled = False
        Command4.Caption = "&Batal"
        SiapIsi
        NoAnggota.SetFocus
        Call BukaDB
        RSAnggota.Open "TBLAnggota", Conn
        NoAnggota.Clear
        Do Until RSAnggota.EOF
            NoAnggota.AddItem RSAnggota!no_anggota
            RSAnggota.MoveNext
        Loop
    Else
        If NoAnggota = "" Or Nama = "" Or SmpWajib = "" Or SmpPokok = "" Or Saldo = "" Or FolderFoto = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update TBLAnggota Set Nama= '" & Nama & "',wajib='" & SmpWajib & "',Pokok='" & SmpPokok & "',Saldo='" & Saldo & "',lokasi='" & FolderFoto & "',foto='" & Picture1.Picture & "' where No_Anggota='" & NoAnggota & "'"
            Conn.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub command3_Click()
    If Command3.Caption = "&Hapus" Then
        Command1.Enabled = False
        Command2.Enabled = False
        Command4.Caption = "&Batal"
        KosongkanText
        SiapIsi
        NoAnggota.SetFocus
        Call BukaDB
        RSAnggota.Open "TBLAnggota", Conn
        NoAnggota.Clear
        Do Until RSAnggota.EOF
            NoAnggota.AddItem RSAnggota!no_anggota
            RSAnggota.MoveNext
        Loop
    End If
End Sub

Private Sub command4_Click()
    Select Case Command4.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub NoAnggota_Click()
Call CariData
Call TampilkanData
 If Command3.Enabled = True Then
        Call CariData
        If Not RSAnggota.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From TBLAnggota where No_Anggota= '" & NoAnggota & "'"
                Conn.Execute SQLHapus
                Form_Activate
                Call KondisiAwal
            Else
                Form_Activate
                Call KondisiAwal
                Command3.SetFocus
            End If
        Else
            MsgBox "Data Tidak ditemukan"
            NoAnggota.SetFocus
        End If
    End If
End Sub

Private Sub NoAnggota_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(NoAnggota) < 3 Or Len(NoAnggota) > 3 Then
        MsgBox "Kode Harus 3 Digit, Contoh 'A01'"
        NoAnggota.SetFocus
        Exit Sub
    Else
        Nama.SetFocus
    End If

    If Command1.Caption = "&Simpan" Then
        Call CariData
        If Not RSAnggota.EOF Then
            TampilkanData
            MsgBox "Kode Anggota Sudah Ada"
            KosongkanText
            NoAnggota.SetFocus
        Else
            Nama.SetFocus
        End If
    End If
    
    If Command2.Caption = "&Simpan" Then
        Call CariData
        If Not RSAnggota.EOF Then
            TampilkanData
            NoAnggota.Enabled = False
            Nama.SetFocus
        Else
            MsgBox "Kode Anggota Tidak Ada"
            NoAnggota = ""
            NoAnggota.SetFocus
        End If
    End If
    
'    If Command3.Enabled = True Then
'        Call CariData
'        If Not RSAnggota.EOF Then
'            TampilkanData
'            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
'            If Pesan = vbYes Then
'                Dim SQLHapus As String
'                SQLHapus = "Delete From TBLAnggota where No_Anggota= '" & NoAnggota & "'"
'                Conn.Execute SQLHapus
'                Form_Activate
'                Call KondisiAwal
'            Else
'                Form_Activate
'                Call KondisiAwal
'                Command3.SetFocus
'            End If
'        Else
'            MsgBox "Data Tidak ditemukan"
'            NoAnggota.SetFocus
'        End If
'    End If
End If
End Sub

Private Sub Picture1_Click()
Cdlg1.ShowOpen
FolderFoto = Cdlg1.FileName
'FolderFoto = Cdlg1.InitDir
'FolderFoto = UCase(Left(Cdlg1.FileName, Len(Cdlg1.FileName) - Len(Cdlg1.FileTitle)))
End Sub

Private Sub Nama_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then SmpWajib.SetFocus
End Sub

Private Sub smpwajib_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then SmpPokok.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub smpPokok_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        Saldo = Val(SmpWajib) + Val(SmpPokok)
        Saldo.Enabled = False
        If Command1.Enabled = True Then
            Command1.SetFocus
        ElseIf Command2.Enabled = True Then
            Command2.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


Private Sub FolderFoto_Change()
Picture1.Picture = LoadPicture(FolderFoto)
End Sub
