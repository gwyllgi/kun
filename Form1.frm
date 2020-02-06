VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   600
      TabIndex        =   23
      Top             =   5400
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16776960
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TUTUP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   22
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UBAH"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   19
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0013
      TabIndex        =   17
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":003E
      Left            =   2640
      List            =   "Form1.frx":0048
      TabIndex        =   16
      Top             =   2760
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108003329
      CurrentDate     =   43866
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108003329
      CurrentDate     =   43866
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   2640
      TabIndex        =   13
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   2640
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2640
      TabIndex        =   9
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Aktif_Sampai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "NoHp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Agama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis_Kelamin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Tanggal_Lahir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Tempat_Lahir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "NIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim koneksi As New ADODB.Connection
Dim koneksirecord As New ADODB.Recordset

Private Sub Command1_Click()
Set koneksirecord = New ADODB.Recordset
simpan = "insert into kunana value('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & DTPicker1 & "','" & Combo1 & "','" & Combo2 & "','" & Text4 & "','" & Text5 & "','" & DTPicker2 & "')"
koneksi.Execute simpan
MsgBox "Data Berhasil Disimpan"
Call KondisiAwal

End Sub

Private Sub Command2_Click()
Set koneksirecord = New ADODB.Recordset
Delete = "delete from kunana where NIS='" & Text1 & "'"
koneksi.Execute Delete
MsgBox "Data Berhasil Dihapus"
Call KondisiAwal

End Sub

Private Sub Command3_Click()
Set koneksirecord = New ADODB.Recordset
Update = "update kunana set Nama='" & Text2 & "',Tempat_Lahir='" & Text3 & "',Tanggal_Lahir='" & DTPicker1 & "',Jenis_Kelamin='" & Combo1 & "',agama='" & Combo2 & "',NoHp='" & Text4 & "',Aktif_Sampai='" & DTPicker2 & "' where NIS='" & Text1 & "'"
koneksi.Execute Update
MsgBox "Data Berhasil Diubah"
Call KondisiAwal

End Sub

Private Sub Command5_Click()
End
End Sub
Private Sub bukakoneksi()
Set koneksi = Nothing
Set koneksi = New ADODB.Connection
koneksi.Open "DSN=kun"
End Sub
Sub KondisiAwal()
Call bukakoneksi
Set koneksirecord = New ADODB.Recordset
koneksirecord.CursorLocation = adUseClient
koneksirecord.Open "select NIS,Nama,Tempat_Lahir,Tanggal_Lahir,Jenis_Kelamin,Agama,Alamat,NoHp,Aktif_Sampai from kunana order by nis asc", koneksi, adOpenDynamic, adLockBatchOptimistic
Set DataGrid1.DataSource = koneksirecord.DataSource
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Combo1.AddItem "Laki-Laki"
Combo1.AddItem "Perempuan"
Combo2.AddItem "ISLAM"
Combo2.AddItem "KRISTEN"
Combo2.AddItem "KATOLIK"
Combo2.AddItem "HINDU"
Combo2.AddItem "BUDHA"

Call bukakoneksi
Set koneksirecord = Nothing
Set koneksirecord = New ADODB.Recordset
koneksirecord.CursorLocation = adUseClient
koneksirecord.Open "select NIS,Nama,Tempat_Lahir,Tanggal_Lahir,Jenis_Kelamin,Agama,Alamat,NoHp,Aktif_Sampai from kunana order by NIS asc", koneksi, adOpenDynamic, adLockBatchOptimistic

Set DataGrid1.DataSource = koneksirecord.DataSource
DataGrid1.Refresh

End Sub

