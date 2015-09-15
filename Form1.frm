VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Allium"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox cmdNota 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdCalendario 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   6000
      Picture         =   "Form1.frx":029B
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdCalculadora 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   5400
      Picture         =   "Form1.frx":0F13
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdExport 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   4800
      Picture         =   "Form1.frx":1A95
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdConfig 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":2488
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdActualizar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   3600
      Picture         =   "Form1.frx":30B1
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdLogin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   3000
      Picture         =   "Form1.frx":3B4F
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdBuscar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   2400
      Picture         =   "Form1.frx":446E
      ScaleHeight     =   551.724
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdUsuarios 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1800
      Picture         =   "Form1.frx":4D7A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdSaldo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1200
      Picture         =   "Form1.frx":5854
      ScaleHeight     =   551.724
      ScaleMode       =   0  'User
      ScaleWidth      =   551.725
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdBoletas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   600
      Picture         =   "Form1.frx":639D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox cmdEmpresas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":6DD2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   0
      Width           =   480
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5760
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      _Version        =   393216
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.ListBox List1 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   -360
      X2              =   11160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      DrawMode        =   7  'Invert
      X1              =   0
      X2              =   11160
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBoletas_Click()
frmBoletas.Show
End Sub

Private Sub cmdCalculadora_Click()
frmCalc.Show
End Sub

Private Sub cmdCalendario_Click()
frmCalendario.Show
End Sub

Private Sub cmdEmpresas_Click()
frmNuevoProveedor.Show
frmNuevoProveedor.txtBuscar.SetFocus
End Sub

Private Sub Form_Load()

direccion = "C:\Users\Anthony\Desktop\Proyecto\2ºEntrega\proyecto.mdb"

conectar

Dim rs As New ADODB.Recordset
rs.ActiveConnection = DBconnect
rs.Source = "SELECT * FROM Proveedores"
rs.Open

'Esta parte recorre la tabla proveedores, mostrando todos sus elementos
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    While rs.EOF = False
        List1.AddItem rs!Nombre
        rs.MoveNext
    Wend
End If

rs.Close
DBconnect.Close

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Picture4_Click()

End Sub
