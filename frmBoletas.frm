VERSION 5.00
Begin VB.Form frmBoletas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boletas"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtFecha 
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtUID 
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtRUC 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtImporte 
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   1215
      Left            =   5160
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtNumeroBoleta 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblUID 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblRUC 
      Caption         =   "RUC"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblImporte 
      Caption         =   "Importe"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Descripción"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblNumero 
      Caption         =   "Número de boleta"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
Adodc1.ConnectionString = "Provider= Microsoft.Jet.OLEDB.4.0;Data Source=""" & direccion & """;Persist Security Info=False"

Adodc1.RecordSource = "SELECT * FROM Compras WHERE Fecha LIKE '%" & txtBuscar.Text & "%'"

Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider= Microsoft.Jet.OLEDB.4.0;Data Source=""" & direccion & """;Persist Security Info=False"

Adodc1.RecordSource = "SELECT * FROM Compras"

Adodc1.Refresh

End Sub

