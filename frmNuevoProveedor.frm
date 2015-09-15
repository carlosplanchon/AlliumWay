VERSION 5.00
Begin VB.Form frmNuevoProveedor 
   Caption         =   "Proveedores"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtBuscar 
      Height          =   405
      Left            =   1560
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtContacto 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtRUC 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtTelefono 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtDireccion 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Contacto"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblRUC 
      Caption         =   "RUC"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblTelefono 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblDireccion 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmNuevoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBorrar_Click()

txtBuscar.Text = ""
txtNombre.Text = ""
txtDireccion.Text = ""
txtTelefono.Text = ""
txtRUC.Text = ""
txtContacto.Text = ""

cmdAgregar.Enabled = True

End Sub

Private Sub cmdModificar_Click()

conectar

DBconnect.Execute " UPDATE Proveedores SET Nombre = '" & txtNombre.Text & "', Direccion = '" & txtDireccion.Text & "', Telefono = '" & txtTelefono.Text & "', RUC = '" & txtRUC.Text & "', Contacto = '" & txtContacto.Text & "' WHERE Nombre = '" & txtBuscar.Text & "' "

DBconnect.Close

End Sub

Private Sub cmdAgregar_Click()

If txtNombre.Text = "" And txtRUC.Text = "" Then

    MsgBox "Completar los campos Nombre y RUC obligatoriamente", vbCritical, "Error!"

Else

    conectar
    
    Dim rs As New ADODB.Recordset
    rs.ActiveConnection = DBconnect
    rs.Source = "SELECT RUC FROM Proveedores WHERE RUC = '" & txtRUC.Text & "' "
    
    
    
    rs.Open
    
    If (rs.BOF And rs.EOF) Then
        DBconnect.Execute " INSERT INTO Proveedores (Nombre, Direccion, Telefono, RUC, Contacto ) VALUES ( '" & txtNombre.Text & "','" & txtDireccion.Text & "','" & txtTelefono.Text & "','" & txtRUC.Text & "','" & txtContacto.Text & "');"
        MsgBox "Proveedor agregado correctamente"
    
        
        Else
        MsgBox "Proveedor existente", vbCritical, "Error"
        
    End If
    rs.Close
    DBconnect.Close

End If


End Sub

Private Sub cmdBuscar_Click()

Dim rs As New ADODB.Recordset

conectar

rs.ActiveConnection = DBconnect
rs.Source = "SELECT * FROM Proveedores WHERE Nombre = '" & txtBuscar.Text & "' "

rs.Open

If (rs.BOF And rs.EOF) Then
    
   MsgBox "No existe proveedor", vbCritical, "Error"

    
    Else
   
        txtNombre.Text = rs!Nombre
        txtDireccion.Text = rs!direccion
        txtTelefono.Text = rs!telefono
        txtRUC.Text = rs!RUC
        txtContacto.Text = rs!Contacto
        cmdModificar.Enabled = True
End If

rs.Close
DBconnect.Close


If txtBuscar.Text = "" And txtNombre.Text = "" And txtDireccion.Text = "" And txtTelefono.Text = "" And txtRUC.Text = "" And txtContacto.Text = "" Then

cmdModificar.Enabled = False

End If

cmdAgregar.Enabled = False


End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
mayus KeyAscii
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
mayus KeyAscii
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
mayus KeyAscii
End Sub

