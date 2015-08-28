VERSION 5.00
Begin VB.Form frmCalc 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton n1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton n2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton n3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton n4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton n5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton n6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton n7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton n8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton n9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MaskColor       =   &H8000000A&
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton lblEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton n0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdResta 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSuma 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdProd 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Limpiar 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'VARABLES GLOBALES

Dim Operacion As String
Dim A As Double 'Numero que el usuario va a ingrasar
Dim B As Double '1:Suma / 2:Resta / 3:Multip / 4:Division
Dim Resultado As Double

Private Sub cmdDiv_Click() 'divicion

A = txtOp.Text
txtOp.Text = ""
Operacion = "/"
 
 End Sub

Private Sub cmdProd_Click() 'multiplicacion

A = txtOp.Text
txtOp.Text = ""
Operacion = "*"

End Sub

Private Sub cmdResta_Click() 'resta

A = txtOp.Text
txtOp.Text = ""
Operacion = "-"

 End Sub

Private Sub cmdSuma_Click() 'suma
A = txtOp.Text
txtOp.Text = ""
Operacion = "+"
 End Sub


Private Sub lblEqual_Click() 'igual
 B = txtOp.Text
 txtOp.Text = ""
 If Operacion = "+" Then
   txtOp.Text = A + B
 ElseIf Operacion = "-" Then
   txtOp.Text = A - B
 ElseIf Operacion = "*" Then
   txtOp.Text = A * B
 ElseIf Operacion = "/" Then
   txtOp.Text = A / B
 End If

End Sub

Private Sub Limpiar_Click() 'limpieza de pantalla
 txtOp.Text = ""
End Sub

'NUMEROS QUE EL USUARIO INGRESA(1,2,3,4,5,6,7,8,9,0)
Private Sub n0_Click()
 txtOp.Text = txtOp.Text & "0"
End Sub

Private Sub n1_Click()
 txtOp.Text = txtOp.Text & "1"
End Sub

Private Sub n2_Click()
 txtOp.Text = txtOp.Text & "2"
End Sub

Private Sub n3_Click()
 txtOp.Text = txtOp.Text & "3"
End Sub

Private Sub n4_Click()
 txtOp.Text = txtOp.Text & "4"
End Sub

Private Sub n5_Click()
 txtOp.Text = txtOp.Text & "5"
End Sub

Private Sub n6_Click()
 txtOp.Text = txtOp.Text & "6"
End Sub

Private Sub n7_Click()
 txtOp.Text = txtOp.Text & "7"
End Sub

Private Sub n8_Click()
 txtOp.Text = txtOp.Text & "8"
End Sub

Private Sub n9_Click()
 txtOp.Text = txtOp.Text & "9"
End Sub

