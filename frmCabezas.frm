VERSION 5.00
Begin VB.Form frmCabezas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabezas"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Agregar nueva"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.ListBox HeadList 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Arriba:"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Izquierda:"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Derecha:"
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Abajo:"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmCabezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
ReDim Preserve MisCabezas(1 To UBound(MisCabezas) + 1) As tIndiceCabeza
Call HeadList.AddItem(CStr(UBound(MisCabezas)))
HeadList.ListIndex = HeadList.ListCount - 1
End Sub

Private Sub Direccion_Change(index As Integer)
    'Prevent non numeric characters
    If Not IsNumeric(Direccion(index).Text) Then
        Direccion(index).Text = Val(Direccion(index).Text)
    End If
    
    'Prevent overflow
    If Val(Direccion(index).Text) > UBound(GrhData) Then
        Direccion(index).Text = UBound(GrhData)
    End If
    
    'Prevent underrflow
    If Val(Direccion(index).Text) < 0 Then
        Direccion(index).Text = "0"
    End If
        
    'If grh is valid, change the number of grh for the head
    If Val(Direccion(index).Text) Then
        If GrhData(Val(Direccion(index).Text)).NumFrames > 0 Then
            MisCabezas(Val(HeadList.Text)).Head(index) = Val(Direccion(index).Text)
        End If
    Else
        MisCabezas(Val(HeadList.Text)).Head(index) = 0
    End If
End Sub

Private Sub HeadList_Click()
Dim i As Byte

For i = 1 To 4
    Direccion(i).Text = MisCabezas(Val(HeadList.Text)).Head(i)
Next i
End Sub

Private Sub Form_Load()
Dim i As Long

'Fill the lists
For i = 1 To UBound(MisCabezas())
    Call HeadList.AddItem(CStr(i))
Next i

HeadList.ListIndex = 0
End Sub
