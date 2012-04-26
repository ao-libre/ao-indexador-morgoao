VERSION 5.00
Begin VB.Form frmPersonajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personajes"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar nuevo cuerpo"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox OffsetY 
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox OffsetX 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox BodiesList 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Direccion 
      Height          =   285
      Index           =   2
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
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HeadOffsetY:"
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "HeadOffsetX:"
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Abajo:"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Derecha:"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Izquierda:"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Arriba:"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   450
   End
End
Attribute VB_Name = "frmPersonajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BodiesList_Click()
Dim i As Byte

For i = 1 To 4
    Direccion(i).Text = MisCuerpos(Val(BodiesList.Text)).Body(i)
Next i

OffsetX.Text = MisCuerpos(Val(BodiesList.Text)).HeadOffsetX
OffsetY.Text = MisCuerpos(Val(BodiesList.Text)).HeadOffsetY
End Sub

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Private Sub cmdAgregar_Click()
ReDim Preserve MisCuerpos(1 To UBound(MisCuerpos) + 1) As tIndiceCuerpo

Call BodiesList.AddItem(UBound(MisCuerpos))
BodiesList.ListIndex = BodiesList.ListCount - 1
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
        
    'If grh is valid, change the number of grh for the body
    If Val(Direccion(index).Text) Then
        If GrhData(Val(Direccion(index).Text)).NumFrames > 0 Then
            MisCuerpos(Val(BodiesList.Text)).Body(index) = Val(Direccion(index).Text)
        End If
    Else
        MisCuerpos(Val(BodiesList.Text)).Body(index) = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Long

'Fill the lists
For i = 1 To UBound(MisCuerpos())
    Call BodiesList.AddItem(CStr(i))
Next i

BodiesList.ListIndex = 0
End Sub

Private Sub OffsetX_Change()
    'Prevent non numeric characters
    If Not IsNumeric(OffsetX.Text) Then
        OffsetX.Text = Val(OffsetX.Text)
    End If
    
    'Prevent overflow
    If Val(OffsetX.Text) > &H7FFF Then
        OffsetX.Text = &H7FFF
    End If
    
    'Prevent underrflow
    If Val(OffsetX.Text) < &HFFFF8000 Then
        OffsetX.Text = &HFFFF8000
    End If
    
    'Change te OffsetX
    MisCuerpos(Val(BodiesList.Text)).HeadOffsetX = Val(OffsetX.Text)
End Sub

Private Sub OffsetY_Change()
    'Prevent non numeric characters
    If Not IsNumeric(OffsetY.Text) Then
        OffsetY.Text = Val(OffsetY.Text)
    End If
    
    'Prevent overflow
    If Val(OffsetY.Text) > &H7FFF Then
        OffsetY.Text = &H7FFF
    End If
    
    'Prevent underrflow
    If Val(OffsetY.Text) < &HFFFF8000 Then
        OffsetY.Text = &HFFFF8000
    End If
    
    'Change te OffsetX
    MisCuerpos(Val(BodiesList.Text)).HeadOffsetY = Val(OffsetY.Text)
End Sub
