VERSION 5.00
Begin VB.Form frmFxs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fxs"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Agregar nuevo"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox AnimTxt 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox OffsetX 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox OffsetY 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.ListBox FxsList 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Animación:"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "HeadOffsetX:"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HeadOffsetY:"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   960
   End
End
Attribute VB_Name = "frmFxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnimTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(AnimTxt.Text) Then
        AnimTxt.Text = Val(AnimTxt.Text)
    End If
    
    'Prevent overflow
    If Val(AnimTxt.Text) > UBound(GrhData) Then
        AnimTxt.Text = UBound(GrhData)
    End If
    
    'Prevent underrflow
    If Val(AnimTxt.Text) < 0 Then
        AnimTxt.Text = "0"
    End If
        
    'If grh is valid, change the number of grh for the fx
    If Val(AnimTxt.Text) Then
        If GrhData(Val(AnimTxt.Text)).NumFrames > 0 Then
            FxData(Val(FxsList.Text)).Animacion = Val(AnimTxt.Text)
        End If
    Else
        FxData(Val(FxsList.Text)).Animacion = 0
    End If
End Sub

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
ReDim Preserve FxData(1 To UBound(FxData) + 1) As tIndiceFx
Call FxsList.AddItem(CStr(UBound(FxData)))
FxsList.ListIndex = FxsList.ListCount - 1
End Sub

Private Sub Form_Load()
Dim i As Long

'Fill the lists
For i = 1 To UBound(FxData())
    Call FxsList.AddItem(CStr(i))
Next i

FxsList.ListIndex = 0
End Sub

Private Sub FxsList_Click()
AnimTxt.Text = FxData(Val(FxsList.Text)).Animacion
OffsetX.Text = FxData(Val(FxsList.Text)).OffsetX
OffsetY.Text = FxData(Val(FxsList.Text)).OffsetY
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
    FxData(Val(FxsList.Text)).OffsetX = Val(OffsetX.Text)
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
    FxData(Val(FxsList.Text)).OffsetY = Val(OffsetY.Text)
End Sub
