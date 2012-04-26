VERSION 5.00
Begin VB.Form frmAutoIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto-Index"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Género"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   3255
      Begin VB.OptionButton Gender 
         Caption         =   "Barca"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Gender 
         Caption         =   "Mujer"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Gender 
         Caption         =   "Hombre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Raza"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton Race 
         Caption         =   "Gnomo"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Race 
         Caption         =   "Enano"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Race 
         Caption         =   "Elfo oscuro"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Race 
         Caption         =   "Humano"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Race 
         Caption         =   "Elfo"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdIndex 
      Caption         =   "Indexar"
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de indexación"
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3255
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   35
         Top             =   1300
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frames por fila"
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3015
         Begin VB.TextBox txtRow 
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   22
            Text            =   "5"
            Top             =   690
            Width           =   495
         End
         Begin VB.TextBox txtRow 
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   21
            Text            =   "5"
            Top             =   330
            Width           =   495
         End
         Begin VB.TextBox txtRow 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   20
            Text            =   "6"
            Top             =   690
            Width           =   495
         End
         Begin VB.TextBox txtRow 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Text            =   "6"
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fila 4:"
            Height          =   195
            Left            =   1680
            TabIndex        =   18
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fila 3:"
            Height          =   195
            Left            =   1680
            TabIndex        =   17
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fila 2:"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fila 1:"
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Text            =   "25"
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Text            =   "45"
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox txtFileNum 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "0"
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero de gráfico:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de indexación"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton IndexType 
         Caption         =   "Cuerpos"
         Height          =   195
         Index           =   6
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton IndexType 
         Caption         =   "FXs"
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton IndexType 
         Caption         =   "Escudos"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton IndexType 
         Caption         =   "Cascos"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton IndexType 
         Caption         =   "Cabezas"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton IndexType 
         Caption         =   "Armas"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAutoIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Function AutoIndex(ByVal FileNum As Long, ByVal Width As Integer, ByVal Height As Integer) As Integer
Const Count As Byte = 4
Dim Inicio As Long
Dim k As Long
Dim i As Long
Dim Pos As Long

If FileNum = 0 Then Exit Function

For i = 1 To UBound(GrhData)
    If GrhData(i).NumFrames = 0 Then
        k = k + 1
        If k = Count Then
            Inicio = i - Count
            Exit For
        End If
    Else
        k = 0
        Inicio = 0
    End If
Next i

If Inicio = 0 Then
    Inicio = UBound(GrhData)

    'Resize array
    ReDim Preserve GrhData(1 To Inicio + Count) As GrhData
End If

'Search where to place the grh....
For Pos = 0 To frmMain.grhList.ListCount - 1
    If Val(frmMain.grhList.List(Pos)) >= Inicio Then
        Exit For
    End If
Next Pos
    
For k = Inicio + 1 To Inicio + Count
    'Make sure he is not overwritting anything
    If k <= UBound(GrhData()) Then
        If GrhData(k).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Function
            End If
        End If
    End If
    
    Pos = Pos + 1
    
    If GrhData(k).NumFrames = 0 Then
        'Add it!
        Call frmMain.grhList.AddItem(k, Pos)
    Else
        frmMain.grhList.List(Pos) = k
    End If
    
    'Fill in grh data
    With GrhData(k)
        .FileNum = FileNum
        
        .NumFrames = 1
        ReDim .Frames(1 To .NumFrames) As Long
        
        .pixelHeight = Height
        .pixelWidth = Width
        
        .Speed = 0
        .Frames(1) = k
        
        .sX = (k - Inicio - 1) * .pixelWidth
        .sY = 0
        
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = Pos
    DoEvents
Next k

AutoIndex = Inicio + 1

End Function

Private Function AutoIndexWithAnimation(ByVal FileNum As Long, ByVal Width As Integer, ByVal Height As Integer, ByRef Rows() As Byte) As Integer
Dim Count As Byte
Dim Inicio As Long
Dim k As Long
Dim i As Long
Dim Pos As Long
Dim Cant As Byte

If FileNum = 0 Then Exit Function

For i = 1 To 4
    Count = Count + Rows(i) + 1
Next i

For i = 1 To UBound(GrhData)
    If GrhData(i).NumFrames = 0 Then
        k = k + 1
        If k = Count Then
            Inicio = i - Count
            Exit For
        End If
    Else
        k = 0
        Inicio = 0
    End If
Next i

If Inicio = 0 Then
    Inicio = UBound(GrhData)

    'Resize array
    ReDim Preserve GrhData(1 To Inicio + Count) As GrhData
End If

'Search where to place the grh....
For Pos = 0 To frmMain.grhList.ListCount - 1
    If Val(frmMain.grhList.List(Pos)) >= Inicio Then
        Exit For
    End If
Next Pos
    
For k = Inicio + 1 To Inicio + Count
    If k - Inicio <= (Count - 4) Then
        Cant = 1
    Else
        Cant = Rows((k - Inicio) - (Count - 4))
    End If
    
    'Make sure he is not overwritting anything
    If k <= UBound(GrhData()) Then
        If GrhData(k).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Function
            End If
        End If
    End If
    
    Pos = Pos + 1
    
    If GrhData(k).NumFrames = 0 Then
        'Add it!
        If Cant > 1 Then
            Call frmMain.grhList.AddItem(k & " (ANIMACIÓN)", Pos)
        Else
            Call frmMain.grhList.AddItem(k, Pos)
        End If
    Else
        If Cant > 1 Then
            frmMain.grhList.List(Pos) = k & " (ANIMACIÓN)"
        Else
            frmMain.grhList.List(Pos) = k
        End If
    End If
    
    'Fill in grh data
    With GrhData(k)
        .FileNum = FileNum
        
        .NumFrames = Cant
        ReDim .Frames(1 To .NumFrames) As Long
        
        .pixelHeight = Height
        .pixelWidth = Width
        
        Cant = 0
        
        If .NumFrames = 1 Then
            .Speed = 0
            .Frames(1) = k
            
            For i = 1 To 4
                If k - Inicio <= Cant + Rows(i) Then
                    .sX = (k - Inicio - Cant - 1) * .pixelWidth
                    .sY = .pixelHeight * (i - 1)
                    Exit For
                End If
                
                Cant = Cant + Rows(i)
            Next i
        Else
            For i = 1 To (k - Inicio) - (Count - 4) - 1
                Cant = Cant + Rows(i)
            Next i
            
            For i = 1 To .NumFrames
                .Frames(i) = Inicio + i + Cant
            Next i
            
            .Speed = .NumFrames * 1000 / 18
        End If
        
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = Pos
    DoEvents
Next k

AutoIndexWithAnimation = Inicio + Count - 3

End Function

Private Sub cmdIndex_Click()
Dim Rows(1 To 4) As Byte
Dim i As Byte
Dim Cant As Integer
Dim path As String
Dim hFile As Integer
Dim index As Long
Dim Raza As Byte
Dim Genero As Byte
Dim IniCabezas As Integer

For i = 1 To 4
    Rows(i) = Val(txtRow(i))
Next i

For Raza = Race.LBound To Race.UBound
    If Race(Raza).value Then Exit For
Next Raza

For Genero = Gender.LBound To Gender.UBound
    If Gender(Genero).value Then Exit For
Next Genero

For i = IndexType.LBound To IndexType.UBound
    If IndexType(i).value Then Exit For
Next i

path = Config.initPath
If Right$(path, 1) <> "\" Then path = path & "\"
hFile = FreeFile()

Select Case i
    Case 1  'Armas
        index = AutoIndexWithAnimation(Val(txtFileNum), Val(txtWidth), Val(txtHeight), Rows())
        path = path & "Armas.dat"
        
        Cant = GetVar(path, "INIT", "NumArmas") + 1
        Call WriteVar(path, "INIT", "NumArmas", Cant)
        
        Open path For Append Shared As #hFile
            Print #hFile, ""
            Print #hFile, "'" & txtName
            Print #hFile, "[Arma" & Cant & "]"
            Print #hFile, "Dir3=" & index
            Print #hFile, "Dir1=" & index + 1
            Print #hFile, "Dir4=" & index + 2
            Print #hFile, "Dir2=" & index + 3
        Close #hFile
    Case 2  'Cabezas
        index = AutoIndex(Val(txtFileNum), Val(txtWidth), Val(txtHeight))
        
        IniCabezas = (Raza * 100) + (69 * Genero) + 1

        Do While MisCabezas(IniCabezas).Head(1) > 0
            IniCabezas = IniCabezas + 1
            
            If IniCabezas > UBound(MisCabezas) Then
                ReDim Preserve MisCabezas(1 To IniCabezas) As tIndiceCabeza
            End If
        Loop
            
        MisCabezas(IniCabezas).Head(3) = index
        MisCabezas(IniCabezas).Head(2) = index + 1
        MisCabezas(IniCabezas).Head(4) = index + 2
        MisCabezas(IniCabezas).Head(1) = index + 3
    Case 3  'Cascos
        index = AutoIndex(Val(txtFileNum), Val(txtWidth), Val(txtHeight))
        
        IniCabezas = 3  'Iniciamos acá porque el 2 es no animation.

        Do While MisCascos(IniCabezas).Head(1) > 0
            IniCabezas = IniCabezas + 1
            
            If IniCabezas > UBound(MisCascos) Then
                ReDim Preserve MisCascos(1 To IniCabezas) As tIndiceCabeza
            End If
        Loop
        
        MisCascos(IniCabezas).Head(3) = index
        MisCascos(IniCabezas).Head(2) = index + 1
        MisCascos(IniCabezas).Head(4) = index + 2
        MisCascos(IniCabezas).Head(1) = index + 3
    Case 4  'Escudos
        index = AutoIndexWithAnimation(Val(txtFileNum), Val(txtWidth), Val(txtHeight), Rows())
        path = path & "Escudos.dat"
        
        Cant = GetVar(path, "INIT", "NumEscudos") + 1
        Call WriteVar(path, "INIT", "NumEscudos", Cant)
        
        Open path For Append Shared As #hFile
            Print #hFile, ""
            Print #hFile, "'" & txtName
            Print #hFile, "[ESC" & Cant & "]"
            Print #hFile, "Dir3=" & index
            Print #hFile, "Dir1=" & index + 1
            Print #hFile, "Dir4=" & index + 2
            Print #hFile, "Dir2=" & index + 3
        Close #hFile
    Case 5  'FXs
        'Por ahora a manopla
    Case 6  'Cuerpos
        index = AutoIndexWithAnimation(Val(txtFileNum), Val(txtWidth), Val(txtHeight), Rows())
        path = path & "Personajes.ind"
        
        ReDim Preserve MisCuerpos(1 To UBound(MisCuerpos) + 1) As tIndiceCuerpo
        
        MisCuerpos(UBound(MisCuerpos)).Body(3) = index
        MisCuerpos(UBound(MisCuerpos)).Body(1) = index + 1
        MisCuerpos(UBound(MisCuerpos)).Body(4) = index + 2
        MisCuerpos(UBound(MisCuerpos)).Body(2) = index + 3
        
        If Genero <> 2 Then
            MisCuerpos(UBound(MisCuerpos)).HeadOffsetY = IIf(Raza = 3 Or Raza = 4, 6, -4)
        End If
End Select

End Sub

Private Sub IndexType_Click(index As Integer)
Frame4.Enabled = (index = 2 Or index = 6)
Frame5.Enabled = (index = 2 Or index = 6)
End Sub
