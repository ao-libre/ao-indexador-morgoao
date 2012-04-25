VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexador Alkon"
   ClientHeight    =   9690
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   646
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   736
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer animation 
      Enabled         =   0   'False
      Left            =   240
      Top             =   240
   End
   Begin VB.ListBox grhList 
      Height          =   4935
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame AnimationFrame 
      Caption         =   "Animation"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   8880
      Width           =   8295
      Begin VB.TextBox FrameTxt 
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   ">"
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "<"
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox FramesTxt 
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox DurationTxt 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label FrameLbl 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4245
         TabIndex        =   31
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Frame:"
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Frames:"
         Height          =   195
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Duración:"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame grhFrame 
      Caption         =   "Grh"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   8160
      Width           =   8295
      Begin VB.TextBox grhXTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhYTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhHeightTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhWidthTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox bmpTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   5040
         TabIndex        =   22
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bmp:"
         Height          =   195
         Left            =   6960
         TabIndex        =   20
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.CheckBox grhOnly 
      Caption         =   "Mostrar solamente el Grh"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zoom"
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
      Begin VB.CommandButton ZoomOut 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton ZoomIn 
         Caption         =   "+"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox ZoomTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   285
         Width           =   120
      End
   End
   Begin VB.HScrollBar picScrollH 
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      TabIndex        =   4
      Top             =   7800
      Width           =   7695
   End
   Begin VB.VScrollBar picScrollV 
      Height          =   7695
      LargeChange     =   10
      Left            =   10680
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox fileList 
      Height          =   2205
      ItemData        =   "frmMain.frx":0004
      Left            =   120
      List            =   "frmMain.frx":0006
      TabIndex        =   2
      Top             =   5760
      Width           =   2655
   End
   Begin VB.PictureBox previewer 
      AutoRedraw      =   -1  'True
      Height          =   7680
      Left            =   3000
      ScaleHeight     =   508
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   0
      Top             =   120
      Width           =   7680
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save grhs"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveOldMnu 
         Caption         =   "Save grhs in &old format"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Save grhs in &new format"
      End
      Begin VB.Menu SaveInAllMnu 
         Caption         =   "Save grhs in all the files"
      End
      Begin VB.Menu SaveHeadsMnu 
         Caption         =   "Save Heads"
      End
      Begin VB.Menu SaveHelmetsMnu 
         Caption         =   "Save Helmets"
      End
      Begin VB.Menu SaveBodiesMnu 
         Caption         =   "Save Bodies"
      End
      Begin VB.Menu SaveFXsMnu 
         Caption         =   "Save FXs"
      End
   End
   Begin VB.Menu GrhMnu 
      Caption         =   "&Grh"
      Begin VB.Menu AddGrhMnu 
         Caption         =   "&Agregar Grh..."
         Shortcut        =   ^N
      End
      Begin VB.Menu RemoveGrhMnu 
         Caption         =   "&Remover Grh"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu OpenMnu 
      Caption         =   "&Open"
      Begin VB.Menu HeadsMnu 
         Caption         =   "Heads(Cabezas.ind)"
      End
      Begin VB.Menu HelmetsMnu 
         Caption         =   "Helmets(Cascos.ind)"
      End
      Begin VB.Menu BodiesMnu 
         Caption         =   "Bodies(Personajes.ind)"
      End
      Begin VB.Menu FXsMnu 
         Caption         =   "Fxs(Fxs.ind)"
      End
      Begin VB.Menu OpenNormalMnu 
         Caption         =   "Árboles normales(Graficos3.ind)"
      End
      Begin VB.Menu OpenTransparentMnu 
         Caption         =   "Árboles transparentes(Graficos2.ind)"
      End
      Begin VB.Menu OpenSmallMnu 
         Caption         =   "Árboles chicos(Graficos1.ind)"
      End
   End
   Begin VB.Menu UtilidadesMnu 
      Caption         =   "&Utilidades"
      Begin VB.Menu AutoIndexarMnu 
         Caption         =   "&Auto-Indexar"
      End
      Begin VB.Menu SearchBMPMnu 
         Caption         =   "Buscar BMP"
      End
      Begin VB.Menu SearchErroresMnu 
         Caption         =   "Buscar errores"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Default zoom, 100%
Private Const DEFAULT_ZOOM As Integer = 100

''
' Maximum zoom possible, 10 times bigger.
Private Const MAX_ZOOM As Integer = DEFAULT_ZOOM * 10

''
' Minimum zoom possible, 10 times smaller.
Private Const MIN_ZOOM As Integer = DEFAULT_ZOOM / 10

''
' Step by which zoom is altered.
Private Const ZOOM_STEP As Integer = 10

''
' Means no grh is being rendered.
Private Const NO_GRH As Long = -1


''
' Defines the different points of the selection box that are being edited.
'
' @param    sbpeNone            No coord is being modified.
' @param    sbpeStartX          Starting x coord is being modified.
' @param    sbpeStartY          Starting y coord is being modified.
' @param    sbpeEndX            Ending x coord is being modified.
' @param    sbpeEndY            Ending y coord is being modified.
' @param    sbpeStartXStartY    Starting x coord and starting y coord are being modified.
' @param    sbpeEndXEndY        Ending x coord and ending y coord are being modified.
' @param    sbpeStartXEndY      Starting x coord and ending y coord are being modified.
' @param    sbpeEndXStartY      Ending x coord and starting y coord are being modified.

Private Enum eSelectionBoxPointEdition
    sbpeNone
    sbpeStartX
    sbpeStartY
    sbpeEndX
    sbpeEndY
    sbpeStartXStartY
    sbpeEndXEndY
    sbpeStartXEndY
    sbpeEndXStartY
End Enum

Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type

Private Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits      As Long
End Type

''
' The current zoom, 1 == 100%
Private zoom As Single

''
'Currently loaded picture. Used to render avoiding to reload everytime zoom or scroll happens.
Private currentPic As StdPicture

''
' X coord where a selection started.
Private selectionAreaStartX As Single

''
' Y coord where a selection started.
Private selectionAreaStartY As Single

''
' X coord where a selection ended.
Private selectionAreaEndX As Single

''
' Y coord where a selection ended.
Private selectionAreaEndY As Single

''
' Cord currently being edited.
Private editionCoord As eSelectionBoxPointEdition

''
' The grh currently being displayed
Private currentGrh As Long

''
' The current frame of the grh being displayed
Private currentFrame As Long

''
' Flag used to ignore calls to RenderSelectionBox.
Private ignoreSelectionBoxRender As Boolean

''
' Flag used to ignore update events to grh' data textboxes.
Private ignoreGrhTextUpdate As Boolean

Private FileHeaderBMP As BITMAPFILEHEADER
Private InfoHeaderBMP As BITMAPINFOHEADER

Private Sub AddGrhMnu_Click()
    Call frmAddGrh.Show(vbModal, Me)
End Sub

Private Sub animation_Timer()
    Dim path As String
    
    'If an animated grh is chosen, animate!
    If currentGrh <> NO_GRH Then
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            currentFrame = currentFrame + 1
            
            If currentFrame > GrhData(currentGrh).NumFrames Then
                currentFrame = 1
            End If
            
            'Load new bitmap
            If Right$(Config.bmpPath, 1) <> "\" Then
                path = Config.bmpPath & "\" & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
            Else
                path = Config.bmpPath & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
            End If
            
            'Prevent memory leaks
            Set currentPic = Nothing
            Set currentPic = LoadPicture(path)
            
            Call RedrawPicture(currentGrh, currentFrame)
        End If
    End If
End Sub

Private Sub ArmasMnu_Click()
Const Count As Long = 26
Dim Inicio As Long
Dim k As Long
Dim i As Long
Dim j As Long
Dim Cant As Byte
Dim FileNum As String

FileNum = InputBox("Ingrese el número del gráfico.")

If StrPtr(FileNum) = 0 Then Exit Sub
If FileNum = 0 Then Exit Sub

For i = 1 To UBound(GrhData)
    If GrhData(i).NumFrames = 0 Then
        k = k + 1
        If k = Count Then
            Inicio = i - (Count - 2)
            Exit For
        End If
    Else
        k = 0
        Inicio = 0
    End If
Next i

If Inicio = 0 Then
    Inicio = UBound(GrhData) + 1

    'Resize array
    ReDim Preserve GrhData(1 To Inicio + (Count - 1)) As GrhData
End If
    
For k = Inicio To Inicio + (Count - 1)
    If k - (Inicio - 1) <= 22 Then
        Cant = 1
    ElseIf k - (Inicio - 1) <= 24 Then
        Cant = 6
    Else
        Cant = 5
    End If
    
    'Make sure he is not overwritting anything
    If k <= UBound(GrhData()) Then
        If GrhData(k).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    If GrhData(k).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > k Then
                Exit For
            End If
        Next i
        
        'Add it!
        If Cant > 1 Then
            Call frmMain.grhList.AddItem(k & " (ANIMACIÓN)", i)
        Else
            Call frmMain.grhList.AddItem(k, i)
        End If
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = k Then
                If Cant > 1 Then
                    frmMain.grhList.List(i) = k & " (ANIMACIÓN)"
                Else
                    frmMain.grhList.List(i) = k
                End If
                
                Exit For
            End If
        Next i
    End If
    
    'Fill in grh data
    With GrhData(k)
        .FileNum = FileNum
        
        .NumFrames = Cant
        ReDim .Frames(1 To .NumFrames) As Long
        
        If Cant = 1 Then
            .Speed = 0
            .Frames(1) = k
        Else
            For j = 1 To .NumFrames
                Select Case k - (Inicio - 1)
                    Case 23
                        .Frames(j) = (Inicio - 1) + j
                    Case 24
                        .Frames(j) = (Inicio - 1) + 6 + j
                    Case 25
                        .Frames(j) = (Inicio - 1) + 12 + j
                    Case 26
                        .Frames(j) = (Inicio - 1) + 17 + j
                End Select
            Next j
            
            .Speed = .NumFrames * 1000 / 18
        End If
                
        .pixelHeight = 45
        .pixelWidth = 25
        
        If k - Inicio < 6 Then
            .sX = (k - Inicio) * .pixelWidth
            .sY = 0
        ElseIf k - Inicio < 12 Then
            .sX = (k - Inicio - 6) * .pixelWidth
            .sY = 45
        ElseIf k - Inicio < 17 Then
            .sX = (k - Inicio - 12) * .pixelWidth
            .sY = 90
        ElseIf k - Inicio < 22 Then
            .sX = (k - Inicio - 17) * .pixelWidth
            .sY = 135
        Else
            .sX = 0
            .sY = 0
        End If
            
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = i
    DoEvents
Next k

End Sub

Private Sub AutoIndexarMnu_Click()
Call frmAutoIndex.Show(vbModeless, Me)
End Sub

Private Sub bmpTxt_Change()
    Dim path As String
    
    'Prevent non numeric characters
    If Not IsNumeric(bmpTxt.Text) Then
        bmpTxt.Text = Val(bmpTxt.Text)
    End If
    
    'Prevent overflow
    If Val(bmpTxt.Text) > &H7FFFFFFF Then
        bmpTxt.Text = &H7FFFFFFF
    End If
    
    'Prevent underrflow
    If Val(bmpTxt.Text) < 1 Then
        bmpTxt.Text = "1"
    End If
    
    
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\" & bmpTxt.Text & ".bmp"
    Else
        path = Config.bmpPath & bmpTxt.Text & ".bmp"
    End If
    
    'If file exists, load it
    If FileExists(path) And currentGrh <> NO_GRH Then
        GrhData(currentGrh).FileNum = CLng(bmpTxt.Text)
        
        'Prevent memory leaks
        Set currentPic = Nothing
        Set currentPic = LoadPicture(path)
        
        'Set scrollers!
        Call SetScrollers
        
        'Display the grh!
        Call RedrawPicture(currentGrh, currentFrame)
        
        'Show selection box (if needed)
        ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
        Call RenderSelectionBox
    End If
End Sub

Private Sub BodiesMnu_Click()
Call frmPersonajes.Show(vbModeless, Me)
End Sub

Private Sub Command1_Click()
Const Count As Long = 12
Dim Inicio As Long
Dim k As Long

Dim index As Long
Dim i As Long
Dim j As Long

Dim Cant As Byte

Dim file As Integer
Dim Indexados As Long

Dim Alto As Boolean

For Indexados = 11117 To 11204 Step 2
    For i = 1 To UBound(GrhData)
        If GrhData(i).NumFrames = 0 Then
            k = k + 1
            If k = Count Then
                Inicio = i - (Count - 2)
                Exit For
            End If
        Else
            k = 0
            Inicio = 0
        End If
    Next i
    
    If Inicio = 0 Then Inicio = UBound(GrhData) + 2
    Cant = 1
    Alto = Not Alto
    
    'Make sure he is not overwritting anything
    If (Inicio - 1) <= UBound(GrhData()) Then
        If GrhData(Inicio - 1).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
    Else
        'Resize array
        ReDim Preserve GrhData(1 To (Inicio - 1)) As GrhData
    End If
    
    If GrhData(Inicio - 1).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > (Inicio - 1) Then
                Exit For
            End If
        Next i
        
        'Add it!
        If Cant > 1 Then
            Call frmMain.grhList.AddItem((Inicio - 1) & " (ANIMACIÓN)", i)
        Else
            Call frmMain.grhList.AddItem((Inicio - 1), i)
        End If
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = (Inicio - 1) Then
                If Cant > 1 Then
                    frmMain.grhList.List(i) = (Inicio - 1) & " (ANIMACIÓN)"
                Else
                    frmMain.grhList.List(i) = (Inicio - 1)
                End If
                
                Exit For
            End If
        Next i
    End If
    
    'Fill in grh data
    With GrhData(Inicio - 1)
        .FileNum = Indexados
        
        .NumFrames = 1
        ReDim .Frames(1 To .NumFrames) As Long
        
        .Speed = 0
        .Frames(1) = Inicio - 1
                
        .pixelHeight = 32
        .pixelWidth = 32
        .sX = 0
        .sY = 0
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    For k = Inicio To Inicio + (Count - 1)
        index = k
        
        If k - (Inicio - 1) <= 22 Then
            Cant = 1
        ElseIf k - (Inicio - 1) <= 24 Then
            Cant = 6
        Else
            Cant = 5
        End If
        
        'Make sure he is not overwritting anything
        If index <= UBound(GrhData()) Then
            If GrhData(index).NumFrames > 0 Then
                If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                    Exit Sub
                End If
            End If
        Else
            'Resize array
            ReDim Preserve GrhData(1 To index) As GrhData
        End If
        
        If GrhData(index).NumFrames = 0 Then
            'Search where to place the grh....
            For i = 0 To frmMain.grhList.ListCount - 1
                If Val(frmMain.grhList.List(i)) > index Then
                    Exit For
                End If
            Next i
            
            'Add it!
            If Cant > 1 Then
                Call frmMain.grhList.AddItem(index & " (ANIMACIÓN)", i)
            Else
                Call frmMain.grhList.AddItem(index, i)
            End If
        Else
            'Search for the grh index within the grhList
            For i = 0 To frmMain.grhList.ListCount - 1
                If Val(frmMain.grhList.List(i)) = index Then
                    If Cant > 1 Then
                        frmMain.grhList.List(i) = index & " (ANIMACIÓN)"
                    Else
                        frmMain.grhList.List(i) = index
                    End If
                    
                    Exit For
                End If
            Next i
        End If
        
        'Fill in grh data
        With GrhData(index)
            .FileNum = Indexados + 1
            
            .NumFrames = Cant
            ReDim .Frames(1 To .NumFrames) As Long
            
            If Cant = 1 Then
                .Speed = 0
                .Frames(1) = index
            Else
                .Speed = .NumFrames * 1000 / 18
                ReDim Preserve MisCuerpos(1 To UBound(MisCuerpos) + 1) As tIndiceCuerpo
                
                If k - (Inicio - 1) = 23 Then
                    MisCuerpos(UBound(MisCuerpos)).Body(3) = index
                    MisCuerpos(UBound(MisCuerpos)).Body(1) = index + 1
                    MisCuerpos(UBound(MisCuerpos)).Body(4) = index + 2
                    MisCuerpos(UBound(MisCuerpos)).Body(2) = index + 3
                End If
                    
                For j = 1 To .NumFrames
                    Select Case k - (Inicio - 1)
                        Case 23
                            .Frames(j) = (Inicio - 1) + j
                        Case 24
                            .Frames(j) = (Inicio - 1) + 6 + j
                        Case 25
                            .Frames(j) = (Inicio - 1) + 12 + j
                        Case 26
                            .Frames(j) = (Inicio - 1) + 17 + j
                    End Select
                Next j
                
                If Alto Then
                    MisCuerpos(UBound(MisCuerpos)).HeadOffsetY = -38
                Else
                    MisCuerpos(UBound(MisCuerpos)).HeadOffsetY = -28
                End If
            End If
                    
            .pixelHeight = 45
            .pixelWidth = 25
            
            If k - Inicio < 6 Then
                .sX = (k - Inicio) * .pixelWidth
                .sY = 0
            ElseIf k - Inicio < 12 Then
                .sX = (k - Inicio - 6) * .pixelWidth
                .sY = 45
            ElseIf k - Inicio < 17 Then
                .sX = (k - Inicio - 12) * .pixelWidth
                .sY = 90
            ElseIf k - Inicio < 22 Then
                .sX = (k - Inicio - 17) * .pixelWidth
                .sY = 135
            Else
                .sX = 0
                .sY = 0
            End If
                
            .TileHeight = .pixelHeight / Config.TilePixelHeight
            .TileWidth = .pixelWidth / Config.TilePixelWidth
        End With
        
        'Now select it in the list
        frmMain.grhList.ListIndex = i
        DoEvents
    Next k
Next Indexados
End Sub

Private Sub Command2_Click()
Dim i As Long
Dim j As Long

For i = 1 To UBound(GrhData)
    With GrhData(i)
        If .NumFrames > 1 Then
            For j = 2 To .NumFrames
                If GrhData(.Frames(j)).pixelHeight <> GrhData(.Frames(1)).pixelHeight Then
                    Debug.Print "Grh " & i & " con índices bugueados"
                    Debug.Print "Grh inicial(" & .Frames(1) & ") con alto de " & GrhData(.Frames(1)).pixelHeight
                    Debug.Print "Frame " & j & "(" & .Frames(j) & ") con alto de " & GrhData(.Frames(j)).pixelHeight
                    Debug.Print ""
                ElseIf GrhData(.Frames(j)).pixelWidth <> GrhData(.Frames(1)).pixelWidth Then
                    Debug.Print "Grh " & i & " con índices bugueados"
                    Debug.Print "Grh inicial(" & .Frames(1) & ") con ancho de " & GrhData(.Frames(1)).pixelWidth
                    Debug.Print "Frame " & j & "(" & .Frames(j) & ") con ancho de " & GrhData(.Frames(j)).pixelWidth
                    Debug.Print ""
                End If
            Next j
        End If
    End With
Next i
End Sub

Private Sub CabezasMnu_Click()
Const Count As Long = 4
Dim Inicio As Long
Dim k As Long
Dim i As Long
Dim FileNum As Long
Dim Raza As Byte
Dim Genero As Byte
Dim IniCabezas As Integer

FileNum = Val(InputBox("Ingrese el número del gráfico."))

Raza = Val(InputBox("Ingrese el número de la raza." & vbCrLf & _
    "1. Humanos" & vbCrLf & _
    "2. Elfos" & vbCrLf & _
    "3. Elfos oscuros" & vbCrLf & _
    "4. Enanos" & vbCrLf & _
    "5. Gnomos"))
    
Genero = Val(InputBox("Ingrese el género." & vbCrLf & _
    "1. Hombres" & vbCrLf & _
    "2. Mujeres"))
    
If FileNum = 0 Then Exit Sub
If Raza = 0 Then Exit Sub
If Genero = 0 Then Exit Sub

Raza = Raza - 1
Genero = Genero - 1

For i = 1 To UBound(GrhData)
    If GrhData(i).NumFrames = 0 Then
        k = k + 1
        If k = Count Then
            Inicio = i - (Count - 2)
            Exit For
        End If
    Else
        k = 0
        Inicio = 0
    End If
Next i

If Inicio = 0 Then
    Inicio = UBound(GrhData) + 1

    'Resize array
    ReDim Preserve GrhData(1 To Inicio + (Count - 1)) As GrhData
End If

For k = Inicio To Inicio + (Count - 1)
    If GrhData(k).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > k Then
                Exit For
            End If
        Next i
        
        'Add it!
        Call frmMain.grhList.AddItem(k, i)
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = k Then
                frmMain.grhList.List(i) = k
                
                Exit For
            End If
        Next i
    End If
    
    'Fill in grh data
    With GrhData(k)
        .FileNum = FileNum
        
        .NumFrames = 1
        ReDim .Frames(1 To 1) As Long
        
        .Speed = 0
        .Frames(1) = k
                
        .pixelHeight = 50
        .pixelWidth = 17
        
        .sY = 0
        
        .sX = 17 * (k - Inicio)
            
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = i
    DoEvents
Next k

IniCabezas = (Raza * 100) + (69 * Genero) + 1

Do While MisCabezas(IniCabezas).Head(1) > 0
    IniCabezas = IniCabezas + 1
    
    If IniCabezas > UBound(MisCabezas) Then
        ReDim Preserve MisCabezas(1 To IniCabezas) As tIndiceCabeza
    End If
Loop
    
MisCabezas(IniCabezas).Head(3) = Inicio
MisCabezas(IniCabezas).Head(2) = Inicio + 1
MisCabezas(IniCabezas).Head(4) = Inicio + 2
MisCabezas(IniCabezas).Head(1) = Inicio + 3
End Sub

Private Sub DurationTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(DurationTxt.Text) Then
        DurationTxt.Text = Val(DurationTxt.Text)
    End If
    
    'Prevent overflow
    If Val(DurationTxt.Text) > &H9FFFA Then
        DurationTxt.Text = &H9FFFA
    End If
    
    'Prevent negative values
    If Val(DurationTxt.Text) < 0 Then
        DurationTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        If Val(DurationTxt.Text) < 5 Then
            GrhData(currentGrh).Speed = CSng(DurationTxt.Text) * GrhData(currentGrh).NumFrames * 1000 / 18
        Else
            GrhData(currentGrh).Speed = DurationTxt.Text
        End If
    End If
End Sub

Private Sub DurationTxt_LostFocus()
Call grhList_Click
End Sub

Private Sub fileList_Click()
    Dim path As String
    
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\" & fileList.Text & ".bmp"
    Else
        path = Config.bmpPath & fileList.Text & ".bmp"
    End If
    
    'Prevent memory leaks
    Set currentPic = Nothing
    Set currentPic = LoadPicture(path)
    
    'Reset selection box
    selectionAreaEndX = 0
    selectionAreaEndY = 0
    selectionAreaStartX = 0
    selectionAreaStartY = 0
    
    'Set scrollers!
    Call SetScrollers
    
    currentGrh = NO_GRH
    
    bmpTxt.Text = fileList.Text
    
    'Draw!
    Call RedrawPicture(NO_GRH, 0)
    
    ignoreSelectionBoxRender = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim fileName As String
    Dim path As String
    
    If Not LoadConfig() Then
        'Show config form
        Call frmConfig.Show(vbModal, Me)
    End If
    
    'Load Grhs!
    Call LoadGrhData(Config.initPath)
    Call CargarCabezas(Config.initPath)
    Call CargarCascos(Config.initPath)
    Call CargarCuerpos(Config.initPath)
    Call CargarFxs(Config.initPath)
    
    'Fill the lists
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            If GrhData(i).NumFrames = 1 Then
                Call grhList.AddItem(CStr(i))
            Else
                Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
            End If
        End If
    Next i
    
    'Set up bmp search path
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\*.bmp"
    Else
        path = Config.bmpPath & "*.bmp"
    End If
    
    fileName = Dir$(path, vbArchive)
    
    While fileName <> ""
        'Add it!
        fileName = Left$(fileName, InStr(1, fileName, ".") - 1)
        
        'Make usre it's numeric
        If IsNumeric(fileName) Then
            Call fileList.AddItem(fileName)
        End If
        
        fileName = Dir()
    Wend
    
    'Set default zoom value
    ZoomTxt.Text = DEFAULT_ZOOM
    
    editionCoord = sbpeNone
    
    currentGrh = NO_GRH
    
    'By default update events are not ignored
    ignoreGrhTextUpdate = False
    
    'Show first grh by default
    If grhList.ListCount > 0 Then
        grhList.ListIndex = 0
    ElseIf fileList.ListCount > 0 Then
        fileList.ListIndex = 0
    End If
End Sub

Private Sub FramesTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(FramesTxt.Text) Then
        FramesTxt.Text = Val(FramesTxt.Text)
    End If
    
    'Prevent overflow
    If Val(FramesTxt.Text) > &H7FFFFFFF Then
        FramesTxt.Text = &H7FFFFFFF
    End If
End Sub

Private Sub FramesTxt_LostFocus()
    Dim i As Long
    
    'Prevent negative values and animations with one frame
    If Val(FramesTxt.Text) < 2 Then
        FramesTxt.Text = "2"
    End If
    
    With GrhData(Val(grhList.Text))
        ReDim Preserve .Frames(1 To Val(FramesTxt.Text)) As Long
        
        If Val(FramesTxt.Text) > .NumFrames Then
            For i = .NumFrames + 1 To Val(FramesTxt.Text)
                .Frames(i) = .Frames(.NumFrames)
            Next i
        End If
        
        .NumFrames = Val(FramesTxt.Text)
    End With
    
    Call grhList_Click
End Sub

Private Sub FrameTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(FrameTxt.Text) Then
        FrameTxt.Text = Val(FrameTxt.Text)
    End If
    
    'Prevent overflow
    If Val(FrameTxt.Text) >= currentGrh Then
        If currentGrh > 1 Then
            FrameTxt.Text = currentGrh - 1
        Else
            FrameTxt.Text = currentGrh
        End If
    End If
    
    'Prevent underrflow
    If Val(FrameTxt.Text) < 1 Then
        FrameTxt.Text = "1"
    End If
End Sub

Private Sub FrameTxt_LostFocus()
GrhData(currentGrh).Frames(FrameLbl) = FrameTxt
End Sub

Private Sub FXsMnu_Click()
Call frmFxs.Show(vbModeless, Me)
End Sub

Private Sub grhHeightTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhHeightTxt.Text) Then
        grhHeightTxt.Text = Val(grhHeightTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhHeightTxt.Text) > &H7FFF Then
        grhHeightTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text) Then
        grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text))
    End If
    
    'Prevent negative values
    If CInt(grhHeightTxt.Text) < 0 Then
        grhHeightTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelHeight = CInt(grhHeightTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhList_Click()
    Dim path As String
    
    ' Set current grh and reset frame
    currentGrh = Val(grhList.Text)
    currentFrame = 1
    
    'Should grh controls be enabled?
    Call SetGrhControlsEnabled(grhList.Text = CStr(currentGrh))
    
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\" & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
    Else
        path = Config.bmpPath & GrhData(GrhData(currentGrh).Frames(currentFrame)).FileNum & ".bmp"
    End If
    
    'Prevent memory leaks
    Set currentPic = Nothing
    Set currentPic = LoadPicture(path)
    
    'Enable animations if necessary
    If GrhData(currentGrh).NumFrames > 1 Then
        animation.Enabled = False
        animation.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
        animation.Enabled = True
        
        grhOnly.value = vbChecked
        grhOnly.Enabled = False
    Else
        animation.Enabled = False
        
        If Not grhOnly.Enabled Then
            grhOnly.Enabled = True
            
            grhOnly.value = vbChecked
        ElseIf grhOnly.value = vbUnchecked Then
            'Set selection box!
            Call SelectGrhArea(currentGrh)
        End If
        
        'Show bmp
        bmpTxt.Text = GrhData(currentGrh).FileNum
        
        'Filelist will reset the currentGrh, restore it!
        currentGrh = Val(grhList.Text)
        
        'Set selection box!
        Call SelectGrhArea(currentGrh)
        
        'Display grh info
        grhXTxt.Text = GrhData(currentGrh).sX
        grhYTxt.Text = GrhData(currentGrh).sY
        grhWidthTxt.Text = GrhData(currentGrh).pixelWidth
        grhHeightTxt.Text = GrhData(currentGrh).pixelHeight
    End If
    
    DurationTxt.Text = GrhData(currentGrh).Speed
    FramesTxt.Text = GrhData(currentGrh).NumFrames
    FrameLbl.Caption = 1
    FrameTxt.Text = GrhData(currentGrh).Frames(1)
    
    'Set scrollers!
    Call SetScrollers
    
    'Display the grh!
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box (if needed)
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    Call RenderSelectionBox
End Sub

Private Sub grhOnly_Click()
    If currentGrh = NO_GRH Then Exit Sub
    
    Call RedrawPicture(currentGrh, currentFrame)
    
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    
    'Set selection box!
    Call SelectGrhArea(currentGrh)
    
    Call RenderSelectionBox
End Sub

Private Sub grhWidthTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhWidthTxt.Text) Then
        grhWidthTxt.Text = Val(grhWidthTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhWidthTxt.Text) > &H7FFF Then
        grhWidthTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhWidthTxt.Text) > previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text) Then
        grhWidthTxt.Text = Round(previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text))
    End If
    
    'Prevent negative values
    If CInt(grhWidthTxt.Text) < 0 Then
        grhWidthTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelWidth = CInt(grhWidthTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndX = selectionAreaStartX + CInt(grhWidthTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhXTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhXTxt.Text) Then
        grhXTxt.Text = Val(grhXTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhXTxt.Text) > &H7FFF Then
        grhXTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhXTxt.Text) > previewer.ScaleX(currentPic.Width) Then
        grhXTxt.Text = Round(previewer.ScaleX(currentPic.Width))
    End If
    
    'Prevent negative values
    If CInt(grhXTxt.Text) < 0 Then
        grhXTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sX = CInt(grhXTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartX = CInt(grhXTxt.Text)
    selectionAreaEndX = selectionAreaStartX + Val(grhWidthTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhYTxt_Change()
    'Prevent non numeric characters
    If Not IsNumeric(grhYTxt.Text) Then
        grhYTxt.Text = Val(grhYTxt.Text)
    End If
    
    'Prevent overflow
    If Val(grhYTxt.Text) > &H7FFF Then
        grhYTxt.Text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    If CInt(grhYTxt.Text) > previewer.ScaleY(currentPic.Height) Then
        grhYTxt.Text = Round(previewer.ScaleY(currentPic.Height))
    End If
    
    'Trim height to prevent invalid values
    If CInt(grhYTxt.Text) + Val(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) Then
        grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height)) - CInt(grhYTxt.Text)
    End If
    
    'Prevent negative values
    If CInt(grhYTxt.Text) < 0 Then
        grhYTxt.Text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sY = CInt(grhYTxt.Text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, currentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartY = Val(grhYTxt.Text)
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.Text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub HeadsMnu_Click()
Call frmCabezas.Show(vbModeless, Me)
End Sub

Private Sub HelmetsMnu_Click()
Call frmCascos.Show(vbModeless, Me)
End Sub

Private Sub NextCmd_Click()
If Val(FrameLbl.Caption) = GrhData(currentGrh).NumFrames Then Exit Sub
FrameLbl.Caption = Val(FrameLbl.Caption) + 1
FrameTxt = GrhData(currentGrh).Frames(Val(FrameLbl.Caption))
End Sub

Private Sub OpenNormalMnu_Click()
Dim i As Long

animation.Enabled = False
Call grhList.Clear

'Load Grhs!
Call LoadGrhData(Config.initPath)

'Fill the lists
For i = 1 To UBound(GrhData())
    If GrhData(i).NumFrames > 0 Then
        If GrhData(i).NumFrames = 1 Then
            Call grhList.AddItem(CStr(i))
        Else
            Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
        End If
    End If
Next i
End Sub

Private Sub OpenSmallMnu_Click()
Dim i As Long

animation.Enabled = False
Call grhList.Clear

'Load Grhs!
Call LoadGrhData(Config.initPath, 1)

'Fill the lists
For i = 1 To UBound(GrhData())
    If GrhData(i).NumFrames > 0 Then
        If GrhData(i).NumFrames = 1 Then
            Call grhList.AddItem(CStr(i))
        Else
            Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
        End If
    End If
Next i
End Sub

Private Sub OpenTransparentMnu_Click()
Dim i As Long

animation.Enabled = False
Call grhList.Clear

'Load Grhs!
Call LoadGrhData(Config.initPath, 2)

'Fill the lists
For i = 1 To UBound(GrhData())
    If GrhData(i).NumFrames > 0 Then
        If GrhData(i).NumFrames = 1 Then
            Call grhList.AddItem(CStr(i))
        Else
            Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
        End If
    End If
Next i
End Sub

Private Sub picScrollH_Change()
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub picScrollV_Change()
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub previewer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If no picture is loaded, there is nothing to be done
    If currentPic Is Nothing Then Exit Sub
    
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        Select Case Me.MousePointer
            Case vbDefault
                'A new box is being created, we are fixing start x-y coord and moving end x-y
                editionCoord = sbpeEndXEndY
                
                'Make sure selection box doesn't go beyond bmp
                If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
                    x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
                End If
                
                If ViewPortToBmpPosY(y) > previewer.ScaleY(currentPic.Height) Then
                    y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
                End If
                
                
                'Convert mouse pos to pixel pos of origin
                selectionAreaStartX = ViewPortToBmpPosX(x)
                selectionAreaStartY = ViewPortToBmpPosY(y)
                
                'Reset end area, we are starting a new rectangle
                selectionAreaEndX = selectionAreaStartX
                selectionAreaEndY = selectionAreaStartY
                
                'Show selection box!
                Call RenderSelectionBox
            
            Case vbSizeNS
                If Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 2 Then
                    editionCoord = sbpeStartY
                ElseIf Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 2 Then
                    editionCoord = sbpeEndY
                End If
            
            Case vbSizeWE
                If Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 2 Then
                    editionCoord = sbpeStartX
                ElseIf Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 2 Then
                    editionCoord = sbpeEndX
                End If
            
            Case vbSizeNWSE
                If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 5) Then
                    editionCoord = sbpeStartXStartY
                ElseIf (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 5) Then
                    editionCoord = sbpeEndXEndY
                End If
            
            Case vbSizeNESW
                If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 5) Then
                    editionCoord = sbpeStartXEndY
                ElseIf (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 5) Then
                    editionCoord = sbpeEndXStartY
                End If
        End Select
    End If
End Sub

Private Sub previewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        'If we got past the border, we scroll!!
        If x < 0 Then
            x = 0
            
            If picScrollH.value > 0 And picScrollH.Enabled Then
                picScrollH.value = picScrollH.value - 1
            End If
        ElseIf x > previewer.Width Then
            x = previewer.Width
            
            If picScrollH.value < picScrollH.max And picScrollH.Enabled Then
                picScrollH.value = picScrollH.value + 1
            End If
        End If
        
        If y < 0 Then
            y = 0
            
            If picScrollV.value > 0 And picScrollV.Enabled Then
                picScrollV.value = picScrollV.value - 1
            End If
        ElseIf y > previewer.Height Then
            y = previewer.Height
            
            If picScrollV.value < picScrollV.max And picScrollV.Enabled Then
                picScrollV.value = picScrollV.value + 1
            End If
        End If
        
        
        'Make sure selection box doesn't go beyond bmp
        If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
            x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
        End If
        
        If ViewPortToBmpPosY(y) > previewer.ScaleY(currentPic.Height) Then
            y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
        End If
        
        
        'Update coords
        Call UpdateSelectionBox(x, y)
        
        'Show selection box!
        Call RenderSelectionBox
    ElseIf Not ignoreSelectionBoxRender And selectionAreaStartX <> selectionAreaEndX And selectionAreaStartY <> selectionAreaEndY Then
        'Allow the user to resize the selection box!
        
        'Set mouse pointer appropiately
        If (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 5) _
                Or (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 5) Then
            Me.MousePointer = vbSizeNWSE
        
        ElseIf (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 5) _
                Or (Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 5 And Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 5) Then
            Me.MousePointer = vbSizeNESW
        
        ElseIf (Abs(selectionAreaStartX - ViewPortToBmpPosX(x)) < 2 Or Abs(selectionAreaEndX - ViewPortToBmpPosX(x)) < 2) _
                And ViewPortToBmpPosY(y) > selectionAreaStartY And ViewPortToBmpPosY(y) < selectionAreaEndY Then
            Me.MousePointer = vbSizeWE
        
        ElseIf (Abs(selectionAreaStartY - ViewPortToBmpPosY(y)) < 2 Or Abs(selectionAreaEndY - ViewPortToBmpPosY(y)) < 2) _
                And ViewPortToBmpPosX(x) > selectionAreaStartX And ViewPortToBmpPosX(x) < selectionAreaEndX Then
            Me.MousePointer = vbSizeNS
        
        Else
            Me.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub previewer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        
        'Make sure selection box doesn't go beyond bmp
        If ViewPortToBmpPosX(x) > previewer.ScaleX(currentPic.Width) Then
            x = BmpToViewPortPosX(previewer.ScaleX(currentPic.Width))
        End If
        
        If ViewPortToBmpPosY(y) > previewer.ScaleY(currentPic.Height) Then
            y = BmpToViewPortPosY(previewer.ScaleY(currentPic.Height))
        End If
        
        'Update selection box
        Call UpdateSelectionBox(x, y)
        
        'Show selection box!
        Call RenderSelectionBox
    End If
End Sub

Private Sub PreviousCmd_Click()
If Val(FrameLbl.Caption) = 1 Then Exit Sub
FrameLbl.Caption = Val(FrameLbl.Caption) - 1
FrameTxt = GrhData(currentGrh).Frames(Val(FrameLbl.Caption))
End Sub

Private Sub RemoveGrhMnu_Click()
    Dim i As Long
    
    If currentGrh = NO_GRH Then
        MsgBox "There is no grh selected."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the grh " & currentGrh & "?" & vbCrLf & "This change can't be undone.", vbOKCancel) = vbOK Then
        'Reset it
        With GrhData(currentGrh)
            .FileNum = 0
            ReDim .Frames(0)
            .NumFrames = 0
            .pixelHeight = 0
            .pixelWidth = 0
            .Speed = 0
            .sX = 0
            .sY = 0
            .TileHeight = 0
            .TileWidth = 0
        End With
        
        'Remove it
        For i = 0 To grhList.ListCount - 1
            If Val(grhList.List(i)) = currentGrh Then
                grhList.RemoveItem (i)
                Exit For
            End If
        Next i
        
        'Select next grh
        If i < grhList.ListCount Then
            grhList.ListIndex = i
        Else
            grhList.ListIndex = grhList.ListCount - 1
        End If
    End If
End Sub

Private Sub SaveBodiesMnu_Click()
    If Not grh.SaveBodies(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveFXsMnu_Click()
    If Not grh.SaveFXs(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveHeadsMnu_Click()
    If Not grh.SaveHeads(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveHelmetsMnu_Click()
    If Not grh.SaveHelmets(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveInAllMnu_Click()
    If Not grh.SaveInAllGraphics(Config.initPath, (grh.fileVersion <> -1)) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveMnu_Click()
    'Detect the original file format and save it
    If grh.fileVersion = -1 Then
        If Not grh.SaveGrhDataOld(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    Else
        If Not grh.SaveGrhDataNew(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
End Sub

Private Sub SaveNewMnu_Click()
    If Not grh.SaveGrhDataNew(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveOldMnu_Click()
    If MsgBox("The old file format speed system is FPS based, animation's speed may be altered. Do you want to proceed?", vbYesNo) = vbYes Then
        If Not grh.SaveGrhDataOld(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
End Sub

Private Sub SearchBMPMnu_Click()
Dim FileNum As String
Dim i As Long
Dim Str As String

FileNum = InputBox("Ingrese el gráfico a buscar.")
If StrPtr(FileNum) = 0 Then Exit Sub

FileNum = Val(FileNum)

For i = 1 To UBound(GrhData)
    If GrhData(i).FileNum = CLng(FileNum) Then
        If Str = vbNullString Then
            Str = i
        Else
            Str = Str & ", " & i
        End If
    End If
Next i

MsgBox "El BMP es utilizado por el/los grh/s:" & vbCrLf & Str & "."
End Sub

Private Sub SearchErroresMnu_Click()
Dim i As Long
Dim j As Long
Dim hFile As Long
Dim path As String

hFile = FreeFile

'Load new bitmap
If Right$(Config.bmpPath, 1) <> "\" Then
    path = Config.bmpPath & "\"
Else
    path = Config.bmpPath
End If
            
For i = 1 To UBound(GrhData)
    With GrhData(i)
        If .NumFrames > 0 Then
            If .NumFrames > 1 Then
                For j = 2 To .NumFrames
                    If GrhData(.Frames(j)).pixelHeight <> GrhData(.Frames(1)).pixelHeight Then
                        Debug.Print "Grh " & i & " con índices bugueados"
                        Debug.Print "Grh inicial(" & .Frames(1) & ") con alto de " & GrhData(.Frames(1)).pixelHeight
                        Debug.Print "Frame " & j & "(" & .Frames(j) & ") con alto de " & GrhData(.Frames(j)).pixelHeight
                        Debug.Print ""
                    End If
                    
                    If GrhData(.Frames(j)).pixelWidth <> GrhData(.Frames(1)).pixelWidth Then
                        Debug.Print "Grh " & i & " con índices bugueados"
                        Debug.Print "Grh inicial(" & .Frames(1) & ") con ancho de " & GrhData(.Frames(1)).pixelWidth
                        Debug.Print "Frame " & j & "(" & .Frames(j) & ") con ancho de " & GrhData(.Frames(j)).pixelWidth
                        Debug.Print ""
                    End If
                Next j
            Else
                Open path & .FileNum & ".bmp" For Binary Access Read As hFile
                    Get hFile, , FileHeaderBMP
                    Get hFile, , InfoHeaderBMP
                Close hFile
                
                If .pixelHeight > InfoHeaderBMP.biHeight - .sY Then
                    .pixelHeight = Round(InfoHeaderBMP.biHeight) - .sY
                    
                    Debug.Print "Grh " & i & " sobre pasa en el ancho las medidas del BMP."
                    Debug.Print ""
                End If
                
                If .pixelWidth > InfoHeaderBMP.biWidth - .sX Then
                    .pixelWidth = Round(InfoHeaderBMP.biWidth) - .sX
                    
                    Debug.Print "Grh " & i & " sobre pasa en el alto las medidas del BMP."
                    Debug.Print ""
                End If
            End If
        End If
    End With
Next i
End Sub

Private Sub ZoomIn_Click()
    ZoomTxt.Text = Val(ZoomTxt.Text) + ZOOM_STEP
End Sub

Private Sub ZoomOut_Click()
    ZoomTxt.Text = Val(ZoomTxt.Text) - ZOOM_STEP
End Sub

Private Sub ZoomTxt_Change()
    'Validate
    If Not IsNumeric(ZoomTxt.Text) Then
        ZoomTxt.Text = DEFAULT_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.Text) > MAX_ZOOM Then
        ZoomTxt.Text = MAX_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.Text) < MIN_ZOOM Then
        ZoomTxt.Text = MIN_ZOOM
        Exit Sub
    End If
    
    'Recompute zoom
    zoom = CSng(ZoomTxt.Text) / DEFAULT_ZOOM
    
    
    'Reset scrollbars
    Call SetScrollers
    
    'Redraw
    Call RedrawPicture(currentGrh, currentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

''
' Sets the scrollers' properties appropiately for the current picture loaded, zoom and value.

Private Sub SetScrollers()
    Dim oldMax As Integer
    
    If currentPic Is Nothing Then
        picScrollH.Enabled = False
        picScrollV.Enabled = False
        Exit Sub
    End If
    
    'Set up scrollers
    If previewer.Width < previewer.ScaleX(currentPic.Width) * zoom Then
        oldMax = IIf(picScrollH.max > 0, picScrollH.max, 1)
        
        picScrollH.max = previewer.ScaleX(currentPic.Width) - previewer.Width / zoom
        picScrollH.value = picScrollH.value * picScrollH.max / oldMax
        picScrollH.Enabled = True
    Else
        picScrollH.value = 0
        picScrollH.Enabled = False
    End If
    
    If previewer.Height < previewer.ScaleY(currentPic.Height) * zoom Then
        oldMax = IIf(picScrollV.max > 0, picScrollV.max, 1)
        
        picScrollV.max = previewer.ScaleX(currentPic.Height) - previewer.Height / zoom
        picScrollV.value = picScrollV.value * picScrollV.max / oldMax
        picScrollV.Enabled = True
    Else
        picScrollV.value = 0
        picScrollV.Enabled = False
    End If
End Sub

''
' Renders the last laoded picture.
'
' @param    grh     The grh to be rendered within the loaded picture. Can be @code NO_GRH
' @param    frame   The frame of the grh to be rendered. Only important if grh is not @code NO_GRH

Private Sub RedrawPicture(ByVal grh As Long, ByVal frame As Long)
    If currentPic Is Nothing Then Exit Sub
    
    'Clear picturebox
    Set previewer.Picture = Nothing
    previewer.Picture = LoadPicture("")
    
    'Render!
    If grh <> NO_GRH And grhOnly.value = vbChecked Then
        'Transform grh to actual frame grh.
        grh = GrhData(grh).Frames(frame)
        
        Call previewer.PaintPicture(currentPic, -picScrollH.value * zoom, -picScrollV.value * zoom, _
                                    GrhData(grh).pixelWidth * zoom, _
                                    GrhData(grh).pixelHeight * zoom, _
                                    GrhData(grh).sX, GrhData(grh).sY, _
                                    GrhData(grh).pixelWidth, GrhData(grh).pixelHeight)
    Else
        Call previewer.PaintPicture(currentPic, -picScrollH.value * zoom, -picScrollV.value * zoom, _
                                    previewer.ScaleX(currentPic.Width) * zoom, _
                                    previewer.ScaleY(currentPic.Height) * zoom)
    End If
End Sub

''
' Renders the selection box.

Private Sub RenderSelectionBox()
    Dim startX As Long
    Dim startY As Long
    Dim endX As Long
    Dim endY As Long
    
    If ignoreSelectionBoxRender Then Exit Sub
    
    'Transform origin coord to those in the picturebox
    startX = BmpToViewPortPosX(selectionAreaStartX)
    startY = BmpToViewPortPosY(selectionAreaStartY)
    
    'Transform end coord to those in the picturebox
    endX = BmpToViewPortPosX(selectionAreaEndX)
    endY = BmpToViewPortPosY(selectionAreaEndY)
    
    previewer.AutoRedraw = False
    previewer.Cls
    previewer.Line (startX, startY)-(endX, endY), vbRed, B
    previewer.AutoRedraw = True
End Sub

''
' Converts a bmp absolute pixel pos in the x axis to the picturebox's view area coord.
'
' @param    x   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function BmpToViewPortPosX(ByVal x As Long) As Long
    BmpToViewPortPosX = (x - picScrollH.value) * zoom
End Function

''
' Converts a bmp absolute pixel pos in the y axis to the picturebox's view area coord.
'
' @param    y   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function BmpToViewPortPosY(ByVal y As Long) As Long
    BmpToViewPortPosY = (y - picScrollV.value) * zoom
End Function

''
' Converts a picturebox's view area pos in the x axis to the bmp absolute pixel coord.
'
' @param    x   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function ViewPortToBmpPosX(ByVal x As Long) As Long
    ViewPortToBmpPosX = picScrollH.value + Fix(x / zoom)
End Function

''
' Converts a picturebox's view area pos in the y axis to the bmp absolute pixel coord.
'
' @param    y   The pixel position to be transformed.
' @return   The coord within the picturebox matching the bmp pixel pos.

Private Function ViewPortToBmpPosY(ByVal y As Long) As Long
    ViewPortToBmpPosY = picScrollV.value + Fix(y / zoom)
End Function

''
' Updates the appropiate selection box coords according to the current value of @code editionCoord.
'
' @param    x   The mouse pos in the x coord within the previewer.
' @param    y   The mouse pos in the y coord within the previewer.

Private Sub UpdateSelectionBox(ByVal x As Long, ByVal y As Long)
    Dim tmp As Long
    
    Select Case editionCoord
        Case sbpeNone
            'Convert mouse pos to pixel pos of end
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(y)
        
        Case sbpeStartX
            selectionAreaStartX = ViewPortToBmpPosX(x)
        
        Case sbpeStartY
            selectionAreaStartY = ViewPortToBmpPosY(y)
        
        Case sbpeEndX
            selectionAreaEndX = ViewPortToBmpPosX(x)
        
        Case sbpeEndY
            selectionAreaEndY = ViewPortToBmpPosY(y)
        
        Case sbpeStartXStartY
            selectionAreaStartX = ViewPortToBmpPosX(x)
            selectionAreaStartY = ViewPortToBmpPosY(y)
        
        Case sbpeEndXEndY
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(y)
        
        Case sbpeStartXEndY
            selectionAreaStartX = ViewPortToBmpPosX(x)
            selectionAreaEndY = ViewPortToBmpPosY(y)
        
        Case sbpeEndXStartY
            selectionAreaEndX = ViewPortToBmpPosX(x)
            selectionAreaStartY = ViewPortToBmpPosY(y)
    End Select
    
    'Invert coordinates if needed to prevent pointer from going crazy on corners.
    If selectionAreaStartX > selectionAreaEndX Then
        tmp = selectionAreaEndX
        selectionAreaEndX = selectionAreaStartX
        selectionAreaStartX = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndX
                editionCoord = sbpeStartX
            
            Case sbpeEndXEndY
                editionCoord = sbpeStartXEndY
            
            Case sbpeEndXStartY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartX
                editionCoord = sbpeEndX
            
            Case sbpeStartXEndY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartXStartY
                editionCoord = sbpeEndXStartY
        End Select
    End If
    
    If selectionAreaStartY > selectionAreaEndY Then
        tmp = selectionAreaEndY
        selectionAreaEndY = selectionAreaStartY
        selectionAreaStartY = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndY
                editionCoord = sbpeStartY
            
            Case sbpeEndXEndY
                editionCoord = sbpeEndXStartY
            
            Case sbpeEndXStartY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartY
                editionCoord = sbpeEndY
            
            Case sbpeStartXEndY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartXStartY
                editionCoord = sbpeStartXEndY
        End Select
    End If
    
    'Display data at the bottom
    ignoreGrhTextUpdate = True
    
    grhHeightTxt.Text = selectionAreaEndY - selectionAreaStartY
    grhWidthTxt.Text = selectionAreaEndX - selectionAreaStartX
    grhXTxt.Text = selectionAreaStartX
    grhYTxt.Text = selectionAreaStartY
    
    ignoreGrhTextUpdate = False
End Sub

''
' Sets up the selection area around the given grh within it's bmp.
'
' @param    grh     The grh to be selected.

Private Sub SelectGrhArea(ByVal grh As Long)
    selectionAreaStartX = GrhData(grh).sX
    selectionAreaStartY = GrhData(grh).sY
    selectionAreaEndX = selectionAreaStartX + GrhData(grh).pixelWidth
    selectionAreaEndY = selectionAreaStartY + GrhData(grh).pixelHeight
End Sub

''
'Enables / disables the grh controls (those within the grhFrame control).
'
' @param    enable  True if controls should be enabled, False otherwise.

Private Sub SetGrhControlsEnabled(ByVal enable As Boolean)
    Dim i As Long
    
    For i = 0 To frmMain.Controls.Count - 1
        If Not TypeOf frmMain.Controls(i) Is Timer And Not TypeOf frmMain.Controls(i) Is Menu Then
            If frmMain.Controls(i).Container Is grhFrame Then
                frmMain.Controls(i).Enabled = enable
            ElseIf frmMain.Controls(i).Container Is AnimationFrame Then
                frmMain.Controls(i).Enabled = Not enable
            End If
        End If
    Next i
    
    grhFrame.Enabled = enable
    AnimationFrame.Enabled = Not enable
End Sub
