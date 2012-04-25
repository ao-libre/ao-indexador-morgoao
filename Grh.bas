Attribute VB_Name = "Grh"
Option Explicit

Private GRH_DAT_FILE As String  '= "Graficos.ind"
Private Const OLD_FORMAT_HEADER As String = "Argentum Online by Noland-Studios."
Private Const OLD_FORMAT_INIT_FILE As String = "Inicio.con"

Public Const GRAFICOS_DIFERENTES As String = "*9*427*444*641*643*644*647*735*1529*1530*1531*1532*1533*" & _
                                                "1534*3584*4503*4876*4885*4893*4899*4940*5591*5598*5601*6566*" & _
                                                "7000*7001*7002*7222*7223*7224*7225*7226*7231*9141*9142*9143*" & _
                                                "9144*9145*9146*9147*9148*9149*9150*9151*9152*9153*9154*9155*" & _
                                                "9156*9157*9158*9159*9160*9161*9162*9163*9164*9165*9166*9167*" & _
                                                "9168*9169*9170*9171*9172*9173*9174*9175*9176*9177*9178*9179*" & _
                                                "9180*9181*9182*9183*9184*9185*9186*9187*9188*9189*9190*9191*" & _
                                                "9192*12334*13953*18478*18479*18480*18481*18802*18803*"

Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Private Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public GrhData() As GrhData
Public MisCabezas() As tIndiceCabeza
Public MisCascos() As tIndiceCabeza
Public MisCuerpos() As tIndiceCuerpo
Public FxData() As tIndiceFx

Public fileVersion As Long

Public Function LoadGrhData(ByVal path As String, Optional ByVal IndAUsarse As Byte = 3) As Boolean
On Error GoTo ErrHandler
    Dim handle As Integer
    Dim MiCabecera As tCabecera
    Dim i As Long
    
    'Set initial size
    ReDim GrhData(0) As GrhData
    
    handle = FreeFile()
    
    If path = vbNullString Then Exit Function
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    GRH_DAT_FILE = "Graficos" & IndAUsarse & ".ind"
    
    If Not FileExists(path & GRH_DAT_FILE) Then
        MsgBox "The file " & path & GRH_DAT_FILE & " does not exist. A new one will be created with your work."
        Exit Function
    End If
    
    frmMain.Caption = "Indexador Alkon (" & GRH_DAT_FILE & ")"
    
    Open path & GRH_DAT_FILE For Binary Access Read Lock Write As handle
    
    'Check file format! (The crappy header had to have some use after all!)
    Get handle, , MiCabecera
    
    If Left$(MiCabecera.desc, Len(OLD_FORMAT_HEADER)) = OLD_FORMAT_HEADER Then
        LoadGrhData = LoadGrhDataOld(handle, NumberOfGrhs(path))
        
        'No version available in old file format
        fileVersion = -1
    Else
        'We dont' have header, move back to the beginning
        Seek handle, 1
        
        LoadGrhData = LoadGrhDataNew(handle)
    End If
    
    Close handle
Exit Function

ErrHandler:
    Close handle
    
    MsgBox "An error occured while loading the grh data." & vbCrLf _
        & "Make sure file format is valid, and in case of using the old format, make sure the " _
        & OLD_FORMAT_INIT_FILE & " file is in the init folder"
End Function

''
' Old crappy format loading. Restricted to 2^15-1 grhs,
' stores animation speed in frames and other crappy stuff.
' Coded just for backwards compatibility, users should avoid using this format.
'
' @param    handle      Handle to the open file containing the grh data.
'                       The header should have allready been removed.
' @param    totalGrhs   The total number of grhs that could exist.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataOld(ByVal handle As Integer, ByVal totalGrhs As Long) As Boolean
On Error GoTo ErrorHandler
    Dim grh As Integer
    Dim frame As Long
    Dim tempint As Integer
    Dim max As Integer
    
    max = -1
    
    'Resize array
    ReDim GrhData(1 To totalGrhs) As GrhData
    
    'Open files
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Fill Grh List
    
    'Get first Grh Number
    Get handle, , grh
    
    Do Until grh <= 0
        'Get highest grh number being used
        If grh > max Then
            max = grh
        End If
        
        With GrhData(grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'Resize animation array
            ReDim .Frames(1 To .NumFrames) As Long
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For frame = 1 To .NumFrames
                
                    Get handle, , tempint
                    
                    'Old format uses integers
                    .Frames(frame) = tempint
                    
                    If .Frames(frame) <= 0 Or .Frames(frame) > totalGrhs Then
                        GoTo ErrorHandler
                    End If
                Next frame
                
                Get handle, , tempint
                
                'Convert old speed to new one (time based)!
                .Speed = CSng(tempint) * .NumFrames * 1000 / 18
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , tempint
                
                'Old format used ints, not longs.
                .FileNum = tempint
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                    
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = grh
            End If
        End With
        
        'Get Next Grh Number
        Get handle, , grh
    Loop
    
    'Trim array
    ReDim Preserve GrhData(1 To max) As GrhData
    
    LoadGrhDataOld = True
Exit Function

ErrorHandler:
    LoadGrhDataOld = False
End Function

''
' Finds out the number of grhs for the old file format
'
' @param    path    The path to the folder in which the init file is stored.
'
' @return   The number of grhs that can exist at most.

Private Function NumberOfGrhs(ByVal path As String) As Long
    Dim N As Integer
    Dim GameIni As tGameIni
    Dim MiCabecera As tCabecera
    
    N = FreeFile
    
    Open path & OLD_FORMAT_INIT_FILE For Binary As #N
    
    Get N, , MiCabecera
    
    Get N, , GameIni
    
    Close N
    
    NumberOfGrhs = GameIni.NumeroDeBMPs
End Function

''
' Loads grh data using the new file format.
'
' @param    handle      Handle to the open file containing the grh data.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataNew(ByVal handle As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim frame As Long
    Dim grhCount As Long
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , grh
        
        With GrhData(grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For frame = 1 To .NumFrames
                    Get handle, , .Frames(frame)
                    If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = grh
            End If
        End With
    Wend
    
    LoadGrhDataNew = True
Exit Function

ErrorHandler:
    LoadGrhDataNew = False
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise (data can't be stored in the old file format, use new one).

Public Function SaveGrhDataOld(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            'Index too big for this file format?
            If i > &H7FFF& Then
                Close handle
                Kill path
                Exit Function
            End If
            
            Put handle, , CInt(i)
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Put handle, , CInt(.Frames(frame))
                    Next frame
                    
                    Put handle, , CInt(.Speed * 0.018 / .NumFrames)
                Else
                    'Write in normal GRH data
                    Put handle, , CInt(.FileNum)
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataOld = True
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise.

Public Function SaveGrhDataNew(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim frame As Long
    Dim i As Long
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    'Increment file version
    fileVersion = fileVersion + 1
    
    Put handle, , fileVersion
    
    Put handle, , CLng(UBound(GrhData()))
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            Put handle, , i
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Put handle, , .Frames(frame)
                    Next frame
                    
                    Put handle, , .Speed
                Else
                    'Write in normal GRH data
                    Put handle, , .FileNum
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataNew = True
End Function

Public Function SaveInAllGraphics(ByVal path As String, Optional ByVal NewFormat As Boolean = True) As Boolean
Dim i As Long
Dim j As Long
Dim TempGrhData() As GrhData
Dim Temp As String

On Error GoTo ErrHandler
    TempGrhData = GrhData
    Temp = GRH_DAT_FILE
    
    For j = 1 To 3
        Call LoadGrhData(Config.initPath, j)
        
        ReDim Preserve GrhData(1 To UBound(TempGrhData)) As GrhData
        
        For i = 1 To UBound(GrhData)
            If InStr(1, GRAFICOS_DIFERENTES, "*" & i & "*") = 0 Then
                GrhData(i) = TempGrhData(i)
            End If
        Next i
        
        If NewFormat Then
            Call SaveGrhDataNew(path)
        Else
            Call SaveGrhDataOld(path)
        End If
    Next j
    
    GrhData = TempGrhData
    GRH_DAT_FILE = Temp
    frmMain.Caption = "Indexador Alkon (" & GRH_DAT_FILE & ")"
    
    SaveInAllGraphics = True
    Exit Function
    
ErrHandler:
    SaveInAllGraphics = False
End Function

Public Sub CargarCabezas(ByVal path As String)
    Dim N As Integer
    Dim i As Long
    Dim NumHeads As Integer
    Dim MiCabecera As tCabecera
    
    N = FreeFile()
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    Open path & "Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumHeads
    
    'Resize array
    ReDim MisCabezas(1 To NumHeads) As tIndiceCabeza
    
    For i = 1 To NumHeads
        Get #N, , MisCabezas(i)
    Next i
    
    Close #N
End Sub

Public Sub CargarCascos(ByVal path As String)
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer
    Dim MiCabecera As tCabecera
    
    N = FreeFile()
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    Open path & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim MisCascos(1 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , MisCascos(i)
    Next i
    
    Close #N
End Sub

Public Sub CargarCuerpos(ByVal path As String)
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MiCabecera As tCabecera
    
    N = FreeFile()
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    Open path & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim MisCuerpos(1 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
    Next i
    
    Close #N
End Sub

Public Sub CargarFxs(ByVal path As String)
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    Dim MiCabecera As tCabecera
    
    N = FreeFile()
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    Open path & "Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Public Function SaveHeads(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim i As Long
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & "Cabezas.ind"
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , CInt(UBound(MisCabezas))
    
    For i = 1 To UBound(MisCabezas)
        Put handle, , MisCabezas(i)
    Next i
    
    Close handle
    
    SaveHeads = True
End Function

Public Function SaveHelmets(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim i As Long
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & "Cascos.ind"
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , CInt(UBound(MisCascos))
    
    For i = 1 To UBound(MisCascos)
        Put handle, , MisCascos(i)
    Next i
    
    Close handle
    
    SaveHelmets = True
End Function

Public Function SaveBodies(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim i As Long
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & "Personajes.ind"
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , CInt(UBound(MisCuerpos))
    
    For i = 1 To UBound(MisCuerpos)
        Put handle, , MisCuerpos(i)
    Next i
    
    Close handle
    
    SaveBodies = True
End Function

Public Function SaveFXs(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim i As Long
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & "Fxs.ind"
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , CInt(UBound(FxData))
    
    For i = 1 To UBound(FxData)
        Put handle, , FxData(i)
    Next i
    
    Close handle
    
    SaveFXs = True
End Function
