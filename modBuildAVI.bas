Attribute VB_Name = "modBuildAVI"
Option Explicit

Public OutputAVIName   As String


Public Sub BUILD_AVI(fPATH As String, OutFPS As Long, Me_hWnd As Long, ByRef INFO As Label)

    OutputAVIName = ""


    Dim s              As String

    Dim fLIST()        As String
    Dim C              As Long

    s = Dir(fPATH & "*.bmp")

    If s = "" Then Exit Sub

    ReDim Preserve fLIST(1)
    Do
        fLIST(C) = fPATH & s
        C = C + 1
        ReDim Preserve fLIST(0 To C)
        s = Dir
    Loop While s <> ""

    '----------------------------------------------------------------------------------------
    Dim file           As cFileDlg
    Dim InitDir        As String
    Dim szOutputAVIFile As String
    Dim res            As Long
    Dim pfile          As Long    'ptr PAVIFILE
    Dim bmp            As cDIB
    Dim ps             As Long    'ptr PAVISTREAM
    Dim psCompressed   As Long    'ptr PAVISTREAM
    Dim strhdr         As AVI_STREAM_INFO
    Dim BI             As BITMAPINFOHEADER
    Dim opts           As AVI_COMPRESS_OPTIONS
    Dim pOpts          As Long
    Dim I              As Long
    Dim I2             As Long

    Dim Perc           As Single

    Dim EXTRA          As Integer    'Extra Frame

    Debug.Print
    Set file = New cFileDlg
    'get an avi filename from user
    With file
        .InitDirectory = App.Path & "\VIDEO\"
        .DefaultExt = "avi"
        .DlgTitle = "Choose a filename to save AVI to..."
        .Filter = "AVI Files|*.avi"
        .OwnerHwnd = frmMain.hWnd
    End With
    szOutputAVIFile = "MyAVI.avi"
    If file.VBGetSaveFileName(szOutputAVIFile) <> True Then Exit Sub


    OutputAVIName = szOutputAVIFile
    'Stop

    '    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    If bmp.CreateFromFile(fLIST(1)) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.title
        GoTo error
    End If
    'Stop

    '   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&     '// default AVI handler
        .dwScale = 1
        .dwRate = OutFPS     '* (Val(cmbEXTRA) + 1) '// fps
        .dwSuggestedBufferSize = bmp.SizeImage    '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)    '// rectangle for stream
    End With

    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

    '   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me_hWnd, _
                         ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                         1, _
                         ps, _
                         pOpts)    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then         'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If

    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error

    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With

    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    '   Now write out each video frame
    I2 = 0
    For I = 0 To C - 1

        Perc = I / (C - 1) * 100
        INFO.BackColor = RGB(255 - Perc * 2.55, Perc * 2.55, 0)
        INFO = "Frame " & I & " of " & C - 1 & "  (" & Int(Perc) & "%)"
        DoEvents


        'For EXTRA = 0 To Val(cmbEXTRA)

        bmp.CreateFromFile (fLIST(I))    'load the bitmap (ignore errors)


        res = AVIStreamWrite(psCompressed, _
                             I2, _
                             1, _
                             bmp.PointerToBits, _
                             bmp.SizeImage, _
                             AVIIF_KEYFRAME, _
                             ByVal 0&, _
                             ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
        'Show user feedback
        'imgPreview.Picture = LoadPicture(lstDIBList.Text)
        'imgPreview.Refresh
        'lblStatus = "Frame number " & i & " saved"
        'lblStatus.Refresh
        I2 = I2 + 1

        'Next EXTRA


    Next


    INFO = "Avi file Created!"



error:
    '   Now close the file
    Set file = Nothing
    Set bmp = Nothing

    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.title
    End If

    'Stop


End Sub

