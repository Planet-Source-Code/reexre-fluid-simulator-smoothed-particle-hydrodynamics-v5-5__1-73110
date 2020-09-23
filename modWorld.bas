Attribute VB_Name = "modWorld"
Option Explicit


Public MaxX            As Single
Public MaxY            As Single

Public MaxXStick       As Single
Public MaxYStick       As Single


Public W               As New clsFLuid

Public Const PI        As Single = 3.14159265358979
Public Const PI2       As Single = 6.28318530717959


Public Const AirResistence As Single = 0.999    '0.9985    '0.996    ' 0.994

Public Const BounceNess As Single = 0.85    ' 0.994

Public Const Gravity = 0.006    '0.0075    ' 0.0125        '0.05        '0.1 '0.125  v2       '0.25
Public GravDirX        As Single
Public GravDirY        As Single
Public GravANGFrom     As Single
Public GravANGTo       As Single
Public GravANG         As Single
Public StepGravA       As Single
Public GravMag         As Single
Public GravMagFrom     As Single
Public GravMagTo       As Single
Public StepGravM       As Single




Public Frame           As Long
Public pFrame          As Long
Public Running         As Boolean
Public Const EveryFrame = 5  '6  '5  '4


Public Const Sepa      As String = vbCrLf

Public INTRO           As String

Public Wx              As Long
Public Wy              As Long

Public BLUR            As New clsBLUR


Public Sub CreateINTRO(ByRef P As PictureBox, fStep As Single)

    Dim I              As Long
    Dim s()            As String

    Dim KH             As Single
    Dim H              As Single
    Dim oH             As Single


    INTRO = Now & Sepa & _
            "Based on this Paper:  http://www.iro.umontreal.ca/labs/infographie/papers/Clavet-2005-PVFS/pvfs.pdf" & Sepa & _
            "Author: Roberto Mior - reexre@gmail.com" & Sepa & "PARAMETERS" & Sepa & INTRO

    s = Split(INTRO, Sepa)

    For I = 0 To UBound(s)
        GoSub PrintToVideo2
        DoEvents
    Next
    GoTo ENDSUB

    '******
PrintToVideo:

    P.FontName = "Courier New"

    P.FontBold = True

    KH = 50
    oH = 0
    For H = 1 To KH Step fStep


        If oH = 0 Then oH = 1
        P.Cls
        P.ForeColor = RGB(0, 200, 0)
        P.Font.Size = oH * 1.2
        P.CurrentX = MaxX * 0.5 - oH * 0.5 * Len(s(I))
        P.CurrentY = MaxY * 0.5 - oH * 1
        P.Print s(I)

        P.ForeColor = vbGreen
        P.Font.Size = H * 1.2
        P.CurrentX = MaxX * 0.5 - H * 0.5 * Len(s(I))
        P.CurrentY = MaxY * 0.5 - H * 1
        P.Print s(I)
        Frame = Frame + 1
        P.Refresh
        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        oH = H
    Next
    For H = KH To 1 Step -fStep * 3
        If oH = 0 Then oH = 1
        P.Cls
        P.ForeColor = RGB(0, 200, 0)
        P.Font.Size = oH * 1.2
        P.CurrentX = MaxX * 0.5 - oH * 0.5 * Len(s(I))
        P.CurrentY = MaxY * 0.5 - oH * 1
        P.Print s(I)

        P.ForeColor = vbGreen
        P.Font.Size = H * 1.2
        P.CurrentX = MaxX * 0.5 - H * 0.5 * Len(s(I))
        P.CurrentY = MaxY * 0.5 - H * 1
        P.Print s(I)
        Frame = Frame + 1
        P.Refresh
        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        oH = H
    Next

    '******
    Return


    '**************************************************************
PrintToVideo2:
    P.FontName = "Courier New"

    P.FontBold = True

    P.ForeColor = RGB(0, 200, 0)
    P.Font.Size = 20

    P.ForeColor = vbGreen
    For H = MaxX To MaxX * 0.5 - 0.5 * Len(s(I)) * P.FontSize Step -fStep
        P.Cls
        P.CurrentX = H
        P.CurrentY = MaxY * 0.5
        P.Print s(I)
        Frame = Frame + 1
        P.Refresh
        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        DoEvents
    Next

    For H = MaxX * 0.5 - 0.5 * Len(s(I)) * P.FontSize To -Len(s(I)) * P.FontSize Step -fStep * 3
        P.Cls
        P.CurrentX = H
        P.CurrentY = MaxY * 0.5
        P.Print s(I)
        Frame = Frame + 1
        P.Refresh
        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        DoEvents
    Next


    Return
    '**************************************************************

ENDSUB:

End Sub

Public Sub CreateINTRO2(ByRef P As PictureBox, fStep As Single)


    Dim s()            As String
    Dim Y              As Single
    Dim I              As Long
    Dim S1             As String
    Dim S2             As String
    Dim SI             As String



    INTRO = "[VERSION " & App.Major & "." & App.Minor & "]  " & Now & Sepa & _
            "Based on this Paper:  http://www.iro.umontreal.ca/labs/infographie/papers/Clavet-2005-PVFS/pvfs.pdf" & Sepa & _
            "Author: Roberto Mior - reexre@gmail.com" & Sepa & "PARAMETERS" & Sepa & INTRO



    P.FontName = "Courier New"
    P.FontBold = True
    P.Font.Size = 14
    P.ForeColor = vbWhite

    s = Split(INTRO, Sepa)
    INTRO = ""
    For I = 0 To UBound(s)
        S1 = s(I)
        SI = S1
        S2 = ""
        While Len(S1) * P.Font.Size > MaxX * 1.2
            S1 = Left$(S1, MaxX * 1.2 / P.Font.Size)
            INTRO = INTRO & S1 & vbCrLf
            S2 = "       " & Right$(SI, Len(SI) - Len(S1))
            S1 = S2
            SI = S1
        Wend
        INTRO = INTRO & S1 & vbCrLf

    Next

    '-------------------------------------------------
    s = Split(INTRO, Sepa)
    Y = MaxY

    For Y = MaxY To -1 Step -fStep
        P.Cls
        P.CurrentY = Y
        P.Print INTRO
        Frame = Frame + 1
        P.Refresh
        'If frmMain.sMotionBlurred <> 0 Then
        '    BLUR.GetSource P.Image.Handle
        '    BLUR.BLUR frmMain.sMotionBlurred
        '    BLUR.TagertTO P.Image.Handle
        '    P.Refresh
        '    DoEvents
        'End If

        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        DoEvents
    Next

    For Y = 0 To Val(frmMain.txtFPS) * 2
        Frame = Frame + 1
        'If frmMain.sMotionBlurred <> 0 Then
        '    P.Cls
        '    P.Print INTRO
        '    BLUR.GetSource P.Image.Handle
        '    BLUR.BLUR frmMain.sMotionBlurred * 0.75
        '    BLUR.TagertTO P.Image.Handle
        '    P.Refresh
        '    DoEvents
        'End If

        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        DoEvents
    Next

    For Y = -1 To -UBound(s) * P.Font.Size * 1.5 Step -fStep * 2.5
        P.Cls
        P.CurrentY = Y
        P.Print INTRO
        Frame = Frame + 1
        P.Refresh
        'If frmMain.sMotionBlurred <> 0 Then
        '    BLUR.GetSource P.Image.Handle
        '    BLUR.BLUR 5
        '    BLUR.TagertTO P.Image.Handle
        '    P.Refresh
        '    DoEvents
        'End If

        SavePicture P.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
        DoEvents
    Next
  
End Sub


Public Sub DoLOOP()

    With frmMain

        Do

            If Abs(GravMagTo - GravMag) > Abs(StepGravM) Then
                GravMag = GravMag + StepGravM
                W.GX = GravDirX * Gravity * GravMag
                W.GY = GravDirY * Gravity * GravMag
                .LineG.X2 = .LineG.X1 + GravDirX * (.picG.ScaleWidth \ 2) * GravMag
                .LineG.Y2 = .LineG.Y1 + GravDirY * (.picG.ScaleWidth \ 2) * GravMag
            End If

            If Abs(AngleDiff(GravANG, GravANGTo)) > Abs(StepGravA) Then

                GravANG = GravANG + StepGravA
                GravDirX = Cos(GravANG)
                GravDirY = Sin(GravANG)

                .LineG.X2 = .LineG.X1 + GravDirX * (.picG.ScaleWidth \ 2) * GravMag
                .LineG.Y2 = .LineG.Y1 + GravDirY * (.picG.ScaleWidth \ 2) * GravMag
                W.GX = GravDirX * Gravity * GravMag
                W.GY = GravDirY * Gravity * GravMag


            End If




            Frame = Frame + 1

            If Frame Mod EveryFrame = 0 Then
                .DRAWALL

                If .sMotionBlurred <> 0 Then
                    '
                    BLUR.GetSource .PIC.Image.Handle
                    BLUR.BLUR .sMotionBlurred
                    BLUR.TagertTO .PIC.Image.Handle

                    .PIC.Refresh
                    DoEvents
                End If


                If .hRndDIR <> 0 Then
                    If (Frame \ EveryFrame) Mod .hRndDIR = 0 Then
                        .cmdGRAV_Click (Int(Rnd * 9))
                        DoEvents
                    End If

                End If

                .PIC.Refresh
                If .chSaveFrame Then SavePicture .PIC.Image, App.Path & "\VideoFrames\P" & Format(Frame, "0000000") & ".bmp"
                .lVIDEOINFO = Frame & " [" & Frame \ EveryFrame & "]"
                .lPTS = "Pts:" & W.NP & " Pairs:" & W.N_Springs
                DoEvents
            End If

            W.AAASimulation2

            '.drawall



            ' W.DRAWMetaBallCont .PIC.Hdc



            If Rnd < .hFaucet * 0.01 Then
                W.ADDPoint Wx + Rnd * 8, Wy + Rnd * 8
            End If

        Loop While Running
    End With

End Sub



Public Function Atan2(ByVal dX As Single, ByVal dY As Single) As Single
'This Should return Angle

    Dim theta          As Single

    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            theta = 0#
        ElseIf (dY > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(dY / dX)

        If (dX < 0) Then
            If (dY >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If


    Atan2 = theta

    If Atan2 < 0 Then Atan2 = Atan2 + PI * 2

End Function

Public Function AngleDiff(A1 As Single, A2 As Single) As Single
'double difference = secondAngle - firstAngle;
'while (difference < -180) difference += 360;
'while (difference > 180) difference -= 360;
'return difference;

    AngleDiff = A2 - A1
    While AngleDiff < -PI
        AngleDiff = AngleDiff + PI2
    Wend
    While AngleDiff > PI
        AngleDiff = AngleDiff - PI2
    Wend


End Function

