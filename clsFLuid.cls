VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFLuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' BASED ON this Paper
' http://www.iro.umontreal.ca/labs/infographie/papers/Clavet-2005-PVFS/pvfs.pdf
'
' Author: Roberto Mior       reexre@gmail.com
'
'
' If you use or modify this code remember to cite the author
'


Option Explicit


Private Type TPair

    I                  As Long
    J                  As Long

    IsPair             As Boolean

    '    Rij                As single    'Distance Between P(i) and p(j)
    '    DistX              As single    'XDistance Between P(i) and p(j)
    '    DistY              As single    'YDistance Between P(i) and p(j)

    NrmlzdDX           As Single    'Normalized DistX
    NrmlzdDY           As Single    'Normalized DistX

    OneMinusQ1         As Single    '1-Rij/H
    OneMinusQ2         As Single    '(1-Rij/H)^2
    OneMinusQ3         As Single    '(1-Rij/H)^3

    SpringRestL        As Single    'Current Spring Length


End Type


Private Type tP

    X                  As Single
    Y                  As Single
    OldX               As Single
    OldY               As Single
    Vx                 As Single
    Vy                 As Single

    Density            As Single
    DensityNEAR        As Single
    Pressure           As Single
    PressureNEAR       As Single
    XGrid              As Long
    YGrid              As Long

    MoveX              As Single
    MoveY              As Single


End Type

Private P()            As tP
Public NP              As Long

Private PAIR()         As TPair
Public Npairs          As Long
Public MaxNPairs As Long


Private IJtoPairIndex() As Long


Public H               As Single    'Interaction Radius
Public Hcell           As Single
Public InvH As Single


Public HSQ             As Single    'Interaction Radius Squared
Public SpringL         As Single    'Spring Normal Length
Public Dstick          As Single    'Stick Distance


Public N_Springs       As Long

Public RestDensity     As Single
Public KStiff          As Single
Public KNear           As Single
Public KspringELA      As Single
Public KspringINC      As Single

Public VISCOLinear     As Single
Public VISCOSquared    As Single


Public dT              As Single
Private dT2            As Single
Public InvDT            As Single


Public GX              As Single
Public GY              As Single

Public KPlastYield     As Single    '0-0.2
Public KPlasticity     As Single    'Plasticity

Public KStickness      As Single

Public Initialized     As Boolean

Private Dtable()       As Single

Private YCell          As Long


'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, _
 source As Any, ByVal Bytes As Long)

Private Cells()        As New Collection



Public Property Let SetX(I As Long, Value As Single)
    P(I).X = Value
End Property
Public Property Let SetY(I As Long, Value As Single)
    P(I).Y = Value
End Property
Public Property Let SetOldX(I As Long, Value As Single)
    P(I).OldX = Value
End Property
Public Property Let SetOldY(I As Long, Value As Single)
    P(I).OldY = Value
End Property
Public Property Get GetX(I As Long) As Single
    GetX = P(I).X
End Property
Public Property Get GetY(I As Long) As Single
    GetY = P(I).Y
End Property
Public Property Get GetOldX(I As Long) As Single
    GetOldX = P(I).OldX
End Property
Public Property Get GetOldY(I As Long) As Single
    GetOldY = P(I).OldY
End Property
Public Property Get GetVX(I As Long) As Single
    GetVX = P(I).Vx
End Property
Public Property Get GetVY(I As Long) As Single
    GetVY = P(I).Vy
End Property
Public Property Let SetVX(I As Long, Value As Single)
    P(I).Vx = Value
End Property
Public Property Let SetVY(I As Long, Value As Single)
    P(I).Vy = Value
End Property
Public Property Let SetXgrid(I As Long, Value As Single)
    P(I).XGrid = Value
End Property
Public Property Let SetYgrid(I As Long, Value As Single)
    P(I).YGrid = Value
End Property
Public Property Get GetXgrid(I As Long) As Long
    GetXgrid = P(I).XGrid
End Property
Public Property Get GetYgrid(I As Long) As Long
    GetYgrid = P(I).YGrid
End Property
'Public Property Get GetSpringRestL(I, J) As single
'    GetSpringRestL = P(I).SpringRestL(J)
'End Property
Public Property Get GetDensity(I As Long) As Single
    GetDensity = P(I).Density

End Property

Public Sub AAAINITFluid(Optional Npoints = 0, _
                        Optional IterRadius = 6, _
                        Optional DeltaTime = 1.3, _
                        Optional Stiffness = 0.055, _
                        Optional K_Near = 0.02, _
                        Optional Rest_Density = 0.6, _
                        Optional Plasticity = 0.0005, _
                        Optional PlastYield = 0.1, _
                        Optional ViscosityLinear = 0.055, _
                        Optional ViscositySquared = 0.0001, _
                        Optional SpringElasticity = 0.0017, _
                        Optional Stickiness = 0.2)

     'Optional IterRadius = 6, _
     'Optional DeltaTime = 1.25, _
     'Optional Stiffness = 0.055, _
     'Optional K_Near = 0.02, _
     'Optional Rest_Density = 0.6, _
     'Optional Plasticity = 0.0005, _
     'Optional PlastYield = 0.1, _
     'Optional ViscosityLinear = 0.055, _
     'Optional ViscositySquared = 0.0001, _
     'Optional SpringElasticity = 0.0017, _
     'Optional Stickiness = 0.175)




    ' Optional ViscositySquared = 0.0001, _

      'BLOBBY
    'Optional IterRadius = 6, _
     '        Optional DeltaTime = 1.125, _
     '        Optional Stiffness = 0.165, _
     '        Optional K_Near = 0.0125, _
     '        Optional Rest_Density = 0.5, _
     '        Optional Plasticity = 0.001, _
     '        Optional PlastYield = 0.1, _
     '        Optional ViscosityLinear = 0.025, _
     '        Optional ViscositySquared = 0.001, _
     '        Optional SpringElasticity = 0.0525, _
     '        Optional Stickiness = 0.3)


    'viscolenaer 0.025
    'Plasticity = 0.00025

    'Tweaking with this parameters may produce better results...
    'Please Help me to find better values,
    'If you got good ones email me at reexre@gmail.com or post them as PSC Entry Comment
    '( I'm gone crazy trying to find best values ;) )

    '
    'http://www.iro.umontreal.ca/labs/infographie/papers/Clavet-2005-PVFS/pvfs.pdf


    Dim I          As Long
    Dim J          As Long


    H = IterRadius
    HSQ = H * H
    InvH = 1 / H


    Hcell = H * 2


    Dstick = H * 1.05             ' 0.9 '9 ' 0.9        ' 0.9         ' 0.8


    INTRO = "Interaction Radius=" & IterRadius & Sepa & "Delta Time=" & DeltaTime & Sepa & _
            "Stiffness=" & Stiffness & Sepa & "K Near=" & K_Near & Sepa & _
            "Density=" & Rest_Density & Sepa & "Plasticity=" & Plasticity & Sepa & _
            "Plasticity Yield=" & PlastYield & Sepa & "ViscosityLinear=" & ViscosityLinear & _
            Sepa & "ViscositySquared=" & ViscositySquared & "       (DStick=" & Dstick & ")" & Sepa & _
            "SpringElasticity=" & SpringElasticity & _
            "      (Garvity=" & Gravity & ")" & Sepa & _
            "Stickiness=" & Stickiness & "               (AirResistence=" & AirResistence & ")"


    CreateINTRO2 frmMain.PIC, MaxY / 25    '8


    NP = Npoints

    ReDim P(NP)

    SpringL = H * 0.5

    VISCOLinear = ViscosityLinear
    VISCOSquared = ViscositySquared

    dT = DeltaTime
    dT2 = dT * dT
    InvDT = 1 / dT


    KPlasticity = dT * Plasticity
    KPlastYield = PlastYield

    KStiff = Stiffness
    KNear = K_Near
    RestDensity = Rest_Density

    KspringELA = dT2 * SpringElasticity

    'KStickness = 1 / (dT * Stickiness)
    KStickness = (dT * Stickiness)


    Npairs = 1
    ReDim PAIR(1)
    PAIR(1).IsPair = False



    ReDim Dtable(MaxX, MaxX)
    For I = 0 To MaxX
        For J = 0 To MaxX
            Dtable(I, J) = Sqr(I * I + J * J)
        Next
    Next


    YCell = MaxX \ Hcell
    'ReDim Cells((MaxX \ Hcell) * (MaxY \ Hcell))
    ReDim Cells(MaxX \ Hcell, MaxY \ Hcell)
    Initialized = True

End Sub




Private Function ptDistance(ByRef II As Long, ByRef JJ As Long, ByRef dX As Single, ByRef dY As Single) As Single

    dX = P(JJ).X - P(II).X
    dY = P(JJ).Y - P(II).Y
    ptDistance = Sqr(dX * dX + dY * dY)


End Function
Private Function ptDistanceSQUARED(ByRef II As Long, ByRef JJ As Long, ByRef dX As Single, ByRef dY As Single) As Single

    dX = P(JJ).X - P(II).X
    dY = P(JJ).Y - P(II).Y
    ptDistanceSQUARED = (dX * dX + dY * dY)

End Function

'float DistanceToPointSquared(float X, float Y)
'    {
'      float vx = x1-X, vy = y1-Y, ux = x2-x1, uy = y2-y1;
'      float length = ux*ux+uy*uy;
'
'      float det = (-vx*ux)+(-vy*uy); //if this is < 0 or > length then its outside the line segment
'      if(det<0 || det>length)
'      {
'        ux=x2-X;
'        uy=y2-Y;
'        return std::min<float>(vx*vx+vy*vy, ux*ux+uy*uy);
'      }
'
'      det = ux*vy-uy*vx;
'      return (det*det)/length;
'    }
Public Function ptDistFromLine(PtI As Long, X1 As Single, Y1 As Single, _
                               X2 As Single, Y2 As Single) As Single
    Dim Vx             As Single
    Dim Vy             As Single
    Dim UX             As Single
    Dim UY             As Single
    Dim L              As Single
    Dim Det            As Single
    ptDistFromLine = 99999

    Vx = X1 - P(PtI).X
    Vy = Y1 - P(PtI).Y
    UX = X2 - X1
    UY = Y2 - Y1

    L = UX * UX + UY * UY
    Det = (-Vx * UX) + (-Vy * UY)
    'if this is < 0 or > length then its outside the line segment
    If (Det < 0 Or Det > L) Then
        'ux=x2-X;
        'uy=y2-Y;
        'return std::min<float>(vx*vx+vy*vy, ux*ux+uy*uy);
    Else
        Det = UX * Vy - UY * Vx
        'return (det*det)/length;
        ptDistFromLine = Sqr((Det * Det) / L)
    End If



End Function
Public Function PtDistFromLine2(PtI As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByRef RetX As Single, ByRef RetY As Single) As Single
' p(pti).x,p(pti).y is the point to test.
' x1,y1,x2,y2 is the line to check distance.
'
' Returns distance from the line, or if the intersecting point on the line nearest
' the point tested is outside the endpoints of the line, the distance to the
' nearest endpoint.
'
' Returns 99999 on 0 denominator conditions.
'Dim RetX As single, RetY As single ' intersecting point

    Dim LineMag        As Single
    Dim LineMagSQRD    As Single

    Dim U              As Single

    Dim XX             As Single
    Dim YY             As Single
    XX = X2 - X1
    YY = Y2 - Y1

    PtDistFromLine2 = 99999

    LineMagSQRD = (XX * XX + YY * YY)    'Distance(x1, y1, x2, y2)
    If LineMagSQRD < 0.001 Then PtDistFromLine2 = 99999: Exit Function

    U = (((P(PtI).X - X1) * (XX)) + ((P(PtI).Y - Y1) * (YY)))
    U = U / (LineMagSQRD)
    If U < 0.0001 Or U > 1 Then
        '// closest point does not fall within the line segment, take the shorter distance
        '// to an endpoint
        ' RetX = Distance(P(PtI).x, P(PtI).y, x1, y1)
        ' RetY = Distance(P(PtI).x, P(PtI).y, x2, y2)
        ' If RetX > RetY Then DistancePointLine = RetY Else DistancePointLine = RetX
    Else
        ' Intersecting point is on the line, use the formula
        RetX = X1 + U * XX
        RetY = Y1 + U * YY
        PtDistFromLine2 = Distance(P(PtI).X, P(PtI).Y, RetX, RetY)
    End If

End Function




Public Sub ADDPoint(X As Single, Y As Single, Optional PairsReset As Boolean = True)
    Dim I              As Long

    NP = NP + 1
    ReDim Preserve P(NP)

    With P(NP)
        .X = X
        .Y = Y
        .OldX = X
        .OldY = Y
        .XGrid = .X \ H
        .YGrid = .Y \ H

    End With

    If PairsReset Then ResetPairs


End Sub
Public Sub RemovePoint(II As Long)
    Dim I              As Long
    Dim J              As Long

    For I = II To NP - 1
        P(I) = P(I + 1)
    Next
    NP = NP - 1


    ResetPairs



End Sub


Public Sub DRAW(picHdc As Long, DrawSprings As Boolean)
    Dim I              As Long
    Dim RedColor       As Long
    Dim GreenColor     As Long
    Dim BlueColor      As Long
    Dim color          As Long

    Dim X1             As Long
    Dim Y1             As Long
    Dim X2             As Long
    Dim Y2             As Long


    If DrawSprings Then
        color = RGB(100, 127, 255)
        For I = 1 To Npairs
            With PAIR(I)
                If .IsPair Then
                    'If .SpringRestL > 0 Then

                    'BlueColor = 255 * .SpringRestL / SpringL
                    'RedColor = 255 - BlueColor * 0.5
                    'FastLine picHdc, X1, Y1, X2, Y2, 1, RGB(RedColor, 127, BlueColor)

                    FastLine picHdc, P(.I).X \ 1, P(.I).Y \ 1, P(.J).X \ 1, P(.J).Y \ 1, 1, color

                End If
            End With

        Next
    End If
    'GoTo SKIP

    For I = 1 To NP
        With P(I)

            'RedColor = 2000 * .Pressure
            RedColor = 50 + 1500 * .Pressure    'NEAR
            'RedColor = 100 * .Density
            If RedColor < 0 Then RedColor = 0

            'GreenColor = 120 + (Abs(.Vx) + Abs(.Vy)) * 60
            'GreenColor = 100 + Sqr(.Vx * .Vx + .Vy * .Vy) * 70
            GreenColor = 100 + (.Vx * .Vx + .Vy * .Vy) * 20

            'If I Mod 2 = 0 Then
            color = RGB(RedColor, GreenColor, 255)

            ''        FastLine picHdc, CLng(.X), CLng(.Y), _
                     CLng(.OldX), CLng(.OldY), 3, GreenColor    'vbWhite

            'FastLine picHdc, .X \ 1, .Y \ 1, _
                     .OldX \ 1, .OldY \ 1, 3, color    'vbWhite
            ''        MyCircle picHdc, CLng(P(I).X), CLng(P(I).Y), 2, 3, GreenColor
            'X1 = .X
            'Y1 = .Y
            'SetPixel picHdc, X1, Y1, color
            'SetPixel picHdc, X1, Y1 + 1, color
            'X1 = X1 + 1
            'SetPixel picHdc, X1, Y1, color
            'SetPixel picHdc, X1, Y1 + 1, color
            
            bLOCK picHdc, .X - 1, .Y - 1, 2, color
            
        End With
    Next


SKIP:
    X1 = 35
    Y1 = 35
    X2 = X1 + GravDirX * 30 * GravMag
    Y2 = Y1 + GravDirY * 30 * GravMag

    FastLine picHdc, X1, Y1, X2, Y2, 1, vbYellow
    MyCircle picHdc, X2, Y2, 2, 2, vbYellow
    MyCircle picHdc, X1, Y1, 30, 1, vbYellow



End Sub

Public Sub DRAWMetaBall(picHdc As Long)
    Dim MAX            As Single
    Dim MIN            As Single
    Dim X1             As Long
    Dim X2             As Long
    Dim Y1             As Long
    Dim Y2             As Long

    Dim X              As Single
    Dim Y              As Single
    Dim SUM            As Single
    Dim SUMp           As Single
    Dim SumV           As Single

    Dim I              As Long
    Dim D              As Single
    Dim C              As Single
    Dim dX             As Single
    Dim dY             As Single

    Dim color          As Long

    MAX = 30                 '25
    MIN = 20                 '3 '10
    For X = 0 To MaxX Step 2

        For Y = 0 To MaxY Step 2

            SUM = 0
            SUMp = 0
            SumV = 0
            For I = 1 To NP

                dX = Abs(P(I).X - X)

                If dX <= 24 Then
                    dY = Abs(P(I).Y - Y)
                    If dY <= 24 Then
                        'D = Dtable(DX, Dy) + 0.01
                        D = (dX * dX + dY * dY) + 0.01
                        SUM = SUM + 50 / D
                        SUMp = SUMp + (35 * P(I).Pressure)
                        SumV = SumV + Abs(P(I).Vx) + Abs(P(I).Vy)
                    End If
                End If

            Next
            SumV = SumV * 0.8    '0.6
            'MsgBox SUM
            If SUM > MIN And SUM < MAX Then
                C = (SUM - MIN) * 255 / (MAX - MIN)
                '                SetPixel picHdc, X, Y, RGB(C, 255 - C, 0)
                color = RGB(C * 0.25, C * 0.5, C)
                SetPixel picHdc, X, Y, color
                SetPixel picHdc, X + 1, Y, color
                SetPixel picHdc, X, Y + 1, color
                SetPixel picHdc, X + 1, Y + 1, color

            End If
            If SUM > MAX Then
                If SUMp < 0 Then SUMp = 0
                If SUMp > 255 Then SUMp = 255

                If SUMp > 0 Then
                    color = RGB(30 + SUMp, 128 + SumV, 255)
                Else
                    If SUMp < -51 Then SUMp = -51
                    color = RGB(30, 128 + SumV, 255 + SUMp * 5)
                End If

                SetPixel picHdc, X, Y, color
                SetPixel picHdc, X + 1, Y, color
                SetPixel picHdc, X, Y + 1, color
                SetPixel picHdc, X + 1, Y + 1, color



            End If
        Next
    Next

  X1 = 35
    Y1 = 35
    X2 = X1 + GravDirX * 30 * GravMag
    Y2 = Y1 + GravDirY * 30 * GravMag

    FastLine picHdc, X1, Y1, X2, Y2, 1, vbYellow
    MyCircle picHdc, X2, Y2, 2, 2, vbYellow
    MyCircle picHdc, X1, Y1, 30, 1, vbYellow
    DoEvents

End Sub

Public Sub DRAWMetaBallCont(picHdc As Long)
    Dim MAX            As Single
    Dim MIN            As Single
    Dim X1             As Long
    Dim X2             As Long
    Dim Y1             As Long
    Dim Y2             As Long

    Dim X              As Single
    Dim Y              As Single


    Dim I              As Long
    Dim D              As Single
    Dim C              As Single
    Dim dX             As Single
    Dim dY             As Single

    Dim color          As Long
    Dim SUM            As Single


    MAX = 70
    MIN = 45
    For X = 0 To MaxX        'Step 2

        For Y = 0 To MaxY    'Step 2
            SUM = 0
            For I = 1 To NP

                dX = Abs(P(I).X - X)
                If dX <= 24 Then
                    dY = Abs(P(I).Y - Y)
                    If dY <= 24 Then
                        'D = Dtable(DX, Dy) + 0.01
                        D = (dX * dX + dY * dY) + 0.01
                        SUM = SUM + 100 / D
                    End If
                End If

            Next

            'If SUM >= MAX Then
            '    SUM = MIN + (MAX - MIN) - (SUM - MAX)
            'End If
            'MsgBox SUM
            If SUM > MIN And SUM < MAX Then
                C = (SUM - MIN) * 255 / (MAX - MIN)
                '                SetPixel picHdc, X, Y, RGB(C, 255 - C, 0)
                'color = RGB(C * 0.25, C * 0.5, C)
                color = vbCyan
                SetPixel picHdc, X, Y, color
                SetPixel picHdc, X + 1, Y, color
                SetPixel picHdc, X, Y + 1, color
                SetPixel picHdc, X + 1, Y + 1, color

            End If

        Next
        DoEvents
    Next

    'Stop
    X1 = 35
    Y1 = 35
    X2 = X1 + GravDirX * 30 * GravMag
    Y2 = Y1 + GravDirY * 30 * GravMag

    FastLine picHdc, X1, Y1, X2, Y2, 1, vbYellow
    MyCircle picHdc, X2, Y2, 2, 2, vbYellow
    MyCircle picHdc, X1, Y1, 30, 1, vbYellow
    DoEvents
End Sub

Public Sub SaveWater(Optional Filename = "WATER")
    Dim I              As Long
    Dim J              As Long
    Dim FN             As String

    FN = App.Path & "\Scenes\" & Filename

    If Right$(FN, 4) <> ".txt" Then FN = FN + ".txt"

    If Dir(FN) <> "" Then Kill FN

    Open FN For Output As 1

    Print #1, NP

    For I = 1 To NP
        Print #1, Replace(P(I).X, ",", ".")
        Print #1, Replace(P(I).Y, ",", ".")
        Print #1, Replace(P(I).OldX, ",", ".")
        Print #1, Replace(P(I).OldY, ",", ".")
        Print #1, Replace(P(I).Vx, ",", ".")
        Print #1, Replace(P(I).Vy, ",", ".")
        Print #1, Replace(P(I).XGrid, ",", ".")
        Print #1, Replace(P(I).YGrid, ",", ".")

        Print #1, Replace(P(I).Density, ",", ".")
        Print #1, Replace(P(I).DensityNEAR, ",", ".")
        Print #1, Replace(P(I).Pressure, ",", ".")
        Print #1, Replace(P(I).PressureNEAR, ",", ".")

        'For J = 1 To NP
        '    Print #1, Replace(P(I).SpringRestL(J), ",", ".")
        'Next
    Next

    Print #1, Nlines
    For I = 1 To Nlines
        Print #1, Replace(LineO(I).X1, ",", ".")
        Print #1, Replace(LineO(I).Y1, ",", ".")
        Print #1, Replace(LineO(I).X2, ",", ".")
        Print #1, Replace(LineO(I).Y2, ",", ".")
        Print #1, Replace(LineO(I).NormalX, ",", ".")
        Print #1, Replace(LineO(I).NormalY, ",", ".")
    Next

    Print #1, Nballs
    For I = 1 To Nballs
        Print #1, Replace(BallO(I).X, ",", ".")
        Print #1, Replace(BallO(I).Y, ",", ".")
        Print #1, Replace(BallO(I).Radius, ",", ".")
    Next

    Close 1
End Sub

Public Sub LoadWater(Optional Filename = "WATER.txt")
    Dim I              As Long
    Dim J              As Long
    Dim FN             As String
    FN = App.Path & "\Scenes\" & Filename

    Open FN For Input As 1

    Input #1, NP
    ReDim P(NP)
    ResetPairs



    For I = 1 To NP
        Input #1, P(I).X
        Input #1, P(I).Y
        Input #1, P(I).OldX
        Input #1, P(I).OldY
        Input #1, P(I).Vx
        Input #1, P(I).Vy
        Input #1, P(I).XGrid
        Input #1, P(I).YGrid
        Input #1, P(I).Density
        Input #1, P(I).DensityNEAR
        Input #1, P(I).Pressure
        Input #1, P(I).PressureNEAR
        'For J = 1 To NP
        '    input #1, P(I).SpringRestL(J)
        'Next
    Next

    Input #1, Nlines
    If Nlines > 0 Then
        ReDim LineO(Nlines)
        For I = 1 To Nlines
            Input #1, LineO(I).X1
            Input #1, LineO(I).Y1
            Input #1, LineO(I).X2
            Input #1, LineO(I).Y2
            Input #1, LineO(I).NormalX
            Input #1, LineO(I).NormalY
        Next
    End If

    If Not (EOF(1)) Then     'Conpatiblity with version <4.0
        Input #1, Nballs
        If Nballs > 0 Then
            ReDim BallO(Nballs)
            For I = 1 To Nballs
                Input #1, BallO(I).X
                Input #1, BallO(I).Y
                Input #1, BallO(I).Radius
            Next
        End If
    Else
        Nballs = 0
    End If


    Close 1

End Sub
Private Function Distance(ByRef X1 As Single, ByRef Y1 As Single, _
                          ByRef X2 As Single, ByRef Y2 As Single) As Single

    Dim dX             As Single
    Dim dY             As Single

    dX = X1 - X2
    dY = Y1 - Y2
    Distance = Sqr(dX * dX + dY * dY)


End Function

Public Sub StickApply(PtI As Long, Dist As Single, npx As Single, npy As Single)

'  Optional ByRef fX As single, Optional ByRef fY As single)
    Dim NewVX          As Single
    Dim NewVY          As Single
    Dim Vx             As Single
    Dim Vy             As Single

    Dim NX             As Single
    Dim NY             As Single
    Dim N              As Single
    Dim F              As Single
    Dim fX             As Single
    Dim fy             As Single

    With P(PtI)

        '.Vx = .Vx - (1 - Dist * invh) * KStickness * (.x - npx)
        '.Vy = .Vy - (1 - Dist * invh) * KStickness * (.y - npy)

        'Attraction force  [kstickness=stickness]
        '        .x = .x - (1 - Dist / (dstick)) * KStickness * (.x - npx)
        '        .y = .y - (1 - Dist / (dstick)) * KStickness * (.y - npy)


        'Speed reduce [kstickness=1/stickness] 'Speed reduced proportionaly to speed
        'Vx = (.x - .OldX)
        'Vy = (.y - .OldY)
        'NewVX = Vx * (Dist / (dstick)) * KStickness
        'NewVY = Vy * (Dist / (dstick)) * KStickness
        'If Abs(NewVX) > Abs(Vx) Then NewVX = Vx
        'If Abs(NewVY) > Abs(Vy) Then NewVY = Vy
        '.x = .OldX + NewVX
        '.y = .OldY + NewVY

        'Speed reduce [kstickness=1/stickness]  'Speed reduced by constant
        'Vx = (.x - .OldX)
        'Vy = (.y - .OldY)
        'NewVX = Vx - Sgn(Vx) * (1 - Dist / (dstick)) * KStickness * 0.0025
        'NewVY = Vy - Sgn(Vy) * (1 - Dist / (dstick)) * KStickness * 0.0025
        'If Abs(NewVX) < 0 Then NewVX = 0
        'If Abs(NewVY) < 0 Then NewVY = 0
        '.x = .OldX + NewVX
        '.y = .OldY + NewVY

        'LIKE PAPER [kstickness=stickness]

        NX = npx - .X
        NY = npy - .Y
        N = Sqr(NX * NX + NY * NY)
        NX = NX / N
        NY = NY / N

        Vx = (.X - .OldX)
        Vy = (.Y - .OldY)

        F = KStickness * Dist * (1 - Dist / (Dstick))
        fX = F * NX
        fy = F * NY

        NewVX = Vx - fX
        NewVY = Vy - fy

        .X = .OldX + NewVX
        .Y = .OldY + NewVY



    End With
End Sub





Public Sub AAAApplyViscosity()    'ALGO 5
    Dim I              As Long
    Dim J              As Long
    Dim vRij           As Single
    Dim NdX            As Single
    Dim NdY            As Single

    Dim dVX            As Single
    Dim dVY            As Single

    Dim UU             As Single
    Dim UUQ            As Single

    Dim ViscXJ         As Single
    Dim ViscYJ         As Single
    Dim TViscXI        As Single
    Dim TViscYI        As Single

    Dim K              As Single

    Dim NPm1           As Single

    Dim MUL            As Single

    Dim dX             As Single
    Dim dY             As Single


    NPm1 = NP - 1

    For I = 1 To NPm1

        TViscXI = 0
        TViscYI = 0


        With P(I)
            For J = I + 1 To NP
                If Abs(.XGrid - P(J).XGrid) <= 1 Then
                    If Abs(.YGrid - P(J).YGrid) <= 1 Then

                        vRij = ptDistanceSQUARED(I, J, dX, dY)

                        If vRij <= HSQ Then
                        If vRij > 0 Then

                            vRij = Sqr(vRij)

                            NdX = dX / vRij
                            NdY = dY / vRij
                            dVX = (.Vx - P(J).Vx)
                            dVY = (.Vy - P(J).Vy)

                            UUQ = (dVX * dVX + dVY * dVY)
                            UU = Sqr(UUQ)

                            K = dT * (1 - vRij * InvH)

                            ViscXJ = K * (VISCOLinear * UU + VISCOSquared * UUQ) * NdX
                            ViscYJ = K * (VISCOLinear * UU + VISCOSquared * UUQ) * NdY

                            TViscXI = TViscXI + ViscXJ
                            TViscYI = TViscYI + ViscYJ

                            P(J).Vx = P(J).Vx + ViscXJ
                            P(J).Vy = P(J).Vy + ViscYJ
                        End If
                        End If
                        
                    End If
                End If

            Next

            .Vx = .Vx - TViscXI
            .Vy = .Vy - TViscYI

        End With
    Next

End Sub



Public Sub AAASimulation2()
    'http://www.iro.umontreal.ca/labs/infographie/papers/Clavet-2005-PVFS/pvfs.pdf

    Dim Rij        As Single

    Dim I          As Long
    Dim J          As Long
    Dim II         As Long

    Dim Deform     As Single

    Dim ELAsticity As Single
    Dim ElaXJ      As Single
    Dim ElaYJ      As Single
    Dim TElaXI     As Single
    Dim TElaYI     As Single

    Dim Dens       As Single
    Dim DensNear   As Single
    Dim Press      As Single
    Dim PressNear  As Single

    Dim q          As Single
    Dim Omq        As Single


    Dim Displace   As Single
    Dim dX         As Single
    Dim dY         As Single
    Dim NdX        As Single
    Dim NdY        As Single

    Dim NPm1       As Long

    Dim MUL        As Single

    Dim CurPAIR    As Long

    Dim dVX        As Single
    Dim dVY        As Single
    Dim UU         As Single
    Dim UUQ        As Single
    Dim ViscXJ     As Single
    Dim ViscYJ     As Single
    Dim K          As Single

    Dim IGridX     As Long
    Dim IGridY     As Long
    Dim Ip1 As Long
    
    NPm1 = NP - 1

    ' apply gravity (MOVED IN NEXT CYCLE BECAUSE aPPLYvISCOSITY iS iNSIDE pair lOOP)
    'For I = 1 To NP
    '    With P(I)
    '        'MUL = IIf(I Mod 2 = 0, 2, 1)
    '        .Vy = .Vy + GY * dT    '* MUL
    '        .Vx = .Vx + GX * dT    '* MUL
    '        .Vx = .Vx * AirResistence
    '        .Vy = .Vy * AirResistence
    '    End With
    'Next

    'modify velocities with pairwise viscosity impulses
    '   AAAApplyViscosity
    'Put inside for speed up


    For I = 1 To NP
        With P(I)

            '----------------------- MOVED HERE FROM Bottom
            .Vx = (.X - .OldX) * InvDT    '/ dT
            .Vy = (.Y - .OldY) * InvDT    '/ dT
            '-----------------------

            '-----------------------
            .Vy = .Vy + GY * dT
            .Vx = .Vx + GX * dT
            .Vx = .Vx * AirResistence
            .Vy = .Vy * AirResistence
            '-----------------------

            .OldX = .X
            .OldY = .Y
            .X = .X + dT * .Vx
            .Y = .Y + dT * .Vy
            .XGrid = .X \ H
            .YGrid = .Y \ H

            .MoveX = 0
            .MoveY = 0

            '-----------------------
            .Density = 0
            .DensityNEAR = 0
            '------------------------
        End With
    Next


    Dens = 0
    DensNear = 0
    '    For I = 1 To NP 'Moved up
    '        With P(I)
    '            .Density = 0
    '            .DensityNEAR = 0
    '        End With
    '    Next I


    '****************************************************************************************
    '**********************       PAIRS
    N_Springs = 0

    For I = 1 To NPm1

        IGridX = P(I).XGrid
        IGridY = P(I).YGrid

        Ip1 = I + 1
        For J = Ip1 To NP

            'CurPAIR = IJtoPairIndex(I, J)

            If Abs(IGridX - P(J).XGrid) <= 1 Then
                If Abs(IGridY - P(J).YGrid) <= 1 Then

                    Rij = ptDistanceSQUARED(I, J, dX, dY)

                    If Rij <= HSQ Then    ' And Rij > 0 Then



                        N_Springs = N_Springs + 1

                        Rij = Sqr(Rij)

                        CurPAIR = IJtoPairIndex(I, J)
                        If CurPAIR = 0 Then
                            CurPAIR = ADDPair(I, J)
                        End If

                        With PAIR(CurPAIR)

                            '.Rij = Rij
                            '.DistX = dX
                            '.DistY = dY
                            NdX = dX / Rij
                            NdY = dY / Rij
                            .NrmlzdDX = NdX
                            .NrmlzdDY = NdY


                            q = Rij * InvH
                            Omq = 1 - q

                            '*********** Viscosity ***************************
                            dVX = (P(I).Vx - P(J).Vx)
                            dVY = (P(I).Vy - P(J).Vy)

                            UUQ = (dVX * dVX + dVY * dVY)
                            UU = Sqr(UUQ)

                            'DT2 because compute Result for MoveX not VX
                            K = dT2 * (Omq) * (VISCOLinear * UU + VISCOSquared * UUQ)

                            ViscXJ = K * NdX
                            ViscYJ = K * NdY

                            P(J).MoveX = P(J).MoveX + ViscXJ
                            P(J).MoveY = P(J).MoveY + ViscYJ

                            P(I).MoveX = P(I).MoveX - ViscXJ
                            P(I).MoveY = P(I).MoveY - ViscYJ

                            '*****************************************


                            '*********** PRESSURE DENSITY


                            .OneMinusQ1 = Omq
                            .OneMinusQ2 = Omq * Omq
                            .OneMinusQ3 = .OneMinusQ2 * Omq

                            P(I).Density = P(I).Density + .OneMinusQ2    '(1)
                            P(I).DensityNEAR = P(I).DensityNEAR + .OneMinusQ3    '(4)

                            P(J).Density = P(J).Density + .OneMinusQ2
                            P(J).DensityNEAR = P(J).DensityNEAR + .OneMinusQ3
                            '****************************************



                            '**** Spring Adjustment     ALGO (4)
                            'if there is no spring i j
                            'If .SpringRestL <= 0 Then .SpringRestL = SpringL    ' Rij 'H     'Rij
                            '
                            Deform = KPlastYield * .SpringRestL
                            If Rij > SpringL + Deform Then
                                '  stretch
                                .SpringRestL = .SpringRestL + KPlasticity * (Rij - SpringL - Deform)
                            ElseIf Rij < SpringL - Deform Then
                                ' compress
                                .SpringRestL = .SpringRestL - KPlasticity * (SpringL - Deform - Rij)
                            End If
                            '*************************


                            '********************* Spring displacements    ALGO  (3)
                            If .SpringRestL <= H And .SpringRestL > 0 Then

                                ELAsticity = KspringELA * (1 - .SpringRestL * InvH) * (.SpringRestL - Rij)

                                ElaXJ = ELAsticity * NdX
                                ElaYJ = ELAsticity * NdY

                                P(J).MoveX = P(J).MoveX + ElaXJ
                                P(J).MoveY = P(J).MoveY + ElaYJ

                                P(I).MoveX = P(I).MoveX - ElaXJ
                                P(I).MoveY = P(I).MoveY - ElaYJ

                            Else
                                'Remove Spring
                                .SpringRestL = 0
                                .IsPair = False
                                IJtoPairIndex(I, J) = 0
                            End If
                            '*****************************************
                        End With
                    Else
                        'Remove Spring
                        If IJtoPairIndex(I, J) <> 0 Then
                            CurPAIR = IJtoPairIndex(I, J)
                            PAIR(CurPAIR).IsPair = 0    'False
                            IJtoPairIndex(I, J) = 0
                        End If
                    End If
                Else
                    'Remove Spring
                    If IJtoPairIndex(I, J) <> 0 Then
                        CurPAIR = IJtoPairIndex(I, J)
                        PAIR(CurPAIR).IsPair = 0    'False
                        IJtoPairIndex(I, J) = 0
                    End If
                End If
            Else
                'Remove Spring
                If IJtoPairIndex(I, J) <> 0 Then
                    CurPAIR = IJtoPairIndex(I, J)
                    PAIR(CurPAIR).IsPair = 0    'False
                    IJtoPairIndex(I, J) = 0
                End If
            End If
        Next J

    Next I
    '****************************************************************************************





    '************************** single Density Relaxation
    '************************** Algorithm (2)

    ' Moved UP
    'compute density and near-density
    '    Dens = 0
    '    DensNear = 0
    '
    '    For I = 1 To NP
    '        With P(I)
    '            .Density = 0
    '            .DensityNEAR = 0
    '        End With
    '    Next I


    'Moved inside "PAIR" cycle
    'For I = 1 To Npairs
    '    With PAIR(I)'
    '
    '            If PAIR(I).IsPair Then
    '
    '                q = .Rij * invh
    '                Omq = 1 - q
    '                .OneMinusQ1 = Omq
    '                .OneMinusQ2 = Omq * Omq
    '                .OneMinusQ3 = .OneMinusQ2 * Omq
    '
    '                P(.I).Density = P(.I).Density + .OneMinusQ2
    '                P(.I).DensityNEAR = P(.I).DensityNEAR + .OneMinusQ3
    '
    '                P(.J).Density = P(.J).Density + .OneMinusQ2
    '                P(.J).DensityNEAR = P(.J).DensityNEAR + .OneMinusQ3
    '
    '            End If'
    '
    '        End With
    '    Next

    For I = 1 To NP
        'compute pressure and near-pressure
        With P(I)
            .Pressure = KStiff * (.Density - RestDensity)    '(2)
            .PressureNEAR = KNear * .DensityNEAR    '(5)
        End With
    Next

    ' apply displacements
    For II = 1 To Npairs

        With PAIR(II)

            If PAIR(II).IsPair Then

                I = .I
                J = .J

                'Press = P(J).Pressure
                'PressNear = P(J).PressureNEAR
                'Press = P(I).Pressure
                'PressNear = P(I).PressureNEAR

                Press = (P(I).Pressure + P(J).Pressure)    ' * 0.5
                PressNear = (P(I).PressureNEAR + P(J).PressureNEAR)    '* 0.5

                Displace = dT2 * (Press * .OneMinusQ1 + PressNear * .OneMinusQ2)

                dX = Displace * .NrmlzdDX
                dY = Displace * .NrmlzdDY

                P(J).MoveX = P(J).MoveX + dX
                P(J).MoveY = P(J).MoveY + dY

                P(I).MoveX = P(I).MoveX - dX
                P(I).MoveY = P(I).MoveY - dY

            End If

        End With
    Next




    '**************************
    For I = 1 To NP
        With P(I)
            .X = .X + .MoveX
            .Y = .Y + .MoveY
        End With
    Next


    LINE_TESTCollisionAndStick
    BALL_TESTCollisionAndStick


    '*********************************************+ Moved at simulation Beginning
    '// use previous position to compute next velocity
    'For I = 1 To NP
    '    With P(I)
    '        .Vx = (.X - .OldX) * InvDT    '/ dT
    '        .Vy = (.Y - .OldY) * InvDT    '/ dT
    '    End With
    'Next

End Sub

Public Sub AAASimulation3()


End Sub

Private Function ADDPair(ByVal I As Long, ByVal J As Long) As Long

    Dim NotFree    As Boolean
    Dim pIDX       As Long

    NotFree = True

    pIDX = 1
    Do
        'If IJtoPairIndex(PAIR(pIDX).I, PAIR(pIDX).J) = 0 Then NotFree = False
        If PAIR(pIDX).IsPair = 0 Then Exit Do    'NotFree = False
        pIDX = pIDX + 1
        '    Loop While NotFree And (pIDX <= Npairs)
    Loop While (pIDX <= Npairs)
    If pIDX <= Npairs Then pIDX = pIDX + 1: NotFree = False

    If (NotFree) Then
        Npairs = Npairs + 1
        If Npairs > MaxNPairs Then MaxNPairs = Npairs: ReDim Preserve PAIR(MaxNPairs)
        pIDX = Npairs
    Else

        pIDX = pIDX - 1
    End If

    With PAIR(pIDX)
        .I = I
        .J = J
        'PAIR(CurPAIR).SpringRestL = Rij
        .SpringRestL = SpringL
        .IsPair = True

    End With
    ADDPair = pIDX
    IJtoPairIndex(I, J) = pIDX

End Function


Private Sub ResetPairs()
    Npairs = 1
    MaxNPairs = 1
    ReDim PAIR(1)
    PAIR(1).IsPair = False
    ReDim IJtoPairIndex(NP, NP)

End Sub

Public Sub ADDFilledRect(ByVal X1, ByVal Y1, ByVal X2, ByVal Y2)
    If Not (Initialized) Then Exit Sub

    Dim X              As Single
    Dim Y              As Single

    If X1 > X2 Then X = X1: X1 = X2: X2 = X
    If Y1 > Y2 Then Y = Y1: Y1 = Y2: Y2 = Y
    'Stop

    frmMain.lVIDEOINFO = "Adding Rectangle"
    DoEvents
    For X = X1 To X2 Step SpringL * 1.3
        For Y = Y1 To Y2 Step SpringL * 1.3

            ADDPoint X, Y, False

        Next
    Next

    ResetPairs
    frmMain.lVIDEOINFO = "Computing new pairs..."
    DoEvents

End Sub

Public Sub ADDFilledCircle(ByVal X1, ByVal Y1, ByVal R)

    Dim X              As Single
    Dim Y              As Single
    Dim dX             As Single
    Dim dY             As Single
    Dim D              As Single

    If Not (Initialized) Then Exit Sub


    frmMain.lVIDEOINFO = "Adding Rectangle"
    DoEvents
    For X = X1 - R To X1 + R Step SpringL * 1.3
        dX = X - X1
        For Y = Y1 - R To Y1 + R Step SpringL * 1.3

            dY = Y - Y1
            D = Sqr(dX * dX + dY * dY)

            If D <= R Then ADDPoint X, Y, False

        Next
    Next

    ResetPairs
    frmMain.lVIDEOINFO = "Computing new pairs..."
    DoEvents

End Sub


