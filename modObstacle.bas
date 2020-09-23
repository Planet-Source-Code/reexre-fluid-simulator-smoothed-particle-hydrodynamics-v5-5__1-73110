Attribute VB_Name = "modObstacle"
'
' Author: Roberto Mior       reexre@gmail.com
'
'
' If you use or modify this code remember to cite the author
'


Option Explicit

Private Type tStick
    NearX              As Single
    NearY              As Single
    Dist               As Single
    J                  As Long

End Type


Private Type tLine

    X1                 As Single
    Y1                 As Single
    X2                 As Single
    Y2                 As Single

    NormalX            As Single
    NormalY            As Single

End Type

Private Type tBall
    X                  As Single
    Y                  As Single
    OldX               As Single
    OldY               As Single
    Mass               As Single
    Radius             As Single
End Type






Private Type tCollResp
    O                  As Long
    cX                 As Single
    cY                 As Single
End Type

Public LineO()         As tLine
Public Nlines          As Long

Public BallO()         As tBall
Public Nballs          As Long


Public Function Distance(ByRef X1 As Single, ByRef Y1 As Single, _
                         ByRef X2 As Single, ByRef Y2 As Single) As Single

    Dim dX             As Single
    Dim dY             As Single

    dX = X1 - X2
    dY = Y1 - Y2
    Distance = Sqr(dX * dX + dY * dY)
End Function

Public Sub ADDLineOBS(ByVal X1, ByVal Y1, ByVal X2, ByVal Y2)
    Dim Wx             As Single
    Dim Wy             As Single
    Dim L              As Single


    If X1 < 5 Then X1 = 0
    If X2 < 5 Then X2 = 0
    If Y1 < 5 Then Y1 = 0
    If Y2 < 5 Then Y2 = 0
    If X1 > MaxX - 5 Then X1 = MaxX + 1
    If X2 > MaxX - 5 Then X2 = MaxX + 1
    If Y1 > MaxY - 5 Then Y1 = MaxY + 1
    If Y2 > MaxY - 5 Then Y2 = MaxY + 1


    Nlines = Nlines + 1
    ReDim Preserve LineO(Nlines)
    With LineO(Nlines)

        .X1 = X1
        .Y1 = Y1
        .X2 = X2
        .Y2 = Y2

        Wx = .X2 - .X1
        Wy = .Y2 - .Y1
        L = Sqr(Wx * Wx + Wy * Wy)

        .NormalX = Wx / L
        .NormalY = Wy / L


    End With

End Sub

Public Sub ADDBallOBS(cX, cY, Radius)
    Nballs = Nballs + 1
    ReDim Preserve BallO(Nballs)
    With BallO(Nballs)
        .X = cX
        .Y = cY
        .OldX = cX
        .OldX = cY
        .Radius = Radius
        .Mass = .Radius * .Radius * PI

    End With

End Sub

Public Sub DRAWLineObs()
    Dim I              As Long
    '    Stop

    For I = 1 To Nlines
        With LineO(I)

            FastLine frmMain.PIC.Hdc, CLng(.X1), CLng(.Y1), CLng(.X2), CLng(.Y2), 2, vbWhite

        End With
    Next

End Sub
Public Sub DRAWBallObs()
    Dim I              As Long
    '    Stop
    'Stop

    For I = 1 To Nballs
        With BallO(I)

            MyCircle frmMain.PIC.Hdc, CLng(.X), CLng(.Y), CLng(.Radius), 2, vbWhite

        End With
    Next

End Sub

Public Sub LinesIntersect(Ax1 As Single, Ay1 As Single, Ax2 As Single, Ay2 As Single, _
                          Bx1 As Single, By1 As Single, Bx2 As Single, By2 As Single, ByRef RetX As Single, ByRef RetY As Single)
'Stop
    Dim R              As Single
    Dim s              As Single
    Dim InvD              As Single
    Dim D As Single
    
    Dim AX2mAX1        As Single
    Dim AY2mAY1        As Single

    AX2mAX1 = (Ax2 - Ax1)
    AY2mAY1 = (Ay2 - Ay1)
    RetX = -999
    RetY = -999
    D = AX2mAX1 * (By2 - By1) - AY2mAY1 * (Bx2 - Bx1)

    If D <> 0 Then
        InvD = 1 / D
        
        R = ((Ay1 - By1) * (Bx2 - Bx1) - (Ax1 - Bx1) * (By2 - By1)) * InvD '/ D

        If R >= 0 And R <= 1 Then

            s = ((Ay1 - By1) * AX2mAX1 - (Ax1 - Bx1) * AY2mAY1) * InvD '/ D

            If s <= 1 Then
                If s >= 0 Then
                    RetX = Ax1 + R * AX2mAX1
                    RetY = Ay1 + R * AY2mAY1
                End If
            End If

        End If

    End If



End Sub



Public Sub LINE_TESTCollisionAndStick()
'Stop

    Dim I              As Long
    Dim J              As Long

    Dim Rx             As Single
    Dim Ry             As Single

    Dim Coll()         As tCollResp
    Dim Ncoll          As Long
    Dim JJ             As Long
    Dim D              As Single
    Dim Dmin           As Single
    Dim DfromLine      As Single
    Dim WNP            As Long

    Dim s              As tStick

    Dim HH             As Single

    Dim xI             As Single
    Dim yI             As Single

    HH = W.Dstick


    WNP = W.NP

    'GoTo SKipStick


    '*******************           STICK
    For I = 1 To WNP

        xI = W.GetX(I)
        yI = W.GetY(I)

        '****** All near lines
        'For J = 1 To Nlines
        '    DfromLine = W.PtDistFromLine2(I, LineO(J).X1, LineO(J).Y1, LineO(J).X2, LineO(J).Y2, Rx, Ry)
        '    If DfromLine < HH Then
        '        If DfromLine > 0 Then
        '            W.StickApply I, DfromLine, Rx, Ry
        '       End If
        '    End If
        'Next
        '***********

        '******** Only 1 nearest line
        s.Dist = -1
        Dmin = 99999999999#
        For J = 1 To Nlines
            '            Stop

            DfromLine = W.PtDistFromLine2(I, LineO(J).X1, LineO(J).Y1, LineO(J).X2, LineO(J).Y2, Rx, Ry)
            If DfromLine < HH Then
                '                If DfromLine > 0 Then
                If DfromLine < Dmin Then
                    Dmin = DfromLine
                    s.NearX = Rx
                    s.NearY = Ry
                    s.Dist = DfromLine
                End If
                '                End If
            End If
            DfromLine = Distance(LineO(J).X1, LineO(J).Y1, xI, yI)
            If DfromLine < HH Then
                If DfromLine < Dmin Then
                    Dmin = DfromLine
                    s.NearX = LineO(J).X1
                    s.NearY = LineO(J).Y1
                    s.Dist = DfromLine
                End If
            End If
            DfromLine = Distance(LineO(J).X2, LineO(J).Y2, xI, yI)
            If DfromLine < HH Then
                If DfromLine < Dmin Then
                    Dmin = DfromLine
                    s.NearX = LineO(J).X2
                    s.NearY = LineO(J).Y2
                    s.Dist = DfromLine
                End If
            End If

        Next
        If s.Dist > 0 Then W.StickApply I, s.Dist, s.NearX, s.NearY
        '************


        '****** FIELD STICK
        '        If W.GetX(I) < W.H Then If W.GetX(I) >= 1 Then W.StickApply I, Abs(W.GetX(I)), 0, W.GetY(I)
        '        If W.GetY(I) < W.H Then If W.GetY(I) >= 1 Then W.StickApply I, Abs(W.GetY(I)), W.GetX(I), 0
        '        If W.GetX(I) > MaxXStick Then If W.GetX(I) <= MaxX - 1 Then W.StickApply I, Abs(MaxX - W.GetX(I)), MaxX, W.GetY(I)
        '        If W.GetY(I) > MaxYStick Then If W.GetY(I) <= MaxY - 1 Then W.StickApply I, Abs(MaxY - W.GetY(I)), W.GetX(I), MaxY
        If W.GetX(I) < W.Dstick Then If W.GetX(I) > 0 Then W.StickApply I, Abs(W.GetX(I)), 0, W.GetY(I)
        If W.GetY(I) < W.Dstick Then If W.GetY(I) > 0 Then W.StickApply I, Abs(W.GetY(I)), W.GetX(I), 0
        If W.GetX(I) > MaxXStick Then If W.GetX(I) < MaxX Then W.StickApply I, Abs(MaxX - W.GetX(I)), MaxX, W.GetY(I)
        If W.GetY(I) > MaxYStick Then If W.GetY(I) < MaxY Then W.StickApply I, Abs(MaxY - W.GetY(I)), W.GetX(I), MaxY

    Next
    '******************************

SKipStick:

    For I = 1 To WNP
AG:
        Ncoll = 0
        JJ = 0
        For J = 1 To Nlines

            LinesIntersect LineO(J).X1, LineO(J).Y1, LineO(J).X2, LineO(J).Y2, _
                           W.GetX(I), W.GetY(I), W.GetOldX(I), W.GetOldY(I), Rx, Ry

            If Rx <> -999 Then
                Ncoll = Ncoll + 1
                ReDim Preserve Coll(Ncoll)
                Coll(Ncoll).O = J
                Coll(Ncoll).cX = Rx
                Coll(Ncoll).cY = Ry
            End If
        Next

        Dmin = 999999
        For J = 1 To Ncoll
            D = Distance(W.GetOldX(I), W.GetOldY(I), Coll(J).cX, Coll(J).cY)

            If D < Dmin Then Dmin = D: JJ = Coll(J).O: Rx = Coll(J).cX: Ry = Coll(J).cY
        Next
        If JJ <> 0 Then CollisionResponse I, LineO(JJ).NormalX, LineO(JJ).NormalY, Rx, Ry
        'If JJ <> 0 Then CollisionResponse I, JJ, Rx, Ry


        '****** FIELD Collision
        If W.GetX(I) < 0 Then CollisionResponse I, 0, 1, 0, W.GetY(I)
        If W.GetY(I) < 0 Then CollisionResponse I, 1, 0, W.GetX(I), 0
        If W.GetX(I) > MaxX Then CollisionResponse I, 0, 1, MaxX, W.GetY(I)
        If W.GetY(I) > MaxY Then CollisionResponse I, 1, 0, W.GetX(I), MaxY
        If Ncoll <> 0 Then GoTo AG

    Next I



End Sub

'Private Sub CollisionResponse(WaterPoint, ObsLine, PcollX As single, PcollY As single)
Public Sub CollisionResponse(WaterPoint As Long, ByVal OBSnormalX As Single, ByVal OBSnormalY As Single, PcollX As Single, PcollY As Single, _
                             Optional ByRef EX As Single, Optional ByRef Ey As Single)

    Dim Vx             As Single
    Dim Vy             As Single

    '    Dim OBSnormalX             As single
    '    Dim OBSnormalY             As single

    Dim L              As Single
    Dim C1             As Single
    Dim nVx            As Single
    Dim nVy            As Single
    Dim tX             As Single
    Dim tY             As Single


    '    Vx = W.GetVX(WaterPoint)
    '    Vy = W.GetVY(WaterPoint)
    Vx = (W.GetX(WaterPoint) - W.GetOldX(WaterPoint)) * W.InvDT '/ W.dT
    Vy = (W.GetY(WaterPoint) - W.GetOldY(WaterPoint)) * W.InvDT '/ W.dT



    L = Sqr(Vx * Vx + Vy * Vy)
    nVx = Vx / L
    nVy = Vy / L


    '    OBSnormalX = LineO(ObsLine).NormalX
    '    OBSnormalY = LineO(ObsLine).NormalY



    C1 = OBSnormalX * Vx + OBSnormalY * Vy    'DOT

    OBSnormalX = OBSnormalX * C1 * 2
    OBSnormalY = OBSnormalY * C1 * 2

    EX = OBSnormalY
    Ey = -OBSnormalX

    Vx = -Vx + OBSnormalX
    Vy = -Vy + OBSnormalY

    '***** Old Way
    'W.SetX(WaterPoint) = W.GetOldX(WaterPoint)
    'W.SetY(WaterPoint) = W.GetOldY(WaterPoint)
    '
    'W.SetVX(WaterPoint) = Vx * BounceNess
    'W.SetVY(WaterPoint) = Vy * BounceNess
    '*****
    'Stop

    '-------------------------------
    tX = (PcollX + W.GetOldX(WaterPoint)) * 0.5
    tY = (PcollY + W.GetOldY(WaterPoint)) * 0.5


    W.SetOldX(WaterPoint) = tX
    W.SetOldY(WaterPoint) = tY

    W.SetX(WaterPoint) = tX + Vx * BounceNess
    W.SetY(WaterPoint) = tY + Vy * BounceNess

    'W.SetVX(WaterPoint) = Vx    '* BounceNess
    'W.SetVY(WaterPoint) = Vy    '* BounceNess

    '---------------------------
    W.SetXgrid(WaterPoint) = W.GetX(WaterPoint) \ W.H
    W.SetYgrid(WaterPoint) = W.GetY(WaterPoint) \ W.H

End Sub

Public Sub BALL_TESTCollisionAndStick()
    Dim I              As Long
    Dim J              As Long


    Dim Dmin           As Single
    Dim DfromBALL      As Single

    Dim WNP            As Long

    Dim s              As tStick

    Dim HH             As Single

    Dim dX             As Single
    Dim dY             As Single

    Dim D              As Single
    Dim InvD As Single
    
    Dim EX             As Single

    Dim Ey             As Single
    Dim fX             As Single

    Dim fy             As Single

    HH = W.Dstick

    WNP = W.NP

    For I = 1 To WNP

        '********
        s.Dist = -1
        Dmin = 99999999999#
        For J = 1 To Nballs
            '            Stop

            dX = W.GetX(I) - BallO(J).X
            dY = W.GetY(I) - BallO(J).Y

            If Abs(dX) < BallO(J).Radius + HH Then
                If Abs(dY) < BallO(J).Radius + HH Then

                    D = Sqr(dX * dX + dY * dY)
                    InvD = 1 / D
                    DfromBALL = D - BallO(J).Radius
                    
                    If DfromBALL < 0 Then

                        '*****************************************
                        'BALL COLLISION
                        '*********************************
                        '            Stop
                        
                        dX = dX * InvD '/ D
                        dY = dY * InvD '/ D
                        CollisionResponse I, -dY, dX, BallO(J).X + dX * BallO(J).Radius, BallO(J).Y + dY * BallO(J).Radius, EX, Ey

                        'BallO(J).X = BallO(J).X + EX
                        'BallO(J).Y = BallO(J).Y + Ey

                    ElseIf DfromBALL < HH Then

                        '*******************           STICK
                        If DfromBALL < Dmin Then
                            Dmin = DfromBALL
                            dX = dX * InvD '/ D
                            dY = dY * InvD '/ D
                            '                    Stop

                            s.NearX = BallO(J).X + dX * BallO(J).Radius
                            s.NearY = BallO(J).Y + dY * BallO(J).Radius
                            s.Dist = DfromBALL
                            s.J = J
                        End If

                    End If

                End If
            End If

        Next
        If s.Dist > 0 Then
            W.StickApply I, s.Dist, s.NearX, s.NearY    ', fX, fy
            ' BallO(s.J).X = BallO(s.J).X + 50 * fX / BallO(s.J).Mass
            ' BallO(s.J).Y = BallO(s.J).Y + 50 * fY / BallO(s.J).Mass
        End If

        '************
    Next

End Sub

