Attribute VB_Name = "modBound"
Private Type tB

    X1                 As Single
    Y1                 As Single
    X2                 As Single
    Y2                 As Single

    Enab               As Boolean


    II                 As Long
    JJ                 As Long

End Type


Public Type tPB
    Num                As Single


End Type


Public B()             As tB
Public NB              As Long

Public PB()            As tPB




Public Sub DRAWBound(picHdc As Long)
    Dim I              As Long
    For I = 1 To NB
        With B(I)

            If .Enab Then

                FastLine picHdc, CLng(.X1), CLng(.Y1), CLng(.X2), CLng(.Y2), 2, vbWhite

            End If
        End With

    Next

End Sub



Public Sub FindBound2()



    Dim I              As Long
    Dim J              As Long
    Dim Rx             As Single
    Dim Ry             As Single

    ReDim PB(W.NP)

    NB = 0
    For I = 1 To W.NP
        PB(I).Num = 0
    Next

    For I = 1 To W.NP - 1
        For J = I + 1 To W.NP - 1

            If Abs(W.GetXgrid(I) - W.GetXgrid(J)) <= 2 Then
                If Abs(W.GetYgrid(I) - W.GetYgrid(J)) <= 2 Then
                    'If W.GetSpringRestL(I, J) > 0 Then
                    If Distance(W.GetX(I), W.GetY(I), W.GetX(J), W.GetY(J)) < W.H Then
                        NB = NB + 1
                        ReDim Preserve B(NB)
                        With B(NB)
                            .X1 = W.GetX(I)
                            .Y1 = W.GetY(I)
                            .X2 = W.GetX(J)
                            .Y2 = W.GetY(J)
                            .Enab = True
                            .II = I
                            .JJ = J
                            PB(I).Num = PB(I).Num + 1 * W.GetDensity(J)
                            PB(J).Num = PB(J).Num + 1 * W.GetDensity(I)
                        End With
                    End If
                End If
            End If
        Next
    Next

    For I = 1 To W.NP
        If PB(I).Num = 0 Then
            PB(I).Num = 1
        Else
            PB(I).Num = 1 / PB(I).Num
        End If
    Next

    For I = 1 To NB

        If PB(B(I).II).Num < 0.2 Then B(I).Enab = False
        If PB(B(I).JJ).Num < 0.2 Then B(I).Enab = False

    Next

End Sub

