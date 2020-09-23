Attribute VB_Name = "BrushLine"
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------

Public Type POINTAPI
    X                  As Long
    Y                  As Long
End Type

Public poi             As POINTAPI


Public Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function Arc Lib "gdi32" (ByVal Hdc As Long, _
                                         ByVal xInizioRettangolo As Long, _
                                         ByVal yInizioRettangolo As Long, _
                                         ByVal xFineRettangolo As Long, _
                                         ByVal yFineRettangolo As Long, _
                                         ByVal xInizioArco As Long, _
                                         ByVal yInizioArco As Long, _
                                         ByVal xFineArco As Long, _
                                         ByVal yFineArco As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal Hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal Hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


'Declare Function Arc Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, _
 ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
 ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long


Public Sub SetBrush(ByVal Hdc As Long, ByVal PenWidth As Long, ByVal PenColor As Long)


    DeleteObject (SelectObject(Hdc, CreatePen(vbSolid, PenWidth, PenColor)))
    'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
    'SetBrush = kOBJ


End Sub



Public Sub FastLine(ByRef Hdc As Long, ByRef X1 As Long, ByRef Y1 As Long, _
                    ByRef X2 As Long, ByRef Y2 As Long, ByRef W As Long, ByRef color As Long)
Attribute FastLine.VB_Description = "disegna line veloce"

    Dim poi            As POINTAPI

    'SetBrush hdc, W, color
    DeleteObject (SelectObject(Hdc, CreatePen(vbSolid, W, color)))

    MoveToEx Hdc, X1, Y1, poi
    LineTo Hdc, X2, Y2

End Sub

Sub MyCircle(ByRef Hdc As Long, ByRef X As Long, ByRef Y As Long, ByRef R As Long, ByRef W As Long, ByRef color As Long)
    Dim XpR            As Long

    'SetBrush hdc, W, color
    DeleteObject (SelectObject(Hdc, CreatePen(vbSolid, W, color)))

    XpR = X + R

    Arc Hdc, X - R, Y - R, XpR, Y + R, XpR, Y, XpR, Y

End Sub


Public Sub bLOCK(ByRef Hdc As Long, X As Long, Y As Long, ByRef W As Long, ByRef color As Long)

    DeleteObject (SelectObject(Hdc, CreatePen(vbSolid, 2, color)))

    Rectangle Hdc, X, Y, X + W, Y + W

End Sub
