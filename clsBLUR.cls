VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBLUR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type Bitmap
    bmType             As Long
    bmWidth            As Long
    bmHeight           As Long
    bmWidthBytes       As Long
    bmPlanes           As Integer
    bmBitsPixel        As Integer
    bmBits             As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private TargetBytes()  As Byte
Private SourceBytes()  As Byte

Private hBmp           As Bitmap

Private iRET           As Long

Dim pW                 As Long
Dim pH                 As Long


Public Sub InitTarget(pBoxPicturehandle As Long)

    iRET = GetObject(pBoxPicturehandle, Len(hBmp), hBmp)

    ReDim TargetBytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte

    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1

End Sub
Public Sub GetSource(pBoxPicturehandle As Long)
    Dim iRET           As Long
    'Get the bitmap header
    iRET = GetObject(pBoxPicturehandle, Len(hBmp), hBmp)
    'Resize to hold image data


    ReDim SourceBytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into SourceBytes array
    iRET = GetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, SourceBytes(0, 0, 0))

    '    ReDim TargetBytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into TargetBytes array
    '  iRet = GetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, TargetBytes(0, 0, 0))


    'ReDim SINGLEb(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Single

End Sub

Public Sub TagertTO(pBoxPicturehandle As Long)
    Dim iRET           As Long
    'Set the new image data back onto pBox
    iRET = SetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, TargetBytes(0, 0, 0))
    'Erase TargetBytes because we finished with it now
    'Erase TargetBytes
End Sub




Public Sub BLUR(nFrames As Integer)
    Dim X              As Long
    Dim Y              As Long
    Dim B              As Long
    Dim Per            As Integer
    Dim iBA            As Integer
    Dim C              As Long


    Per = 256 / (nFrames + 1)


    For X = 0 To pW
        For Y = 0 To pH

            C = RGB(SourceBytes(0, X, Y), SourceBytes(1, X, Y), SourceBytes(2, X, Y))

            If C = 0 Then

                For B = 0 To 2
                    '

                    iBA = CInt(TargetBytes(B, X, Y)) - Per
                    If iBA < 0 Then iBA = 0

                    TargetBytes(B, X, Y) = CByte(iBA)
                Next
            Else

                For B = 0 To 2
                    TargetBytes(B, X, Y) = SourceBytes(B, X, Y)
                Next
            End If

        Next
    Next
'Stop

End Sub



Public Sub Termin()
    Erase TargetBytes
    Erase SourceBytes


End Sub
