VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fluid Simulator (Smoothed particle hydrodynamics)"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   613
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCLRCircles 
      Caption         =   "Clear CIRCLES"
      Height          =   555
      Left            =   12720
      TabIndex        =   38
      Top             =   120
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   4770
      Left            =   6480
      TabIndex        =   36
      Top             =   2760
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "LOAD Scene"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10320
      TabIndex        =   35
      ToolTipText     =   "Load Scene  (Water and lines)"
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE Scene"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10320
      TabIndex        =   34
      ToolTipText     =   "Save Scene as  (water and Lines)"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   22
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "|>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   21
      ToolTipText     =   "Play / Resume simulation"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "| |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      ToolTipText     =   "Pause Simulation"
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame fGRAVITY 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GRAVITY DIRECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   11880
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
      Begin VB.HScrollBar StepSPEED 
         Height          =   255
         Left            =   1920
         Max             =   500
         Min             =   6
         TabIndex        =   48
         Top             =   2520
         Value           =   25
         Width           =   1215
      End
      Begin VB.PictureBox picG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   2160
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   47
         ToolTipText     =   "Move mouse and then Click to Set Gravity."
         Top             =   360
         Width           =   975
         Begin VB.Line LineG2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   0
            X2              =   56
            Y1              =   48
            Y2              =   0
         End
         Begin VB.Line LineG 
            BorderWidth     =   2
            X1              =   0
            X2              =   56
            Y1              =   56
            Y2              =   8
         End
         Begin VB.Shape GShape 
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.HScrollBar hRndDIR 
         Height          =   255
         Left            =   120
         Max             =   750
         TabIndex        =   30
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   19
         ToolTipText     =   "No Gravity"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Left"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   17
         ToolTipText     =   "Down"
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   495
         Index           =   3
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Right"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   495
         Index           =   4
         Left            =   720
         TabIndex        =   15
         ToolTipText     =   "Up"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Left Down"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   13
         ToolTipText     =   "Right Down"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   375
         Index           =   7
         Left            =   1320
         TabIndex        =   12
         ToolTipText     =   "Right Up"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdGRAV 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Left Up"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lCHframe 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N  Frames to Complete Gravity Change"
         Height          =   615
         Left            =   1920
         TabIndex        =   49
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lRND 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Random Direction Every N Frames"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCLRwater 
      Caption         =   "Clear WATER"
      Height          =   555
      Left            =   13680
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCLRLines 
      Caption         =   "Clear LINES"
      Height          =   555
      Left            =   11880
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer TimerFPS 
      Interval        =   5000
      Left            =   13080
      Top             =   1320
   End
   Begin VB.HScrollBar hFaucet 
      Height          =   255
      Left            =   11880
      Max             =   100
      TabIndex        =   5
      Top             =   960
      Value           =   1
      Width           =   1935
   End
   Begin VB.Frame fPictureAction 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MOUSE ACTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   10200
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
      Begin VB.OptionButton oFILLCircle 
         Caption         =   "FILLED Circle"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Draw Rectangle filled with water"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton oFILLREC 
         Caption         =   "FILLED Rectangle"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":0BD4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Draw Rectangle filled with water"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.OptionButton oBALL 
         Caption         =   "Circle"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":0EDE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Add Obstacle BALLs"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton oPP 
         Caption         =   "Push/Pull Water"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":11E8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Mouse Left = Repulsor, Mouse Right = Attractor"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.OptionButton oWater 
         Caption         =   "Add/Delete Water"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":14F2
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Mouse Left = Add Water, Mouse Right = Delete Water"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton oLINE 
         Caption         =   "Line"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":17FC
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add Obstacle Lines"
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Water"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lobs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obstacles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "START"
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      ToolTipText     =   "START Simulation - New Water - Erase All Old Video Frames"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame frmDRAW 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10200
      TabIndex        =   42
      Top             =   5880
      Width           =   1575
      Begin VB.HScrollBar sMotionBlurred 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   51
         Top             =   1830
         Width           =   1215
      End
      Begin VB.OptionButton oSTYLEblobby 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Blobby"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton oStyleSTANDARD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Standard"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox oDRAWSprings 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Draw springs"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton oStyleBlobby2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Blobby 2"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lBLUR 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Motion Blurred = 0 Frames"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame fVIDEO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VIDEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   11880
      TabIndex        =   23
      Top             =   5280
      Width           =   1935
      Begin VB.CommandButton cmdSelectPlayer 
         Caption         =   "Select Player ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Select Your AVI Player (.exe)"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CheckBox chPLAY 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto PLAY Avi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "AutoPlay AVI when It's Created"
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtFPS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   26
         Text            =   "25"
         ToolTipText     =   "Output FPS"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdBUILD_AVI 
         Caption         =   "STOP and Build AVI"
         Height          =   435
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chSaveFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save Video Frames"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lVIDEOINFO 
         BackColor       =   &H0000C000&
         Caption         =   "_"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lOUTFPS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Video FPS (1-30) BEST (20-25)"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label lFPS 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   27
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lPTS 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lFaucet 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private X1             As Long
Private Y1             As Long
Private X2             As Long
Private Y2             As Long


Private CMD            As New cFileDlg

Private AVIPLAYER      As String


Private Sub chPLAY_Click()
    If chPLAY And (AVIPLAYER = "") Then MsgBox "Select Avi Player": cmdSelectPlayer_Click

End Sub

Private Sub cmdCLRCircles_Click()
    Nballs = 0
    DRAWALL
    DoEvents
End Sub

Private Sub cmdCLRLines_Click()
    Nlines = 0
    'ADDLineOBS 2, 2, MaxX, 2
    'ADDLineOBS 2, 2, 2, MaxY
    'ADDLineOBS MaxX, MaxY, 2, MaxY
    'ADDLineOBS MaxX, MaxY, MaxX, 2
    DRAWALL
    DoEvents

End Sub

Private Sub cmdCLRwater_Click()
    W.NP = 0
    W.Npairs = 0

    DRAWALL
    DoEvents
End Sub

Private Sub cmdGO_Click()

    Dim I              As Long

    lVIDEOINFO = "Delating old video Frames...": DoEvents

    If Dir(App.Path & "\VideoFrames\*.BMP") <> "" Then Kill App.Path & "\VideoFrames\*.*"

    Frame = 0

    W.AAAINITFluid 0
    W.GX = GravDirX * Gravity
    W.GY = GravDirY * Gravity

    '    W.AAAINITFluid 2500
    '    I = 0
    '    Do
    '    x = Rnd * MaxX
    '    y = Rnd * MaxY
    '    If Sqr((x - MaxX / 2) ^ 2 + (y - MaxY / 2) ^ 2) < 80 Then
    '    I = I + 1
    '    W.SetX(I) = x
    '    W.SetY(I) = y
    '    End If
    '    Loop While I < 2500


    MaxXStick = MaxX - W.Dstick
    MaxYStick = MaxY - W.Dstick


    'For x = 50 To 100 Step W.SpringL * 2
    'For y = 20 To 100 Step W.SpringL * 2
    'W.ADDPoint x, y
    'Next
    'Me.Caption = x & "  NP" & W.NP
    'DoEvents
    'Next
    'W.SaveWater


    Wx = PIC.Width \ 2
    Wy = 20

    oPP.Enabled = True
    oWater.Enabled = True
    oFILLREC.Enabled = True
    oFILLCircle.Enabled = True

    cmdSave.Enabled = True
    cmdLoad.Enabled = True


    '    For I = 1 To 10
    '        ADDBallOBS MaxX * Rnd, MaxY * Rnd, 10 + Rnd * 80
    '    Next

    '    W.LoadWater "DEFAULT.TXT"

    Running = True

    DoLOOP

End Sub



Public Sub cmdGRAV_Click(Index As Integer)
    GravANGFrom = GravANG
    GravMagFrom = GravMag

    Select Case Index
        Case 0
            'GravDirX = 0
            'GravDirY = 0
            GravMagTo = 0
        Case 1
            GravDirX = -1
            GravDirY = 0
            GravMagTo = 1
        Case 2
            GravDirX = 0
            GravDirY = 1
            GravMagTo = 1
        Case 3
            GravDirX = 1
            GravDirY = 0
            GravMagTo = 1
        Case 4
            GravDirX = 0
            GravDirY = -1
            GravMagTo = 1
        Case 5
            GravDirX = -1 / Sqr(2)
            GravDirY = 1 / Sqr(2)
            GravMagTo = 1
        Case 6
            GravDirX = 1 / Sqr(2)
            GravDirY = 1 / Sqr(2)
            GravMagTo = 1
        Case 7
            GravDirX = 1 / Sqr(2)
            GravDirY = -1 / Sqr(2)
            GravMagTo = 1
        Case 8
            GravDirX = -1 / Sqr(2)
            GravDirY = -1 / Sqr(2)
            GravMagTo = 1

    End Select

    'LineG.X2 = LineG.X1 + GravDirX * picG.ScaleWidth \ 2
    'LineG.Y2 = LineG.Y1 + GravDirY * picG.ScaleWidth \ 2
    'W.GX = GravDirX * Gravity
    'W.GY = GravDirY * Gravity
    GravANGTo = Atan2(GravDirX, GravDirY)
    StepGravA = AngleDiff(GravANGFrom, GravANGTo) / (StepSPEED * EveryFrame)
    StepGravM = (GravMagTo - GravMagFrom) / (StepSPEED * EveryFrame)



End Sub

Private Sub cmdLoad_Click()

    File1.Path = App.Path & "\Scenes\"
    File1.Refresh
    File1.Enabled = Not (File1.Enabled)

    File1.Visible = File1.Enabled
End Sub

Private Sub cmdPause_Click()
    Running = False
    DoEvents
End Sub

Private Sub cmdPlay_Click()
    Running = True
    DoLOOP

End Sub

Private Sub cmdBUILD_AVI_Click()

    DoEvents
    cmdPause_Click


    BUILD_AVI App.Path & "\VideoFrames\", Val(txtFPS), Me.hWnd, lVIDEOINFO
    DoEvents

    If chPLAY.Value = Checked Then
        If (AVIPLAYER <> "") Then
            If OutputAVIName <> "" Then
                Shell AVIPLAYER & " " & Chr$(34) & OutputAVIName, vbNormalFocus
            End If
        Else
            MsgBox "Can't Autoplay: No Avi Player Selected!", vbCritical
        End If

    End If

End Sub

Private Sub cmdSave_Click()

    Dim F              As String

    F = InputBox("Save this scene as: ", "Save Scene", "Water")

    If F <> vbNullString Then W.SaveWater F

End Sub

Private Sub cmdSelectPlayer_Click()
    Dim Filename       As String

    With CMD
        '.filename = ""
        '.InitDir = "c:\"
        '.Filter = "AVI Player|*.EXE" ';*.mpg"
        '.DialogTitle = "Select AVI PLAYER"


        If AVIPLAYER <> "" Then
            .InitDirectory = AVIPLAYER
        Else
            .InitDirectory = "C:\"
        End If
        .DefaultExt = "Exe"
        .DlgTitle = "Select AVI PLAYER"
        .Filter = "AVI Player (.EXE) |*.EXE"
        .OwnerHwnd = frmMain.hWnd



    End With
    'CMD.Action = 1
    CMD.VBGetOpenFileName Filename


    If Asc(Filename) <> 0 Then
        AVIPLAYER = Filename
        Open App.Path & "\Player.txt" For Output As 22
        Print #22, AVIPLAYER
        Close 22
    End If
End Sub

Private Sub File1_DblClick()
    Dim Orun           As Boolean
    Orun = Running

    If Running Then Running = False
    W.LoadWater File1.Filename
    File1.Visible = False
    File1.Enabled = False

    DRAWALL



    Running = Orun


    DoEvents
End Sub

Private Sub Form_Load()


    If Dir(App.Path & "\VideoFrames\", vbDirectory) = "" Then MkDir App.Path & "\VideoFrames\"
    If Dir(App.Path & "\Video\", vbDirectory) = "" Then MkDir App.Path & "\Video\"
    If Dir(App.Path & "\Scenes\", vbDirectory) = "" Then MkDir App.Path & "\Scenes\"


    If Dir(App.Path & "\Player.txt") <> "" Then
        Open App.Path & "\Player.txt" For Input As 22
        Input #22, AVIPLAYER
        Close 22
    End If


    PIC.Height = 360         '360
    PIC.Width = PIC.Height * 16 \ 9    '14 \ 9    '16 / 9

    Randomize Timer

    Me.Caption = Me.Caption & "  V" & App.Major & "." & App.Minor


    MaxX = PIC.Width - 2
    MaxY = PIC.Height - 2



    'ADDLineOBS 3, 3, MaxX, 3
    'ADDLineOBS 3, 3, 3, MaxY
    'ADDLineOBS MaxX, MaxY, 3, MaxY
    'ADDLineOBS MaxX, MaxY, MaxX, 3




    X1 = -99
    Y1 = -99
    oLINE_Click
    hFaucet_Change
    cmdGRAV_Click 2
    StepSPEED_Change



    GShape.Width = picG.ScaleWidth
    GShape.Height = GShape.Width
    LineG.X1 = GShape.Width \ 2
    LineG.Y1 = GShape.Height \ 2
    LineG.X2 = LineG.X1
    LineG.Y2 = GShape.Height
    LineG2.X1 = GShape.Width \ 2
    LineG2.Y1 = GShape.Height \ 2
    LineG2.X2 = LineG.X1
    LineG2.Y2 = GShape.Height

    GravDirX = 0
    GravDirY = 1
    GravANGFrom = Atan2(0, 1)
    GravANG = GravANGFrom

    DoEvents

    BLUR.InitTarget PIC.Image.Handle



End Sub

Private Sub Form_Unload(Cancel As Integer)

    BLUR.Termin
    End

End Sub

Private Sub hFaucet_Change()
    lFaucet = "Faucet " & hFaucet & "%"
End Sub

Private Sub hFaucet_Scroll()
    lFaucet = "Faucet " & hFaucet & "%"
End Sub

Private Sub hRndDIR_Change()
    lRND = "Random Direction Every " & hRndDIR & " Frames"
End Sub

Private Sub hRndDIR_Scroll()
    lRND = "Random Direction Every " & hRndDIR & " Frames"
End Sub

Private Sub oBALL_Click()
    PIC.MouseIcon = oBALL.MouseIcon
End Sub

Private Sub oFILLCircle_Click()
    PIC.MouseIcon = oFILLCircle.MouseIcon
End Sub

Private Sub oFILLREC_Click()
    PIC.MouseIcon = oFILLREC.MouseIcon
End Sub

Private Sub oLINE_Click()
    PIC.MouseIcon = oLINE.MouseIcon

End Sub

Private Sub oPP_Click()
    PIC.MouseIcon = oPP.MouseIcon

End Sub

Private Sub oSTYLEblobby_Click()
    oDRAWSprings.Visible = oStyleSTANDARD
End Sub

Private Sub oStyleBlobby2_Click()
    oDRAWSprings.Visible = oStyleSTANDARD
End Sub

Private Sub oStyleSTANDARD_Click()
    oDRAWSprings.Visible = oStyleSTANDARD
End Sub

Private Sub oWater_Click()
    PIC.MouseIcon = oWater.MouseIcon

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R              As Single
    Dim dX             As Single
    Dim dY             As Single

    If oLINE Then
        If Button = 1 Then
            If X1 <> -99 And Y1 <> -99 Then ADDLineOBS X1, Y1, X, Y
        End If

        X1 = X
        Y1 = Y
        If Button = 2 Then
            X1 = -99: Y1 = -99

            DRAWALL

        End If
    End If

    If oBALL Then
        If Button = 1 And X1 <> -99 Then
            dX = X - X1
            dY = Y - Y1
            R = Sqr(dX * dX + dY * dY)
            If X1 <> -99 And Y1 <> -99 Then
                ADDBallOBS X1, Y1, R
                DRAWALL
                DoEvents
            End If

            X1 = -99: Y1 = -99
        Else
            X1 = X
            Y1 = Y
        End If
        If Button = 2 Then
            X1 = -99: Y1 = -99

            DRAWALL
            DoEvents

        End If
    End If


    If oFILLREC Then
        If Button = 1 And X1 <> -99 Then

            W.ADDFilledRect X, Y, X1, Y1

            DRAWALL
            DoEvents
            X1 = -99: Y1 = -99
        Else
            X1 = X
            Y1 = Y
        End If
        If Button = 2 Then
            X1 = -99: Y1 = -99
            DRAWALL
            DoEvents
        End If
    End If

    If oFILLCircle Then
        If Button = 1 And X1 <> -99 Then

            dX = X - X1
            dY = Y - Y1
            R = Sqr(dX * dX + dY * dY)

            W.ADDFilledCircle X1, Y1, R

            DRAWALL
            DoEvents
            X1 = -99: Y1 = -99
        Else
            X1 = X
            Y1 = Y
        End If
        If Button = 2 Then
            X1 = -99: Y1 = -99
            DRAWALL
            DoEvents
        End If
    End If

    If Button <> 0 Then
        If oWater Then AddRemoveWater Button, X, Y
        If oPP Then PushPullWater Button, X, Y
    End If




End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R              As Single
    Dim dX             As Single
    Dim dY             As Single
    If oLINE Then
        If X1 <> -99 Then
            X2 = X
            Y2 = Y

            DRAWALL
            FastLine PIC.Hdc, X1, Y1, CLng(X), CLng(Y), 1, vbWhite
            DoEvents
        End If
        PIC.Refresh
    End If

    If oBALL Then
        If X1 <> -99 Then



            dX = X - X1
            dY = Y - Y1
            R = Sqr(dX * dX + dY * dY)
            DRAWALL
            MyCircle PIC.Hdc, X1, Y1, CLng(R), 1, vbWhite
            DoEvents

        End If
        PIC.Refresh
    End If

    If oFILLREC Then
        If X1 <> -99 Then
            X2 = X
            Y2 = Y

            DRAWALL
            FastLine PIC.Hdc, X1, Y1, X2, Y1, 1, vbWhite
            FastLine PIC.Hdc, X1, Y2, X2, Y2, 1, vbWhite
            FastLine PIC.Hdc, X1, Y1, X1, Y2, 1, vbWhite
            FastLine PIC.Hdc, X2, Y1, X2, Y2, 1, vbWhite
            DoEvents
        End If
        PIC.Refresh
    End If


    If oFILLCircle Then
        If X1 <> -99 Then
            X2 = X
            Y2 = Y
            dX = X2 - X1
            dY = Y2 - Y1
            R = Sqr(dX * dX + dY * dY)
            DRAWALL
            MyCircle PIC.Hdc, X1, Y1, CLng(R), 1, vbWhite
            DoEvents
        End If
        PIC.Refresh
    End If

    If Button <> 0 Then
        If oWater Then AddRemoveWater Button, X, Y
        If oPP Then PushPullWater Button, X, Y
    End If

End Sub


Private Sub AddRemoveWater(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    Dim I              As Long
    Dim dX             As Single
    Dim dY             As Single
    Dim D              As Single

    If Button = 1 Then
        Wx = X
        Wy = Y
        W.ADDPoint Wx + Rnd, Wy + Rnd
    End If

    If Button = 2 Then
        I = 1
        Do
            dX = X - W.GetX(I)
            dY = Y - W.GetY(I)
            D = Sqr(dX * dX + dY * dY)
            If D < 20 Then
                W.RemovePoint (I): I = I - 1
            End If
            I = I + 1
        Loop While I <= W.NP

    End If

    DRAWALL
    DoEvents




End Sub
Private Sub PushPullWater(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    Dim I              As Long
    Dim dX             As Single
    Dim dY             As Single
    Dim D              As Single
    Dim D2             As Single
    Dim s              As Single

    s = IIf(Button = 2, -1, 1)



    For I = 1 To W.NP
        dX = X - W.GetX(I)
        dY = Y - W.GetY(I)
        D = Sqr(dX * dX + dY * dY)
        If D < 40 Then
            D2 = 1 - D / 40
            'W.SetVX(I) = W.GetVX(I) - dX * 0.05 * D2 * s
            'W.SetVY(I) = W.GetVY(I) - dY * 0.05 * D2 * s
            W.SetX(I) = W.GetX(I) - dX * 0.05 * D2 * s
            W.SetY(I) = W.GetY(I) - dY * 0.05 * D2 * s

        End If
    Next


End Sub



Private Sub picG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX
    Dim dY
    Dim L

    dX = X - LineG2.X1
    dY = Y - LineG2.Y1


    L = Sqr(dX * dX + dY * dY)

    If L < (picG.ScaleWidth \ 2) Then L = picG.ScaleWidth \ 2

    dX = dX / L
    dY = dY / L
    Me.Caption = L

    LineG2.X2 = LineG2.X1 + dX * picG.ScaleWidth \ 2
    LineG2.Y2 = LineG2.Y1 + dY * picG.ScaleWidth \ 2

    GravMagTo = Sqr(dX * dX + dY * dY)


    StepGravM = (GravMagTo - GravMag) / (StepSPEED * EveryFrame)

    GravANGTo = Atan2(dX, dY)
    GravANGFrom = GravANG
    StepGravA = AngleDiff(GravANGFrom, GravANGTo) / (StepSPEED * EveryFrame)



End Sub

Private Sub picG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX
    Dim dY
    Dim L

    dX = X - LineG2.X1
    dY = Y - LineG2.Y1
    L = Sqr(dX * dX + dY * dY)

    If L < (picG.ScaleWidth \ 2) Then L = picG.ScaleWidth \ 2

    dX = dX / L
    dY = dY / L

    LineG2.X2 = LineG2.X1 + dX * picG.ScaleWidth \ 2
    LineG2.Y2 = LineG2.Y1 + dY * picG.ScaleWidth \ 2
    '    W.GX = GravDirX * Gravity
    '    W.GY = GravDirY * Gravity



    ' GravANGTo = Atan2(GravDirX, GravDirY)
    ' GravANGFrom = GravANG
    ' StepGravA = AngleDiff(GravANGFrom, GravANGTo) / (StepSPEED * EveryFrame)



End Sub

Private Sub sMotionBlurred_Change()
    lBLUR = "Motion Blurred = " & sMotionBlurred & " frames"
End Sub

Private Sub sMotionBlurred_Scroll()
   lBLUR = "Motion Blurred = " & sMotionBlurred & " frames"
End Sub

Private Sub StepSPEED_Change()
    lCHframe = "" & StepSPEED & " Frames to Complete Gravity Changes"
End Sub

Private Sub StepSPEED_Scroll()
    lCHframe = "" & StepSPEED & " Frames to Complete Gravity Changes"
End Sub

Private Sub TimerFPS_Timer()
    Dim fps            As Long

    '    fps = (Frame - pFrame) * 0.33333333333    'Interval=3000
    fps = (Frame - pFrame) * 0.2    'Interval=5000


    lFPS = "Computed FPS:" & fps & " [Drawn:" & Int(10 * fps / EveryFrame) * 0.1 & "]"

    pFrame = Frame
End Sub

Public Sub DRAWALL()

    BitBlt PIC.Hdc, 0, 0, PIC.ScaleWidth, PIC.ScaleHeight, PIC.Hdc, 0, 0, vbBlack    'ness

'W.DRAW PIC.Hdc, oDRAWSprings

    If oStyleSTANDARD Then
        W.DRAW PIC.Hdc, oDRAWSprings
    ElseIf oSTYLEblobby Then
        W.DRAWMetaBall PIC.Hdc
    ElseIf oStyleBlobby2 Then W.DRAWMetaBallCont PIC.Hdc

    End If

    DRAWLineObs
    DRAWBallObs
    PIC.Refresh


End Sub
