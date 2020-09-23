VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Terrain 3D"
   ClientHeight    =   5955
   ClientLeft      =   1155
   ClientTop       =   1800
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6825
   Begin VB.CheckBox chkUseBuffer 
      Caption         =   "Use Buffer"
      Height          =   195
      Left            =   3840
      TabIndex        =   11
      Top             =   540
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   2880
      List            =   "frmMain.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox picScr 
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   6735
      TabIndex        =   3
      Top             =   840
      Width           =   6795
   End
   Begin VB.CheckBox chkWireFrame 
      Caption         =   "View Wire Frame"
      Height          =   195
      Left            =   3840
      TabIndex        =   2
      Top             =   300
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkRoll 
      Caption         =   "Enable Roll"
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "New Terrain"
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   6480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   6120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   540
      Width           =   375
   End
   Begin VB.Label lblRoll 
      Caption         =   "lblRoll"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label lblPitch 
      Caption         =   "lblPitch"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   300
      Width           =   2115
   End
   Begin VB.Label lblYaw 
      Caption         =   "lblYaw"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WATER_LEVEL = -5

Private Type ScreenCoord
    X As Single
    Y As Single
End Type
Private Type Coord3D
    X As Single
    Y As Single
    Z As Single
End Type
Private Type HPoint
    H As Single
    bRecurring As Boolean
End Type

Dim pan As Coord3D
Dim altitude() As HPoint
Dim yaw As Single
Dim pitch As Single
Dim roll As Single
Dim FocalDistance As Single
Dim MAX_X As Long
Dim MAX_Y As Long
Dim Initialize As Boolean
Dim UseBuffer As Boolean

Private Declare Function Polygon Lib "GDI32" (ByVal hdc As Long, lpPoint As PointApi, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Private Type PointApi   ' pt
    X As Long
    Y As Long
End Type

Private Sub CalculateMidPoints(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Dim xm As Long
    Dim ym As Long
    Dim d As Single
    Dim Y As Long
    
    xm = (x1 + x2 + 1) \ 2
    ym = (y1 + y2 + 1) \ 2
    If ym > y1 And ym < y2 And xm > x1 And xm < x2 And Not altitude(xm, ym).bRecurring Then
        d = Sqr((x2 - x1 + 1) ^ 2 + (y2 - y1 + 1) ^ 2) / 2
        altitude(xm, ym).H = (altitude(x1, y1).H + altitude(x1, y2).H + altitude(x2, y1).H + altitude(x2, y2).H) / 4 + d / 2 - Rnd * (d)
        altitude(xm, ym).bRecurring = True
        d = (x2 - x1 + 1) / 2
        If x1 <> 0 And y1 <> 0 And x2 <> MAX_Y - 1 And y2 <> MAX_Y - 1 Then
            altitude(x1, ym).H = (altitude(x1, y1).H + altitude(x1, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(x1, ym).bRecurring = True
            altitude(x2, ym).H = (altitude(x2, y1).H + altitude(x2, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(x2, ym).bRecurring = True
            altitude(xm, y1).H = (altitude(x1, y1).H + altitude(x2, y1).H) / 2 + d / 2 - Rnd * (d)
            altitude(xm, y1).bRecurring = True
            altitude(xm, y2).H = (altitude(x1, y2).H + altitude(x2, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(xm, y2).bRecurring = True
        End If
        CalculateMidPoints x1, y1, xm, ym
        CalculateMidPoints x1, ym, xm, y2
        CalculateMidPoints xm, y1, x2, ym
        CalculateMidPoints xm, ym, x2, y2
    End If
End Sub

Private Sub GenerateTerrain()
    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim d As Single

    'Reset array
    For X = 0 To MAX_X - 1
        For Y = 0 To MAX_Y - 1
            altitude(X, Y).H = 0
            altitude(X, Y).bRecurring = False
        Next
    Next
    CalculateMidPoints 0, 0, MAX_X - 1, MAX_Y - 1
    'smooth 3D surface
    For i = 1 To 2
        For X = 1 To MAX_X - 2
            For Y = 1 To MAX_Y - 2
                altitude(X, Y).H = (altitude(X - 1, Y - 1).H + altitude(X + 1, Y - 1).H + altitude(X - 1, Y + 1).H + altitude(X + 1, Y + 1).H + altitude(X, Y).H * 2) / 6 '+ (d / 2 - Rnd * (d / 2))
            Next
        Next
    Next
    'water level
    For X = 1 To MAX_X - 2
        For Y = 1 To MAX_Y - 2
            If altitude(X, Y).H < WATER_LEVEL Then altitude(X, Y).H = WATER_LEVEL
        Next
    Next
End Sub

Private Sub chkRoll_Click()
    If chkRoll.Value = vbUnchecked Then
        roll = 0
        If UseBuffer Then
            Call DrawTerrain(picBuffer)
            PutBufferToScreen
        Else
            picScr.Cls
            Call DrawTerrain(picScr)
        End If
    End If
End Sub

Private Sub chkUseBuffer_Click()
    If chkUseBuffer.Value = vbChecked Then
        UseBuffer = True
    Else
        picScr.Cls
        UseBuffer = False
        Call DrawTerrain(picScr)
    End If
End Sub

Private Sub chkWireFrame_Click()
    If UseBuffer Then
        Call DrawTerrain(picBuffer)
        PutBufferToScreen
    Else
        picScr.Cls
        Call DrawTerrain(picScr)
    End If
End Sub

Private Sub cmbSize_Click()
    MAX_X = cmbSize.Text
    MAX_Y = cmbSize.Text
    ReDim altitude(MAX_X - 1, MAX_Y - 1)
    
    yaw = 0.88
    pitch = -0.6
    'set default pan
    pan.X = 0
    pan.Y = 0
    pan.Z = MAX_Y * 1.5
    'set focal distance
    FocalDistance = 8000
    
    GenerateTerrain
    
    If UseBuffer Then
        ClearBuffer
        Call DrawTerrain(picBuffer)
        PutBufferToScreen
    Else
        picScr.Cls
        Call DrawTerrain(picScr)
    End If
End Sub

Private Sub cmdGenerate_Click()
    GenerateTerrain
    
    If UseBuffer Then
        ClearBuffer
        Call DrawTerrain(picBuffer)
        PutBufferToScreen
    Else
        picScr.Cls
        Call DrawTerrain(picScr)
    End If
End Sub

Private Sub Form_Load()
    Initialize = True
    UseBuffer = True
    picBuffer.Height = picScr.Height
    picBuffer.Width = picScr.Width
    picBlank.Height = picScr.Height
    picBlank.Width = picScr.Width
    Randomize Timer
    cmbSize.Text = 64
End Sub

Private Sub Form_Resize()
    If UseBuffer Then
        ClearBuffer
        Call DrawTerrain(picBuffer)
        PutBufferToScreen
    Else
        picScr.Cls
        Call DrawTerrain(picScr)
    End If
    picScr.Height = Height - 825
    picScr.Width = Width - 150
    picBuffer.Height = picScr.Height
    picBuffer.Width = picScr.Width
    picBlank.Height = picScr.Height
    picBlank.Width = picScr.Width
    If Initialize Then
        picScr.Cls
        Call DrawTerrain(picScr)
        Initialize = False
    End If
End Sub

Private Sub DrawTerrain(DrawTo As PictureBox)
    Dim X As Long
    Dim Y As Long
    Dim p2() As ScreenCoord
    Dim mp As Coord3D
    Dim mc As Long
    Dim minh As Single
    Dim maxh As Single
    Dim point(0 To 5) As PointApi
    ReDim p2(MAX_X - 1, MAX_Y - 1)
    
    'calculate screen coordinates of land points
    For X = 0 To MAX_X - 1
        For Y = 0 To MAX_Y - 1
            mp.X = X
            mp.Y = altitude(X, Y).H
            mp.Z = Y
            p2(X, Y) = PointToScreen(mp, pan, FocalDistance, yaw, pitch, roll)
            If altitude(X, Y).H < minh Then minh = altitude(X, Y).H
            If altitude(X, Y).H > maxh Then maxh = altitude(X, Y).H
        Next
    Next
    'Me.AutoRedraw = False
    'draw every polygon
    For X = 0 To MAX_X - 2
        For Y = MAX_Y - 2 To 0 Step -1
            If altitude(X, Y).H = WATER_LEVEL And altitude(X + 1, Y).H = WATER_LEVEL And altitude(X + 1, Y + 1).H = WATER_LEVEL And altitude(X, Y + 1).H = WATER_LEVEL Then
                'set water color to blue
                mc = RGB(80, 80, 210 + Rnd * 30)
                DrawTo.ForeColor = mc
            Else
                'set land color
                mc = RGB(40, 50 + 205 * (altitude(X, Y).H - minh) / (maxh - minh), 10)
                If chkWireFrame.Value = vbChecked Then
                    DrawTo.ForeColor = vbBlack
                Else
                    DrawTo.ForeColor = mc
                End If
            End If
            DrawTo.FillColor = mc
            DrawTo.FillStyle = vbSolid
            point(0).X = DrawTo.ScaleX(p2(X, Y).X, vbTwips, vbPixels)
            point(0).Y = DrawTo.ScaleY(p2(X, Y).Y, vbTwips, vbPixels)
            point(1).X = DrawTo.ScaleX(p2(X, Y + 1).X, vbTwips, vbPixels)
            point(1).Y = DrawTo.ScaleY(p2(X, Y + 1).Y, vbTwips, vbPixels)
            point(2).X = DrawTo.ScaleX(p2(X + 1, Y + 1).X, vbTwips, vbPixels)
            point(2).Y = DrawTo.ScaleY(p2(X + 1, Y + 1).Y, vbTwips, vbPixels)
            point(3).X = DrawTo.ScaleX(p2(X + 1, Y).X, vbTwips, vbPixels)
            point(3).Y = DrawTo.ScaleY(p2(X + 1, Y).Y, vbTwips, vbPixels)
            point(4).X = point(0).X
            point(4).Y = point(0).Y
            Polygon DrawTo.hdc, point(0), 4
        Next
    Next
    'Me.AutoRedraw = True
    DrawTo.CurrentX = 0
    DrawTo.CurrentY = 0
    DrawTo.ForeColor = vbBlack
    lblYaw = "yaw=" & yaw
    lblPitch = "pitch=" & pitch
    lblRoll = "roll=" & roll
End Sub

Private Sub picScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static oldy As Single
    Static oldx As Single
    
    If oldy <> 0 And oldx <> 0 And Button = vbLeftButton Then
        If UseBuffer Then
            ClearBuffer
        End If
        pitch = pitch - (Y - oldy) / 1000
        yaw = yaw - (X - oldx) / 1000
        If UseBuffer Then
            Call DrawTerrain(picBuffer)
            PutBufferToScreen
        Else
            picScr.Cls
            Call DrawTerrain(picScr)
        End If
    End If
    If oldy <> 0 And oldx <> 0 And Button = vbRightButton Then
        If UseBuffer Then
            ClearBuffer
        End If
        If chkRoll.Value = vbChecked Then roll = roll + (X - oldx) / 1000
        FocalDistance = FocalDistance + (Y - oldy)
        If UseBuffer Then
            Call DrawTerrain(picBuffer)
            PutBufferToScreen
        Else
            picScr.Cls
            Call DrawTerrain(picScr)
        End If
    End If
    oldy = Y
    oldx = X
End Sub

'convert 3d coordinates to screen coordinates
Private Function PointToScreen(p As Coord3D, pan As Coord3D, ByVal FocalDistance As Single, ByVal yaw As Single, ByVal pitch As Single, ByVal roll As Single) As ScreenCoord
    Dim np1 As Coord3D
    Dim np2 As Coord3D
    
    'apply pan to center 3d grid to view position
    np2.X = p.X - MAX_X / 2
    np2.Z = p.Z - MAX_Y / 2
    np2.Y = p.Y
   
    np1.X = np2.Z * Sin(yaw) + np2.X * Cos(yaw)
    np1.Y = np2.Y
    np1.Z = np2.Z * Cos(yaw) - np2.X * Sin(yaw)
    
    np2.X = np1.X
    np2.Y = np1.Y * Cos(pitch) - np1.Z * Sin(pitch)
    np2.Z = np1.Y * Sin(pitch) + np1.Z * Cos(pitch)
    
    np1.X = np2.Y * Sin(roll) + np2.X * Cos(roll)
    np1.Y = np2.Y * Cos(roll) - np2.X * Sin(roll)
    np1.Z = np2.Z
    
    np1.X = np1.X + pan.X
    np1.Y = np1.Y + pan.Y
    np1.Z = np1.Z + pan.Z
    
    If np1.Z <> 0 Then
        PointToScreen.X = np1.X * (FocalDistance) / np1.Z + Me.Width / 2
        PointToScreen.Y = -np1.Y * (FocalDistance) / np1.Z + Me.Height / 2
    End If
End Function

Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub

Sub PutBufferToScreen()
    BitBlt picScr.hdc, 0, 0, picScr.Width, picScr.Height, picBuffer.hdc, 0, 0, vbSrcCopy
End Sub
