VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rippulator"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdc 
      Left            =   840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open 300x300 Image"
      Filter          =   "Still Image (*.bmp; *.jpg)|*.bmp;*.jpg"
   End
   Begin VB.PictureBox stats 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4680
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   13
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Frame frmopt 
      Caption         =   "Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   4455
      Begin VB.CommandButton ApplyTimeInt 
         Caption         =   "Apply"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtti 
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         Text            =   "40"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtfn 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "10"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtrs 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "6"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtrn 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "10"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Create 
         Caption         =   "Create Ripple"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Interval:"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Frames:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Light Refraction:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ripple Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox dest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   0
      Left            =   4680
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   120
      Width           =   4500
   End
   Begin VB.PictureBox source 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      Begin VB.CommandButton cmdnewimage 
         Caption         =   "Load Image"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Line xpos 
         DrawMode        =   6  'Mask Pen Not
         X1              =   72
         X2              =   72
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line ypos 
         DrawMode        =   6  'Mask Pen Not
         X1              =   0
         X2              =   300
         Y1              =   224
         Y2              =   224
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi As Double = 3.14159265358979
Const Pi2 As Double = 6.28318530717958
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim TickCount As Long
Dim t As Long, fn As Long

Private Sub cmdnewimage_Click()
    cdc.ShowOpen
    If cdc.FileName <> "" Then source.Picture = LoadPicture(cdc.FileName)
End Sub

Private Sub Create_Click()
    stats.Cls
    DoEvents
    txtti.Text = 400 / txtfn.Text
    Timer.Enabled = False
    If dest.Count > 1 Then
        For tt = 1 To dest.Count - 1
            Unload dest(tt)
        Next tt
    End If
    Dim X As Long, Y As Long, a As Double, r As Double, re As Double, w As Long, s As String, rn As Double, rs As Double, RCX As Single, RCY As Single
    Dim x2 As Single, y2 As Single
    RCX = xpos.x1
    RCY = ypos.y1
    rn = txtrn.Text
    rs = txtrs.Text
    fn = txtfn.Text
    d = 300
    TickCount = GetTickCount
    ReDim SourceBuffer.Bits(3, 299, 299)
    With SourceBuffer.Header
      .biSize = 40
      .biWidth = 300
      .biHeight = -300
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 3 * d * d
    End With
     
    ReDim DestBuffer.Bits(3, 299, 299)
    With DestBuffer.Header
      .biSize = 40
      .biWidth = 300
      .biHeight = -300
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 3 * d * d
    End With

    GetDIBits source.hdc, source.Image.Handle, 0, 300, SourceBuffer.Bits(0, 0, 0), SourceBuffer, 0&
    For w = 1 To fn
            For Y = 0 To 299
                For X = 0 To 299
                    a = ATan2(RCY - Y, RCX - X)
                    r = Sqr((RCX - X) * (RCX - X) + (RCY - Y) * (RCY - Y))
                    re = r + (rs * Sin((Pi2 * (r / 150) * rn) - (Pi2 * (w / fn))))
                    x2 = Within(RCX - (re * Cos(a)), 0, 299)
                    y2 = Within(RCY + (re * Sin(a)), 0, 299)
                    DestBuffer.Bits(0, X, Y) = SourceBuffer.Bits(0, x2, y2)
                    DestBuffer.Bits(1, X, Y) = SourceBuffer.Bits(1, x2, y2)
                    DestBuffer.Bits(2, X, Y) = SourceBuffer.Bits(2, x2, y2)
                Next X
            Next Y
        SetDIBits dest(w - 1).hdc, dest(w - 1).Image.Handle, 0, 300, DestBuffer.Bits(0, 0, 0), DestBuffer, 0&
        Load dest(w)
        dest(w).Visible = True
        dest(w).Move dest(0).Left, dest(0).Top
        dest(w - 1).ZOrder
        stats.Print (w / fn) * 100 & "%  ";
        DoEvents
    Next w
    stats.Print
    stats.Print "Operation took " & (GetTickCount - TickCount) / 1000 & " seconds."
    Timer.Interval = txtti.Text
    Timer.Enabled = True
    t = 0
End Sub

Private Sub ApplyTimeInt_Click()
    Timer.Interval = txtti.Text
End Sub

Private Sub source_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    source_MouseMove Button, Shift, X, Y
End Sub

Private Sub source_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        xpos.x1 = X
        xpos.x2 = X
        ypos.y1 = Y
        ypos.y2 = Y
    End If
End Sub

Private Sub Timer_Timer()
    t = t + 1
    If t = fn Then t = 0
    dest(t).ZOrder
End Sub
