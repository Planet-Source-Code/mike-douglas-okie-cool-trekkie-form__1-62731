VERSION 5.00
Begin VB.Form frmStandAlone 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4395
   ControlBox      =   0   'False
   Icon            =   "frmStandAlone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   30
      ScaleHeight     =   435
      ScaleWidth      =   4605
      TabIndex        =   0
      Top             =   30
      Width           =   4605
      Begin VB.Label cmdClose 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " X "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4350
         TabIndex        =   3
         Top             =   90
         Width           =   285
      End
      Begin VB.Label Capton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cool Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   1
         Top             =   90
         Width           =   1080
      End
   End
   Begin VB.PictureBox TitleBarShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   0
      Width           =   4755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This form has all of the cool form code included in it."
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   870
      Width           =   3660
   End
End
Attribute VB_Name = "frmStandAlone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'captionless dragging
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112


' Region API functins
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub Capton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    TitleBar.Width = Me.Width - 200
    TitleBarShadow.Width = TitleBar.Width + 90
    cmdClose.Left = TitleBar.Width - cmdClose.Width - 150
    RoundShape TitleBar
    RoundShape TitleBarShadow
    RoundShape Me
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As _
    Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Paint()
    RoundShape Me
End Sub

Sub RoundShape(ByVal obj As Object)
    Dim lpRect As RECT
    Dim wi As Long, he As Long
    Dim hRgn As Long
    Dim hWnd As Long
    
    hWnd = obj.hWnd
    GetWindowRect hWnd, lpRect
    wi = lpRect.Right - lpRect.Left
    he = lpRect.Bottom - lpRect.Top
    hRgn = CreateRoundRectRgn(0, 0, wi, he, 20, 20)
    SetWindowRgn hWnd, hRgn, True
    DeleteObject hRgn
End Sub

