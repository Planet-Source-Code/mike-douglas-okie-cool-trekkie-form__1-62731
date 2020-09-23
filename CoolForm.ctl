VERSION 5.00
Begin VB.UserControl CoolForm 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   855
   ScaleWidth      =   5055
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   30
      Top             =   360
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   4725
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4725
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
         TabIndex        =   1
         Top             =   90
         Width           =   285
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
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
         TabIndex        =   2
         Top             =   90
         Width           =   810
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
      ScaleWidth      =   4845
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4845
   End
End
Attribute VB_Name = "CoolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'captionless dragging
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112


' Region API functins
Private Type POINTAPI
    X As Long
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

Dim LastHeight As Long
Dim mCaption As String

Public Property Get tbCaption() As String
    tbCaption = lblCaption.Caption
End Property

Public Property Let tbCaption(Text As String)
    lblCaption.Caption = Text
    mCaption = Text
    PropertyChanged "tbCaption"
End Property

Private Sub Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub cmdClose_Click()
    Unload UserControl.Parent
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Timer1_Timer()
    If UserControl.Parent.Height <> LastHeight Then
        LastHeight = UserControl.Parent.Height
        Refresh
    End If
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
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

Private Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_Resize()
    TitleBar.Width = UserControl.Width - 100
    TitleBarShadow.Width = TitleBar.Width + 60
    cmdClose.Left = TitleBar.Width - cmdClose.Width - 150
    Refresh
End Sub

Public Sub Refresh()
    On Error Resume Next
    lblCaption.Caption = mCaption
    RoundShape TitleBar
    RoundShape TitleBarShadow
    RoundShape UserControl.Parent
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    mURL = "http://www.google.com/"
    mCaption = UserControl.Name
    PropertyChanged "URL"
    PropertyChanged "Caption"
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    tbCaption = PropBag.ReadProperty("tbCaption")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "tbCaption", mCaption, UserControl.Name
End Sub

