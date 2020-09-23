VERSION 5.00
Begin VB.Form frmEmbeddedControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin Project1.CoolForm CoolForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   873
      tbCaption       =   "Embedded is Preffered"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This form utilizes the embedded control for cool form."
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1620
      Width           =   3705
   End
End
Attribute VB_Name = "frmEmbeddedControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CoolForm1.Align = vbAlignTop
End Sub
