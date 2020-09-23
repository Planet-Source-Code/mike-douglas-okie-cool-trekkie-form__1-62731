VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show Form with Embedded Control"
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   1050
      Width           =   4065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Stand-Alone Form"
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   4065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    frmStandAlone.Show
End Sub

Private Sub Command2_Click()
    frmEmbeddedControl.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
