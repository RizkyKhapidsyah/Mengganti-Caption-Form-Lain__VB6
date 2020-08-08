VERSION 5.00
Begin VB.Form frmForm1 
   Caption         =   "Mengganti Caption Form Lain"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowForm 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShowForm_Click()
    Load frmForm2
    frmForm2.Show
    frmForm2.SetReference Me
End Sub

Private Sub Form_Load()
    Load frmForm2
End Sub

