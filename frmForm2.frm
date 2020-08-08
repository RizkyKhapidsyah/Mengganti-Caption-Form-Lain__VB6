VERSION 5.00
Begin VB.Form frmForm2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetCaption 
      Caption         =   "Ganti Caption"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtNewCaption 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ref As Object

Public Sub SetReference(objRef As Object)
    Set ref = objRef
End Sub

Private Sub cmdSetCaption_Click()
    Dim cap As String
    cap = txtNewCaption.Text
    ref.Caption = cap
End Sub

