VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Tipe Kalender"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If DateTime.Calendar = vbCalGreg Then
     MsgBox "Kalender Masehi!", vbInformation, _
            "Masehi"
  Else
     MsgBox "Kalender Hijriah!", vbInformation, _
            "Hijriah"
  End If
End Sub


