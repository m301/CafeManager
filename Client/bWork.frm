VERSION 5.00
Begin VB.Form bWork 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   210
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   210
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "bWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cFet.Create_Restore "Startup : MAD Cafe Manager", 12
Unload Me
End Sub
