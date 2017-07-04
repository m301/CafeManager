VERSION 5.00
Begin VB.Form cMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Title :::"
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "No"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "cMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Function cMessage(message As String, title As String)

End Function

