VERSION 5.00
Begin VB.Form cOpt 
   Caption         =   "::: Optimizer :::"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   360
      TabIndex        =   30
      Top             =   2160
      Width           =   7935
      Begin VB.CommandButton Command22 
         Caption         =   "Create A System Restore Point"
         Height          =   375
         Left            =   480
         TabIndex        =   50
         Top             =   3360
         Width           =   4335
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Disable Shutdown"
         Height          =   375
         Left            =   480
         TabIndex        =   48
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Disable Registry"
         Height          =   375
         Left            =   2760
         TabIndex        =   47
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Disable Log Off"
         Height          =   375
         Left            =   5040
         TabIndex        =   46
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Shutdown"
         Height          =   375
         Left            =   480
         TabIndex        =   45
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Restart"
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Log Off"
         Height          =   375
         Left            =   480
         TabIndex        =   43
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cancel Shutdown"
         Height          =   375
         Left            =   480
         TabIndex        =   42
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Enable Taskmanager"
         Height          =   375
         Left            =   2760
         TabIndex        =   41
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Disable Taskmanager"
         Height          =   375
         Left            =   2760
         TabIndex        =   40
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Enable Run"
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Disable run"
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Enable control Panel"
         Height          =   375
         Left            =   5040
         TabIndex        =   37
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Disable Control Panel"
         Height          =   375
         Left            =   5040
         TabIndex        =   36
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Enable Shell / CMD"
         Height          =   375
         Left            =   5040
         TabIndex        =   35
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Disable Shell / CMD"
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Enable Shutdown"
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Enable Registry"
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Enable Log Off"
         Height          =   375
         Left            =   5040
         TabIndex        =   31
         Top             =   2400
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame Frame8 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      Begin VB.CheckBox Check26 
         Caption         =   "Z"
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Y"
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check24 
         Caption         =   "X"
         Height          =   255
         Left            =   5640
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check13 
         Caption         =   "M"
         Height          =   240
         Left            =   6600
         TabIndex        =   25
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check12 
         Caption         =   "L"
         Height          =   240
         Left            =   6120
         TabIndex        =   24
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check11 
         Caption         =   "K"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check23 
         Caption         =   "W"
         Height          =   240
         Left            =   5160
         TabIndex        =   22
         Top             =   720
         Width           =   525
      End
      Begin VB.CheckBox Check22 
         Caption         =   "V"
         Height          =   240
         Left            =   4680
         TabIndex        =   21
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check21 
         Caption         =   "U"
         Height          =   240
         Left            =   4200
         TabIndex        =   20
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check20 
         Caption         =   "T"
         Height          =   240
         Left            =   3720
         TabIndex        =   19
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check19 
         Caption         =   "S"
         Height          =   240
         Left            =   3240
         TabIndex        =   18
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check18 
         Caption         =   "R"
         Height          =   240
         Left            =   2760
         TabIndex        =   17
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Q"
         Height          =   240
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check16 
         Caption         =   "P"
         Height          =   240
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check15 
         Caption         =   "O"
         Height          =   240
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check14 
         Caption         =   "N"
         Height          =   240
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   400
      End
      Begin VB.CheckBox Check10 
         Caption         =   "J"
         Height          =   255
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check9 
         Caption         =   "I"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check8 
         Caption         =   "H"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check7 
         Caption         =   "G"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check6 
         Caption         =   "F"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check5 
         Caption         =   "E"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check4 
         Caption         =   "D"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check3 
         Caption         =   "C"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   400
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   400
      End
      Begin VB.CommandButton command1 
         Caption         =   "Hide Selected Drives"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Un-Hide All Drives"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Note : Please restart the system to apply changes successfully !"
      Height          =   255
      Left            =   1440
      TabIndex        =   49
      Top             =   6720
      Width           =   5535
   End
End
Attribute VB_Name = "cOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim a1 As Long

If Check1.Value = 1 Then
    Check1.Tag = 1
Else
    Check1.Tag = 0
End If

If Check2.Value = 1 Then
    Check2.Tag = 2
Else
    Check2.Tag = 0
End If
If Check3.Value = 1 Then
    Check3.Tag = 4
Else
    Check3.Tag = 0
End If

If Check4.Value = 1 Then
    Check4.Tag = 8
Else
    Check4.Tag = 0
End If
If Check5.Value = 1 Then
    Check5.Tag = 16
Else
    Check5.Tag = 0
End If

If Check6.Value = 1 Then
    Check6.Tag = 32
Else
    Check6.Tag = 0
End If
If Check7.Value = 1 Then
    Check7.Tag = 64
Else
    Check7.Tag = 0
End If

If Check8.Value = 1 Then
    Check8.Tag = 128
Else
    Check8.Tag = 0
End If
If Check9.Value = 1 Then
    Check9.Tag = 256
Else
    Check9.Tag = 0
End If

If Check10.Value = 1 Then
    Check10.Tag = 512
Else
    Check10.Tag = 0
End If
If Check11.Value = 1 Then
    Check11.Tag = 1024
Else
    Check11.Tag = 0
End If

If Check12.Value = 1 Then
    Check12.Tag = 2048
Else
    Check12.Tag = 0
End If
If Check13.Value = 1 Then
    Check13.Tag = 4096
Else
    Check13.Tag = 0
End If

If Check14.Value = 1 Then
    Check14.Tag = 8192
Else
    Check14.Tag = 0
End If
If Check15.Value = 1 Then
    Check15.Tag = 16384
Else
    Check15.Tag = 0
End If

If Check16.Value = 1 Then
    Check16.Tag = 32768
Else
    Check16.Tag = 0
End If
If Check17.Value = 1 Then
    Check17.Tag = 65536
Else
    Check17.Tag = 0
End If

If Check18.Value = 1 Then
    Check18.Tag = 131072
Else
    Check18.Tag = 0
End If
If Check19.Value = 1 Then
    Check19.Tag = 262144
Else
    Check19.Tag = 0
End If '

If Check20.Value = 1 Then
    Check20.Tag = 524288
Else
    Check20.Tag = 0
End If
If Check21.Value = 1 Then
    Check21.Tag = 1048576
Else
    Check21.Tag = 0
End If

If Check22.Value = 1 Then
    Check22.Tag = 2097152
Else
    Check22.Tag = 0
End If
If Check23.Value = 1 Then
    Check23.Tag = 4194304
Else
    Check23.Tag = 0
End If

If Check24.Value = 1 Then
    Check24.Tag = 8388608
Else
    Check24.Tag = 0
End If
If Check25.Value = 1 Then
    Check25.Tag = 16777216
Else
    Check25.Tag = 0
End If

If Check26.Value = 1 Then
    Check26.Tag = 33554432
Else
    Check26.Tag = 0
End If

a1 = CLng(Check1.Tag) + CLng(Check2.Tag) + CLng(Check3.Tag) _
+ CLng(Check4.Tag) + CLng(Check5.Tag) + CLng(Check6.Tag) + _
CLng(Check7.Tag) + CLng(Check8.Tag) + CLng(Check9.Tag) + _
CLng(Check10.Tag) + CLng(Check11.Tag) + CLng(Check12.Tag) _
+ CLng(Check13.Tag) + CLng(Check14.Tag) + CLng(Check15.Tag) + _
CLng(Check16.Tag) + CLng(Check17.Tag) + CLng(Check18.Tag) _
+ CLng(Check19.Tag) + CLng(Check20.Tag) + CLng(Check21.Tag) _
+ CLng(Check22.Tag) + CLng(Check23.Tag) + CLng(Check24.Tag) _
+ CLng(Check25.Tag) + CLng(Check26.Tag)
cWinsock.wSend "dhde" & a1, cWinsock.SEL
End Sub

Private Sub Command10_Click()
cWinsock.wSend "oren", cWinsock.SEL
End Sub

Private Sub Command11_Click()
cWinsock.wSend "ordi", cWinsock.SEL
End Sub

Private Sub Command12_Click()
cWinsock.wSend "ocen", cWinsock.SEL
End Sub

Private Sub Command13_Click()
cWinsock.wSend "ocdi", cWinsock.SEL
End Sub

Private Sub Command14_Click()
cWinsock.wSend "osen", cWinsock.SEL
End Sub

Private Sub Command15_Click()
cWinsock.wSend "osdi", cWinsock.SEL
End Sub

Private Sub Command16_Click()

cWinsock.wSend "oshe" & a1, cWinsock.SEL
End Sub

Private Sub Command17_Click()
cWinsock.wSend "oshd" & a1, cWinsock.SEL
End Sub

Private Sub Command18_Click()
cWinsock.wSend "edis", cWinsock.SEL
End Sub

Private Sub Command19_Click()
cWinsock.wSend "rena", cWinsock.SEL
End Sub

Private Sub Command2_Click()
cWinsock.wSend "dude" & a1, cWinsock.SEL
End Sub

Private Sub Command20_Click()
cWinsock.wSend "olen", cWinsock.SEL
End Sub

Private Sub Command21_Click()
cWinsock.wSend "oldi", cWinsock.SEL

End Sub

Private Sub Command22_Click()
cWinsock.wSend "ocre", cWinsock.SEL

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
cWinsock.wSend "oshu", cWinsock.SEL
End Sub

Private Sub Command5_Click()
cWinsock.wSend "ores", cWinsock.SEL
End Sub

Private Sub Command6_Click()
cWinsock.wSend "olog", cWinsock.SEL
End Sub

Private Sub Command7_Click()
cWinsock.wSend "ocsh", cWinsock.SEL
End Sub

Private Sub Command8_Click()
cWinsock.wSend "oten", cWinsock.SEL
End Sub

Private Sub Command9_Click()
cWinsock.wSend "otdi", cWinsock.SEL
End Sub

