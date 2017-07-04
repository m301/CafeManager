VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cWinsock 
   BackColor       =   &H8000000E&
   Caption         =   "MAD Cafe Manager - Server"
   ClientHeight    =   9345
   ClientLeft      =   1500
   ClientTop       =   2010
   ClientWidth     =   15405
   Icon            =   "Winsock.frx":0000
   LinkTopic       =   "cwinsock"
   ScaleHeight     =   9345
   ScaleWidth      =   15405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command44 
      Caption         =   "Command44"
      Height          =   495
      Left            =   12360
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   10560
      Top             =   360
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   2040
      TabIndex        =   120
      Top             =   6120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Window Manager "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clipboard Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bandwith Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Website Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Others"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   3840
      Top             =   0
   End
   Begin VB.ComboBox cmbChooseView 
      Height          =   315
      ItemData        =   "Winsock.frx":E859
      Left            =   8880
      List            =   "Winsock.frx":E863
      TabIndex        =   10
      Text            =   "Icon View"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main"
      Height          =   6855
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      Begin MServer.chameleonButton Others 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   5520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Search PC"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E87E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton Others 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E89A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Price List"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E8B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton cHome 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Home"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E8D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox aLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Login"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CheckBox aRenew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Renew"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Minimize on close"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   6480
         Width           =   1695
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Token"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E8EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Printer Management"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E90A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Users"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E926
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Custom Timers"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E942
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Advanced"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E95E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Employee"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E97A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton mComCol 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Reports"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":E996
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Index           =   0
      Left            =   240
      TabIndex        =   116
      Top             =   4800
      Width           =   375
      Begin VB.TextBox txtLog 
         Height          =   1335
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   119
         Top             =   480
         Width           =   8175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   0
         TabIndex        =   117
         Text            =   "dockit"
         Top             =   0
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   375
         Left            =   5040
         TabIndex        =   118
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock SckUdp 
      Left            =   3000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmd 
      Caption         =   "0000"
      Height          =   975
      Index           =   0
      Left            =   13440
      Picture         =   "Winsock.frx":E9B2
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sock1 
      Index           =   0
      Left            =   3000
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   80
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":F282
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":10C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":122FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":13C27
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":15101
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":16381
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":1779E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":1934E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":1AE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Winsock.frx":1C784
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stat"
      Height          =   6615
      Left            =   12120
      TabIndex        =   121
      Top             =   1440
      Width           =   2415
      Begin MServer.chameleonButton Others 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   136
         Top             =   6120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "LogOut"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":1E082
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MServer.chameleonButton Others 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   129
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Show Details"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Winsock.frx":1E09E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox AppToAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply to all terminals"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label ctLeft 
         Caption         =   "Time Remaining"
         Height          =   255
         Left            =   240
         TabIndex        =   126
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lcDownloaded 
         Caption         =   "Downloaded"
         Height          =   255
         Left            =   240
         TabIndex        =   127
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lStart 
         Caption         =   "Start Time"
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lTamount 
         Caption         =   "Total Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   128
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lTTme 
         Caption         =   "Total Time"
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lcName 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   123
         Top             =   840
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         Height          =   3375
         Left            =   120
         Top             =   720
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   120
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label cInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Unlocked :"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   134
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label cInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Free System:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   133
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label cInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Offline :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   132
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label cInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   130
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label cInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Online :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   131
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblsEl 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Terminal "
         Height          =   375
         Left            =   240
         TabIndex        =   122
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      Begin VB.Label sTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MAD Cafe Manager"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label cfeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MASAC Cyber Cafe (India) U.P."
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   4695
      End
   End
   Begin MSComctlLib.ListView lvwDB 
      Height          =   4695
      Left            =   12960
      TabIndex        =   188
      Top             =   6840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8281
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Index           =   1
      Left            =   1200
      TabIndex        =   154
      Top             =   6600
      Width           =   11775
      Begin VB.CommandButton Command23 
         Caption         =   "Close All Applications"
         Height          =   375
         Left            =   6360
         TabIndex        =   157
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Remote Desktop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   160
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Timer atTitle 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5520
         Top             =   120
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Refresh Title"
         Height          =   255
         Left            =   2760
         TabIndex        =   156
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Get Active Window Title"
         Height          =   375
         Left            =   2760
         TabIndex        =   162
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox caTitle 
         Height          =   405
         Left            =   2760
         TabIndex        =   159
         Text            =   "Active Window Title"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton uDisable 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable USB"
         Height          =   375
         Left            =   360
         TabIndex        =   161
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton uEnable 
         Caption         =   "Enable USB"
         Height          =   375
         Left            =   360
         TabIndex        =   158
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox cUSB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USB Always Enabled "
         Height          =   255
         Left            =   360
         TabIndex        =   155
         Top             =   600
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H8000000A&
         Height          =   1575
         Left            =   240
         Top             =   480
         Width           =   2175
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H8000000A&
         Height          =   1575
         Left            =   6240
         Top             =   480
         Width           =   2055
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H8000000A&
         Height          =   1575
         Left            =   2640
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Index           =   3
      Left            =   2160
      TabIndex        =   45
      Top             =   1440
      Width           =   9735
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   240
         TabIndex        =   46
         Top             =   -1080
         Width           =   8415
         Begin VB.CommandButton Command28 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3360
            TabIndex        =   72
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmddOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   71
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   1920
            TabIndex        =   70
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox sNo 
            Enabled         =   0   'False
            Height          =   350
            Left            =   1200
            TabIndex        =   48
            Text            =   "S No."
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox rePass 
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   4440
            PasswordChar    =   "*"
            TabIndex        =   64
            Text            =   "Re-Password"
            Top             =   2160
            Width           =   1815
         End
         Begin VB.VScrollBar tsAmount 
            Height          =   350
            LargeChange     =   30
            Left            =   2760
            Max             =   0
            Min             =   -1440
            SmallChange     =   5
            TabIndex        =   66
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox cAddress 
            Height          =   350
            Left            =   4440
            TabIndex        =   60
            Text            =   "Address"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.ComboBox uType 
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox idType 
            Height          =   315
            ItemData        =   "Winsock.frx":1E0BA
            Left            =   1200
            List            =   "Winsock.frx":1E0BC
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox pass 
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   61
            Text            =   "Password"
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox Mob 
            Height          =   350
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   58
            Text            =   "Mobile Number"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox idNum 
            Height          =   350
            Left            =   4440
            TabIndex        =   56
            Text            =   "Id Number"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox aUsed 
            Height          =   350
            Left            =   4560
            TabIndex        =   69
            Text            =   "0"
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox tAmount 
            Height          =   350
            Left            =   2040
            TabIndex        =   67
            Text            =   "0"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox cName 
            Height          =   350
            Left            =   1200
            TabIndex        =   50
            Text            =   "Name"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Line Line3 
            X1              =   6720
            X2              =   6720
            Y1              =   360
            Y2              =   3600
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password  :"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   63
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Re Password :"
            Height          =   255
            Index           =   7
            Left            =   3240
            TabIndex        =   62
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile  :"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address  :"
            Height          =   255
            Index           =   5
            Left            =   3360
            TabIndex        =   59
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ID Number :"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   55
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ID Type  :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UserType :"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   51
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "User Name :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblDET 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "S.No. :"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   47
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Used  :"
            Height          =   255
            Left            =   3240
            TabIndex        =   68
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount :"
            Height          =   375
            Left            =   840
            TabIndex        =   65
            Top             =   2760
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Show Passwords"
         Height          =   375
         Left            =   5400
         TabIndex        =   77
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add User"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit User"
         Height          =   375
         Left            =   1920
         TabIndex        =   75
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete User"
         Height          =   375
         Left            =   3600
         TabIndex        =   76
         Top             =   3720
         Width           =   1575
      End
      Begin MSComctlLib.ListView uDb 
         Height          =   3255
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5741
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Index           =   0
      Left            =   2160
      TabIndex        =   26
      Top             =   1320
      Width           =   9135
      Begin VB.CommandButton Command21 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3840
         TabIndex        =   34
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Add"
         Height          =   375
         Left            =   3840
         TabIndex        =   32
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3840
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox teCha 
         Height          =   375
         Left            =   4800
         TabIndex        =   29
         Text            =   "10"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox teMin 
         Height          =   375
         Left            =   3600
         TabIndex        =   28
         Text            =   "60"
         Top             =   1080
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwPrice 
         Height          =   3975
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Line Line2 
         X1              =   3480
         X2              =   5880
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   3480
         X2              =   5880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape6 
         Height          =   2535
         Left            =   3360
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   255
         Left            =   5520
         TabIndex        =   31
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Min."
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Index           =   1
      Left            =   2160
      TabIndex        =   35
      Top             =   1320
      Width           =   12855
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   5880
         TabIndex        =   36
         Top             =   120
         Width           =   3135
         Begin VB.TextBox tPass 
            Height          =   285
            Left            =   480
            TabIndex        =   38
            Text            =   "Token Password"
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox tPassword 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Passwords"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Add Token"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Delete Token"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Shape Shape7 
            Height          =   975
            Left            =   240
            Top             =   360
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView lvwToken 
         Height          =   3975
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Index           =   4
      Left            =   2280
      TabIndex        =   78
      Top             =   1440
      Width           =   9375
      Begin MSComctlLib.ListView lvwTime 
         Height          =   3855
         Left            =   240
         TabIndex        =   79
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   3735
         Left            =   4920
         TabIndex        =   80
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton Command34 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   360
            TabIndex        =   87
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Ok"
            Height          =   375
            Left            =   360
            TabIndex        =   86
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox tmrTot 
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Text            =   "60"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox tmrName 
            Height          =   375
            Left            =   240
            TabIndex        =   82
            Text            =   "User Name"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Time :"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Min."
            Height          =   375
            Left            =   1920
            TabIndex        =   85
            Top             =   1440
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   3735
         Left            =   4800
         TabIndex        =   88
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton Command33 
            Caption         =   "Re-Start Timer"
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Pause / Resume Timer"
            Height          =   375
            Left            =   120
            TabIndex        =   92
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Add Timer"
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Delete Timer"
            Height          =   375
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Edit Timer"
            Height          =   375
            Left            =   120
            TabIndex        =   91
            Top             =   1320
            Width           =   2295
         End
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   5
      Left            =   2040
      TabIndex        =   94
      Top             =   1320
      Width           =   9855
      Begin VB.CommandButton Command38 
         Caption         =   "Exit Client"
         Height          =   375
         Left            =   4320
         TabIndex        =   100
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Save Setting"
         Height          =   375
         Left            =   4320
         TabIndex        =   97
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Load Setting"
         Height          =   375
         Left            =   4320
         TabIndex        =   96
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command36 
         Caption         =   "Shutdown All Free System"
         Height          =   375
         Left            =   480
         TabIndex        =   102
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Set Currency"
         Height          =   375
         Left            =   480
         TabIndex        =   101
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Disconnect Selected Terminal"
         Height          =   435
         Left            =   480
         TabIndex        =   98
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton disALL 
         Caption         =   "Disconnect All Terminals"
         Height          =   435
         Left            =   480
         TabIndex        =   99
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Change Details"
         Height          =   375
         Left            =   480
         TabIndex        =   95
         Top             =   480
         Width           =   3015
      End
      Begin VB.Shape Shape17 
         Height          =   1095
         Index           =   1
         Left            =   4200
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Shape Shape19 
         Height          =   3135
         Left            =   360
         Top             =   360
         Width           =   3255
      End
      Begin VB.Shape Shape17 
         Height          =   1095
         Index           =   0
         Left            =   4200
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Index           =   7
      Left            =   2280
      TabIndex        =   114
      Top             =   1560
      Width           =   9615
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   6
      Left            =   2160
      TabIndex        =   103
      Top             =   1560
      Width           =   9615
      Begin VB.Frame addEmployee 
         BackColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   3960
         TabIndex        =   104
         Top             =   360
         Width           =   2895
         Begin VB.ComboBox eRights 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton Command49 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   480
            TabIndex        =   110
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CommandButton Command48 
            Caption         =   "Add"
            Height          =   375
            Left            =   480
            TabIndex        =   109
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox eRepin 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   60
            PasswordChar    =   "*"
            TabIndex        =   108
            Text            =   "re-pin"
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox ePin 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   60
            PasswordChar    =   "*"
            TabIndex        =   107
            Text            =   "pin"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox eName 
            Height          =   375
            Left            =   240
            MaxLength       =   60
            TabIndex        =   105
            Text            =   "Name"
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Remove Employee"
         Height          =   375
         Left            =   4440
         TabIndex        =   113
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Add Employee"
         Height          =   375
         Left            =   4440
         TabIndex        =   112
         Top             =   840
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvwAdmins 
         Height          =   3375
         Left            =   240
         TabIndex        =   111
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Shape Shape20 
         Height          =   1575
         Left            =   3960
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Index           =   0
      Left            =   1680
      TabIndex        =   137
      Top             =   6840
      Width           =   9975
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   7440
         TabIndex        =   146
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton Command52 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   150
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   149
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   360
            TabIndex        =   148
            Text            =   "10"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount :"
            Height          =   255
            Left            =   360
            TabIndex        =   147
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Remote Desktop"
         Height          =   375
         Left            =   5160
         TabIndex        =   142
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command42 
         Caption         =   "File Transfer"
         Height          =   375
         Left            =   5160
         TabIndex        =   153
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Add Charge"
         Height          =   375
         Left            =   7560
         TabIndex        =   143
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Add Discount"
         Height          =   375
         Left            =   7560
         TabIndex        =   151
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Transfer Session"
         Height          =   375
         Left            =   5160
         TabIndex        =   145
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Identify Terminal"
         Height          =   375
         Left            =   5160
         TabIndex        =   140
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Generate Invoice"
         Height          =   375
         Left            =   7560
         TabIndex        =   141
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Lock Terminal / Close Session"
         Height          =   375
         Left            =   600
         TabIndex        =   152
         Top             =   1800
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Unlock Terminal / Unlimited"
         Height          =   375
         Left            =   600
         TabIndex        =   144
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New User"
         Height          =   375
         Left            =   600
         TabIndex        =   138
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Time"
         Height          =   375
         Left            =   2760
         TabIndex        =   139
         Top             =   360
         Width           =   1935
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H8000000A&
         Height          =   1935
         Index           =   1
         Left            =   7440
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H8000000A&
         Height          =   2055
         Index           =   0
         Left            =   5040
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H8000000A&
         Height          =   1095
         Left            =   480
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H8000000A&
         Height          =   615
         Left            =   480
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Index           =   2
      Left            =   1920
      TabIndex        =   163
      Top             =   6600
      Width           =   9375
      Begin VB.CommandButton clipBrdGet 
         Caption         =   "Get From Clipboard"
         Height          =   375
         Left            =   5520
         TabIndex        =   166
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton clipBrdSet 
         Caption         =   "Paste To Clipboard"
         Height          =   375
         Left            =   5520
         TabIndex        =   165
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox clipBrd 
         Height          =   1335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   164
         Text            =   "Winsock.frx":1E0BE
         Top             =   600
         Width           =   5055
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H8000000A&
         Height          =   1575
         Left            =   240
         Top             =   480
         Width           =   7695
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Index           =   5
      Left            =   1440
      TabIndex        =   186
      Top             =   6480
      Width           =   9495
      Begin VB.CommandButton Command24 
         Caption         =   "Optimize System"
         Height          =   495
         Left            =   480
         TabIndex        =   187
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Index           =   3
      Left            =   1680
      TabIndex        =   167
      Top             =   6720
      Width           =   9495
      Begin VB.CheckBox wDL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warn Me If a User Exceeds Download Limit"
         Height          =   255
         Left            =   480
         TabIndex        =   170
         Top             =   840
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CommandButton nDlSel 
         Caption         =   "No Download Limit For Selected PC"
         Height          =   375
         Left            =   960
         TabIndex        =   178
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox sDlUnlock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Limit For Unlocked PC"
         Height          =   315
         Left            =   480
         TabIndex        =   174
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox uDllimit 
         Height          =   285
         Left            =   6120
         TabIndex        =   176
         Text            =   "100"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Set Download Limit For  All PC's"
         Height          =   375
         Left            =   5520
         TabIndex        =   179
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CheckBox sDlLock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock On Download Limit Exceed"
         Height          =   315
         Left            =   480
         TabIndex        =   175
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000C&
         Height          =   1455
         Left            =   240
         Top             =   720
         Width           =   4575
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000C&
         Height          =   1455
         Left            =   4800
         Top             =   720
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   495
         Left            =   3360
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mb"
         Height          =   255
         Left            =   7440
         TabIndex        =   173
         Top             =   960
         Width           =   855
      End
      Begin VB.Label cDLLimit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "100"
         Height          =   255
         Left            =   6960
         TabIndex        =   172
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Limit  :"
         Height          =   255
         Left            =   5880
         TabIndex        =   171
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mb"
         Height          =   255
         Left            =   7320
         TabIndex        =   177
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label uDL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4680
         TabIndex        =   168
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Downloaded :"
         Height          =   255
         Left            =   3600
         TabIndex        =   169
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Index           =   4
      Left            =   1800
      TabIndex        =   180
      Top             =   6360
      Width           =   9255
      Begin VB.CommandButton Command18 
         Caption         =   "Remove All Blocks"
         Height          =   375
         Left            =   1080
         TabIndex        =   184
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Block Website"
         Height          =   375
         Left            =   3480
         TabIndex        =   185
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox bWeb 
         Height          =   375
         Left            =   600
         TabIndex        =   183
         Text            =   "http://BlockThisWebsite.com"
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Open "
         Height          =   375
         Left            =   4320
         TabIndex        =   182
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox cWebsite 
         Height          =   375
         Left            =   600
         TabIndex        =   181
         Text            =   "http://madsacsoft.com"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H8000000A&
         Height          =   1095
         Left            =   480
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H8000000A&
         Height          =   615
         Left            =   480
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   8
      Left            =   2040
      TabIndex        =   115
      Top             =   1200
      Width           =   9855
   End
   Begin VB.Frame mfrm 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Index           =   2
      Left            =   2280
      TabIndex        =   42
      Top             =   1560
      Width           =   9855
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pause Printing"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   2640
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvwPrinters 
         Height          =   1455
         Left            =   480
         TabIndex        =   43
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList32"
         SmallIcons      =   "ImageList16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Shape Shape18 
         Height          =   2415
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   5175
      End
   End
   Begin VB.Label cfePhone 
      Caption         =   "phone"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label CfeOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label cfeAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   615
      Left            =   11400
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label cfeCode 
      Caption         =   "Code"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   360
      Picture         =   "Winsock.frx":1E108
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "cWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form within form dec
'
'Option Explicit
Public uCurrency As String
Public SocketCounter As Integer
Dim wD(10), DBPath As String
Dim clCLount, sockB, timeCounter As Integer
Private udbPath As String
Private Const txtport As String = 1996
Public rCount As Integer
Public SEL As Integer
Private m_X As Single
Private m_Y As Single

Private Sub atTitle_Timer()
wSend "atit", CInt(SEL)
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
atTitle.Enabled = True
Else
atTitle.Enabled = False
End If
End Sub

Private Sub Check3_Click()
FPrinters.pAll
End Sub

Private Sub cHome_Click()
lvwDB.Visible = True
For i = 0 To mfrm.Count - 1
mfrm(i).Visible = False
Next i
FPrinters.Hide
End Sub

Private Sub clipBrdGet_Click()
wSend "gcli", CInt(SEL)
End Sub

Private Sub clipBrdSet_Click()
wSend "scli" & clipBrd.Text, CInt(SEL)
End Sub

Private Sub cmbChooseView_Click()
If cmbChooseView.ListIndex = 0 Then
lvwDB.View = 0
Else
lvwDB.View = 3
End If
End Sub

Private Sub Command10_Click()
If SEL <> 0 Then
wSend "iden", CInt(SEL)
Else
MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End If
End Sub

Private Sub Command12_Click()
If SEL <> 0 Then sTransfer.Show vbModal Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End Sub

Private Sub Command13_Click()
wSend "atit", CInt(SEL)
End Sub

Private Sub Command11_Click()
If SEL <> 0 Then cReciept.Show vbModal Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"

End Sub


Private Sub Command14_Click()
sock1(SEL).Close
End Sub



Private Sub Command16_Click()
wSend "wopn" & cWebsite.Text, CInt(SEL)
End Sub

Private Sub Command17_Click()
wSend "wblo" & bWeb.Text, CInt(SEL)
End Sub

Private Sub Command18_Click()
wSend "cblo", CInt(SEL)
End Sub

Private Sub Command19_Click()
If Command19.Caption = "Ok" Then
lvwPrice.ListItems(lvwPrice.SelectedItem.Index).SubItems(1) = teMin.Text
lvwPrice.ListItems(lvwPrice.SelectedItem.Index).SubItems(2) = teCha.Text
Command19.Caption = "Edit"
Command20.Enabled = True
Command21.Enabled = True
lstSave
Else
Command19.Caption = "Ok"
Command20.Enabled = False
Command21.Enabled = False
End If
End Sub

Private Sub Command2_Click()
If SEL <> 0 Then tmAdd.Show vbModal Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End Sub

Private Sub Command20_Click()
If Command20.Caption = "Ok" Then
Dim n As Integer
Dim cCon As Boolean
For n = 1 To lvwPrice.ListItems.Count
If lvwPrice.ListItems(n).SubItems(1) = teMin.Text Then cCon = True
Next n

If cCon = False Then
lvwPrice.ListItems.Add , , lvwPrice.ListItems.Count
lvwPrice.ListItems(lvwPrice.ListItems.Count).SubItems(1) = teMin.Text
lvwPrice.ListItems(lvwPrice.ListItems.Count).SubItems(2) = teCha.Text
Command20.Caption = "Add"
Command19.Enabled = True
Command21.Enabled = True
lstSave
Else
MsgBox "Price for this value already exist ! ", vbOKOnly + vbInformation, "Entry already exist "
End If

Else
Command20.Caption = "Ok"
Command19.Enabled = False
Command21.Enabled = False
End If

End Sub

Private Sub Command21_Click()
Dim Mres As Integer
If lvwPrice.ListItems.Count - 1 > 0 Then
Mres = MsgBox(" Are you sure to remove selected item from PriceList ! ", vbInformation + vbYesNo, "Delete ?")
Select Case Mres
    Case 6
       lvwPrice.ListItems.Remove (lvwPrice.SelectedItem.Index)
       lstSave
End Select
Else
MsgBox " No Item will be left,hence you cannot delete selected item ! ", vbCritical + vbOKOnly, "Cannot Delete !"
End If
End Sub



Private Sub Command22_Click()
cToken.Show
End Sub

Private Sub Command23_Click()
If SEL <> 0 Then wSend "pkil", CInt(SEL) Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"

End Sub

Private Sub Command24_Click()
cOpt.Show vbModal
End Sub
Private Sub Command25_Click()
cCheck.Show
End Sub

Private Sub Command26_Click()
cWinsock.LoadAll
End Sub

Private Sub Command27_Click()
Frame6.Visible = True
Command31.Caption = "Ok"
Label10.Visible = True
tmrName.Visible = True
End Sub

Private Sub Command28_Click()
Frame4.Visible = False
bEnabled "edit", True
bEnabled "add", False

End Sub

Private Sub Command29_Click()
Dim Mres As Integer
 Mres = MsgBox(" Are you sure to delete selected timer ?", vbInformation + vbYesNo, "Delete Timer")
 If Mres = 6 Then
 Unload tmr(lvwTime.SelectedItem.Index)
 lvwTime.ListItems.Remove lvwTime.SelectedItem.Index

 End If
End Sub

Private Sub Command3_Click()

If SEL <> 0 Then addTm.Show vbModal Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End Sub


Private Sub Command30_Click()
Frame6.Visible = True
Command31.Caption = "Done"
End Sub

Private Sub Command31_Click()
If Command31.Caption = "Ok" Then
Dim i, cVal As Integer

cVal = 1
For i = 1 To lvwTime.ListItems.Count
If CInt(lvwTime.ListItems(i).Text) = cVal Then
cVal = cVal + 1
i = 1
End If
Next i

lvwTime.ListItems.Add , , cVal
Load tmr(cVal)
tmr(cVal).Enabled = True
lvwTime.ListItems(cVal).SubItems(1) = tmrName.Text
lvwTime.ListItems(cVal).SubItems(2) = Format$(Now, "hh:mm:ss AM/PM")
lvwTime.ListItems(cVal).SubItems(3) = tmrTot.Text
lvwTime.ListItems(cVal).SubItems(4) = 0
lvwTime.ListItems(cVal).SubItems(5) = 0
Else
lvwTime.ListItems(lvwTime.SelectedItem.Index).SubItems(1) = tmrName.Text
lvwTime.ListItems(lvwTime.SelectedItem.Index).SubItems(3) = tmrTot.Text
End If
Frame6.Visible = False
End Sub

Private Sub Command32_Click()
If tmr(lvwTime.SelectedItem.Index).Enabled = True Then
tmr(lvwTime.SelectedItem.Index).Enabled = False
Command32.Caption = "Resume Timer"
Else
tmr(lvwTime.SelectedItem.Index).Enabled = True
Command32.Caption = "Pause Timer"
End If

End Sub

Private Sub Command33_Click()
Dim Mres As Integer
 Mres = MsgBox(" Are you sure to restart selected timer ?", vbInformation + vbYesNo, "Restart Timer")
 If Mres = 6 Then
lvwTime.SelectedItem.SubItems(5) = 0
tmr(lvwTime.SelectedItem.Index).Enabled = True
 End If

End Sub

Private Sub Command34_Click()
Frame6.Visible = False
End Sub

Private Sub Command35_Click()
uCurrency = InputBox("Enter your currency ?", "Currency")
End Sub

Private Sub Command36_Click()
Dim i As Integer
   For i = 1 To lvwDB.ListItems.Count
        Select Case lvwDB.ListItems.Item(i).SubItems(11)
        Case 0
            cWinsock.wSend "oshu", i
        Case 2
            cWinsock.wSend "oshu", i
        End Select
    Next i

End Sub

Private Sub Command37_Click()
cWinsock.SaveAll
End Sub

Private Sub Command38_Click()
wSend "exit", CInt(SEL)
End Sub

Private Sub Command39_Click()
Dim lwidth As Integer
lwidth = 9000
If Command39.Caption = "Show Passwords" Then
uDb.ColumnHeaders(9).Width = lwidth * 0.2
Command39.Caption = "Hide Password"
Else
uDb.ColumnHeaders(9).Width = lwidth * 0
Command39.Caption = "Show Passwords"
End If
End Sub

Private Sub Command4_Click()
 Dim Mres As Integer
 Mres = MsgBox(" Are you sure to delete selected token ?", vbInformation + vbYesNo, "Delete Token")
 If Mres = 6 Then
 lvwToken.ListItems.Remove lvwToken.SelectedItem.Index
 ReNumList lvwToken
 End If
End Sub

Private Sub Command40_Click()
Frame8.Visible = True
Command50.Caption = "Add Discount"
End Sub

Private Sub Command41_Click()
Frame8.Visible = True
Command50.Caption = "Add Charge"
End Sub

Private Sub Command44_Click()
cfeCode.Caption = "test"
End Sub

Private Sub Command45_Click()
wSend "prog", CInt(SEL)
mShow 7
End Sub

Private Sub Command46_Click()
addEmployee.Visible = True
eName.Text = "Employee Name"
ePin.Text = "pin"
eRepin.Text = "repin"
End Sub

Private Sub Command47_Click()
 Select Case MsgBox(" Delete selected employee from employee list ? ", vbInformation + vbYesNo, "Remove Employee !")
    Case 6
    Dim i, cO As Integer
    cO = 0
    For i = 1 To lvwAdmins.ListItems.Count
    If lvwAdmins.ListItems(i).SubItems(2) = "Administrator" Then cO = cO + 1
    Next i
    
    If cO < 2 And lvwAdmins.SelectedItem.SubItems(2) = "Administrator" Then
    MsgBox "You cannot delete selected employee as no administrator will be left !", vbOKOnly + vbCritical, "Cannot delete employee !"
    Else
         lvwAdmins.ListItems.Remove lvwAdmins.SelectedItem.Index
    End If
    Case 7
End Select
End Sub

Private Sub Command48_Click()
Dim eExist As Boolean

If ePin.Text = eRepin Then
    If eRights.Text <> "" Then
    
    eExist = False
    
    For i = 1 To lvwAdmins.ListItems.Count
        If eName.Text = lvwAdmins.ListItems(1).SubItems(1) Then eExist = True
    Next i
    
    If eExist = False Then
        lvwAdmins.ListItems.Add , , lvwAdmins.ListItems.Count + 1
        lvwAdmins.ListItems(lvwAdmins.ListItems.Count).SubItems(1) = eName.Text
        lvwAdmins.ListItems(lvwAdmins.ListItems.Count).SubItems(2) = eRights.Text
        lvwAdmins.ListItems(lvwAdmins.ListItems.Count).SubItems(3) = ePin.Text
        cWinsock.SaveAll
        addEmployee.Visible = False
    Else
        MsgBox "Employee name already exist !", vbInformation + vbOKOnly, "Employee exist !"
    End If
    Else
    MsgBox "Please choose valid employee rights !", vbInformation + vbOKOnly, "Invalid employee rights !"
    End If
Else
MsgBox "Please enter confirmaton pin correctely !", vbInformation + vbOKOnly, "Invalid confirmation pin !"
End If
End Sub

Private Sub Command49_Click()
addEmployee.Visible = False
End Sub

Private Sub Command5_Click()
 Dim Mres As Integer
 If SEL <> 0 Then
 If sock1(CInt(GETSOCK(CInt(SEL)))).State = 7 Then
 Mres = MsgBox(" Are you sure to Unlock  " & lblsEl.Caption & " ? ", vbInformation + vbYesNo, "Not connected")
 Select Case Mres
    Case 6

        lsub14 Val(SEL), 0
        lClear Val(SEL)
       wSend "uock", Val(SEL)
       cWinsock.lvwDB.ListItems.Item(cWinsock.SEL).SubItems(11) = 4
    Case 7
End Select

 Else

 Mres = MsgBox(" Selected terminal is not connected. Do you want to continue ? ", vbInformation + vbYesNo, "Not connected")
 Select Case Mres
    Case 6
     cWinsock.lvwDB.ListItems.Item(cWinsock.SEL).SubItems(11) = 4
    Case 7
End Select
End If
Else
MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End If

End Sub

Private Sub Command50_Click()

Select Case Command50.Caption

Case "Add Charge"
    If MsgBox("Add charge of " & cWinsock.uCurrency & " " & Text2.Text & " ? ", vbInformation + vbYesNo, "Add Charge") = 6 Then lvwDB.SelectedItem.SubItems(14) = Text2.Text
Case "Add Discount"
    If MsgBox("Add discount of " & cWinsock.uCurrency & " " & Text2.Text & " ? ", vbInformation + vbYesNo, "Add Discount") = 6 Then lvwDB.SelectedItem.SubItems(16) = Text2.Text
End Select
Frame8.Visible = False
End Sub


Private Sub Command52_Click()
Frame8.Visible = False
End Sub

Private Sub Command7_Click()
 Dim Mres As Integer
 If SEL <> 0 Then
 If sock1(cWinsock.GETSOCK(cWinsock.SEL)).State = 7 Then
 Mres = MsgBox(" Are you sure to Close current session of  " & lblsEl.Caption & " ? ", vbInformation + vbYesNo, "Not connected")
 Select Case Mres
    Case 6
        wSend "cpay", CInt(SEL)
         lsub14 Val(SEL), 0
         lClear Val(SEL)

    Case 7
End Select

 Else

 Mres = MsgBox(" Selected terminal is not connected. Do you want to continue ? ", vbInformation + vbYesNo, "Not connected")
 Select Case Mres
    Case 6
    lsub14 CInt(SEL), 0
    lClear CInt(SEL)
End Select
End If
Else
MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"
End If
End Sub
Private Sub Command8_Click()
Dim i As Integer
For i = 1 To lvwDB.ListItems.Count
wSend "sedl" & uDllimit.Text, CInt(i)
Next i
cDLLimit.Caption = uDllimit.Text
End Sub

Private Sub Command9_Click()

If SEL <> 0 Then cRDP.Show Else MsgBox " Please select a Terminal ", vbExclamation + vbOKOnly, " No terminal Selected"

End Sub

Function Encrypt(Text As String, pw As String, type_of As Boolean) As String
Dim X As Integer
Dim i As Integer
Dim a As Integer
Dim text_chr As String
Dim text_asc As Integer
Dim pw_chr As String
Dim pw_asc As Integer
Dim fin As String
Dim fin_chr As String
Dim fin_asc As Integer

'Written by Dan Andersen (Gozer)
'if type_of = true then Encrypt
'if type_of = False then Decrypt

'making sure there is text to encrypt and a password
'to go with it
If Len(Text$) = 0 Then Exit Function
If Len(pw$) = 0 Then Exit Function

X% = 1
'the X variable is the loop that goes through the password
'characters individually through the encrpytion processes
'X = 1 to set the loop at the first character
For i% = 1 To Len(Text$) 'start of encrpyt loop
    'taking out characters from text to encrypt
    'the single character
    text_chr$ = Mid(Text$, i, 1)
    'changing the character to its ASCII value to
    'easily change the character for encrypting
    text_asc% = Asc(text_chr$)
    'doing the same process with the password
    'using the X variable
    pw_chr$ = Mid(pw$, X, 1)
    pw_asc% = Asc(pw_chr$)
    'adding up variable to continue loop through different
    'characters within the password
    X% = X% + 1
    If X% > Len(pw$) Then X% = 1 'restarting password loop
    'Case to check if the user is Encrypting or Decrypting the text
    Select Case type_of
    Case True: 'Encrypting
        'adding the characters of both string and password
        fin_asc% = text_asc% + pw_asc%
        'making sure the final_asc will equal a valid ASCII character
        If fin_asc% > 255 Then
            'Character was an invalid character so we modify it
            'to equal a valid character
            a% = fin_asc% - 255
            fin_chr$ = Chr$(a%)
        Else
            'character was valid ;D
            fin_chr$ = Chr$(fin_asc%)
        End If
    Case False: 'Decrypting
        'here we subtract the characters...does the opposite of
        'what encrypting does to put it back in its
        'original state, which is why it's called Decrypting
        fin_asc% = text_asc% - pw_asc% 'subtracting character values
        If fin_asc% < 1 Then   'checking for invalid character
            'invalid character..fixing problem =)
            a% = fin_asc% + 255
            fin_chr$ = Chr$(a%)
        Else
            'it was all good.
            fin_chr$ = Chr$(fin_asc%)
        End If
    End Select 'End of case
    'adding the final encrypted character to a string
    'to be later shown in its final state at the end
    fin$ = fin$ & fin_chr$
    'thought i'd be mr. fancy pants by adding a little
    'percentage bar =)

Next 'continuing loop =)
'finalizing function to equal the final encrypted string
Encrypt$ = fin$
DoEvents
'i'm mr. fancy pants

End Function

Private Sub disALL_Click()
For i = 1 To SocketCounter
Unload sock1(i)
Next i
SocketCounter = 0
lvwDB.ListItems.Clear
Broadcast
End Sub

Private Sub Form_Load()
SEL = 0
addEmployee.Visible = False
FPrinters.Show
FPrinters.Hide
lvwToken.Left = 120
lvwToken.Top = 220
Dim lwidth As Integer
lwidth = 7000
Frm(0).ZOrder
    
    
    With lvwPrice
        lwidth = .Width
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Index", 0                       '0
        .ColumnHeaders.Add , , "Time ( Min. )", lwidth * 0.49   '1
        .ColumnHeaders.Add , , "Price ( Unit )", lwidth * 0.49  '2
    End With
    lwidth = 9000
  ' With lvwProcess
   '     .ColumnHeaders.Clear
    '    .ColumnHeaders.Add , , "Index", 0                       '0
     '   .ColumnHeaders.Add , , "Program Name", lwidth * 0.49   '1
      '  .ColumnHeaders.Add , , "Process Name", lwidth * 0.49  '2
   ' End With
   
    With lvwDB
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Terminal Name", lwidth * 0.2        '0
        .ColumnHeaders.Add , , "State", lwidth * 0.23               '1
        .ColumnHeaders.Add , , "Name", lwidth * 0.2                 '2
        .ColumnHeaders.Add , , "Start Time", lwidth * 0.15          '3
        .ColumnHeaders.Add , , "Total Time (min)", lwidth * 0.15    '4
        .ColumnHeaders.Add , , "Time Left (min)", lwidth * 0.15     '5
        .ColumnHeaders.Add , , "Downloaded", lwidth * 0.13          '6
        .ColumnHeaders.Add , , "Uploaded", lwidth * 0               '7
        .ColumnHeaders.Add , , "ID Type", lwidth * 0.11             '8
        .ColumnHeaders.Add , , "ID Number", lwidth * 0.11           '9
        .ColumnHeaders.Add , , "Mob. Number", lwidth * 0.15         '10
        .ColumnHeaders.Add , , "Status", lwidth * 0                 '11
        .ColumnHeaders.Add , , "Time used", lwidth * 0              '12
        .ColumnHeaders.Add , , "Total Charge ", lwidth * 0.15       '13
        .ColumnHeaders.Add , , "Charge Added ", lwidth * 0.15       '14
        .ColumnHeaders.Add , , "UIN", lwidth * 0                    '15
        .ColumnHeaders.Add , , "Discount", lwidth * 0.1             '16
    End With
    With uDb
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No.", lwidth * 0.08           '0
        .ColumnHeaders.Add , , "Name", lwidth * 0.2             '1
        .ColumnHeaders.Add , , "User Type", lwidth * 0.15       '2
        .ColumnHeaders.Add , , "Total Amount", lwidth * 0.13    '3
        .ColumnHeaders.Add , , "Amount Used", lwidth * 0.13     '4
        .ColumnHeaders.Add , , "ID Type", lwidth * 0.11         '5
        .ColumnHeaders.Add , , "ID Number", lwidth * 0.11       '6
        .ColumnHeaders.Add , , "Mob. Number", lwidth * 0.14     '7
        .ColumnHeaders.Add , , "Password", lwidth * 0           '8
        .ColumnHeaders.Add , , "Address", lwidth * 0.25         '9
    End With
    With lvwToken.ColumnHeaders
        .Clear
        .Add , , "S.No.", lwidth * 0.08             '0
        .Add , , "Token Name/Number", lwidth * 0.2  '1
        .Add , , " Amount (Min.) ", lwidth * 0.15   '2
        .Add , , " User Name ", lwidth * 0.2        '3
        .Add , , "Password", 0                      '4
        .Add , , " ID Type ", lwidth * 0.15         '5
        .Add , , " ID Number ", lwidth * 0.15       '6
        .Add , , " Phone ", lwidth * 0.15           '7
    End With
    With lvwTime.ColumnHeaders
    .Clear
        .Add , , "S.No.", lwidth * 0                '0
        .Add , , "User Name", lwidth * 0.2          '1
        .Add , , "Start Time", lwidth * 0.15        '2
        .Add , , "Total Time", lwidth * 0.15        '3
        .Add , , "Time Remaining", lwidth * 0.15    '4
        .Add , , "Time Used", lwidth * 0.15         '5
    End With
    With lvwAdmins.ColumnHeaders
        .Clear
        .Add , , "S.No.", lwidth * 0                '0
        .Add , , "Employee Name", lwidth * 0.2          '1
        .Add , , "Employee Type", lwidth * 0.15        '2
        .Add , , "Password", lwidth * 0      '3

    End With

Listen
SckUdp.Close
Broadcast
sResize

DBPath = "c:\db.txt"

eRights.Clear
'eRights.Style = 2
eRights.AddItem "Administrator"
eRights.AddItem "Viewer"
eRights.AddItem "Moderator"


idType.Clear
'idType.Style = 2
idType.AddItem "Voter ID"
idType.AddItem "Telephone Bill"
idType.AddItem "Electric Bill"
idType.AddItem "Adhar Card/UID"
idType.AddItem "Identity Card"
idType.AddItem "Passport"
idType.AddItem "Credit/Debit Card"
idType.AddItem "PAN Card"

uType.Clear
'uType.Style = 2
uType.AddItem "Normal"
uType.AddItem "Gamer"
uType.AddItem "Office Worker"
uType.AddItem "Student"
uType.AddItem "Old Age"

Frame6.Visible = False

LoadAll
dEdit

bEnabled "edit", True
bEnabled "add", False
mShow 0
mfrm(0).Visible = False
lvwDB.Visible = True

'''Price List
'teMin.Text = lvwPrice.ListItems(lvwPrice.ListItems.Count).SubItems(1)
'teCha.Text = lvwPrice.ListItems(lvwPrice.ListItems.Count).SubItems(2)
Dim i As Integer
For i = 0 To lblDET.Count - 1
lblDET(i).Top = lblDET(i).Top + 50
Next i





For i = 0 To mComCol.Count - 1
mComCol(i).ButtonType = [Windows XP]
Next i
For i = 0 To Others.Count - 1
Others(i).ButtonType = [Windows XP]
Next i
cHome.ButtonType = [Windows XP]



End Sub





Private Sub lvwPrinters_DblClick()
Dim Frm As Form
 Dim inf As CPrinterInfo
Set inf = FPrinters.GetPrinter(lvwPrinters.SelectedItem.Tag)

If lvwPrinters.SelectedItem.Tag = m_NewPrn Then
      Call Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter")
   Else

               ' Check to see if form is already loaded.
         For Each Frm In Forms
            If TypeOf Frm Is FPrinter Then
               If Frm.DeviceName = inf.DeviceName Then
                  Frm.SetFocus
                  Exit For
               End If
            End If
         Next Frm
         ' Create new form is one wasn't found.
         If Screen.ActiveForm Is Me Then
            Set Frm = New FPrinter
            Frm.DeviceName = inf.DeviceName
            Frm.Show , Me
         End If
  End If
End Sub

Private Sub lvwTime_ItemClick(ByVal Item As MSComctlLib.ListItem)
tmrName.Text = lvwTime.SelectedItem.SubItems(1)
tmrTot.Text = lvwTime.SelectedItem.SubItems(3)
If tmr(lvwTime.SelectedItem.Index).Enabled = True Then
Command32.Caption = "Pause Timer"
Else
Command32.Caption = "Resume Timer"
End If

End Sub

Private Sub mComCol_Click(Index As Integer)
mShow Index
End Sub

Private Sub Mob_Change()
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
    End Select
End Sub

Private Sub Others_Click(Index As Integer)

Select Case Others(Index).Caption

Case "Refresh"
Dim Xi As Integer
For Xi = 1 To lvwDB.ListItems.Count
If sock1(Xi).State = 7 Then wSend "atat", CInt(Xi)
Next Xi

Case "Search PC"
Broadcast

Case "LogOut"
cWinsock.Visible = False
cWinsock.Enabled = False
frmLogin.Show

End Select


Select Case Index

Case 2
If Others(2).Caption = "Normal" Then
lvwDB.View = lvwIcon
Others(2).Caption = "Show Detail"
cmbChooseView.Text = "Icon View"
Else
lvwDB.View = 3
cmbChooseView.Text = "Table View"
Others(2).Caption = "Normal"
End If


End Select

End Sub

Private Sub tAmount_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
    End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select

End Sub

Private Sub tmr_Timer(Index As Integer)

    lvwTime.ListItems(Index).SubItems(5) = CInt(lvwTime.ListItems(Index).SubItems(5)) + 1
    lvwTime.ListItems(Index).SubItems(4) = CInt(lvwTime.ListItems(Index).SubItems(3)) - CInt(lvwTime.ListItems(Index).SubItems(5))
If lvwTime.ListItems(Index).SubItems(4) = 0 Then
MsgBox "Time Over of : " & lvwTime.ListItems(Index).SubItems(1), vbOKOnly + vbInformation + vbSystemModal, "Time Over"
tmr(Index).Enabled = False
End If
End Sub

Private Sub tsAmount_Change()
tAmount.Text = tsAmount.Max - tsAmount.Value
End Sub

Private Sub uDb_ItemClick(ByVal Item As MSComctlLib.ListItem)
dEdit
bEnabled "edit", True
bEnabled "add", False
sNo.Text = uDb.SelectedItem.Text
cName.Text = uDb.SelectedItem.SubItems(1)
'uType.Text = uDb.SelectedItem.SubItems(2)
tAmount.Text = uDb.SelectedItem.SubItems(3)
aUsed.Text = uDb.SelectedItem.SubItems(4)
idNum.Text = uDb.SelectedItem.SubItems(6)
Mob.Text = uDb.SelectedItem.SubItems(7)
pass.Text = uDb.SelectedItem.SubItems(8)
cAddress.Text = uDb.SelectedItem.SubItems(9)

End Sub
Private Sub cmdAdd_Click()
eEdit
bEnabled "ok", True
cmddOK.Visible = True
cmdSave.Visible = False
sNo.Enabled = False

End Sub
Private Sub cmdDelete_Click()
If uDb.ListItems.Count <> 0 Then
uDb.ListItems.Remove uDb.SelectedItem.Index
reNum
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
cWinsock.SaveAll

If Check1.Value = 1 Then
Cancel = 1
Me.Hide
Else


Main.CloseAll
End If
End Sub

Private Sub Form_Resize()
sResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.CloseAll
End Sub



Private Sub lvwDB_Click()
Call clCheck
End Sub
Private Sub lvwDB_GotFocus()
Call clCheck
End Sub

Private Sub lvwDB_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call clCheck
End Sub

Private Sub lvwPrice_Click()
teMin.Text = lvwPrice.ListItems(lvwPrice.SelectedItem.Index).SubItems(1)
teCha.Text = lvwPrice.ListItems(lvwPrice.SelectedItem.Index).SubItems(2)
End Sub

Private Sub lvwToken_ItemClick(ByVal Item As MSComctlLib.ListItem)
If tPassword.Value = 0 Then
   tPass.Text = "Password"
 Else
  tPass.Text = lvwToken.SelectedItem.SubItems(4)
 End If
End Sub

Private Sub nDlSel_Click()
wSend "nodl", CInt(SEL)
End Sub

Private Sub sock1_Close(Index As Integer)
sock1(Index).Close
txtLog = txtLog & "Client" & Index & " -> *** Disconnected" & vbCrLf

End Sub
Private Sub sock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'**************************************************************************'
'                           SOCK DATA Arrival                              '
'
'*************************************************************************'

Dim mReply As Integer

Dim dat, aRR() As String
Dim cHrr() As String
sock1(Index).GetData dat, vbString
cHrr = Split(dat, "|½|")
dat = cHrr(0)
wD(0) = Left$(dat, 4)
wD(0) = LCase(wD(0))
wD(1) = Right(dat, Len(dat) - 4)

If wD(0) = "cfen" Then
Dim found As Boolean
found = False

For i = 1 To lvwDB.ListItems.Count
    If wD(1) = lvwDB.ListItems(i).Text Then

        If sock1(CInt(lvwDB.ListItems(i).SubItems(15))).RemoteHostIP = sock1(Index).RemoteHostIP Then
            lvwDB.ListItems.Item(i).SubItems(15) = Index
            found = True
        End If
    End If
Next i

If found = False Then
    lvwDB.ListItems.Add , , wD(1), 8
    lvwDB.ListItems(lvwDB.ListItems.Count).SubItems(11) = 0
    lvwDB.ListItems(lvwDB.ListItems.Count).SubItems(15) = Index
    lClear lvwDB.ListItems.Count
    lsub14 lvwDB.ListItems.Count, 0
End If
wSend "cfen" & cfeName.Caption, Index

Else
For i = 1 To lvwDB.ListItems.Count
    If Index = CInt(lvwDB.ListItems.Item(i).SubItems(15)) Then
        Index = i
    End If
Next i
End If

Select Case wD(0)

Case "hnam" 'send terminal name
        lvwDB.ListItems.Item(Index).Text = wD(1)
        cWinsock.wSend "odat" & lvwDB.ListItems.Item(Index).SubItems(11) & " | " & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lvwDB.ListItems.Item(Index).SubItems(2) & " | " & cWinsock.lvwDB.ListItems.Item(Index).SubItems(8) & " | " & cWinsock.lvwDB.ListItems.Item(Index).SubItems(9) & " | " & cWinsock.lvwDB.ListItems.Item(Index).SubItems(10) & " | " & cWinsock.lvwDB.ListItems.Item(Index).SubItems(4) & " | " & lvwDB.ListItems.Item(Index).SubItems(3) & " | " & lvwDB.ListItems.Item(Index).SubItems(12), Index
        If lvwDB.ListItems.Item(Index).SubItems(11) = "1" Then lvwDB.ListItems.Item(Index).SubItems(11) = "3"
Case "stat" 'simple stat
        aRR = Split(wD(1), "|")
        lvwDB.ListItems.Item(Index).SubItems(6) = aRR(0)
        lvwDB.ListItems.Item(Index).SubItems(7) = aRR(1)
        lvwDB.ListItems.Item(Index).SubItems(5) = aRR(2)
Case "atat" 'full stat
        aRR = Split(wD(1), "|")
        lvwDB.ListItems.Item(Index).SubItems(6) = aRR(0)
        lvwDB.ListItems.Item(Index).SubItems(7) = aRR(1)
        lvwDB.ListItems.Item(Index).SubItems(2) = aRR(2)
        lvwDB.ListItems.Item(Index).SubItems(8) = aRR(3)
        lvwDB.ListItems.Item(Index).SubItems(9) = aRR(4)
        lvwDB.ListItems.Item(Index).SubItems(3) = aRR(5)
        lvwDB.ListItems.Item(Index).SubItems(4) = aRR(6)
        lvwDB.ListItems.Item(Index).SubItems(12) = aRR(7)
        lvwDB.ListItems.Item(Index).SubItems(5) = aRR(8)
        lvwDB.ListItems.Item(Index).SubItems(10) = aRR(9)
       If aRR(10) = "True" Then lvwDB.ListItems.Item(Index).SubItems(11) = 3
Case "sfin" 'time over
lvwDB.ListItems.Item(Index).SubItems(11) = 2
Case "crdp" 'remootedesktop
'cRDP.sGet cWinsock.sock1(Index).RemoteHostIP, wD(1)
Case "tmsk" 'user ask for time
    If aRenew.Value = 1 Then
        data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | Increased : " & wD(1) & " min."
        cWinsock.cLog (data)
        cWinsock.wSend "tadd" & wD(1), Index
    Else
        Me.WindowState = vbMaximized
        mReply = MsgBox(cWinsock.lvwDB.ListItems.Item(Index) & " requested to increase " & wD(1) & " min. Increase ?", vbYesNoCancel + vbQuestion, "Increase Time")
            If mReply = 6 Then
                data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | Increased : " & wD(1) & " min."
                cWinsock.cLog (data)
cWinsock.wSend "tadd" & wD(1), Index
Else
cWinsock.wSend "cmsgYour request to increase time was disapproved ! |0|Request Disapproved", Index
End If
End If

Case "nlog" 'unknown user ask login
aRR = Split(wD(1), "|")
    If aLogin.Value = 1 Then
       data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & wD(1)
                cWinsock.cLog (data)
                
                cWinsock.wSend "tset" & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & aRR(0) & " | " & aRR(1) & " | " & aRR(2) & " | " & aRR(3) & " | " & aRR(4), Index
                dSince = Format$(Now, "hh:mm:ss AM/PM")

                    With cWinsock.lvwDB.ListItems.Item(Index)
                        .SubItems(11) = 3
                        .SubItems(12) = 0
                        .SubItems(2) = aRR(0)
                        .SubItems(3) = dSince
                        .SubItems(4) = aRR(4)
                        .SubItems(6) = " 0 "
                        .SubItems(7) = " 0 "
                        .SubItems(8) = aRR(1)
                        .SubItems(9) = aRR(2)
                        .SubItems(10) = aRR(3)
                    End With

    Else
        Me.WindowState = vbMaximized
        mReply = MsgBox(aRR(0) & " at " & cWinsock.lvwDB.ListItems.Item(Index) & " requested to login for " & aRR(4) & " min. Login ?", vbYesNoCancel + vbQuestion, "Increase Time")
            If mReply = 6 Then
                data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & wD(1)
                cWinsock.cLog (data)
                
                cWinsock.wSend "tset" & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & aRR(0) & " | " & aRR(1) & " | " & aRR(2) & " | " & aRR(3) & " | " & aRR(4), Index
                dSince = Format$(Now, "hh:mm:ss AM/PM")

                    With cWinsock.lvwDB.ListItems.Item(Index)
                        .SubItems(11) = 3
                        .SubItems(12) = 0
                        .SubItems(2) = aRR(0)
                        .SubItems(3) = dSince
                        .SubItems(4) = aRR(4)
                        .SubItems(6) = " 0 "
                        .SubItems(7) = " 0 "
                        .SubItems(8) = aRR(1)
                        .SubItems(9) = aRR(2)
                        .SubItems(10) = aRR(3)
                    End With

            Else
                cWinsock.wSend "lmsg", Index
            End If
    End If
Case "1loc" 'lock

         cWinsock.lsub14 Index, 0
         cWinsock.lClear Index
         
Case "gcli" 'clipoard
clipBrd.Text = wD(1)
Case "atit" 'active window title
caTitle.Text = wD(1)
Case "sett" 'asks for setting
wSend "sett" & cUSB.Value & "|" & wDL.Value & "|" & _
sDlUnlock.Value & "|" & sDlLock.Value & "|" & cDLLimit.Caption, Index
Case "dlex" 'download limit exceed
If lvwDB.ListItems.Item(Index).SubItems(11) <> 4 Then MsgBox cWinsock.lvwDB.ListItems.Item(Index) & " Exceeded his download limit ", vbOKOnly + vbInformation, "Download limit Exceeded"
Case "usrl" 'user logout
        wSend "cpay", Index
        lsub14 Index, 0
        lClear Index
Case "ulog" 'username login
aRR = Split(wD(1), "|")
If ChPass(aRR(0), aRR(1)) <> 0 Then
    With uDb.ListItems(ChPass(aRR(0), aRR(1)))
        wSend "ulda" & .SubItems(1) & "|" & .SubItems(3) & "|" & .SubItems(4) & "|", Index
    End With
Else
    wSend "ulda" & "0|0|", Index
End If

Case "ulti" 'username ask time
aRR = Split(wD(1), "|")

        If Val(RateCalc(CInt(aRR(0)))) <= Val(uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(3)) - Val(uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(4)) Then
                uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(4) = CInt(uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(4)) + RateCalc(CInt(aRR(0)))
                data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & wD(1)
                cWinsock.cLog (data)
                cWinsock.wSend "tset" & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(1) & " | " & cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(5) & " | " & cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(6) & " | " & cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(7) & " | " & aRR(0), Index
                dSince = Format$(Now, "hh:mm:ss AM/PM")
                    With cWinsock.lvwDB.ListItems.Item(Index)
                        .SubItems(11) = 3
                        .SubItems(12) = 0
                        .SubItems(2) = cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(1)
                        .SubItems(3) = dSince
                        .SubItems(4) = aRR(0)
                        .SubItems(6) = " 0 "
                        .SubItems(7) = " 0 "
                        .SubItems(8) = cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(5)
                        .SubItems(9) = cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(6)
                        .SubItems(10) = cWinsock.uDb.ListItems(ChPass(aRR(1), aRR(2))).SubItems(7)
                    End With
                    SaveAll
        Else
            wSend "ulti", Index
        End If

Case "toke" 'token login
Dim userIndex As Integer
aRR = Split(wD(1), "|")
userIndex = ChPassToken(aRR(0), aRR(1))
        If userIndex <> 0 Then
                data = Now & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & wD(1)
                cWinsock.cLog (data)
                
                cWinsock.wSend "tset" & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lvwDB.ListItems.Item(Index) & " | " & cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(3) & " | " & cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(5) & " | " & cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(6) & " | " & cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(7) & " | " & cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(2), Index
                dSince = Format$(Now, "hh:mm:ss AM/PM")
                    With cWinsock.lvwDB.ListItems.Item(Index)
                        .SubItems(11) = 3
                        .SubItems(12) = 0
                        .SubItems(2) = cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(3)
                        .SubItems(3) = dSince
                        .SubItems(4) = cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(2)
                        .SubItems(6) = " 0 "
                        .SubItems(7) = " 0 "
                        .SubItems(8) = cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(5)
                        .SubItems(9) = cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(6)
                        .SubItems(10) = cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(7)
                    End With
                    cWinsock.lvwToken.ListItems(CInt(userIndex)).SubItems(2) = 0
                    SaveAll
        Else
            wSend "tnot", Index
        End If
Case "empl"
aRR = Split(wD(1), "|")
wSend "empl" & cWinsock.CheckAdmin(aRR(0), aRR(1)), Index
End Select
txtLog = txtLog & "Client" & Index & " : " & dat & vbCrLf
End Sub
Private Sub sock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim i As Integer
Dim X As Integer
X = 0
For i = 1 To lvwDB.ListItems.Count
If sock1(GETSOCK(i)).State <> 7 Then X = i
Next i
If X > 0 Then
sock1(X).Close
sock1(X).Accept requestID
sock1(X).SendData "Your Nick is ""Client" & X & """"
txtLog = "Client Connected. IP : " & sock1(X).RemoteHostIP & " , Client Nick : Client" & X & vbCrLf
Else
SocketCounter = SocketCounter + 1
Load sock1(SocketCounter)
sock1(SocketCounter).Accept requestID
txtLog = "Client Connected. IP : " & sock1(SocketCounter).RemoteHostIP & " , Client Nick : Client" & sockcounter & vbCrLf
sock1(SocketCounter).SendData "Your Nick is ""Client" & SocketCounter & """"

End If

End Sub

Private Sub sock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtLog = txtLog & "*** Error ( Client" & Index & ") : " & Description & vbCrLf
sock1_Close Index
Listen
End Sub

Public Sub Listen()
On Error Resume Next
Dim n As Integer
For n = 1 To SocketCounter
    sock1(n).Close
 
Next

On Error GoTo T
sock1(0).Close
sock1(0).LocalPort = txtport
sock1(0).Listen
txtLog = "Listening on Port " & txtport
Exit Sub
T:
MsgBox "Error : " & err.Number, vbCritical
If err.Number = 10048 Then Main.CloseAll

'Listen
End Sub

Private Sub Broadcast()
       On Error Resume Next
       Dim X As Integer
        For X = 4420 To 4440
            SckUdp.Close
            SckUdp.RemoteHost = "255.255.255.255"
            SckUdp.RemotePort = X
            SckUdp.SendData SckUdp.LocalIP & "|" & cWinsock.cfeCode.Caption & "|"
        Next
End Sub

Private Sub TabStrip1_Click()
Frm(TabStrip1.SelectedItem.Index - 1).ZOrder
End Sub

Private Sub teCha_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub

Private Sub teMin_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Dim SOnline As Integer
   Dim sFree As Integer
   Dim sEmployee As Integer
   cInfo(0).Caption = "Total PC's        : " & SocketCounter

  If sockB = 5 Then
  Broadcast
  sockB = 0
  Else
  sockB = sockB + 1
  End If
Dim i As Integer
   For i = 1 To SocketCounter
   With lvwDB.ListItems.Item(i)

   If sock1(CInt(lvwDB.ListItems.Item(i).SubItems(15))).State = 7 Then
    
    SOnline = SOnline + 1
    lvwDB.ListItems.Item(i).Ghosted = False
    
        Select Case lvwDB.ListItems.Item(i).SubItems(11)
        Case 0
        
        If .Icon <> 8 Then
            .Icon = 8
            .SubItems(1) = "Connected - Idle"
        End If
        
            sFree = sFree + 1
        Case 2
        If .Icon <> 8 Then
            .Icon = 8
           .SubItems(1) = "Connected - Idle"
           End If
            sFree = sFree + 1
        Case 1
        If .Icon <> 4 Then
            .Icon = 4
            .SubItems(1) = "Connected - Active"
           End If
        Case 3
        If .Icon <> 4 Then
            .Icon = 4
            .SubItems(1) = "Connected - Active"
            End If
        Case 4
        If .Icon <> 7 Then
            .Icon = 7
            .SubItems(1) = "Connected - Unlocked"
            End If
            sEmployee = sEmployee + 1
        End Select

   Else
   lvwDB.ListItems.Item(i).Ghosted = True
    Select Case lvwDB.ListItems.Item(i).SubItems(11)
        Case 0
        If .Icon <> 10 Then
            .Icon = 10
            .SubItems(1) = "Not Connected"
            End If
        Case 2
        If .Icon <> 8 Then
            .Icon = 8
            .SubItems(1) = "Not Connected - Idle"
            End If
            sFree = sFree + 1
       Case 3
       If .Icon <> 6 Then
            .Icon = 6
            .SubItems(1) = "Not Connected - Active"
            End If
        Case 1
        If .Icon <> 10 Then
            .Icon = 10
            .SubItems(1) = "Not Connected - Waiting"
            End If
        Case 4
        If .Icon <> 7 Then
            .Icon = 7
            .SubItems(1) = "Not Connected - Unlocked"
            End If
            sEmployee = sEmployee + 1
        End Select
        
   End If
End With
   Next i

cInfo(1).Caption = "Online PC's      : " & SOnline
cInfo(2).Caption = "Offline PC's      : " & SocketCounter - SOnline
cInfo(3).Caption = "Free PC's         : " & sFree
cInfo(4).Caption = "Unlocked PC's : " & sEmployee
Call clCheck


End Sub
Public Function wSend(data As String, sTo As Integer)
sTo = cWinsock.GETSOCK(sTo)
On Error Resume Next
   If AppToAll.Value = 1 Then
        For i = 1 To SocketCounter
            If sock1(i).State = 7 Then sock1(i).SendData data & "|½|"
        Next i
    AppToAll.Value = 0
   Else
      If sock1(sTo).State = 7 Then sock1(sTo).SendData data & "|½|"
   End If
End Function

Private Sub Timer2_Timer()
Dim i As Integer
For i = 1 To lvwDB.ListItems.Count
wSend "atat", i
Next i
If timeCounter >= 2 Then
SaveAll
timeCounter = 0
Else
timeCounter = timeCounter + 1
End If
End Sub

Private Sub cmddOK_Click()
Dim uExist, COndi As Boolean
uExist = False
COndi = True
For i = 1 To uDb.ListItems.Count
    If cName.Text = uDb.ListItems(i).SubItems(1) Then
    uExist = True
    COndi = False
    End If
Next i
If uExist = True Then
MsgBox "Username already exist !", vbOKOnly + vbInformation, "User exist !"
COndi = False
ElseIf idType.Text = "" Then
MsgBox "Invalid ID Type !", vbOKOnly + vbInformation, "Invalid ID !"
COndi = False
ElseIf uType.Text = "" Then
MsgBox "Invalid User Type !", vbOKOnly + vbInformation, "Invalid User Type !"
COndi = False
ElseIf idNum.Text = "Id Number" Then
MsgBox "Invalid ID Number !", vbOKOnly + vbInformation, "Invalid Number!"
COndi = False
ElseIf Len(Mob.Text) < 10 Then
MsgBox "Invalid Mobile Number !", vbOKOnly + vbInformation, "Invalid Mobile Number!"
COndi = False
End If

If checkCond And COndi Then
Dim lC As Integer
lC = uDb.ListItems.Count + 1
uDb.ListItems.Add , , lC
With uDb.ListItems(lC)
.SubItems(1) = cName.Text
.SubItems(2) = uType.Text
.SubItems(3) = tAmount.Text
.SubItems(4) = aUsed.Text
.SubItems(5) = idType.Text
.SubItems(6) = idNum.Text
.SubItems(7) = Mob.Text
.SubItems(8) = pass.Text
.SubItems(9) = cAddress.Text
End With
SaveAll
dEdit
bEnabled "edit", True
bEnabled "add", False
End If
End Sub

Private Sub cmdEdit_Click()
eEdit
bEnabled "save", True
sNo.Enabled = False
End Sub

Private Sub cmdSave_Click()
If checkCond Then
Dim lC As Integer
lC = uDb.SelectedItem.Index
With uDb.ListItems(lC)
.SubItems(1) = cName.Text
.SubItems(2) = uType.Text
.SubItems(3) = tAmount.Text
.SubItems(4) = aUsed.Text
.SubItems(5) = idType.Text
.SubItems(6) = idNum.Text
.SubItems(7) = Mob.Text
.SubItems(8) = pass.Text
.SubItems(9) = cAddress.Text
End With
'udbSave
dEdit
bEnabled "edit", True
bEnabled "add", False
End If
End Sub

Private Sub tPassword_Click()
tPass.Text = ""
End Sub

Private Sub uDisable_Click()
wSend "dusb", CInt(SEL)
End Sub

Private Sub uEnable_Click()
wSend "eusb", CInt(SEL)
End Sub

'*************************************************************************************************************************'
'******************************************** |||  Functions  |||  *****************************************************************'
'*************************************************************************************************************************'

Public Function cCharge()
For i = 1 To SocketCounter
With lvwDB.ListItems.Item(i)
    Select Case .SubItems(11)
    
    Case "3"
    For plisT = 1 To lvwPrice.ListItems.Count
        If Val(.SubItems(4)) < Val(lvwPrice.ListItems.Item(plisT).SubItems(1)) Then
            If Val(.SubItems(4)) > lvwPrice.ListItems.Item(plisT).SubItems(2) Then
                .SubItems(13) = Val(lvwPrice.ListItems.Item(plisT).SubItems(3)) + Val(.SubItems(14))
                            Text2.Text = Round(Val(.SubItems(4)) / (Val(lvwPrice.ListItems.Item(plisT).SubItems(2)) - Val(lvwPrice.ListItems.Item(plisT).SubItems(1))))

            Else
                .SubItems(13) = Round(Val(.SubItems(4)) / (Val(lvwPrice.ListItems.Item(plisT).SubItems(2)) - Val(lvwPrice.ListItems.Item(plisT).SubItems(1)))) * Val(lvwPrice.ListItems.Item(plisT).SubItems(3)) + Val(.SubItems(14))
                Text2.Text = Round(Val(.SubItems(4)) / (Val(lvwPrice.ListItems.Item(plisT).SubItems(2)) - Val(lvwPrice.ListItems.Item(plisT).SubItems(1))))
            End If
        End If
    Next plisT

    Case "1"
    For plisT = 1 To lvwPrice.ListItems.Count
        If Val(.SubItems(12)) > Val(lvwPrice.ListItems.Item(plisT).SubItems(1)) And Val(.SubItems(12)) < Val(lvwPrice.ListItems.Item(plisT).SubItems(2)) = True Then
        .SubItems(13) = Round(Val(.SubItems(12)) / 60) * Val(lvwPrice.ListItems.Item(plisT).SubItems(3)) + Val(.SubItems(14))
        End If
    Next plisT
    
    End Select
    
End With
Next i
End Function

Public Function RateCalc(Requested_Time As Integer) As Integer
 Dim tTot, tTme, i, n As Integer
 lvwPrice.SortKey = 1
 lvwPrice.SortOrder = lvwDescending
 lvwPrice.Sorted = True

n = lvwPrice.ListItems.Count

i = 0
tTot = 0

tTme = Requested_Time
While tTme > 0
    i = i + 1
    If i = 1 Then
        If tTme / Val(lvwPrice.ListItems(i).SubItems(1)) >= 1 Then
            While tTme / Val(lvwPrice.ListItems(i).SubItems(1)) >= 1
                tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
               tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
            Wend
            i = 0
        ElseIf tTme > Val(lvwPrice.ListItems(i + 1).SubItems(1)) And tTme < Val(lvwPrice.ListItems(i).SubItems(1)) Then
            tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
            tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
        End If
    ElseIf i = n Then
       If tTme > 0 And tTme <= Val(lvwPrice.ListItems(n).SubItems(1)) Then
          tTot = tTot + Val(lvwPrice.ListItems(n).SubItems(2))
          tTme = tTme - Val(lvwPrice.ListItems(n).SubItems(1))
        Else
         tTme = -1
        End If
    ElseIf tTme > Val(lvwPrice.ListItems(i + 1).SubItems(1)) And tTme <= Val(lvwPrice.ListItems(i).SubItems(1)) Then
        tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
        tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
        i = 1
    End If

    'Else
    'Tme = -1
    'End If
Wend
RateCalc = tTot
lvwPrice.SortKey = 0
lvwPrice.SortOrder = lvwAscending
lvwPrice.Sorted = True

End Function
Public Function aCalc(uSr As Integer)
 Dim tTot, tTme, i, n As Integer
 lvwPrice.SortKey = 1
 lvwPrice.SortOrder = lvwDescending
 lvwPrice.Sorted = True

n = lvwPrice.ListItems.Count

i = 0
tTot = 0

tTme = Val(lvwDB.ListItems(uSr).SubItems(4))
While tTme > 0
    i = i + 1
    If i = 1 Then
        If tTme / Val(lvwPrice.ListItems(i).SubItems(1)) >= 1 Then
            While tTme / Val(lvwPrice.ListItems(i).SubItems(1)) >= 1
                tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
               tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
            Wend
            i = 0
        ElseIf tTme > Val(lvwPrice.ListItems(i + 1).SubItems(1)) And tTme < Val(lvwPrice.ListItems(i).SubItems(1)) Then
            tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
            tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
        End If
    ElseIf i = n Then
       If tTme > 0 And tTme <= Val(lvwPrice.ListItems(n).SubItems(1)) Then
          tTot = tTot + Val(lvwPrice.ListItems(n).SubItems(2))
          tTme = tTme - Val(lvwPrice.ListItems(n).SubItems(1))
        Else
         tTme = -1
        End If
    ElseIf tTme > Val(lvwPrice.ListItems(i + 1).SubItems(1)) And tTme <= Val(lvwPrice.ListItems(i).SubItems(1)) Then
        tTot = tTot + Val(lvwPrice.ListItems(i).SubItems(2))
        tTme = tTme - Val(lvwPrice.ListItems(i).SubItems(1))
        i = 1
    End If

    'Else
    'Tme = -1
    'End If
Wend
lvwDB.ListItems(uSr).SubItems(13) = tTot - Val(lvwDB.ListItems(uSr).SubItems(16)) + Val(lvwDB.ListItems(uSr).SubItems(14))
lvwPrice.SortKey = 0
lvwPrice.SortOrder = lvwAscending
lvwPrice.Sorted = True

End Function

Private Sub clCheck()
On Error GoTo CerroR

aCalc lvwDB.SelectedItem.Index
With lvwDB.SelectedItem
lblsEl.Caption = .Text
lcName.Caption = .SubItems(2)
lStart.Caption = .SubItems(3)
lTTme.Caption = .SubItems(4) & " Min."
ctLeft.Caption = .SubItems(5)
lcDownloaded.Caption = .SubItems(6)
lTamount.Caption = .SubItems(13) & " " & uCurrency
uDL.Caption = .SubItems(6)
End With



SEL = lvwDB.SelectedItem.Index
Exit Sub

CerroR:
SEL = 0
End Sub

Public Function lClear(cSel As Integer)
With cWinsock.lvwDB.ListItems.Item(cSel)
        .SubItems(11) = 2
        .SubItems(12) = 0
        .SubItems(2) = "  ---  "
        .SubItems(3) = "  ---  "
        .SubItems(4) = "  ---  "
        .SubItems(5) = "  ---  "
        .SubItems(6) = "  ---  "
        .SubItems(7) = "  ---  "
        .SubItems(8) = "  ---  "
        .SubItems(9) = "  ---  "
        .SubItems(10) = "  ---  "
        
End With
       
End Function
Public Function lsub14(cSel As Integer, Value As Integer)
With cWinsock.lvwDB.ListItems.Item(cSel)
        .SubItems(11) = Value
End With
End Function

Private Function mShow(bIndex As Integer)
lvwDB.Visible = False
Dim i As Integer
For i = 0 To mfrm.Count - 1
mfrm(i).Visible = False
mfrm(i).Top = mfrm(1).Top
mfrm(i).Left = mfrm(1).Left
Next i
mfrm(bIndex).Visible = True
FPrinters.Hide
End Function

Public Function FormWithinForm(Parent As Object, Child As Object)  ' Makes a form Child

On Error Resume Next
SetParent Child.hwnd, Parent.hwnd
FormWithinForm = (err.Number = 0 And err.LastDllError = 0)

End Function
Public Sub lstSave() 'Save all lists
Dim pData As String
Dim X As Integer
pData = ""

Open DBPath & "x" For Output As #6

Print #6, "***price"
For X = 1 To lvwPrice.ListItems.Count
Print #6, MakeData(lvwPrice, X, 2)
Next X

Print #6, "***token"
For X = 1 To lvwToken.ListItems.Count
Print #6, MakeData(lvwToken, X, lvwToken.ColumnHeaders.Count - 1)
Next X

Print #6, "***admin"
For X = 1 To lvwAdmins.ListItems.Count
Print #6, MakeData(lvwAdmins, X, lvwAdmins.ColumnHeaders.Count - 1)
Next X

Print #6, "***users"
For X = 1 To uDb.ListItems.Count
Print #6, MakeData(uDb, X, uDb.ColumnHeaders.Count - 1)
Next X

Print #6, "***ENDOF"
Close #6
Main.EDecrypt DBPath & "x", DBPath, "m1a9d9s6@MADHAXER$DATABASEMANAGER"
If Dir(DBPath & "x") <> "" Then Kill DBPath & "x"
lstRef
End Sub
Private Function MakeData(ListName As ListView, ListNum As Integer, SubItemCount As Integer) As String
Dim Temp As String
Dim i As Integer


    Temp = "|¤|"
    For i = 1 To SubItemCount
        Temp = Temp & ListName.ListItems(ListNum).SubItems(i) & "|¤|"
    Next i
    
    Temp = Temp & "½¤"
    
MakeData = Temp
End Function
Public Sub SaveAll()
Main.saveSetting
lstSave
LoadAll
End Sub
Public Sub LoadAll()
Main.loadSetting
lstRef
End Sub
Public Sub lstRef() 'Refresh all lists
Dim lstType, cInput As String
Dim uCount As Integer

uCount = 0

If Dir(DBPath) <> "" Then

Main.EDecrypt DBPath, DBPath & "i", "m1a9d9s6@MADHAXER$DATABASEMANAGER"
Open DBPath & "i" For Input As #5

    Do Until EOF(5) = True
        
        Line Input #5, cInput  'Takes Line Input
        
        If Left(cInput, 3) = "***" Then
                lstType = Right(cInput, 5)
                
                If lstType = "ENDOF" Then
                Close #5
                Exit Sub
                End If
                Select Case lstType
                Case "token"
                    lvwToken.ListItems.Clear
                Case "price"
                    lvwPrice.ListItems.Clear
                Case "admin"
                    lvwAdmins.ListItems.Clear
                Case "users"
                    uDb.ListItems.Clear
                End Select
                uCount = 0 'Line Count as 0
        Else
        
            uCount = uCount + 1 'Increase Line Count
        
            Select Case lstType
                Case "token"
                    lstRefresh lvwToken, uCount, cInput, 7
                Case "price"
                    lstRefresh lvwPrice, uCount, cInput, 2
                Case "admin"
                    lstRefresh lvwAdmins, uCount, cInput, 3
                Case "users"
                    lstRefresh uDb, uCount, cInput, 9
            End Select
        End If
    Loop
    
Close #5
 '''''''''''''''''''''''what next ? file not deleting ? why ?
Kill DBPath & "i"
End If
If Dir(DBPath & "i") <> "" Then Kill DBPath & "i"
'Shell "del " & DBPath & "i >>e:\a.txt", vbNormalFocus
End Sub


Private Function lstRefresh(ListName As ListView, Unit_Count As Integer, cInput As String, Total_SubItems As Integer)
Dim MsplT() As String
Dim uSP() As String
Dim i As Integer
MsplT = Split(cInput, "½¤")
uSP = Split(MsplT(0), "|¤|")
ListName.ListItems.Add , , Unit_Count
For i = 1 To Total_SubItems
ListName.ListItems.Item(Unit_Count).SubItems(i) = uSP(i)
Next i
End Function

Public Function ReNumList(ListName As ListView)
Dim i As Integer
For i = 1 To ListName.ListItems.Count
ListName.ListItems(i).Text = i
Next i
lstSave
End Function
Private Function eEdit()
sNo.Enabled = True
cName.Enabled = True
uType.Enabled = True
tAmount.Enabled = True
aUsed.Enabled = True
idType.Enabled = True
idNum.Enabled = True
Mob.Enabled = True
pass.Enabled = True
cAddress.Enabled = True
cmdSave.Enabled = True
End Function

Private Function dEdit()
sNo.Enabled = False
cName.Enabled = False
uType.Enabled = False
tAmount.Enabled = False
aUsed.Enabled = False
idType.Enabled = False
idNum.Enabled = False
Mob.Enabled = False
pass.Enabled = False
cAddress.Enabled = False
cmdSave.Enabled = False
End Function

Private Function bEnabled(num As String, all As Boolean)
If all = True Then
cmddOK.Enabled = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
cmdSave.Enabled = False
tsAmount.Enabled = False
End If
Select Case num
Case "add"
cmdAdd.Enabled = True
Frame4.Visible = False

Case "ok"
cmddOK.Enabled = True
tsAmount.Enabled = True
Frame4.Visible = True
cmddOK.Visible = True
cmdSave.Visible = False

Case "edit"
cmdEdit.Enabled = True
Frame4.Visible = False

Case "save"
cmddOK.Visible = False
cmdSave.Visible = True
cmdSave.Enabled = True
tsAmount.Enabled = True
Frame4.Visible = True

End Select
End Function

Private Function reNum()
Dim i As Integer
For i = 1 To uDb.ListItems.Count
uDb.ListItems(i).Text = i
Next i
SaveAll
End Function
Private Function ChPassToken(cUser As String, cPassword As String) As Integer
Dim i As Integer
ChPassToken = 0
For i = 1 To lvwToken.ListItems.Count
If (LCase(lvwToken.ListItems(i).SubItems(1)) = LCase(cUser)) And (lvwToken.ListItems(i).SubItems(4) = cPassword) Then
ChPassToken = i
i = lvwToken.ListItems.Count
End If

Next i
End Function

Private Function ChPass(cUser As String, cPassword As String) As Integer
Dim i As Integer
ChPass = 0
For i = 1 To uDb.ListItems.Count
If (LCase(uDb.ListItems(i).SubItems(1)) = LCase(cUser)) And (uDb.ListItems(i).SubItems(8) = cPassword) Then
ChPass = i
i = uDb.ListItems.Count
End If
Next i
End Function


Private Function checkCond() As Boolean

checkCond = False
If checkPass Then checkCond = True
End Function

Private Function checkPass() As Boolean
If rePass.Text = pass.Text Then checkPass = True Else MsgBox "Incorrect confirmation password !", vbInformation + vbOKOnly, "Re-Enter Password"
End Function

Public Function sResize()
On Error Resume Next
Dim itmX As ListItem
Dim lwidth As Integer
'*****************************************************'
'====================================================='
'*****************************************************'
lwidth = 9000

Frame1(1).Left = 100
Frame2.Left = cWinsock.Width - (Frame2.Width + 300)
Frame2.Top = Frame1(1).Top

lvwDB.Left = Frame1(1).Width + 100 + Frame1(1).Left
lvwDB.Top = Frame1(1).Top + 100
lvwDB.Width = Frame2.Left - (lvwDB.Left + 100)

cmbChooseView.Left = lvwDB.Left + lvwDB.Width - cmbChooseView.Width
Frm(0).Width = lvwDB.Width
Frm(0).Left = lvwDB.Left


TabStrip1.Left = lvwDB.Left
TabStrip1.Width = lvwDB.Width
Dim i As Integer
For i = 0 To Frm.Count
Frm(i).Top = TabStrip1.Top + TabStrip1.Height - 50
Frm(i).Width = lvwDB.Width
Frm(i).Height = Frm(0).Height
Frm(i).Left = lvwDB.Left

Next i

Frame3.Left = Me.Width / 2 - Frame3.Width / 2
 
lvwToken.Width = lvwDB.Width - lvwToken.Left * 3 - Frame5.Width
Frame5.Left = lvwToken.Left * 2 + lvwToken.Width
For i = 0 To mfrm.Count
mfrm(i).Left = lvwDB.Left
mfrm(i).Width = lvwDB.Width
mfrm(i).Height = lvwDB.Height
mfrm(i).Top = lvwDB.Top - 100
Next i
 'mFrm(0).Left = Frame1(1).Width + 100 + Frame1(1).Left
 'mFrm(0).Top = Frame1(1).Top + 100
 'mFrm(0).Width = Frame2.Left - (mFrm(0).Left + 100)
 'mFrm(0).Height = lvwDB.Height
'*****************************************************'

lvwPrinters.Width = lvwDB.Width - 200
lvwPrinters.Height = Shape18.Height
lvwPrinters.Left = 100
lvwPrinters.Top = 200

uDb.Width = mfrm(3).Width - 480
Frame4.Top = uDb.Top - 200
Frame4.Width = uDb.Width
Frame4.Height = mfrm(3).Height - 250
Frame4.Left = uDb.Left

lvwTime.Width = mfrm(4).Width - lvwTime.Left * 3 - Frame6.Width
Frame6.Left = lvwTime.Width + lvwTime.Left * 2
Frame7.Left = lvwTime.Width + lvwTime.Left * 2

End Function

Public Function cLog(data As String)
Open App.Path & "\log.mCafe" For Append As #1
Print #1, data
Close #1
End Function
Public Function GETLIST(SOCK_Index As Integer) As Integer
Dim i As Integer
For i = 1 To lvwDB.ListItems.Count
If CInt(lvwDB.ListItems(i).SubItems(15)) = SOCK_Index Then GETLIST = i
Next i
End Function
Public Function GETSOCK(List_Index As Integer) As Integer
On Error GoTo err
GETSOCK = 0
GETSOCK = CInt(lvwDB.ListItems(List_Index).SubItems(15))
err:
End Function
Public Function CheckAdmin(User_Name As String, User_Password As String) As String
Dim i As Integer
CheckAdmin = "na"
For i = 1 To lvwAdmins.ListItems.Count
If lvwAdmins.ListItems(i).SubItems(1) = User_Name And lvwAdmins.ListItems(i).SubItems(3) = User_Password Then CheckAdmin = lvwAdmins.ListItems(i).SubItems(2)
Next i
End Function
