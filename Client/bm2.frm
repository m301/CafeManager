VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form bm 
   BorderStyle     =   0  'None
   Caption         =   "::: MAD Cafe Manager Client :::"
   ClientHeight    =   1005
   ClientLeft      =   -2625
   ClientTop       =   1275
   ClientWidth     =   1065
   LinkTopic       =   "Form2"
   ScaleHeight     =   1005
   ScaleWidth      =   1065
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "bm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PictureBoxLeft      As Long = 0
Private Const PictureBoxTop       As Long = 0
Private Const PictureBoxRight     As Long = 0
Private Const PictureBoxBottom    As Long = 240   '240 because form has a menu

'Mouse button for grab and drag
Private Const ButtonDrag          As Integer = 1  'Left Mouse
Private PaintLeft           As Long
Private PaintTop            As Long

Private Const TwipsPerPixel       As Long = 15 'Is this ever not true?

Private m_Image                   As New cImage
Private a_Image     As cImage
Private m_Jpeg      As cJpeg
Private m_FileName  As String
Public iMWidth As Integer
Private iFileNum As Integer, lPacketSize As Long
Private Type RECT
left As Long
top As Long
Right As Long
Bottom As Long
End Type
Private Type PICTDESC
cbSize As Long
pictType As Long
hIcon As Long
hPal As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
(lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
IPic As IPicture) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As _
Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, _
ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal lScreenDC As Long, ByVal xSrc As Long, _
ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
lpRect As RECT) As Long
' Capture the contents of a window or the entire screen
Function GetScreenSnapshot(Optional ByVal hwnd As Long) As IPictureDisp
Dim targetDC As Long
Dim hDC As Long
Dim tempPict As Long
Dim oldPict As Long
Dim wndWidth As Long
Dim wndHeight As Long
Dim Pic As PICTDESC
Dim rcWindow As RECT
Dim GUID(3) As Long
' provide the right handle for the desktop window
If hwnd = 0 Then hwnd = GetDesktopWindow
' get window's size
GetWindowRect hwnd, rcWindow
wndWidth = rcWindow.Right - rcWindow.left
wndHeight = rcWindow.Bottom - rcWindow.top
' get window's device context
targetDC = GetWindowDC(hwnd)
' create a compatible DC
hDC = CreateCompatibleDC(targetDC)
' create a memory bitmap in the DC just created
' the has the size of the window we're capturing
tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
oldPict = SelectObject(hDC, tempPict)
' copy the screen image into the DC
BitBlt hDC, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy
' set the old DC image and release the DC
tempPict = SelectObject(hDC, oldPict)
DeleteDC hDC
ReleaseDC GetDesktopWindow, targetDC
' fill the ScreenPic structure
With Pic
.cbSize = Len(Pic)
.pictType = 1 ' means picture
.hIcon = tempPict
.hPal = 0 ' (you can omit this of course)
End With
' convert the image to a IpictureDisp object
' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
' we use an array of Long to initialize it faster
GUID(0) = &H7BF80980
GUID(1) = &H101ABF32
GUID(2) = &HAA00BB8B
GUID(3) = &HAB0C3000
' create the picture,
' return an object reference right into the function result
OleCreatePictureIndirect Pic, GUID(0), True, GetScreenSnapshot
End Function
Private Sub Form_Load()
On Error GoTo eRr
    Winsock1.Close
    Winsock1.LocalPort = 1003
    Winsock1.Listen
    Me.Caption = "Listening: Port 1003"
    Exit Sub
eRr:
    MsgBox "Socket Error!" & vbNewLine & _
            eRr.Description
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub

Private Sub winsock1_Close()
    If Winsock1.State = sckClosing Then
        Winsock1.Close
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.Accept requestID
    rRefresh
    SendFile App.Path & "\tmp.jpg"

End Sub

Public Sub SendFile(FilePath As String, Optional ByVal PacketSize As Long = 1024)
    On Error GoTo errX
    Dim Buffer() As Byte
    
    lPacketSize = PacketSize ' save the PacketSize for the timer
    Timer1.Enabled = False ' make suze timer is not enabled
    
    iFileNum = FreeFile ' get free file number
    Open FilePath For Binary Access Read As iFileNum ' open file
    
    ' if file size is smaller than PacketSize, then send the whole file, but not more
    ReDim Buffer(lngMIN(LOF(iFileNum), PacketSize) - 1)
    Get iFileNum, , Buffer ' read data
    Winsock1.SendData Buffer  ' send data
    Exit Sub
errX:
End Sub

Public Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
    If L1 < L2 Then
        lngMIN = L1
    Else
        lngMIN = L2
    End If
End Function

Private Sub Winsock1_SendComplete()
    Timer1.Enabled = False
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error GoTo eRr
    Dim Buffer() As Byte, BuffSize As Long
    Timer1.Enabled = False
    If iFileNum <= 0 Then Exit Sub
    If Loc(iFileNum) >= LOF(iFileNum) Then ' FILE COMPLETE
        Close iFileNum ' close file
        iFileNum = 0 ' set file number to 0, timer will exit if another timer event
        BuffSize = 0
        Winsock1.Close
        Winsock1.LocalPort = 1003
        Winsock1.Listen
        Me.Caption = "Listening: Port 1003"
        Exit Sub
    End If
    'if the remaining size in the file is smaller then PacketSize, the read only whatever is left
    BuffSize = lngMIN(LOF(iFileNum) - Loc(iFileNum), lPacketSize)
    ReDim Buffer(BuffSize - 1) ' resize buffer
    Get iFileNum, , Buffer ' read data
     Winsock1.SendData Buffer  ' send data
    ' Show progress
    Me.Caption = "Sending: " & Format(Loc(iFileNum) / CDbl(LOF(iFileNum)) * 100#, "#0.00") & "% Done"
    ' timer event will be called again when last packet is sent, close the file then
    Exit Sub
eRr:
    Winsock1.Close
End Sub


Public Function rRefresh()
'On Error Resume Next
Dim std As StdPicture
Dim MyPic As StdPicture
Dim FileName As String


FileName = App.Path & "\mad_cafe.bmp"
SavePicture GetScreenSnapshot, FileName
Set MyPic = LoadPicture(FileName)
Set m_Image = New cImage
 m_Image.CopyStdPicture MyPic
Set MyPic = Nothing
SaveImage m_Image, App.Path & "\tmp.jpg"
i_Save
Kill FileName

End Function


Public Function i_Save()
   Set m_Jpeg = New cJpeg
    'cboSubSample.ListIndex = 3

     m_Jpeg.Quality = iMWidth

       'Sample the cImage by hDC
        m_Jpeg.SampleHDC a_Image.hDC, a_Image.Width, a_Image.Height

       'Delete file if it exists
        RidFile m_FileName

       'Save the JPG file
        m_Jpeg.SaveFile m_FileName
    
    Set a_Image = Nothing
    Set m_Jpeg = Nothing
    
    
End Function


Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set a_Image = TheImage 'Call this before the form loads to initialize it
    m_FileName = FileName
End Sub






