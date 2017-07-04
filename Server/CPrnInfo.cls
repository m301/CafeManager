VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrinterInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2001 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declares
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long

Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetVersion Lib "kernel32" Alias "GetVersionA" (lpVersionInformation As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

' Some calls need to know OS
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

' Platform ID constants
Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

'  SendMessageTimeout values
Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2

' Used with SendMessageTimeout to tell all apps of changes
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_SETTINGCHANGE = &H1A

' Need defaults to OpenPrinter in some cases
Private Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   pDesiredAccess As Long
End Type

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' Used to retrieve last API error text.
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

' The data area passed to a system call is too small.
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

' Used to indicate what to enumerate
Private Const PRINTER_ENUM_DEFAULT         As Long = &H1
Private Const PRINTER_ENUM_LOCAL           As Long = &H2
Private Const PRINTER_ENUM_CONNECTIONS     As Long = &H4
Private Const PRINTER_ENUM_FAVORITE        As Long = &H4
Private Const PRINTER_ENUM_NAME            As Long = &H8
Private Const PRINTER_ENUM_REMOTE          As Long = &H10
Private Const PRINTER_ENUM_SHARED          As Long = &H20
Private Const PRINTER_ENUM_NETWORK         As Long = &H40

' Printer control codes
Private Const PRINTER_CONTROL_PAUSE        As Long = 1
Private Const PRINTER_CONTROL_RESUME       As Long = 2
Private Const PRINTER_CONTROL_PURGE        As Long = 3
Private Const PRINTER_CONTROL_SET_STATUS   As Long = 4

Private Enum PrinterControlCodes
   pcPause = PRINTER_CONTROL_PAUSE
   pcResume = PRINTER_CONTROL_RESUME
   pcPurge = PRINTER_CONTROL_PURGE
End Enum

' Printer status flags used with PRINTER_INFORMATION_2
Private Const PRINTER_STATUS_READY              As Long = &H0
Private Const PRINTER_STATUS_PAUSED             As Long = &H1
Private Const PRINTER_STATUS_ERROR              As Long = &H2
Private Const PRINTER_STATUS_PENDING_DELETION   As Long = &H4
Private Const PRINTER_STATUS_PAPER_JAM          As Long = &H8
Private Const PRINTER_STATUS_PAPER_OUT          As Long = &H10
Private Const PRINTER_STATUS_MANUAL_FEED        As Long = &H20
Private Const PRINTER_STATUS_PAPER_PROBLEM      As Long = &H40
Private Const PRINTER_STATUS_OFFLINE            As Long = &H80
Private Const PRINTER_STATUS_IO_ACTIVE          As Long = &H100
Private Const PRINTER_STATUS_BUSY               As Long = &H200
Private Const PRINTER_STATUS_PRINTING           As Long = &H400
Private Const PRINTER_STATUS_OUTPUT_BIN_FULL    As Long = &H800
Private Const PRINTER_STATUS_NOT_AVAILABLE      As Long = &H1000
Private Const PRINTER_STATUS_WAITING            As Long = &H2000
Private Const PRINTER_STATUS_PROCESSING         As Long = &H4000
Private Const PRINTER_STATUS_INITIALIZING       As Long = &H8000
Private Const PRINTER_STATUS_WARMING_UP         As Long = &H10000
Private Const PRINTER_STATUS_TONER_LOW          As Long = &H20000
Private Const PRINTER_STATUS_NO_TONER           As Long = &H40000
Private Const PRINTER_STATUS_PAGE_PUNT          As Long = &H80000
Private Const PRINTER_STATUS_USER_INTERVENTION  As Long = &H100000
Private Const PRINTER_STATUS_OUT_OF_MEMORY      As Long = &H200000
Private Const PRINTER_STATUS_DOOR_OPEN          As Long = &H400000
Private Const PRINTER_STATUS_SERVER_UNKNOWN     As Long = &H800000
Private Const PRINTER_STATUS_POWER_SAVE         As Long = &H1000000

Public Enum PrinterStatusCodes
   psReady = PRINTER_STATUS_READY
   psPaused = PRINTER_STATUS_PAUSED
   psError = PRINTER_STATUS_ERROR
   psPendingDeletion = PRINTER_STATUS_PENDING_DELETION
   psPaperJam = PRINTER_STATUS_PAPER_JAM
   psPaperOut = PRINTER_STATUS_PAPER_OUT
   psManualFeed = PRINTER_STATUS_MANUAL_FEED
   psPaperProblem = PRINTER_STATUS_PAPER_PROBLEM
   psOffline = PRINTER_STATUS_OFFLINE
   psIoActive = PRINTER_STATUS_IO_ACTIVE
   psBusy = PRINTER_STATUS_BUSY
   psPrinting = PRINTER_STATUS_PRINTING
   psOutputBinFull = PRINTER_STATUS_OUTPUT_BIN_FULL
   psNotAvailable = PRINTER_STATUS_NOT_AVAILABLE
   psWaiting = PRINTER_STATUS_WAITING
   psProcessing = PRINTER_STATUS_PROCESSING
   psInitializing = PRINTER_STATUS_INITIALIZING
   psWarmingUp = PRINTER_STATUS_WARMING_UP
   psTonerLow = PRINTER_STATUS_TONER_LOW
   psNoToner = PRINTER_STATUS_NO_TONER
   psPagePrint = PRINTER_STATUS_PAGE_PUNT
   psUserIntervention = PRINTER_STATUS_USER_INTERVENTION
   psOutOfMemory = PRINTER_STATUS_OUT_OF_MEMORY
   psDoorOpen = PRINTER_STATUS_DOOR_OPEN
   psServerUnknown = PRINTER_STATUS_SERVER_UNKNOWN
End Enum
   
' Printer attribute flags used with PRINTER_INFORMATION_2
Private Const PRINTER_ATTRIBUTE_QUEUED            As Long = &H1
Private Const PRINTER_ATTRIBUTE_DIRECT            As Long = &H2
Private Const PRINTER_ATTRIBUTE_DEFAULT           As Long = &H4
Private Const PRINTER_ATTRIBUTE_SHARED            As Long = &H8
Private Const PRINTER_ATTRIBUTE_NETWORK           As Long = &H10
Private Const PRINTER_ATTRIBUTE_HIDDEN            As Long = &H20
Private Const PRINTER_ATTRIBUTE_LOCAL             As Long = &H40
Private Const PRINTER_ATTRIBUTE_ENABLE_DEVQ       As Long = &H80
Private Const PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS   As Long = &H100
Private Const PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST As Long = &H200
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE      As Long = &H400
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI       As Long = &H800
Private Const PRINTER_ATTRIBUTE_RAW_ONLY          As Long = &H1000
Private Const PRINTER_ATTRIBUTE_PUBLISHED         As Long = &H2000

Public Enum PrinterAttributeCodes
   paQueued = PRINTER_ATTRIBUTE_QUEUED
   paDirect = PRINTER_ATTRIBUTE_DIRECT
   paDefault = PRINTER_ATTRIBUTE_DEFAULT
   paShared = PRINTER_ATTRIBUTE_SHARED
   paNetwork = PRINTER_ATTRIBUTE_NETWORK
   paHidden = PRINTER_ATTRIBUTE_HIDDEN
   paLocal = PRINTER_ATTRIBUTE_LOCAL
   paEnableDevQ
End Enum

' VB5-friendly structure used to cache the values in this class.
Private Type PrinterInfo2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long 'DEVMODE
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

' Structure used to obtain the data from Windows.
Private Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long 'DEVMODE
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

' Member variables
Private m_pi2 As PrinterInfo2
Private m_pi2Null As PrinterInfo2
Private m_dm As CDevMode
Private m_jobs As CPrinterJobs
Private m_DevName As String
Private m_DispName As String
Private m_GetPrinterError As Long

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Initialize member objects
   Set m_dm = New CDevMode
   Set m_jobs = New CPrinterJobs
End Sub

Private Sub Class_Terminate()
   ' Release member objects
   Set m_dm = Nothing
   Set m_jobs = Nothing
End Sub

' *********************************************
'  Public Properties (Read/Write)
' *********************************************
Public Property Get DeviceName() As String
   ' This is the name returned by a VB Printer
   ' object's .DeviceName property
   DeviceName = m_DevName
End Property

Public Property Let DeviceName(ByVal NewVal As String)
   m_DevName = NewVal
   Call Refresh
   'm_jobs.DeviceName = m_DevName
   'm_jobs.Refresh
End Property

Public Property Get IsDefault() As Boolean
   Dim DefPrn As String
   ' PRINTER_ATTRIBUTE_DEFAULT only works to *set* default.
   DefPrn = DefaultPrinterName
   IsDefault = (UCase$(DefPrn) = UCase$(m_DevName))
End Property

Public Property Let IsDefault(ByVal NewVal As Boolean)
   ' Really makes no sense to react to a new value
   ' of False, as something must be default.
   If NewVal = True Then
      Call DefaultPrinterSet
   End If
End Property

' *********************************************
'  Public Properties (Derived, Read-Only)
' *********************************************
Public Property Get CanAdminister() As Boolean
   Dim hPrn As Long
   Dim pd As PRINTER_DEFAULTS
   ' Try to open handle to printer, requesting
   ' administrative priviledges.
   pd.pDesiredAccess = PRINTER_ACCESS_ADMINISTER
   Call OpenPrinter(m_DevName, hPrn, pd)
   CanAdminister = (hPrn <> 0)
   Call ClosePrinter(hPrn)
End Property

Public Property Get DisplayName() As String
   DisplayName = m_DispName
End Property

Public Property Get GetPrinterError() As Long
   ' If the call to GetPrinter doesn't succeed,
   ' client needs to know LastDllError.
   GetPrinterError = m_GetPrinterError
End Property

Public Property Get IsOffline() As Boolean
   ' Is this printer offline?
   IsOffline = CBool(m_pi2.Status And PRINTER_STATUS_OFFLINE)
End Property

Public Property Get IsPaused() As Boolean
   ' Is this printer paused?
   IsPaused = CBool(m_pi2.Status And PRINTER_STATUS_PAUSED)
End Property

Public Property Get IsShared() As Boolean
   ' Is this printer shared?
   IsShared = CBool(m_pi2.Attributes And PRINTER_ATTRIBUTE_SHARED)
End Property

Public Property Get IsToFile() As Boolean
   If InStr(m_pi2.pPortName, "FILE:") Then
      IsToFile = True
   End If
End Property

Public Property Get IsLocal() As Boolean
   ' Is this a local printer?
   IsLocal = CBool(m_pi2.Attributes And PRINTER_ATTRIBUTE_LOCAL)
End Property

Public Property Get IsNetwork() As Boolean
   ' Is this a network printer?
   IsNetwork = CBool(m_pi2.Attributes And PRINTER_ATTRIBUTE_NETWORK)
End Property

Public Property Get StatusText() As String
   If m_GetPrinterError Then
      StatusText = ApiErrorText(m_GetPrinterError)
   Else
      Select Case m_pi2.Status
         Case PRINTER_STATUS_READY
            StatusText = "Ready"
         Case PRINTER_STATUS_PAUSED
            StatusText = "Paused"
         Case PRINTER_STATUS_ERROR
            StatusText = "Error"
         Case PRINTER_STATUS_PENDING_DELETION
            StatusText = "Deleting..."
         Case PRINTER_STATUS_PAPER_JAM
            StatusText = "Paper Jam"
         Case PRINTER_STATUS_PAPER_OUT
            StatusText = "Paper Out"
         Case PRINTER_STATUS_MANUAL_FEED
            StatusText = "Manual Feed Required"
         Case PRINTER_STATUS_PAPER_PROBLEM
            StatusText = "Paper Problem"
         Case PRINTER_STATUS_OFFLINE
            StatusText = "Offline"
         Case PRINTER_STATUS_IO_ACTIVE
            StatusText = "Downloading Job"
         Case PRINTER_STATUS_BUSY
            StatusText = "Busy"
         Case PRINTER_STATUS_PRINTING
            StatusText = "Printing"
         Case PRINTER_STATUS_OUTPUT_BIN_FULL
            StatusText = "Output Bill Full"
         Case PRINTER_STATUS_NOT_AVAILABLE
            StatusText = "Not Available"
         Case PRINTER_STATUS_WAITING
            StatusText = "Waiting"
         Case PRINTER_STATUS_PROCESSING
            StatusText = "Processing Job"
         Case PRINTER_STATUS_INITIALIZING
            StatusText = "Initializing"
         Case PRINTER_STATUS_WARMING_UP
            StatusText = "Warming Up"
         Case PRINTER_STATUS_TONER_LOW
            StatusText = "Toner Low"
         Case PRINTER_STATUS_NO_TONER
            StatusText = "Toner Out"
         Case PRINTER_STATUS_PAGE_PUNT
            StatusText = "Page too Complex"
         Case PRINTER_STATUS_USER_INTERVENTION
            StatusText = "User Intervention Required"
         Case PRINTER_STATUS_OUT_OF_MEMORY
            StatusText = "Out of Memory"
         Case PRINTER_STATUS_DOOR_OPEN
            StatusText = "Door Open"
         Case PRINTER_STATUS_SERVER_UNKNOWN
            StatusText = "Unable to connect"
         Case PRINTER_STATUS_POWER_SAVE
            StatusText = "Power Save Mode"
         Case Else
            StatusText = Hex$(m_pi2.Status)
      End Select
   End If
End Property

' *********************************************
'  Public Properties (Read-Only)
' *********************************************
Public Property Get Attributes() As Long
   ' Specifies the printer attributes. This member can be any
   ' reasonable combination of the PRINTER_ATTRIBUTE_* values.
   Attributes = m_pi2.Attributes
End Property

Public Property Get AveragePPM() As Long
   ' Specifies the average number of pages per minute
   ' that have been printed on the printer.
   AveragePPM = m_pi2.AveragePPM
End Property

Public Property Get Comment() As String
   ' A string that provides a brief description of
   ' the printer.
   Comment = m_pi2.pComment
End Property

Public Property Get Datatype() As String
   ' A string that specifies the data type used to record
   ' the print job. You can use the EnumPrintProcessorDatatypes
   ' function to obtain a list of data types supported by a
   ' specific print processor.
   Datatype = m_pi2.pDatatype
End Property

Public Property Get DefaultPriority() As Long
   ' Specifies the default priority value
   ' assigned to each print job.
   DefaultPriority = m_pi2.DefaultPriority
End Property

Public Property Get DevMode() As CDevMode
   ' Pointer to a DEVMODE structure that defines default printer
   ' data such as the paper orientation and the resolution.
   Set DevMode = m_dm
End Property

Public Property Get DriverName() As String
   ' A string that specifies the name of the printer driver.
   DriverName = m_pi2.pDriverName
End Property

Public Property Get Jobs() As CPrinterJobs
   ' Return collection of printer jobs
   Set Jobs = m_jobs
End Property

Public Property Get Location() As String
   ' A string that specifies the physical location of
   ' the printer (for example, "Bldg. 38, Room 1164").
   Location = m_pi2.pLocation
End Property

Public Property Get Parameters() As String
   ' A string that specifies the default print-processor
   ' parameters.
   Parameters = m_pi2.pParameters
End Property

Public Property Get PortName() As String
   ' A string that identifies the port(s) used to transmit data
   ' to the printer. If a printer is connected to more than one
   ' port, the names of each port must be separated by commas
   ' (for example, "LPT1:,LPT2:,LPT3:").
   ' Windows 95: This member can specify only one port because
   ' multiple ports per printer are not supported.
   PortName = m_pi2.pPortName
End Property

Public Property Get PrinterName() As String
   ' A string that specifies the name of the printer.
   PrinterName = m_pi2.pPrinterName
End Property

Public Property Get PrintProcessor() As String
   ' A string that specifies the name of the print processor
   ' used by the printer. You can use the EnumPrintProcessors
   ' function to obtain a list of print processors installed
   ' on a server.
   PrintProcessor = m_pi2.pPrintProcessor
End Property

Public Property Get Priority() As Long
   ' Specifies a priority value that the spooler
   ' uses to route print jobs.
   Priority = m_pi2.Priority
End Property

Public Property Get SeparatorPageFile() As String
   ' A string that specifies the name of the file used to create
   ' the separator page. This page is used to separate print jobs
   ' sent to the printer.
   SeparatorPageFile = m_pi2.pSepFile
End Property

Public Property Get ServerName() As String
   ' A string identifying the server that controls the printer.
   ' If this string is NULL, the printer is controlled locally.
   ServerName = m_pi2.pServerName
End Property

Public Property Get ShareName() As String
   ' A string that identifies the sharepoint for the printer.
   ' (This string is used only if the PRINTER_ATTRIBUTE_SHARED
   ' constant was set for the Attributes member.)
   ShareName = m_pi2.pShareName
End Property

Public Property Get StartTime() As Long
   ' Specifies the earliest time at which the printer
   ' will print a job. This value is expressed as
   ' minutes elapsed since 12:00 AM GMT.
   StartTime = m_pi2.StartTime
End Property

Public Property Get Status() As Long
   ' Specifies the printer status. This member can be any
   ' reasonable combination of the PRINTER_STATUS_* values.
   Status = m_pi2.Status
End Property

Public Property Get UntilTime() As Long
   ' Specifies the latest time at which the printer
   ' will print a job. This value is expressed as
   ' minutes elapsed since 12:00 AM GMT.
   UntilTime = m_pi2.UntilTime
End Property

' *********************************************
'  Public Methods
' *********************************************
Public Function ControlPause() As Boolean
   ' Attempt to pause printing
   ControlPause = SendControl(pcPause)
End Function

Public Function ControlPurge() As Boolean
   ' Attempt to purge all print jobs
   ControlPurge = SendControl(pcPurge)
End Function

Public Function ControlResume() As Boolean
   ' Attempt to resume printing
   ControlResume = SendControl(pcResume)
End Function

Public Sub Refresh()
   Dim pi2 As PRINTER_INFO_2
   Dim hPrn As Long
   Dim Buffer() As Byte
   Dim BytesNeeded As Long
   Dim BytesUsed As Long
   Dim slash As Long
   
   ' Zero out cached values
   m_pi2 = m_pi2Null
   Set m_dm = New CDevMode
   
   ' Get handle to printer.
   Call OpenPrinter(m_DevName, hPrn, ByVal 0&)
   If hPrn Then
      ' Call once to get proper buffer size.
      Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
         ' Size buffer and get printer data.
         ReDim Buffer(0 To BytesNeeded - 1) As Byte
         If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
            ' Fill local structure with data/pointers.
            Call CopyMemory(pi2, Buffer(0), Len(pi2))
            ' Transfer string data to cached structure.
            m_pi2.pServerName = PointerToStringA(pi2.pServerName)
            m_pi2.pPrinterName = PointerToStringA(pi2.pPrinterName)
            m_pi2.pShareName = PointerToStringA(pi2.pShareName)
            m_pi2.pPortName = PointerToStringA(pi2.pPortName)
            m_pi2.pDriverName = PointerToStringA(pi2.pDriverName)
            m_pi2.pComment = PointerToStringA(pi2.pComment)
            m_pi2.pLocation = PointerToStringA(pi2.pLocation)
            m_pi2.pSepFile = PointerToStringA(pi2.pSepFile)
            m_pi2.pPrintProcessor = PointerToStringA(pi2.pPrintProcessor)
            m_pi2.pDatatype = PointerToStringA(pi2.pDatatype)
            m_pi2.pParameters = PointerToStringA(pi2.pParameters)
            ' Copy two sub-structure pointers.
            m_pi2.pDevMode = pi2.pDevMode
            m_pi2.pSecurityDescriptor = pi2.pSecurityDescriptor
            ' Fill DEVMODE substructure
            m_dm.Initialize m_pi2.pDevMode
            ' Sling remaining dword data to cached structure.
            ' Attributes begins at byte 52 in structure, so
            ' just copy last 32 bytes of structure directly.
            Call CopyMemory(m_pi2.Attributes, Buffer(52), 32)
         End If
         m_GetPrinterError = 0  'clear error value
      Else
         m_GetPrinterError = Err.LastDllError
      End If
      Call ClosePrinter(hPrn)
   End If
   
   ' Build display name, based on whether this
   ' is a network printer, defaulting to devname.
   m_DispName = m_DevName
   If CBool(m_pi2.Attributes And PRINTER_ATTRIBUTE_NETWORK) Then
      If Len(m_pi2.pServerName) > 2 Then
         If InStr(m_pi2.pServerName, "\\") = 1 Then
            ' Find 3rd backslash.
            slash = InStr(3, m_pi2.pPrinterName, "\")
            If slash Then
               m_DispName = Mid$(m_pi2.pPrinterName, slash + 1) & _
                            " on " & Mid$(m_pi2.pServerName, 3)
            End If
         End If
      End If
   End If
   
   ' Update jobs collection
   If m_jobs.DeviceName <> m_DevName Then
      m_jobs.DeviceName = m_DevName
   Else
      m_jobs.Refresh
   End If
End Sub

Public Sub ShowPropertiesDialog(Optional ByVal hWndParent As Long = 0)
   Dim hPrn As Long
   Dim pd As PRINTER_DEFAULTS
   ' HOWTO: Open the Printer Properties Dialog
   ' http://support.microsoft.com/support/kb/articles/Q198/8/60.asp
   pd.pDatatype = vbNullString
   ' Try admin access first
   pd.pDesiredAccess = PRINTER_ALL_ACCESS
   If OpenPrinter(m_DevName, hPrn, pd) = 0 Then
      ' Not an admin, try reduced privileges
      pd.pDesiredAccess = STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE
      Call OpenPrinter(m_DevName, hPrn, pd)
   End If
   ' Show dialog, if we have a handle to printer
   If hPrn Then
      Call PrinterProperties(hWndParent, hPrn)
      Call ClosePrinter(hPrn)
   End If
End Sub

' *********************************************
'  Private Methods
' *********************************************
Private Function ApiErrorText(ByVal ErrNum As Long) As String
   Dim msg As String
   Dim nRet As Long

   msg = Space$(1024)
   nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, msg, Len(msg), ByVal 0&)
   If nRet Then
      ApiErrorText = Left$(msg, nRet - 2) ' account for Cr/Lf
   Else
      ApiErrorText = "Error (" & ErrNum & ") not defined."
   End If
End Function

Private Function DefaultPrinterSet() As Boolean
   Dim os As OSVERSIONINFO

   ' Fork based on what OS we're running on...
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   Select Case os.dwPlatformId
      Case VER_PLATFORM_WIN32_WINDOWS '95/98/ME
         Call DefaultPrinterSet9x
      Case VER_PLATFORM_WIN32_NT
         Call DefaultPrinterSetNT(os.dwMajorVersion)
   End Select
   
   ' Return results based on test in IsDefault.
   Me.Refresh
   DefaultPrinterSet = Me.IsDefault
End Function

Private Sub DefaultPrinterSet9x()
   ' HOWTO: Retrieve and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q246/7/72.asp
   Dim hPrn As Long
   Dim BytesNeeded As Long
   Dim Buffer() As Byte
   Dim BytesUsed As Long
   Dim Attributes As Long
   
   ' Get handle to printer.
   Call OpenPrinter(m_DevName, hPrn, ByVal 0&)
   If hPrn Then
      ' Call once to get proper buffer size.
      Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
         ' Size buffer and get printer data.
         ReDim Buffer(0 To BytesNeeded - 1) As Byte
         If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
            ' Set default printer attribute for this printer...
            ' Attributes is the 14th element in structure
            Const AttribOffset As Long = 13 * 4&
            Call CopyMemory(Attributes, Buffer(AttribOffset), 4&)
            Attributes = Attributes Or PRINTER_ATTRIBUTE_DEFAULT
            Call CopyMemory(Buffer(AttribOffset), Attributes, 4&)
            ' Send back updated structure.
            If SetPrinter(hPrn, 2, Buffer(0), 0) Then
               ' Alert all other running applications,
               ' giving each 1/2 second to react.
               Call SettingChangeAlert(500)
            End If
         End If
      End If
      Call ClosePrinter(hPrn)
   End If
End Sub

Private Sub DefaultPrinterSetNT(ByVal MajorVersion As Long)
   ' HOWTO: Retrieve and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q246/7/72.asp
   Dim os As OSVERSIONINFO
   Dim BufSize As Long
   Dim pPrinterName As Long
   Dim Result As String
   Dim comma As Long

   ' Use either SetDefaultPrinter (2k+) or WIN.INI (NT4).
   If MajorVersion >= 5 Then
      ' Almost so easy as to be boring. <g>
      Call SetDefaultPrinter(m_DevName)

   Else ' (NT4 or less)
      ' Create satisfactory buffer.
      BufSize = 1024
      Result = Space$(BufSize)
      ' In NT, the old WIN.INI [PrinterPorts] section is mapped to
      ' HKCU\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts
      ' and we can just use GetProfileString to extract! :-)
      ' Returns: "driver name,port,timeout1,timeout2"
      If GetProfileString("PrinterPorts", ByVal m_DevName, "", Result, BufSize) Then
         ' Find 2nd comma and truncate
         comma = InStr(Result, ",")
         comma = InStr(comma + 1, Result, ",")
         If comma Then
            Result = Left$(Result, comma - 1)
            ' Prepend devname and write to registry.
            Result = m_DevName & "," & Result
            Call WriteProfileString("Windows", "device", Result)
            ' Alert all other running applications,
            ' giving each 1/2 second to react.
            Call SettingChangeAlert(500)
         End If
      End If
   End If
End Sub

Private Function DefaultPrinterName() As String
   ' HOWTO: Retrieve and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q246/7/72.asp
   ' HOWTO: Get and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q135/3/87.asp
   Dim os As OSVERSIONINFO
   Dim Buffer() As Byte
   Dim BufSize As Long
   Dim pPrinterName As Long
   Dim Returned As Long
   Dim Result As String
   
   ' Get OS version info, so we know which way to fork.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   
   If os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '95/98/ME
      ' Determine how big the buffer needs to be
      Call EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 2, ByVal 0&, 0, BufSize, Returned)
      If BufSize > 0 Then
         ' Size buffer accordingly
         ReDim Buffer(0 To BufSize - 1) As Byte
         ' Call again to retrieve needed info
         Call EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 2, Buffer(0), BufSize, BufSize, Returned)
         ' A pointer to the default printer name is
         ' returned at the 5th byte in the buffer.
         Call CopyMemory(pPrinterName, Buffer(4), 4)
         Result = PointerToStringA(pPrinterName)
      End If
       
   ElseIf os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      ' Create satisfactory buffer.
      BufSize = 1024
      Result = Space$(BufSize)
      
      ' Use either GetDefaultPrinter (2k+) or WIN.INI (NT4).
      If os.dwMajorVersion >= 5 Then
         If GetDefaultPrinter(Result, BufSize) Then
            ' Truncate at first NULL
            Result = Left$(Result, InStr(Result, vbNullChar) - 1)
         End If
      Else 'NT4 or less
         ' The old WIN.INI [Windows] section is mapped to
         ' HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows
         ' and we can just use GetProfileString to extract! :-)
         ' Returns: "printer name,driver name,port"
         If GetProfileString("Windows", ByVal "device", "", Result, BufSize) Then
            ' Truncate buffer at end of name.
            Result = Left$(Result, InStr(Result, ",") - 1)
         End If
      End If
   End If
      
   ' Return default printer name.
   DefaultPrinterName = Result
End Function

Private Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Dim nRet As Long
   If lpDWord Then
      CopyMemory nRet, ByVal lpDWord, 4
      PointerToDWord = nRet
   End If
End Function

Private Function SendControl(ByVal ControlCode As PrinterControlCodes) As Boolean
   Dim hPrn As Long
   Dim pd As PRINTER_DEFAULTS
   
   ' Get handle to printer.
   pd.pDesiredAccess = PRINTER_ACCESS_ADMINISTER
   Call OpenPrinter(m_DevName, hPrn, pd)
   If hPrn Then
      ' Attempt to send control code
      SendControl = CBool(SetPrinter(hPrn, 0, ByVal 0&, ControlCode))
      Call ClosePrinter(hPrn)
      Call Me.Refresh
   End If
End Function

Private Sub SettingChangeAlert(Optional ByVal Delay As Long = 500)
   ' Send out alert to notify all other running applications
   ' that a system setting has been updated.
   Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, _
                           0, 0, SMTO_NORMAL, Delay, ByVal 0&)
End Sub

