VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrinterJobInfo"
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
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As Any, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type JobInfo2
   JobId As Long
   pPrinterName As String
   pMachineName As String
   pUserName As String
   pDocument As String
   pNotifyName As String
   pDatatype As String
   pPrintProcessor As String
   pParameters As String
   pDriverName As String
   pDevMode As Long  'DEVMODE
   pStatus As String
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   Time As Long
   PagesPrinted As Long
End Type

Private Type JOB_INFO_2
   JobId As Long
   pPrinterName As Long
   pMachineName As Long
   pUserName As Long
   pDocument As Long
   pNotifyName As Long
   pDatatype As Long
   pPrintProcessor As Long
   pParameters As Long
   pDriverName As Long
   pDevMode As Long  'DEVMODE
   pStatus As Long
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   Time As Long
   PagesPrinted As Long
End Type

' Status flags
Private Const JOB_STATUS_PAUSED = &H1
Private Const JOB_STATUS_ERROR = &H2
Private Const JOB_STATUS_DELETING = &H4
Private Const JOB_STATUS_SPOOLING = &H8
Private Const JOB_STATUS_PRINTING = &H10
Private Const JOB_STATUS_OFFLINE = &H20
Private Const JOB_STATUS_PAPEROUT = &H40
Private Const JOB_STATUS_PRINTED = &H80
Private Const JOB_STATUS_DELETED = &H100
Private Const JOB_STATUS_BLOCKED_DEVQ = &H200
Private Const JOB_STATUS_USER_INTERVENTION = &H400
Private Const JOB_STATUS_RESTART = &H800

Public Enum JobStatusCodes
   jsPaused = JOB_STATUS_PAUSED
   jsError = JOB_STATUS_ERROR
   jsDeleting = JOB_STATUS_DELETING
   jsSpooling = JOB_STATUS_SPOOLING
   jsPrinting = JOB_STATUS_PRINTING
   jsOffline = JOB_STATUS_OFFLINE
   jsPaperOut = JOB_STATUS_PAPEROUT
   jsPrinted = JOB_STATUS_PRINTED
   jsDeleted = JOB_STATUS_DELETED
   jsBlocked = JOB_STATUS_BLOCKED_DEVQ
   jsUserInt = JOB_STATUS_USER_INTERVENTION
   jsRestart = JOB_STATUS_RESTART
End Enum

' Member variables
Private m_dm As CDevMode
Private m_ji2 As JobInfo2
Private m_ji2Null As JobInfo2
Private m_MaxSize As Long
Private m_MaxPages As Long

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Initialize member objects
   Set m_dm = New CDevMode
End Sub

Private Sub Class_Terminate()
   ' Release member objects
   Set m_dm = Nothing
End Sub

' *********************************************
'  Public Properties (Derived)
' *********************************************
Public Property Get IsPaused() As Boolean
   ' Is this job paused?
   IsPaused = CBool(m_ji2.Status And JOB_STATUS_PAUSED)
End Property

Public Property Get SizeMax() As Long
   ' Specifies the maximum size, in bytes, of the job.
   SizeMax = m_MaxSize
End Property

Friend Property Let SizeMax(ByVal NewVal As Long)
   ' Store max size, as passed from previous instance.
   If NewVal > m_MaxSize Then
      m_MaxSize = NewVal
   End If
End Property

Public Property Get TotalPagesMax() As Long
   ' Specifies the absolute number of pages required
   ' for the job. The TotalPages property will decrement
   ' as each page prints. This holds the max.
   ' This value may be zero if the print job does not
   ' contain page delimiting information.
   TotalPagesMax = m_MaxPages
End Property

Friend Property Let TotalPagesMax(ByVal NewVal As Long)
   ' Store max pages, as passed from previous instance.
   If NewVal > m_MaxPages Then
      m_MaxPages = NewVal
   End If
End Property

' *********************************************
'  Public Properties (Read-Only)
' *********************************************
Public Property Get Datatype() As String
   ' String that specifies the type of data used
   ' to record the print job.
   Datatype = m_ji2.pDatatype
End Property

Public Property Get DevMode() As CDevMode
   ' Pointer to a DEVMODE structure that defines default printer
   ' data such as the paper orientation and the resolution.
   Set DevMode = m_dm
End Property

Public Property Get Document() As String
   ' String that specifies the name of the
   ' print job (for example, "MS-WORD: Review.doc").
   Document = m_ji2.pDocument
End Property

Public Property Get DriverName() As String
   ' String that specifies the name of the printer driver
   ' that should be used to process the print job.
   DriverName = m_ji2.pDriverName
End Property

Public Property Get JobId() As Long
   ' Specifies a job identifier value.
   JobId = m_ji2.JobId
End Property

Public Property Get MachineName() As String
   ' String that specifies the name of the
   ' machine that created the print job.
   MachineName = m_ji2.pMachineName
End Property

Public Property Get NotifyName() As String
   ' String that specifies the name of the user who
   ' should be notified when the job has been printed
   ' or when an error occurs while printing the job.
   NotifyName = m_ji2.pNotifyName
End Property

Public Property Get PagesPrinted() As Long
   ' Specifies the number of pages that have printed.
   ' This value may be zero if the print job does not
   ' contain page delimiting information.
   PagesPrinted = m_ji2.PagesPrinted
End Property

Public Property Get Parameters() As String
   ' String that specifies print-processor parameters.
   Parameters = m_ji2.pParameters
End Property

Public Property Get Position() As Long
   ' Specifies the job's position in the print queue.
   Position = m_ji2.Position
End Property

Public Property Get PrinterName() As String
   ' String that specifies the name of the
   ' printer for which the job is spooled.
   PrinterName = m_ji2.pPrinterName
End Property

Public Property Get PrintProcessor() As String
   ' String that specifies the name of the print
   ' processor that should be used to print the job.
   PrintProcessor = m_ji2.pPrintProcessor
End Property

Public Property Get Priority() As Long
   ' Specifies the job priority. This member can be one
   ' of the *_PRIORITY values or in the range between
   ' 1 through 99 (MIN_PRIORITY through MAX_PRIORITY).
   Priority = m_ji2.Priority
End Property

Public Property Get Size() As Long
   ' Specifies the size, in bytes, of the job.
   Size = m_ji2.Size
End Property

Public Property Get Submitted() As Date
   ' A localized value that specifies the time when
   ' the job was submitted.
   Submitted = CDate(SystemTimeToDouble(m_ji2.Submitted, True))
End Property

Public Property Get StartTime() As Long
   ' Specifies the earliest time that the job
   ' can be printed. This value is expressed as
   ' minutes elapsed since 12:00 AM GMT.
   StartTime = m_ji2.StartTime
End Property

Public Property Get Status() As Long
   ' Specifies the job status. This member can be
   ' one or more of the JOB_STATUS_* values.
   Status = m_ji2.Status
End Property

Public Property Get StatusText() As String
   Dim sRet As String
   Dim sStatus As String
   ' String that specifies the status of the print job.
   ' This member should be checked prior to Status and,
   ' if pStatus is NULL, the status is defined by the
   ' contents of the Status member.
   If Len(m_ji2.pStatus) Then
      sRet = m_ji2.pStatus
   Else
   
      If m_ji2.Status And JOB_STATUS_PAUSED Then
         sRet = "Paused"
      End If
      If m_ji2.Status And JOB_STATUS_ERROR Then
         sRet = "Error"
      End If
      If m_ji2.Status And JOB_STATUS_USER_INTERVENTION Then
         sRet = "User Intervention Required"
      End If
      
      If m_ji2.Status And JOB_STATUS_OFFLINE Then
         sStatus = "Offline"
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_PAPEROUT Then
         sStatus = "Paper Out"
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_DELETING Then
         sStatus = "Deleting..."
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_SPOOLING Then
         sStatus = "Spooling..."
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_PRINTING Then
         sStatus = "Printing..."
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_RESTART Then
         sStatus = "Restarting..."
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_PRINTED Then
         sStatus = "Printed"
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_DELETED Then
         sStatus = "Deleted"
         GoSub AppendStatus
      End If
      If m_ji2.Status And JOB_STATUS_BLOCKED_DEVQ Then
         sStatus = "Cannot Print"
         GoSub AppendStatus
      End If
   End If
   
   ' Return proper string
   StatusText = sRet
Exit Property
AppendStatus:
   ' Use obsolete structure to reduce code repetition
   If Len(sRet) Then
      sRet = sRet & " - " & sStatus
   Else
      sRet = sStatus
   End If
Return
End Property

Public Property Get Time() As Long
   ' Specifies the total time, in seconds, that has
   ' elapsed since the job began printing.
   Time = m_ji2.Time
End Property

Public Property Get TotalPages() As Long
   ' Specifies the number of pages required for the job.
   ' This value may be zero if the print job does not
   ' contain page delimiting information.
   TotalPages = m_ji2.TotalPages
End Property

Public Property Get UntilTime() As Long
   ' Specifies the latest time that the job
   ' can be printed. This value is expressed as
   ' minutes elapsed since 12:00 AM GMT.
   UntilTime = m_ji2.UntilTime
End Property

Public Property Get UserName() As String
   ' String that specifies the name of the
   ' user who owns the print job.
   UserName = m_ji2.pUserName
End Property

' *********************************************
'  Friend Methods
' *********************************************
Friend Sub Initialize(ByVal lpJobInfo2 As Long, ByVal nIndex As Long)
   ' This routine expects to recieve a pointer to an
   ' array of JOB_INFO_2 structures, such as those
   ' returned from an EnumJobs call, and a zero-based
   ' offset into that array indicating which element
   ' to extract for the basis of this class.
   Dim ji2 As JOB_INFO_2
   Dim lpStruct As Long
   
   ' Zero out cached values
   m_ji2 = m_ji2Null
   Set m_dm = New CDevMode
   
   ' Fill local structure with data/pointers.
   lpStruct = lpJobInfo2 + (nIndex * Len(ji2))
   Call CopyMemory(ji2, ByVal lpStruct, Len(ji2))
   
   ' Transfer data from array element structure
   ' to class-cached structure.
   m_ji2.JobId = ji2.JobId
   m_ji2.pPrinterName = PointerToStringA(ji2.pPrinterName)
   m_ji2.pMachineName = PointerToStringA(ji2.pMachineName)
   m_ji2.pUserName = PointerToStringA(ji2.pUserName)
   m_ji2.pDocument = PointerToStringA(ji2.pDocument)
   m_ji2.pNotifyName = PointerToStringA(ji2.pNotifyName)
   m_ji2.pDatatype = PointerToStringA(ji2.pDatatype)
   m_ji2.pPrintProcessor = PointerToStringA(ji2.pPrintProcessor)
   m_ji2.pParameters = PointerToStringA(ji2.pParameters)
   m_ji2.pDriverName = PointerToStringA(ji2.pDriverName)
   m_ji2.pDevMode = ji2.pDevMode
   m_ji2.pStatus = PointerToStringA(ji2.pStatus)
   m_ji2.pSecurityDescriptor = ji2.pSecurityDescriptor
   m_ji2.Status = ji2.Status
   m_ji2.Priority = ji2.Priority
   m_ji2.Position = ji2.Position
   m_ji2.StartTime = ji2.StartTime
   m_ji2.UntilTime = ji2.UntilTime
   m_ji2.TotalPages = ji2.TotalPages
   m_ji2.Size = ji2.Size
   m_ji2.Submitted = ji2.Submitted
   m_ji2.Time = ji2.Time
   m_ji2.PagesPrinted = ji2.PagesPrinted
   
   ' Fill DEVMODE substructure
   m_dm.Initialize m_ji2.pDevMode
   
   ' Assume this is as big as job has gotten, so far.
   ' Size decreases as job is sent to printer.
   If m_ji2.Size > m_MaxSize Then m_MaxSize = m_ji2.Size
   If m_ji2.TotalPages > m_MaxPages Then m_MaxPages = m_ji2.TotalPages
End Sub

' *********************************************
'  Private Methods
' *********************************************
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

Private Function SystemTimeToDouble(stx As SYSTEMTIME, Optional ByVal Localize As Boolean = False)
   Dim st As SYSTEMTIME
   Dim ft As FILETIME
   
   ' Windows NT/2000 *ONLY* Shortcut:
   ' Using NULL for lpTimeZoneInformation converts SYSTEMTIME
   ' structure from UTC to currently active time zone with call
   ' to SystemTimeToTzSpecificLocalTime
   ' If Localize Then
      ' Call SystemTimeToTzSpecificLocalTime(ByVal 0&, stx, st)
      
   If Localize Then
      ' Convert to FILETIME, localize, then convert
      ' back to SYSTEMTIME.
      Call SystemTimeToFileTime(stx, ft)
      Call FileTimeToLocalFileTime(ft, ft)
      Call FileTimeToSystemTime(ft, st)
   Else
      ' Structures can't be passed byval; make copy.  :-(
      st = stx
   End If
   
   ' Convert to VB-style date (double).
   SystemTimeToDouble = DateSerial(st.wYear, st.wMonth, st.wDay) + _
                        TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function
