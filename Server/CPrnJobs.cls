VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrinterJobs"
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
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Any, ByVal Command As Long) As Long
Private Declare Function GetJob Lib "winspool.drv" Alias "GetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'  dwPlatformId defines:
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

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

' The data area passed to a system call is too small.
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

' Job control constants
Private Const JOB_CONTROL_PAUSE = 1
Private Const JOB_CONTROL_RESUME = 2
Private Const JOB_CONTROL_CANCEL = 3
Private Const JOB_CONTROL_RESTART = 4
Private Const JOB_CONTROL_DELETE = 5
Private Const JOB_CONTROL_SENT_TO_PRINTER = 6
Private Const JOB_CONTROL_LAST_PAGE_EJECTED = 7

Private Enum JobControlCodes
   jcPause = JOB_CONTROL_PAUSE
   jcResume = JOB_CONTROL_RESUME
   jcCancel = JOB_CONTROL_CANCEL
   jcRestart = JOB_CONTROL_RESTART
   jcDelete = JOB_CONTROL_DELETE
End Enum

' Member variables
Private m_DevName As String
Private m_GetPrinterError As Long
Private m_jobs As Collection

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Initialize member objects
   Set m_jobs = New Collection
End Sub

Private Sub Class_Terminate()
   ' Release member objects
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
End Property

' *********************************************
'  Public Properties (Derived, Read-Only)
' *********************************************
Public Property Get Count() As Long
   Dim nRet As Long
   ' Return actual number of jobs in queue.
   nRet = JobCount()
   If nRet <> m_jobs.Count Then
      ' Need to update collection!
      Call Refresh
   End If
   Count = nRet
End Property

Public Property Get Item(ByVal JobId As Variant) As CPrinterJobInfo
Attribute Item.VB_UserMemId = 0
   ' Just allow errors to propogate up.
   Set Item = m_jobs(JobId)
End Property

Public Property Get GetPrinterError() As Long
   ' If the call to GetPrinter doesn't succeed,
   ' client needs to know LastDllError.
   GetPrinterError = m_GetPrinterError
End Property

' *********************************************
'  Public Properties (Read-only)
' *********************************************

' *********************************************
'  Public Methods
' *********************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   ' Support enumeration of Item (default).
   Set NewEnum = m_jobs.[_NewEnum]
End Function

Public Function ControlCancel(ByVal JobId As Long) As Boolean
   Dim os As OSVERSIONINFO
   ' NT4 is the dividing line between two different
   ' control codes for this call.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   ' Attempt to cancel passed job.
   If os.dwPlatformId = VER_PLATFORM_WIN32_NT And os.dwMajorVersion >= 4 Then
      ControlCancel = SendControl(JobId, jcDelete)
   Else
      ControlCancel = SendControl(JobId, jcCancel)
   End If
End Function

Public Function ControlRestart(ByVal JobId As Long) As Boolean
   ' Attempt to resume passed job
   ControlRestart = SendControl(JobId, jcRestart)
End Function

Public Function ControlResume(ByVal JobId As Long) As Boolean
   ' Attempt to resume passed job
   ControlResume = SendControl(JobId, jcResume)
End Function

Public Function ControlPause(ByVal JobId As Long) As Boolean
   ' Attempt to pause passed job
   ControlPause = SendControl(JobId, jcPause)
End Function

Public Function PositionMoveDown(ByVal JobId As Long) As Boolean
   ' Attempt to adjust queue
   PositionMoveDown = AdjustJobPosition(JobId, 1)
End Function

Public Function PositionMoveUp(ByVal JobId As Long) As Boolean
   ' Attempt to adjust queue
   PositionMoveUp = AdjustJobPosition(JobId, -1)
End Function

Public Sub Refresh()
   Dim hPrn As Long
   Dim nJobs As Long
   Dim JobIds() As Long
   Dim Buffer() As Byte
   Dim BytesNeeded As Long
   Dim nReturned As Long
   Dim NewJob As CPrinterJobInfo
   Dim OldJob As CPrinterJobInfo
   Dim i As Long

   ' Get handle to printer.
   Call OpenPrinter(m_DevName, hPrn, ByVal 0&)
   If hPrn Then
      ' Retrieve number of jobs.
      nJobs = JobCount(hPrn)
      If nJobs > 0 Then
         ' Setup array to stash job ids so we can remove
         ' dead jobs from collection
         ReDim JobIds(0 To nJobs - 1) As Long
         ' Call once to get proper buffer size.
         Call EnumJobs(hPrn, 0, nJobs, 2, ByVal 0&, 0, BytesNeeded, nReturned)
         If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
            ' Size buffer and get printer data.
            ReDim Buffer(0 To BytesNeeded - 1) As Byte
            If EnumJobs(hPrn, 0, nJobs, 2, Buffer(0), BytesNeeded, BytesNeeded, nReturned) Then
               ' Parse each returned structure into a
               ' class, and add to collection.
               For i = 0 To nReturned - 1
                  ' Update new object using new data.
                  Set NewJob = New CPrinterJobInfo
                  NewJob.Initialize VarPtr(Buffer(0)), i
                  ' Look to see if object exists for this job.
                  Set OldJob = GetJobInfo(NewJob.JobId)
                  If Not OldJob Is Nothing Then
                     ' Transfer previous max job size to new
                     ' object, and remove from collection.
                     NewJob.SizeMax = OldJob.SizeMax
                     NewJob.TotalPagesMax = OldJob.TotalPagesMax
                     m_jobs.Remove Hex$(OldJob.JobId)
                  End If
                  m_jobs.Add NewJob, Hex$(NewJob.JobId)
                  ' Store this job id
                  JobIds(i) = NewJob.JobId
               Next i
            End If
         End If
      End If
      Call ClosePrinter(hPrn)
   End If
   
   ' Clean out jobs collection of old jobs
   For Each OldJob In m_jobs
      If nJobs > 0 Then
         For i = LBound(JobIds) To UBound(JobIds)
            If OldJob.JobId = JobIds(i) Then Exit For
         Next i
         If i = UBound(JobIds) + 1 Then
            ' Didn't find a matching job!
            m_jobs.Remove Hex$(OldJob.JobId)
         End If
      Else
         ' No jobs, remove all
         m_jobs.Remove Hex$(OldJob.JobId)
      End If
   Next OldJob
End Sub

' *********************************************
'  Private Methods
' *********************************************
Private Function AdjustJobPosition(ByVal JobId As Long, ByVal Delta As Long) As Boolean
   Dim pd As PRINTER_DEFAULTS
   Dim hPrn As Long
   Dim Buffer() As Byte
   Dim BytesNeeded As Long
   Dim Position As Long
   
   ' Requires PRINTER_ACCESS_ADMINISTER permissions.
   
   ' Get handle to printer.
   pd.pDesiredAccess = PRINTER_ACCESS_ADMINISTER
   Call OpenPrinter(m_DevName, hPrn, pd)
   If hPrn Then
      ' Call once to get proper buffer size.
      Call GetJob(hPrn, JobId, 1, ByVal 0&, 0&, BytesNeeded)
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
         ' Size buffer and get job data.
         ReDim Buffer(0 To BytesNeeded - 1) As Byte
         If GetJob(hPrn, JobId, 1, Buffer(0), BytesNeeded, BytesNeeded) Then
            ' Position is the 10th element in the structure.
            Const PositionOffset As Long = 9 * 4&
            Call CopyMemory(Position, Buffer(PositionOffset), 4&)
            Position = Position + Delta
            Call CopyMemory(Buffer(PositionOffset), Position, 4&)
            ' Attempt to update job data.
            AdjustJobPosition = CBool(SetJob(hPrn, JobId, 1, Buffer(0), 0))
         End If
      End If
      Call ClosePrinter(hPrn)
   End If
End Function

Private Function GetJobInfo(ByVal JobId As Long) As CPrinterJobInfo
   On Error Resume Next
   Set GetJobInfo = m_jobs(Hex$(JobId))
End Function

Private Function JobCount(Optional ByVal hPrn As Long = 0) As Long
   Dim Buffer() As Byte
   Dim BytesNeeded As Long
   Dim BytesUsed As Long
   Dim CloseHandle As Boolean
   
   ' If caller didn't open printer, we need to!
   If hPrn = 0 Then
      ' Get handle to printer.
      Call OpenPrinter(m_DevName, hPrn, ByVal 0&)
      CloseHandle = True
   End If

   If hPrn Then
      ' Call once to get proper buffer size.
      Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
         ' Size buffer and get printer data.
         ReDim Buffer(0 To BytesNeeded - 1) As Byte
         If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
            ' cJobs is the 20th 4-byte element of the
            ' PRINTER_INFO_2 structure.
            Const JobsOffset As Long = 19 * 4&
            Call CopyMemory(JobCount, Buffer(JobsOffset), 4&)
         End If
         m_GetPrinterError = 0  'clear error value
      Else
         m_GetPrinterError = Err.LastDllError
      End If
   End If
   
   ' If we opened printer, need to close it too.
   If CloseHandle Then Call ClosePrinter(hPrn)
End Function

Private Function SendControl(ByVal JobId As Long, ByVal ControlCode As JobControlCodes) As Boolean
   Dim hPrn As Long
   ' Get handle to printer.
   Call OpenPrinter(m_DevName, hPrn, ByVal 0&)
   If hPrn Then
      ' Send requested control code.
      SendControl = CBool(SetJob(hPrn, JobId, 0, ByVal 0&, ControlCode))
      Call ClosePrinter(hPrn)
      ' Update all object data.
      Call Me.Refresh
   End If
End Function
