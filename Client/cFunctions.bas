Attribute VB_Name = "cFunctions"
'**************************************************************************'
'                           Create Shortcut Declarations                                     '
'**************************************************************************'
Declare Function fCreateShellLink Lib "vb6stkit.dll" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long

'**************************************************************************'
'                          END Create Shortcut Declarations                                     '
'**************************************************************************'
