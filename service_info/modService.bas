Attribute VB_Name = "modService"
Public Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
End Type

Public Type ENUM_SERVICE_STATUS
   lpServiceName As Long
   lpDisplayName As Long
   ServiceStatus As SERVICE_STATUS
End Type

Public Type QUERY_SERVICE_CONFIG
         dwServiceType As Long
         dwStartType As Long
         dwErrorControl As Long
         lpBinaryPathName As Long 'String
         lpLoadOrderGroup As Long ' String
         dwTagId As Long
         lpDependencies As Long 'String
         lpServiceStartName As Long 'String
         lpDisplayName As Long  'String
End Type

'our own constant
Public Const SIZEOF_SERVICE_STATUS As Long = 36

'windows constants
Public Const ERROR_MORE_DATA = 234
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const LB_SETTABSTOPS As Long = &H192
Public Const SERVICE_STATE_ALL = &H3
                                     
'Service Types (Bit Mask)
'corresponds to SERVICE_STATUS.dwServiceType
Public Const SERVICE_KERNEL_DRIVER As Long = &H1
Public Const SERVICE_FILE_SYSTEM_DRIVER As Long = &H2
Public Const SERVICE_ADAPTER As Long = &H4
Public Const SERVICE_RECOGNIZER_DRIVER As Long = &H8
Public Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
Public Const SERVICE_INTERACTIVE_PROCESS As Long = &H100

'******************************************************************
'Services startups
'Correnponds to QUERY_SERVICE_CONFIG.dwStartType
'Public Const SERVICE_AUTO_START As Long = &H2
'Public Const SERVICE_BOOT_START As Long = &H0
'Public Const SERVICE_DEMAND_START As Long = &H3 'only here i can call StartService
'Public Const SERVICE_DISABLED As Long = &H4
'Public Const SERVICE_SYSTEM_START as Long

'Public Const AUTOMATIC_STARTUP As Long = SERVICE_AUTO_START Or SERVICE_BOOT_START
'Public Const MANUAL_STARTUP As Long = SERVICE_DEMAND_START
'Public Const DISABLED As Long = SERVICE_DISABLED
'*********************************************************************

Public Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS Or _
                                     SERVICE_WIN32_SHARE_PROCESS
                                     
Public Const SERVICE_DRIVER As Long = SERVICE_KERNEL_DRIVER Or _
                                      SERVICE_FILE_SYSTEM_DRIVER Or _
                                      SERVICE_RECOGNIZER_DRIVER
                                      
Public Const SERVICE_TYPE_ALL As Long = SERVICE_WIN32 Or _
                                        SERVICE_ADAPTER Or _
                                        SERVICE_DRIVER Or _
                                        SERVICE_INTERACTIVE_PROCESS
                                     
                                     
'Service State
'corresponds to SERVICE_STATUS.dwCurrentState
Public Const SERVICE_STOPPED As Long = &H1
Public Const SERVICE_START_PENDING As Long = &H2
Public Const SERVICE_STOP_PENDING As Long = &H3
Public Const SERVICE_RUNNING As Long = &H4
Public Const SERVICE_CONTINUE_PENDING As Long = &H5
Public Const SERVICE_PAUSE_PENDING As Long = &H6
Public Const SERVICE_PAUSED As Long = &H7

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

Public Declare Function OpenSCManager Lib "advapi32" _
   Alias "OpenSCManagerA" _
  (ByVal lpMachineName As String, _
   ByVal lpDatabaseName As String, _
   ByVal dwDesiredAccess As Long) As Long

Public Declare Function EnumServicesStatus Lib "advapi32" _
   Alias "EnumServicesStatusA" _
  (ByVal hSCManager As Long, _
   ByVal dwServiceType As Long, _
   ByVal dwServiceState As Long, _
   lpServices As Any, _
   ByVal cbBufSize As Long, _
   pcbBytesNeeded As Long, _
   lpServicesReturned As Long, _
   lpResumeHandle As Long) As Long
   
Public Declare Function CloseServiceHandle Lib "advapi32" _
   (ByVal hSCObject As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Public Const GENERIC_READ = &H80000000

Public Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long

Public Declare Function OpenService Lib "advapi32.dll" _
         Alias "OpenServiceA" _
         (ByVal hSCManager As Long, _
         ByVal lpServiceName As String, _
         ByVal dwDesiredAccess As Long) As Long
Public Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Public Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long



