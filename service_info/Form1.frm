VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get info about the service"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   540
      Width           =   6255
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   105
      TabIndex        =   1
      Top             =   1065
      Width           =   6300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get info about the service!!!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   105
      TabIndex        =   0
      Top             =   3225
      Width           =   6315
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the service installed and get the info about:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   6240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Private Declare Function CloseServiceHandle Lib "advapi32.dll" _
         (ByVal hSCObject As Long) As Long

      Private Declare Function QueryServiceStatus Lib "advapi32.dll" _
         (ByVal hService As Long, _
         lpServiceStatus As SERVICE_STATUS) As Long

      Private Declare Function OpenService Lib "advapi32.dll" _
         Alias "OpenServiceA" _
         (ByVal hSCManager As Long, _
         ByVal lpServiceName As String, _
         ByVal dwDesiredAccess As Long) As Long

      Private Declare Function OpenSCManager Lib "advapi32.dll" _
         Alias "OpenSCManagerA" _
         (ByVal lpMachineName As String, _
         ByVal lpDatabaseName As String, _
         ByVal dwDesiredAccess As Long) As Long

      Private Declare Function QueryServiceConfig Lib "advapi32.dll" _
         Alias "QueryServiceConfigA" _
         (ByVal hService As Long, _
         lpServiceConfig As Byte, _
         ByVal cbBufSize As Long, _
         pcbBytesNeeded As Long) As Long

      Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, _
         hpvSource As Any, _
         ByVal cbCopy As Long)

      Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
         (ByVal lpString1 As String, _
         ByVal lpString2 As Long) As Long

      Private Type SERVICE_STATUS
         dwServiceType As Long
         dwCurrentState As Long
         dwControlsAccepted As Long
         dwWin32ExitCode As Long
         dwServiceSpecificExitCode As Long
         dwCheckPoint As Long
         dwWaitHint As Long
      End Type

      Private Type QUERY_SERVICE_CONFIG
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

      Private Const SERVICE_STOPPED = &H1
      Private Const SERVICE_START_PENDING = &H2
      Private Const SERVICE_STOP_PENDING = &H3
      Private Const SERVICE_RUNNING = &H4
      Private Const SERVICE_CONTINUE_PENDING = &H5
      Private Const SERVICE_PAUSE_PENDING = &H6
      Private Const SERVICE_PAUSED = &H7
      Private Const SERVICE_ACCEPT_STOP = &H1
      Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
      Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
      Private Const SC_MANAGER_CONNECT = &H1
      Private Const SERVICE_INTERROGATE = &H80
      Private Const GENERIC_READ = &H80000000
      Private Const ERROR_INSUFFICIENT_BUFFER = 122



Private Sub Command1_Click()
        
Call GetSvcInfo(Combo1.Text)
        
        
End Sub

Private Sub GetSvcInfo(svcname As String)

         Dim pSTATUS As SERVICE_STATUS
         Dim udtConfig As QUERY_SERVICE_CONFIG
         Dim lRet As Long
         Dim lBytesNeeded As Long
         Dim sTemp As String
         Dim pFileName As Long
         Dim success As Long
         Dim hSCManager As Long
         Dim cbBuffer As Long
         Dim hSCM As Long
         Dim hSVC As Long
        ' Dim svcname As String

      List1.Clear

        'getting list of the services



         ' Open The Service Control Manager
         '
         hSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
         If hSCM = 0 Then
            MsgBox "Error - " & Err.LastDllError
         End If

         ' Open the specific Service to obtain a handle
         '
         hSVC = OpenService(hSCM, Trim(Combo1.Text), GENERIC_READ)
            If hSVC = 0 Then
               MsgBox "Error - " & Err.LastDllError
               GoTo CloseHandles
            End If

         ' Fill the Service Status Structure
         '
         lRet = QueryServiceStatus(hSVC, pSTATUS)
         If lRet = 0 Then
            MsgBox "Error - " & Err.LastDllError
            GoTo CloseHandles
         End If

         ' Report the Current State
         '
         Select Case pSTATUS.dwCurrentState
         Case SERVICE_STOPPED
            sTemp = "The Service is Stopped"
         Case SERVICE_START_PENDING
            sTemp = "The Service Being Started"
         Case SERVICE_STOP_PENDING
            sTemp = "The Service is in the process of being stopped"
         Case SERVICE_RUNNING
            sTemp = "The Service is Running"
         Case SERVICE_CONTINUE_PENDING
            sTemp = "The Service is in the process of being Continued"
         Case SERVICE_PAUSE_PENDING
            sTemp = "The Service is in the process of being Paused"
         Case SERVICE_PAUSED
            sTemp = "The Service is Paused"
         Case SERVICE_ACCEPT_STOP
            sTemp = "The Service is Stopped"
         Case SERVICE_ACCEPT_PAUSE_CONTINUE
            sTemp = "The Service is "
         Case SERVICE_ACCEPT_SHUTDOWN
            sTemp = "The Service is being Shutdown"
         End Select

         List1.AddItem "Service Status : " & sTemp

         ' Call QueryServiceConfig with 1 byte buffer to generate an error
         ' that returns the size of a buffer we need.
         '
         ReDim abConfig(0) As Byte
         lRet = QueryServiceConfig(hSVC, abConfig(0), 0&, lBytesNeeded)
         
         Debug.Print hSVC
         
        If lRet = 0 And Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
            MsgBox "Error - " & Err.LastDllError
         End If

         ' Redim our byte array to the size necessary and call
         ' QueryServiceConfig again
         '
         ReDim abConfig(lBytesNeeded) As Byte
         lRet = QueryServiceConfig(hSVC, abConfig(0), lBytesNeeded, _
            lBytesNeeded)
         If lRet = 0 Then
            MsgBox "Error - " & Err.LastDllError
            GoTo CloseHandles
         End If

         ' Fill our Service Config User Defined Type.
         '
         CopyMemory udtConfig, abConfig(0), Len(udtConfig)

         List1.AddItem "Service Type: " & udtConfig.dwServiceType
         'List1.AddItem "Service Start Type: " & udtConfig.dwStartType
         
         Select Case udtConfig.dwStartType
         Case 2
         List1.AddItem "Service Start Type: " & "Automatic"
         Case 3
         List1.AddItem "Service Start Type: " & "Manual"
         Case 4
         List1.AddItem "Service Start Type: " & "Disabled"
         End Select
         
         Debug.Print udtConfig.dwStartType
         
         List1.AddItem "Service Error Control: " & udtConfig.dwErrorControl

         sTemp = Space(255)

         ' Now use the pointer obtained to copy the path into the temporary
         ' String Variable
         '
         lRet = lstrcpy(sTemp, udtConfig.lpBinaryPathName)
         List1.AddItem "Service Binary Path: " & sTemp

         lRet = lstrcpy(sTemp, udtConfig.lpDependencies)
         List1.AddItem "Service Dependencies: " & sTemp

         lRet = lstrcpy(sTemp, udtConfig.lpDisplayName)
         List1.AddItem "Service DisplayName: " & sTemp

         lRet = lstrcpy(sTemp, udtConfig.lpLoadOrderGroup)
         List1.AddItem "Service LoadOrderGroup: " & sTemp

         lRet = lstrcpy(sTemp, udtConfig.lpServiceStartName)
         List1.AddItem "Service Start Name: " & sTemp

CloseHandles:
      ' Close the Handle to the Service
      '
         CloseServiceHandle (hSVC)

      ' Close the Handle to the Service Control Manager
      '
         CloseServiceHandle (hSCM)

      End Sub


Private Sub Form_Load()
 
   Dim hSCManager As Long
   Dim pntr() As ENUM_SERVICE_STATUS
   Dim pSTATUS As SERVICE_STATUS
   Dim cbbuffsize As Long
   
   Dim cbRequired As Long
   Dim dwReturned As Long
   Dim hEnumResume As Long
   Dim cbBuffer As Long
   Dim success As Long
   Dim i As Long
   Dim hSCM  As Long
   
  'just help to keep the code lines
  'below from becoming too long for
  'html display
   Dim ssvcname As String
   Dim sDispName As String
   Dim dwState As Long
  ' Dim lv1 As ListView
  ' Dim li As ListItem
   Dim lRet As Long
   Dim udtConfig As QUERY_SERVICE_CONFIG
   Dim hSVC As Long
   Dim lBytesNeeded As Long
   Dim startup As String
   Dim temp As String
   
   
   'establish a connection to the service control
  'manager on the local computer and open
  'the local service control manager database.
   hSCManager = OpenSCManager(vbNullString, _
                              vbNullString, _
                              SC_MANAGER_ENUMERATE_SERVICE)

   If hSCManager <> 0 Then

     'Get buffer size by calling EnumServicesStatus.
  
     'To determine the required buffer size, call EnumServicesStatus
     'with cbBuffer and hEnumResume set to zero. EnumServicesStatus
     'fails (returns 0), and Err.LastDLLError returns ERROR_MORE_DATA,
     'filling cbRequired with the size, in bytes, of the buffer
     'required to hold the array of structures and their data.
      success = EnumServicesStatus(hSCManager, _
                                   SERVICE_WIN32, _
                                   SERVICE_STATE_ALL, _
                                   ByVal &H0, _
                                   &H0, _
                                   cbRequired, _
                                   dwReturned, _
                                   hEnumResume)

         'If success is 0 and the LastDllError is
     'ERROR_MORE_DATA, use returned info to create
     'the required data buffer
      If success = 0 And Err.LastDllError = ERROR_MORE_DATA Then

      
        'Calculate number of structures needed
        'and redimension the array
         cbBuffer = (cbRequired \ SIZEOF_SERVICE_STATUS) + 1
         ReDim pntr(0 To cbBuffer)
   
        'Set cbBuffSize equal to the size of the buffer
         cbbuffsize = cbBuffer * SIZEOF_SERVICE_STATUS

        'Enumerate the services. If the function succeeds,
        'the return value is nonzero. If the function fails,
        'the return value is zero. In addition, hEnumResume
        'must be set to 0.
         hEnumResume = 0
         If EnumServicesStatus(hSCManager, _
                               SERVICE_WIN32, _
                               SERVICE_STATE_ALL, _
                               pntr(0), _
                               cbbuffsize, _
                               cbRequired, _
                               dwReturned, _
                               hEnumResume) Then

           'pntr() array is now filled with service data,
           'so it is a simple matter of extracting the
           'required information.
            
         
            
     
            
               
               For i = 0 To dwReturned - 1
            
                  sDispName = GetStrFromPtrA(ByVal pntr(i).lpDisplayName)
                  ssvcname = GetStrFromPtrA(ByVal pntr(i).lpServiceName)
                  dwState = pntr(i).ServiceStatus.dwCurrentState
              
              Combo1.AddItem ssvcname
              
              
               Next
       '  Select Case udtConfig.dwStartType
       '  Case 2
       '  startup = "Automatic"
       '  Case 3
       '  startup = "Manual"
       '  Case 4
       '  startup = "Disabled"
        ' End Select
         
              
              
              
              
              
              
        '      Set li = .ListItems.Add(, , sDispName)
        '          li.SubItems(1) = ssvcname
        '          li.SubItems(2) = startup  'tu bude zobrazeny startup type...
        '          li.SubItems(3) = GetServiceState(dwState)
                  
               
        '    End With
            
         Else
            MsgBox "EnumServicesStatus; error " & _
                  CStr(Err.LastDllError)
         End If  'If EnumServicesStatus

   
      Else
         MsgBox "ERROR_MORE_DATA not returned; error " & _
                CStr(Err.LastDllError)
      End If  'If success = 0 And Err.LastDllError
   
   Else
      MsgBox "OpenSCManager failed; error = " & _
            CStr(Err.LastDllError)
   End If  'If hSCManager <> 0
   
  'Clean up
   Call CloseServiceHandle(hSCManager)
   
  'return the number of services
  'returned as a sign of success
'   EnumSystemServices = dwReturned


'EnumSystemServices = dwReturned
End Sub

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function
