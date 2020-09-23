Attribute VB_Name = "Module3"
Option Explicit
'the one and only shell change notification
'handle for the desktop folder
Private m_hSHNotify As Long

'the desktop's pidl
Private m_pidlDesktop As Long

'User defined notification message sent
'to the specified window's window proc.
Public Const WM_SHNOTIFY = &H401

'------------------------------------------------------
Public Type PIDLSTRUCT
  'Fully qualified pidl (relative to
  'the desktop folder) of the folder
  'to monitor changes in. 0 can also
  'be specified for the desktop folder.
   pidl As Long
   
  'Value specifying whether changes in
  'the folder's subfolders trigger a
  'change notification event (it's actually
  'a Boolean, but we'll go Long because
  'of VB's DWORD struct alignment).
   bWatchSubFolders As Long
End Type

Public Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" _
   (ByVal hWnd As Long, _
    ByVal uFlags As SHCN_ItemFlags, _
    ByVal dwEventID As SHCN_EventIDs, _
    ByVal uMsg As Long, _
    ByVal cItems As Long, _
    lpps As PIDLSTRUCT) As Long

'hWnd     - Handle of the window to receive
'          the window message specified in uMsg.
'
'uFlags   - Flag that indicates the meaning of
'           the dwItem1 and dwItem2 members of
'           the SHNOTIFYSTRUCT (which is pointed
'           to by the window procedure's wParam
'           value when the specified window message
'           is received). This parameter can
'           be one of the SHCN_ItemFlags enum
'           values below. This interpretation may
'           be inaccurate as it appears pidls are
'           almost always returned in the SHNOTIFYSTRUCT.
'           See James' site for more info...
'
'dwEventId- Combination of SHCN_EventIDs enum
'           values that specifies what events the
'           specified window will be notified of.
'           See below.
'
'uMsg     - Window message to be used to identify
'           receipt of a shell change notification.
'           The message should *not* be a value that
'           lies within the specified window's
'           message range ( i.e. BM_ messages for
'           a button window) or that window may
'           not receive all (if not any) notifications
'           sent by the shell!!!
'
'cItems   - Count of PIDLSTRUCT structures in the array
'           pointed to by the lpps param.
'
'lpps     - Pointer to an array of PIDLSTRUCT structures
'           indicating what folder(s) to monitor changes in,
'           and whether to watch the specified folder's subfolder.

'If successful, SHChangeNotifyRegister returns a notification
'handle which must be passed to SHChangeNotifyDeregister
'when no longer used. Returns 0 otherwise.

'Once the specified message is registered with SHChangeNotifyRegister,
'the specified window's function proc will be notified by the shell
'of the specified event in (and under) the folder(s) specified in a pidl.
'On message receipt, wParam points to a SHNOTIFYSTRUCT and lParam
'contains the event's ID value.

'The values in dwItem1 and dwItem2 are event specific. See the
'description of the values for the wEventId parameter of the
'documented SHChangeNotify API function.
Public Type SHNOTIFYSTRUCT
  dwItem1 As Long
  dwItem2 As Long
End Type

'...?
'Public Declare Function SHChangeNotifyUpdateEntryList Lib "shell32" Alias "#5" _
'                             (ByVal hNotify As Long, _
'                             ByVal Unknown As Long, _
'                             ByVal cItem As Long, _
'                             lpps As PIDLSTRUCT) As Boolean
'
'Public Declare Function SHChangeNotifyReceive Lib "shell32" Alias "#5" _
'                             (ByVal hNotify As Long, _
'                             ByVal uFlags As SHCN_ItemFlags, _
'                             ByVal dwItem1 As Long, _
'                             ByVal dwItem2 As Long) As Long

'Closes the notification handle returned from a call
'to SHChangeNotifyRegister. Returns True if successful,
'False otherwise.
Public Declare Function SHChangeNotifyDeregister Lib "shell32" _
    Alias "#4" _
   (ByVal hNotify As Long) As Boolean

'------------------------------------------------------
'This function should be called by any app that
'changes anything in the shell. The shell will then
'notify each "notification registered" window of this action.
Public Declare Sub SHChangeNotify Lib "shell32" _
   (ByVal wEventId As SHCN_EventIDs, _
    ByVal uFlags As SHCN_ItemFlags, _
    ByVal dwItem1 As Long, _
    ByVal dwItem2 As Long)

'Shell notification event IDs
Public Enum SHCN_EventIDs
   SHCNE_RENAMEITEM = &H1          '(D) A non-folder item has been renamed.
   SHCNE_CREATE = &H2              '(D) A non-folder item has been created.
   SHCNE_DELETE = &H4              '(D) A non-folder item has been deleted.
   SHCNE_MKDIR = &H8               '(D) A folder item has been created.
   SHCNE_RMDIR = &H10              '(D) A folder item has been removed.
   SHCNE_MEDIAINSERTED = &H20      '(G) Storage media has been inserted into a drive.
   SHCNE_MEDIAREMOVED = &H40       '(G) Storage media has been removed from a drive.
   SHCNE_DRIVEREMOVED = &H80       '(G) A drive has been removed.
   SHCNE_DRIVEADD = &H100          '(G) A drive has been added.
   SHCNE_NETSHARE = &H200          'A folder on the local computer is being
                                   '    shared via the network.
   SHCNE_NETUNSHARE = &H400        'A folder on the local computer is no longer
                                   '    being shared via the network.
   SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
   SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed,
                                   '    but the folder still exists and has not been renamed.
   SHCNE_UPDATEITEM = &H2000       '(D) An existing non-folder item has changed, but the
                                   '    item still exists and has not been renamed.
   SHCNE_SERVERDISCONNECT = &H4000 'The computer has disconnected from a server.
   SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
   SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should
                                   '    create a new window for the drive.
   SHCNE_RENAMEFOLDER = &H20000    '(D) The name of a folder has changed.
   SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.

#If (WIN32_IE >= &H400) Then
   SHCNE_EXTENDED_EVENT = &H4000000 '(G) Not currently used.
#End If

  SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.
  SHCNE_DISKEVENTS = &H2381F       '(D) Specifies a combination of all of the disk
                                   '    event identifiers.
  SHCNE_GLOBALEVENTS = &HC0581E0   '(G) Specifies a combination of all of the global
                                   '    event identifiers.
  SHCNE_ALLEVENTS = &H7FFFFFFF
  SHCNE_INTERRUPT = &H80000000     'The specified event occurred as a result of a system
                                   'interrupt. It is stripped out before the clients
                                   'of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then
   Public Const SHCNEE_ORDERCHANGED = &H2 'dwItem2 is the pidl of the changed folder
#End If

'Notification flags
'uFlags & SHCNF_TYPE is an ID which indicates
'what dwItem1 and dwItem2 mean
Public Enum SHCN_ItemFlags
   SHCNF_IDLIST = &H0         'LPITEMIDLIST
   SHCNF_PATHA = &H1          'path name
   SHCNF_PRINTERA = &H2       'printer friendly name
   SHCNF_DWORD = &H3          'DWORD
   SHCNF_PATHW = &H5          'path name
   SHCNF_PRINTERW = &H6       'printer friendly name
   SHCNF_TYPE = &HFF
  'Flushes the system event buffer. The
  'function does not return until the system
  'is finished processing the given event.
   SHCNF_FLUSH = &H1000
  'Flushes the system event buffer. The function
  'returns immediately regardless of whether
  'the system is finished processing the given event.
   SHCNF_FLUSHNOWAIT = &H2000

#If UNICODE Then
  SHCNF_PATH = SHCNF_PATHW
  SHCNF_PRINTER = SHCNF_PRINTERW
#Else
  SHCNF_PATH = SHCNF_PATHA
  SHCNF_PRINTER = SHCNF_PRINTERA
#End If

End Enum


Public Function SHNotify_Register(hWnd As Long) As Boolean
  
  'Registers the one and only shell change notification.

   Dim ps As PIDLSTRUCT
  
  'If we don't already have a notification going...
   If (m_hSHNotify = 0) Then
  
     'Get the pidl for the desktop folder.
      m_pidlDesktop = GetPIDLFromFolderID(0, CSIDL_DESKTOP)
      
      If m_pidlDesktop Then
      
        'Fill the one and only PIDLSTRUCT, we're
        'watching desktop and all of the it's
        'subfolders, everything...
         ps.pidl = m_pidlDesktop
         ps.bWatchSubFolders = True
         
        'Register the notification, specifying that
        'we want the dwItem1 and dwItem2 members of
        'the SHNOTIFYSTRUCT to be pidls. We're
        'watching all events.
         m_hSHNotify = SHChangeNotifyRegister(hWnd, _
                                              SHCNF_TYPE Or SHCNF_IDLIST, _
                                              SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, _
                                              WM_SHNOTIFY, _
                                              1, _
                                              ps)
                                              
         SHNotify_Register = CBool(m_hSHNotify)
    
    Else
    
        'If something went wrong...
         Call CoTaskMemFree(m_pidlDesktop)
    
    End If
    
  End If
  
End Function


Public Function SHNotify_Unregister() As Boolean
  
  'Unregisters the one and only shell change notification.
  
  'If we have a registered notification handle.
   If m_hSHNotify Then
   
     'Unregister it. If the call is successful,
     'zero the handle's variable, free and zero
     'the the desktop's pidl.
      If SHChangeNotifyDeregister(m_hSHNotify) Then
      
         m_hSHNotify = 0
         Call CoTaskMemFree(m_pidlDesktop)
         m_pidlDesktop = 0
         SHNotify_Unregister = True
         
      End If
      
   End If

End Function


Public Function SHNotify_GetEventStr(dwEventID As Long) As String

  'Returns the event string associated
  'with the specified event ID value.
  
   Dim sEvent As String
   
   Select Case dwEventID
      Case SHCNE_RENAMEITEM:       sEvent = "SHCNE_RENAMEITEM"       '&H1
      Case SHCNE_CREATE:           sEvent = "SHCNE_CREATE"           '&H2
      Case SHCNE_DELETE:           sEvent = "SHCNE_DELETE"           '&H4
      Case SHCNE_MKDIR:            sEvent = "SHCNE_MKDIR"            '&H8
      Case SHCNE_RMDIR:            sEvent = "SHCNE_RMDIR"            '&H10
      Case SHCNE_MEDIAINSERTED:    sEvent = "SHCNE_MEDIAINSERTED"    '&H20
      Case SHCNE_MEDIAREMOVED:     sEvent = "SHCNE_MEDIAREMOVED"     '&H40
      Case SHCNE_DRIVEREMOVED:     sEvent = "SHCNE_DRIVEREMOVED"     '&H80
      Case SHCNE_DRIVEADD:         sEvent = "SHCNE_DRIVEADD"         '&H100
      Case SHCNE_NETSHARE:         sEvent = "SHCNE_NETSHARE"         '&H200
      Case SHCNE_NETUNSHARE:       sEvent = "SHCNE_NETUNSHARE"       '&H400
      Case SHCNE_ATTRIBUTES:       sEvent = "SHCNE_ATTRIBUTES"       '&H800
      Case SHCNE_UPDATEDIR:        sEvent = "SHCNE_UPDATEDIR"        '&H1000
      Case SHCNE_UPDATEITEM:       sEvent = "SHCNE_UPDATEITEM"       '&H2000
      Case SHCNE_SERVERDISCONNECT: sEvent = "SHCNE_SERVERDISCONNECT" '&H4000
      Case SHCNE_UPDATEIMAGE:      sEvent = "SHCNE_UPDATEIMAGE"      '&H8000&
      Case SHCNE_DRIVEADDGUI:      sEvent = "SHCNE_DRIVEADDGUI"      '&H10000
      Case SHCNE_RENAMEFOLDER:     sEvent = "SHCNE_RENAMEFOLDER"     '&H20000
      Case SHCNE_FREESPACE:        sEvent = "SHCNE_FREESPACE"        '&H40000
    
#If (WIN32_IE >= &H400) Then
      Case SHCNE_EXTENDED_EVENT:   sEvent = "SHCNE_EXTENDED_EVENT"   '&H4000000
#End If
    
      Case SHCNE_ASSOCCHANGED:     sEvent = "SHCNE_ASSOCCHANGED"     '&H8000000
    
      Case SHCNE_DISKEVENTS:       sEvent = "SHCNE_DISKEVENTS"       '&H2381F
      Case SHCNE_GLOBALEVENTS:     sEvent = "SHCNE_GLOBALEVENTS"     '&HC0581E0
      Case SHCNE_ALLEVENTS:        sEvent = "SHCNE_ALLEVENTS"        '&H7FFFFFFF
      Case SHCNE_INTERRUPT:        sEvent = "SHCNE_INTERRUPT"        '&H80000000
   End Select
  
   SHNotify_GetEventStr = sEvent

End Function



