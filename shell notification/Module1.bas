Attribute VB_Name = "Module1"
Option Explicit


Public Const MAX_PATH As Long = 260

'Defined as an HRESULT that corresponds
'to S_OK.
Public Const ERROR_SUCCESS As Long = 0

Public Type SHFILEINFO   'shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'If pidl is invalid, SHGetFileInfoPidl can
'very easily blow up when filling the
'szDisplayName and szTypeName string members
'of the SHFILEINFO struct
Public Type SHFILEINFOBYTE   'sfib
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName(1 To MAX_PATH) As Byte
  szTypeName(1 To 80) As Byte
End Type

'Special folder values for
'SHGetSpecialFolderLocation and
'SHGetSpecialFolderPath (Shell32.dll v4.71)
Public Enum SHSpecialFolderIDs
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_ALTSTARTUP = &H1D        ''DBCS
  CSIDL_COMMON_ALTSTARTUP = &H1E ''DBCS
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
End Enum

Enum SHGFI_FLAGS
  SHGFI_LARGEICON = &H0           'sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1           'sfi.hIcon is small icon
  SHGFI_OPENICON = &H2            'sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4       'sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                'pszPath is pidl, rtns BOOL
  SHGFI_USEFILEATTRIBUTES = &H10  'parent pszPath exists, rtns BOOL
  SHGFI_ICON = &H100              'fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200       'isf.szDisplayName is filled, rtns BOOL
  SHGFI_TYPENAME = &H400          'isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800        'rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000     'fills sfi.szDisplayName with filename
                                  '   containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000          'rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000     'sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000&     'add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        'sfi.hIcon is selected icon
End Enum

Public Declare Function FlashWindow Lib "user32" _
   (ByVal hWnd As Long, _
    ByVal bInvert As Long) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

'Frees memory allocated by the shell (pidls)
Public Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)

'Retrieves the location of a special
'(system) folder. Returns ERROR_SUCCESS if
'successful or an OLE-defined error
'result otherwise.
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" _
   (ByVal hwndOwner As Long, _
    ByVal nFolder As SHSpecialFolderIDs, _
    pidl As Long) As Long

'Converts an item identifier list to a
'file system path. Returns TRUE if successful
'or FALSE if an error occurs, for example,
'if the location specified by the pidl
'parameter is not part of the file system.
Public Declare Function SHGetPathFromIDList Lib "shell32" _
    Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

'Retrieves information about an object
'in the file system, such as a file,
'a folder, a directory, or a drive root.
Public Declare Function SHGetFileInfoPidl Lib "shell32" _
    Alias "SHGetFileInfoA" _
   (ByVal pidl As Long, _
    ByVal dwFileAttributes As Long, _
    psfib As SHFILEINFOBYTE, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As SHGFI_FLAGS) As Long

Public Declare Function SHGetFileInfo Lib "shell32" _
    Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As SHGFI_FLAGS) As Long


Public Function GetPIDLFromFolderID(hOwner As Long, _
                                    nFolder As SHSpecialFolderIDs) As Long
                                    
  'Returns an absolute pidl (relative to
  'the desktop) from a special folder's ID.
  '(Calling proc is responsible for freeing
  'the pidl)
  'hOwner - handle of window that will
  '         own any displayed msg boxes
  'nFolder  - special folder ID
 
   Dim pidl As Long
   
   If SHGetSpecialFolderLocation(hOwner, _
                                  nFolder, _
                                  pidl) = ERROR_SUCCESS Then
      GetPIDLFromFolderID = pidl
   End If
   
End Function


Public Function GetDisplayNameFromPIDL(pidl As Long) As String
    
  'If successful returns the specified
  'absolute pidl's displayname, returns
  'an empty string otherwise.

   Dim sfib As SHFILEINFOBYTE
   
   If SHGetFileInfoPidl(pidl, 0, sfib, Len(sfib), _
                        SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
                        
      GetDisplayNameFromPIDL = _
         GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
      
   End If
   
End Function


Public Function GetPathFromPIDL(pidl As Long) As String

  'Returns a path from only an absolute pidl
  '(relative to the desktop).

   Dim sPath As String * MAX_PATH
   
  'SHGetPathFromIDList rtns TRUE (1),
  'if successful, FALSE (0) if not
   If SHGetPathFromIDList(pidl, sPath) Then
      GetPathFromPIDL = GetStrFromBufferA(sPath)
   End If
   
End Function


Public Function GetStrFromBufferA(sz As String) As String

   'Return the string before first null
   'char encountered (if any) from an
   'ANSII string. If no null, return the
   'string passed
  
   If InStr(sz, vbNullChar) Then
         GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
   Else: GetStrFromBufferA = sz
   End If
   
End Function



