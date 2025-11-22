Attribute VB_Name = "modGlobal"
Option Explicit

Public Const MCI_CLOSE = &H804&
Public Const MCI_WAIT = &H2&
Public Const MCI_OPEN = &H803&
Public Const MCI_FORMAT_MILLISECONDS = 0
Public Const MCI_SET = &H80D&
Public Const MCI_OPEN_ELEMENT = &H200&
Public Const MCI_SET_TIME_FORMAT = &H400&
Public Const MCI_STOP = &H808&
Public Const MCI_SEEK = &H807&
Public Const MCI_SEEK_TO_START = &H100&
Public Const MCI_PLAY = &H806&
Public Const MCI_NOTIFY = &H1&
Public Const MCI_STATUS_POSITION = &H2&
Public Const MCI_STATUS = &H814&
Public Const MCI_STATUS_ITEM = &H100&
Public Const MCI_STATUS_LENGTH = &H1&
Public Const MCI_TO = &H8&
Public Const MCI_OPEN_TYPE = &H2000&
Public Const MCI_RECORD = &H80F&
Public Const MCI_SAVE = &H813&
Public Const MCI_SAVE_FILE = &H100&
Public Const MCI_MCIAVI_PLAY_WINDOW = &H1000000
Public Const MCI_PAUSE = &H809&
Public Const MCI_DGV_OPEN_PARENT = &H20000
Public Const MCI_DGV_OPEN_WS = &H10000
Public Const MCI_WINDOW = &H841&
Public Const MCI_DGV_WINDOW_STATE = &H40000
Public Const MCI_DGV_STATUS_HWND = &H4001&
Public Const MCI_WHERE = &H843
Public Const MCI_DGV_WHERE_SOURCE = &H20000

Public Const WS_CHILD = &H40000000
Public Const SW_SHOW = 5

Public Const MAX_PATH As Long = 260

Public Const INVALID_HANDLE_VALUE = -1

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Function FileExists(FileName As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FindFirstFile(FileName, wfd)
  FileExists = Not hFile = INVALID_HANDLE_VALUE
  Call FindClose(hFile)
End Function
