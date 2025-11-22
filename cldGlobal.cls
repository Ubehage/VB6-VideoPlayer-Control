VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cldGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type MCI_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
End Type

Public Type MCI_GENERIC_PARMS
  dwCallback As Long
End Type

Public Type MCI_SET_PARMS
  dwCallback As Long
  dwTimeFormat As Long
  dwAudio As Long
End Type

Public Type MCI_PLAY_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Public Type MCI_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Integer
End Type

Public Type MCI_SEEK_PARMS
  dwCallback As Long
  dwTo As Long
End Type

Public Type MCI_RECORD_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Public Type MCI_SAVE_PARMS
  dwCallback As Long
  lpFileName As String
End Type

Public Type AVIRECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type MCI_DGV_RECT_PARMS
  dwCallback As Long
  rc As AVIRECT
End Type

Public Type MCI_DGV_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
  dwStyle As Long
  hWndParent As Long
End Type

Public Type MCI_DGV_WINDOW_PARMS
  dwCallback As Long
  hWnd As Long
  nCmdShow As Long
  lpstrText As String
End Type

Public Type MCI_DGV_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Long
  lpstrDrive As String
  dwReference As Long
End Type

Public Type VideoSize
  Width As Long
  Height As Long
End Type

Public Enum MULTIMEDIA_COMMAND
  mcOpen = &H1
  mcClose = &H2
  mcPlay = &H4
  mcPause = &H8
  mcStop = &H10
  mcRewind = &H20
End Enum
