Attribute VB_Name = "Variables"
Option Explicit

'--------------------------API Declarations------------------------
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long 'used to time things

'--------------------------Constants-------------------------------
'---System Constants---
Public Const Black As Long = &H80000006
Public Const White As Long = &H80000005
Public Const OFF_BOARD As Byte = 107
Public Const NOT_PIECE As Byte = 108
Public Const ReDimInterval As Long = 5000

'---Finished Constants---
Public Const NORMAL As Long = 0
Public Const LEAVE As Long = 1
Public Const CANNOTMOVE As Long = 2

'---AI Constants---
Public Const CDouble As Long = 3
Public Const CSingle As Long = 1
Public Const CStrategy As Long = 1

'---Registry Constants---
Public Const RegistryKey As String = "Software\Infostrategy\"
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

'--------------------------Self Defined Types-----------------------
Public Type PieceAttr
  Player As Long
  X As Long
  Y As Long
  Index As Long
  Double As Boolean
End Type

Public Type StatePieceAttr
  Index As Byte
  Double As Byte
End Type

Public Type StateSelectedSquare
  Index As Long
  Player As Long
  Piece As Long
  Double As Boolean
End Type

Public Type SelectedSquare
  X As Long
  Y As Long
  Index As Long
  Player As Long
  IsPiece As Boolean
  IsSquare As Boolean
  Piece As Long
  Double As Boolean
End Type

Public Type ArrayPieces
  Pieces(1 To 40) As StatePieceAttr
End Type

Public Type BoardState
  ParentID As Long
  SubParentId As Long
  ArrayNum As Long
  Score As Integer
  Depth As Byte
  Finished As Byte
End Type

Public Type Coordinate
  X As Long
  Y As Long
End Type

Public Type Time
  Minutes As Long
  Seconds As Double
End Type

Public Type MMoves
  ParentID As Long
  From As Long
  Too As Long
End Type

'-------------------------Global Variables------------------------
Public IndexMoves(1 To 8, 1 To 10) As Long 'Array of Index Increments for different directions
Public GameStarted As Boolean
Public XYMoves(1 To 8, 1 To 10) As Coordinate
Public Currentpieces(1 To 40) As PieceAttr
Public Lastpieces(1 To 40) As PieceAttr
Public Score(1 To 2) As Long, Turn As Long, Names(1 To 2) As String
Public ERR7 As Boolean 'Has run out of memory
Public VP1Time As Time, VP2Time As Time, TotalTime As Time, VTurns As Long 'Timing variables
Public MaxThoughtTime As Long
Public P1MultiMode As Boolean
Public PlayerType As Long
Public AutoDebug As Long
Public MoveSpeed As Long
Public ForceMove As Long
Public MaxDepth As Long
Public ABPMode As Long
Public UpperMove As Long
Public UpperS As Long
Public IsAdvanced As Long
Public SMatrix() As BoardState
Public MoveMatrix() As ArrayPieces
Public Reversed As Boolean 'Is board currently reserved
Public AutoSwitch As Long
Public CheatSwitch As Long
Public MemoryLimit As Long 'Maximum upper bound of States
Public PruneThreshold As Long 'Error margin for MemoryLimit (in percent)
