VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function IndexTranslation(Index As Byte) As Byte
  If Reversed Then IndexTranslation = 63 - Index Else IndexTranslation = Index
End Function

Public Sub ResetGame()
Dim Lng1 As Long, Lng2 As Long

Turn = 1
 
For Lng1 = 1 To 7 Step 2
  Lng2 = Lng2 + 1
  Currentpieces(Lng1 + 8).X = Lng1
  Currentpieces(Lng1 + 8).Y = 1
  Currentpieces(Lng1).X = Lng1
  Currentpieces(Lng1).Y = 7
Next

Lng2 = 0

For Lng1 = 2 To 8 Step 2
  Lng2 = Lng2 + 1
  Currentpieces(Lng1 + 8).X = Lng1
  Currentpieces(Lng1 + 8).Y = 2
  Currentpieces(Lng1).X = Lng1
  Currentpieces(Lng1).Y = 8
Next

For Lng1 = 1 To 16
  Currentpieces(Lng1).Index = IConvert(Currentpieces(Lng1).X, Currentpieces(Lng1).Y)
  If Lng1 <= 8 Then Currentpieces(Lng1).Player = 1 Else Currentpieces(Lng1).Player = 2
  Currentpieces(Lng1).Double = False
Next

Call RefreshDisplay
Call RefreshBoard(Currentpieces)

Score(1) = 8
Score(2) = 8
End Sub

Public Sub RefreshDisplay()
  If Turn = 1 Then
    frmMain.lblTurn = Names(1)
  Else
    frmMain.lblTurn = Names(2)
  End If
  
  frmMain.lblP1Points = Names(1) & " - " & 8 - Score(1)
  frmMain.lblP2Points = Names(2) & " - " & 8 - Score(2)
End Sub

Public Function IConvert(X As Byte, Y As Byte) As Byte
  If X > 8 Or X < 1 Or Y > 8 Or Y < 1 Then
    IConvert = 65
  Else
    IConvert = ((Y - 1) * 8) + X - 1
  End If
End Function

Public Sub XYConvert(Index As Byte, ByRef X As Byte, ByRef Y As Byte)
  If Index > 63 Or Index < 0 Then X = 0: Y = 0: Exit Sub
  X = (Index Mod 8) + 1
  Y = ((Index - X) / 8) + 1
End Sub

Public Function SaveSettings()
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Error Margin", ErrorMargin
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Limit", MemoryLimit
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Dynamic Ply", DynamicDepth
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "MaxDepth", MaxDepth
End Function

Public Function GetSettings()
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Error Margin", ErrorMargin
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Limit", MemoryLimit
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Dynamic Ply", DynamicDepth
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "MaxDepth", MaxDepth
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = Val(tmpVal)                                ' Copy String Value
        KeyValStr = tmpVal
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Val(Format$("&h" + KeyVal))                ' Convert Double Word To String
        KeyValStr = CStr(Format$("&h" + KeyVal))
    End Select
    
    GetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = 0                                             ' Set Return Val To Empty String
    KeyValStr = ""
    GetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function SetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If RC <> ERROR_SUCCESS Then
      If RC = ERROR_FILE_NOT_FOUND Then
        RC = RegCreateKey(KeyRoot, KeyName, hKey)
      Else
        GoTo GetKeyError          ' Handle Error...
      End If
    End If
    
    If KeyValStr = "" Then KeyValStr = CStr(KeyVal)
    KeyValSize = Len(KeyValStr) + 1
    KeyValType = REG_SZ
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegSetValueEx(hKey, SubKeyRef, 0, KeyValType, ByVal KeyValStr, KeyValSize)
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    SetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    SetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


