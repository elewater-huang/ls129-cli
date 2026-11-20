Attribute VB_Name = "M_LS129"
Option Explicit
' LS129 SDK V0.3 for 32-bit VBA (Access 2000/2003, Excel 32-bit).
' Import this BAS into your VBA project. No Python or COM registration.
' Copy LS129Core.dll and LS129.ini to your MDB/XLS folder.
'
' Native DLL uses stdcall and stable, undecorated names from LS129Core.def.
' Values have NO unit conversion. Raw energy is returned as Double to
' preserve unsigned 32-bit data. Correction applies only to irradiance.
' IMPORTANT: reset and smoothing write to the physical LS129.

Private Declare Function WinLoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long

Private Declare Function NativeOpen Lib "LS129Core.dll" Alias "LS129_Open" _
    (ByVal iniPath As String, ByRef session As Long) As Long

Private Declare Function NativeRead Lib "LS129Core.dll" Alias "LS129_Read" _
    (ByVal session As Long, ByRef rawValue As Double, _
     ByRef adjustedValue As Double) As Long

Private Declare Function NativeReadAll Lib "LS129Core.dll" Alias "LS129_ReadAll" _
    (ByVal session As Long, ByRef rawValue As Double, _
     ByRef adjustedValue As Double, ByRef peakValue As Double, _
     ByRef energyValue As Double) As Long

Private Declare Function NativeReset Lib "LS129Core.dll" Alias "LS129_Reset" _
    (ByVal session As Long) As Long

Private Declare Function NativeGetSmoothing Lib "LS129Core.dll" Alias "LS129_GetSmoothing" _
    (ByVal session As Long, ByRef mode As Long) As Long

Private Declare Function NativeSetSmoothing Lib "LS129Core.dll" Alias "LS129_SetSmoothing" _
    (ByVal session As Long, ByVal mode As Long) As Long

Private Declare Function NativeClose Lib "LS129Core.dll" Alias "LS129_Close" _
    (ByVal session As Long) As Long

Private Declare Function NativeVersion Lib "LS129Core.dll" Alias "LS129_APIVersion" _
    () As Long

Private mSession As Long
Private mDllHandle As Long

' Pass CurrentProject.Path from Access, ThisWorkbook.Path from Excel.
Public Function LS129_Begin(ByVal projectFolder As String) As Long
    Dim folder As String
    Dim libraryPath As String
    Dim configPath As String
    Dim rc As Long

    If mSession <> 0 Then
        LS129_Begin = 105
        Exit Function
    End If
    If Len(projectFolder) = 0 Then
        LS129_Begin = 101
        Exit Function
    End If

    folder = projectFolder
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    libraryPath = folder & "\LS129Core.dll"
    configPath = folder & "\LS129.ini"

    If Len(Dir$(libraryPath)) = 0 Then
        LS129_Begin = 102
        Exit Function
    End If
    If Len(Dir$(configPath)) = 0 Then
        LS129_Begin = 104
        Exit Function
    End If
    If mDllHandle = 0 Then mDllHandle = WinLoadLibrary(libraryPath)
    If mDllHandle = 0 Then
        LS129_Begin = 103
        Exit Function
    End If

    rc = NativeOpen(configPath, mSession)
    LS129_Begin = rc
End Function

Public Function LS129_ReadCurrent(ByRef rawValue As Double, _
                                 ByRef adjustedValue As Double) As Long
    If mSession = 0 Then
        LS129_ReadCurrent = 6
        Exit Function
    End If
    LS129_ReadCurrent = NativeRead(mSession, rawValue, adjustedValue)
End Function

Public Function LS129_ReadStatistics(ByRef rawValue As Double, _
                                    ByRef adjustedValue As Double, _
                                    ByRef peakValue As Double, _
                                    ByRef energyValue As Double) As Long
    If mSession = 0 Then
        LS129_ReadStatistics = 6
        Exit Function
    End If
    LS129_ReadStatistics = NativeReadAll(mSession, rawValue, adjustedValue, _
                                        peakValue, energyValue)
End Function

' Resets cumulative energy AND peak value on the probe.
Public Function LS129_NewPeriod() As Long
    If mSession = 0 Then
        LS129_NewPeriod = 6
        Exit Function
    End If
    LS129_NewPeriod = NativeReset(mSession)
End Function

' 0=off, 1=50Hz, 2=60Hz. Changes the probe, with readback verification.
Public Function LS129_Smoothing(ByVal mode As Long) As Long
    If mSession = 0 Then
        LS129_Smoothing = 6
        Exit Function
    End If
    LS129_Smoothing = NativeSetSmoothing(mSession, mode)
End Function

Public Function LS129_CurrentSmoothing(ByRef mode As Long) As Long
    If mSession = 0 Then
        LS129_CurrentSmoothing = 6
        Exit Function
    End If
    LS129_CurrentSmoothing = NativeGetSmoothing(mSession, mode)
End Function

Public Function LS129_End() As Long
    Dim rc As Long
    If mSession = 0 Then
        LS129_End = 6
        Exit Function
    End If
    rc = NativeClose(mSession)
    If rc = 0 Then mSession = 0
    LS129_End = rc
End Function

Public Function LS129_Version() As Long
    If mDllHandle = 0 Then
        LS129_Version = 0
    Else
        LS129_Version = NativeVersion()
    End If
End Function

Public Function LS129_ErrorText(ByVal rc As Long) As String
    Select Case rc
        Case 0: LS129_ErrorText = "OK"
        Case 1: LS129_ErrorText = "Invalid argument"
        Case 2: LS129_ErrorText = "INI file missing"
        Case 3: LS129_ErrorText = "Invalid INI configuration"
        Case 4: LS129_ErrorText = "Cannot open COM port"
        Case 5: LS129_ErrorText = "Device already open"
        Case 6: LS129_ErrorText = "No valid device session"
        Case 101: LS129_ErrorText = "Project folder not supplied"
        Case 102: LS129_ErrorText = "LS129Core.dll missing from folder"
        Case 103: LS129_ErrorText = "Failed to load x86 DLL"
        Case 104: LS129_ErrorText = "LS129.ini missing from folder"
        Case 105: LS129_ErrorText = "VBA session already active"
        Case Else: LS129_ErrorText = "Serial/Modbus error code " & CStr(rc)
    End Select
End Function
