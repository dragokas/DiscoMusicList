Attribute VB_Name = "WriteLog"
Option Explicit

Private Const MAX_PATH As Long = 260&

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long

Private Const ERROR_MORE_DATA       As Long = 234&

Public Sub ErrorMsg(ByVal ErrObj As ErrObject, ParamArray CodeModule())
    Dim HRESULT     As String
    Dim Other       As String
    Dim ErrNumber   As Long
    Dim ErrDescr    As String
    Dim ErrLastDll  As Long
    Dim i           As Long
    Dim ErrText     As String
    Dim HRESULT_LastDll As String
    Dim curTime     As Date
    
    curTime = Now()
    
    ErrNumber = ErrObj.Number           'сохраняю изначальные свойства
    ErrDescr = Replace$(Replace$(ErrObj.Description, vbCr, " "), vbLf, " ")
    ErrLastDll = ErrObj.LastDllError
    
    On Error Resume Next
    
    For i = 1 To UBound(CodeModule)
        Other = Other & CodeModule(i) & " "
    Next
    
    If ErrLastDll <> 0 Then
        HRESULT_LastDll = MessageText(ErrLastDll)
    End If
    If ErrNumber <> 0 Then
        HRESULT = MessageText(ErrNumber)
    End If
    
    ErrText = _
        "- " & ParseDateTime(curTime) & _
        " - " & CodeModule(0) & _
        " - #" & ErrNumber
    If ErrNumber <> 0 Then ErrText = ErrText & " (" & ErrDescr & ")" & IIf(Len(HRESULT) <> 0, " (" & HRESULT & ")", "")
    ErrText = ErrText & " LastDllError = " & ErrLastDll
    If ErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(Other) <> 0 Then ErrText = ErrText & " " & Other
    
    Debug.Print ErrText
End Sub

' Преобразовать дату и время в формат DD.MM.YYYY hh:mm:ss
Public Function ParseDateTime(myDate As Date) As String
    ParseDateTime = Right$("0" & Day(myDate), 2) & _
        "." & Right$("0" & Month(myDate), 2) & _
        "." & Year(myDate) & _
        " " & Right$("0" & Hour(myDate), 2) & _
        ":" & Right$("0" & Minute(myDate), 2) & _
        ":" & Right$("0" & Second(myDate), 2)
End Function

Public Function MessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
    
    Dim sRtrnCode   As String
    Dim lRet        As Long

    sRtrnCode = String$(MAX_PATH, 0&)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, lCode, ByVal 0&, sRtrnCode, MAX_PATH, ByVal 0&)
    If lRet > 0 Then
        MessageText = Left$(sRtrnCode, lRet)
        MessageText = Replace$(MessageText, vbCr, " ")
        MessageText = Replace$(MessageText, vbLf, " ")
    End If
End Function

Public Sub AppendErrorLogCustom(ParamArray CodeModule())
End Sub

Public Function ReadIniValue(sPath As String, sSection As String, sParam As String, Optional sDefaultData As String = "") As String
    Dim buf As String
    Dim lr As Long
    
    buf = String$(256&, 0)
    lr = GetPrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(sDefaultData), StrPtr(buf), Len(buf), StrPtr(sPath))
    If Err.LastDllError = ERROR_MORE_DATA Then
        buf = String$(1001&, 0)
        lr = GetPrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(sDefaultData), StrPtr(buf), Len(buf), StrPtr(sPath))
        If Err.LastDllError = ERROR_MORE_DATA Then
            buf = String$(10001&, 0)
            lr = GetPrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(sDefaultData), StrPtr(buf), Len(buf), StrPtr(sPath))
        End If
    End If
    
    If Err.LastDllError = 0 Then
        ReadIniValue = Left$(buf, lr)
    End If
End Function

Public Function WriteIniValue(sPath As String, sSection As String, sParam As String, vData As Variant) As Long
    WriteIniValue = WritePrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(CStr(vData)), StrPtr(sPath))
End Function

