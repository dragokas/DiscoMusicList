Attribute VB_Name = "modHelper"
Option Explicit

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
End Type

Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long

Public Function IsArrDimmed(vArray As Variant) As Boolean
    IsArrDimmed = (GetArrDims(vArray) > 0)
End Function

Public Function GetArrDims(vArray As Variant) As Integer
    Dim ppSA As Long
    Dim pSA As Long
    Dim vt As Long
    Dim sa As SAFEARRAY
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, vt
            If vt And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                memcpy sa, ByVal pSA, LenB(sa)
                If sa.pvData <> 0 Then
                    GetArrDims = sa.cDims
                End If
            End If
        End If
    End If
End Function

Public Function Translit(ByVal sTxt As String) As String
    Dim sRussian As String
    Dim aTranslit()
    Dim lCount As Long
    
    sRussian = "אבגדהו¸זחטיךכלםמןנסעףפץצקרשתûü‎‏ÿ³¿÷"
    aTranslit = Array("", "a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "jj", "k", _
                      "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ch", _
                      "sh", "zch", "''", "'y", "'", "eh", "ju", "ja", "i", "i", "e")
    For lCount = 1 To Len(sRussian)
        sTxt = Replace$(sTxt, Mid$(sRussian, lCount, 1), aTranslit(lCount), , , vbTextCompare)
    Next
    
    Translit = sTxt
End Function
