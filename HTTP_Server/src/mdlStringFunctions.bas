Attribute VB_Name = "mdlStringFunctions"
Option Explicit

Public Enum IfStringNotFound
    ReturnOriginalStr = 0
    ReturnEmptyStr = 1
End Enum

' Search from end to beginning, and return the left side of the string
Public Function RightLeft(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStrRev(Str, RFind, , Compare)
    
    If K = 0 Then
        RightLeft = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightLeft = Left(Str, K - 1)
    End If
End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStrRev(Str, RFind, , Compare)
    
    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightRight = Mid(Str, K + 1, Len(Str))
    End If
End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStr(1, Str, LFind, Compare)
    If K = 0 Then
        LeftLeft = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        LeftLeft = Left(Str, K - 1)
    End If
End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStr(1, Str, LFind, Compare)
    If K = 0 Then
        LeftRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        LeftRight = Right(Str, (Len(Str) - Len(LFind)) - K + 1)
    End If
End Function

' Search from the beginning to end and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function LeftRange(ByRef Str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long, Q As Long
    
    K = InStr(1, Str, StrFrom, Compare)
    If K > 0 Then
        Q = InStr(K + Len(StrFrom), Str, StrTo, Compare)
        
        If Q > K Then
            LeftRange = Mid(Str, K + Len(StrFrom), (Q - K) - Len(StrFrom))
        Else
            LeftRange = IIf(RetError = ReturnOriginalStr, Str, "")
        End If
    Else
        LeftRange = IIf(RetError = ReturnOriginalStr, Str, "")
    End If
End Function

' Search from the end to beginning and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function RightRange(ByRef Str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long, Q As Long
    
    K = InStrRev(Str, StrTo, , Compare)
    If K > 0 Then
        Q = InStrRev(Str, StrFrom, K, Compare)
        
        If Q > 0 Then
            RightRange = Mid(Str, Q + Len(StrFrom), (K - Q) - Len(StrTo))
        Else
            RightRange = IIf(RetError = ReturnOriginalStr, Str, "")
        End If
    Else
        RightRange = IIf(RetError = ReturnOriginalStr, Str, "")
    End If
End Function

' SOUNDEX, used in SQL mostly, and dictionaries
' useful to find words that sound the same
Public Function SOUNDEX(Word As String) As String
    Dim K As Integer, PrevNum As Integer, Num As Integer, LLetter As String
    Dim SoundX As String
    
    For K = 2 To Len(Word)
        LLetter = LCase(Mid$(Word, K, 1))
        Select Case LLetter
        Case "b", "f", "p", "v"
            Num = 1
        Case "c", "e", "g", "j", "k", "q", "s", "x", "z"
            Num = 2
        Case "d", "t"
            Num = 3
        Case "l"
            Num = 4
        Case "m", "n"
            Num = 5
        Case "r"
            Num = 6
        Case "a", "e", "i", "o", "u"
            Num = 7
        End Select
        
        If PrevNum <> Num Then
            PrevNum = Num
            SoundX = SoundX & Num
        End If
    Next K
    
    SoundX = Replace(SoundX, "7", "", , , vbBinaryCompare)
    SOUNDEX = UCase(Left(Word, 1)) & Left(SoundX & "000", 3)
End Function

Public Function SOUNDEX2(Word As String, Optional StrLength As Integer = 8) As String
    Dim K As Integer, PrevNum As Integer, Num As Integer, LLetter As String
    Dim SoundX As String
    
    For K = 2 To Len(Word)
        LLetter = LCase(Mid(Word, K, 1))
        Select Case LLetter
        Case "b", "f", "p", "v"
            Num = 1
        Case "c", "e", "g", "j", "k", "q", "s", "x", "z"
            Num = 2
        Case "d", "t"
            Num = 3
        Case "l"
            Num = 4
        Case "m", "n"
            Num = 5
        Case "r"
            Num = 6
        Case "a", "e", "i", "o", "u"
            Num = 7
        End Select
        
        If PrevNum <> Num Then
            PrevNum = Num
            SoundX = SoundX & Num
        End If
    Next K
    
    SoundX = Replace(SoundX, "7", "", , , vbBinaryCompare)
    SOUNDEX2 = UCase(Left(Word, 1)) & Left(SoundX & String(StrLength - 1, "0"), StrLength - 1)
End Function
