Attribute VB_Name = "basPublicProcs"
Option Explicit

Public Const FAILURE& = 0
Public Const SUCCESS& = 1
'________________________________________________________________________

Public Function MonthStr(ByVal MonthAsNumber As Byte, _
    Optional IsFullName As Boolean = False) As String
    
ReDim r$(11)
Const TempShort$ = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
Const TempLong$ = "January,February,March,April,May,June,July,August,September,October,November,December"

    If MonthAsNumber < 1 Or MonthAsNumber > 12 Then
        Err.Raise vbObjectError, "MonthStr", "Invalid Month Number!"
    End If

    If IsFullName Then
        r$() = Split(TempLong$, ",", -1)
    Else
        r$() = Split(TempShort$, ",", -1)
    End If

    MonthStr = r$(MonthAsNumber - 1)

End Function
'________________________________________________________________________

Public Function AddRemSlash&(PathName$, ByVal IsSlash As Byte)
On Error GoTo Handler

Const BACKSLASH$ = "\"

    AddRemSlash& = FAILURE&
    
    If IsSlash Then       'We want a "\" at end
        If Right(PathName$, 1) <> BACKSLASH$ Then PathName$ = PathName$ & BACKSLASH$
        
    Else                        'We don't want a "\" at end
        If Right(PathName$, 1) = BACKSLASH$ Then
            PathName$ = Mid(PathName$, 1, Len(PathName$) - 1)
        End If
    End If
    
    AddRemSlash& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    AddRemSlash& = Err.Number
    'MsgBox Err.Number & ": " & Err.Description, _
            vbExclamation, "AddRemSlash Error"
    Resume ExitProc
    
End Function



