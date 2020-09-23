Attribute VB_Name = "basSetFileDateTime"
Option Explicit

Private Const OFS_MAXPATHNAME = 128

'Types
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Declarations
'Open a file - get handle
Private Declare Function OpenFile& Lib "kernel32" _
    (ByVal lpFileName$, _
    lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle&)
'Close a handle
Private Declare Function CloseHandle& Lib "kernel32" _
    (ByVal hObject&)
'Read the file date/time stamp
Private Declare Function GetFileTime& Lib "kernel32" _
    (ByVal hFile&, _
    pCreationTime As FileTime, _
    lpLastAccessTime As FileTime, _
    lpLastWriteTime As FileTime)
'Set the file date/time stamp
Private Declare Function SetFileTime& Lib "kernel32" _
    (ByVal hFile&, _
    lpCreationTime As FileTime, _
    lpLastAccessTime As FileTime, _
    lpLastWriteTime As FileTime)
'Convert a UTC time file to time file
Private Declare Function FileTimeToLocalFileTime& Lib "kernel32" _
    (lpFileTime As FileTime, _
    lpLocalFileTime As FileTime)
'Convert a time to a UTC time
Private Declare Function LocalFileTimeToFileTime& Lib "kernel32" _
    (lpLocalFileTime As FileTime, _
    lpFileTime As FileTime)
'Convert a file time to a system time
Private Declare Function FileTimeToSystemTime& Lib "kernel32" _
    (lpFileTime As FileTime, _
    lpSystemTime As SYSTEMTIME)
'Convert a system time to a file time
Private Declare Function SystemTimeToFileTime& Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME, _
    lpFileTime As FileTime)
'Find first file
Private Declare Function FindFirstFile& Lib "kernel32" _
    Alias "FindFirstFileA" _
    (ByVal lpFileName$, _
    lpFindFileData As WIN32_FIND_DATA)
'Get file attributes (RASH)
Private Declare Function GetFileAttributes& Lib "kernel32" _
    Alias "GetFileAttributesA" _
    (ByVal lpFileName$)
'Set file attributes (RASH)
Private Declare Function SetFileAttributes& Lib "kernel32" _
    Alias "SetFileAttributesA" _
    (ByVal lpFileName$, _
    ByVal dwFileAttributes&)
    
'Constants
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4

Private Const OF_READ = &H0
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_WRITE = &H1
'________________________________________________________________________

Public Function SetFileDateTime&(ByVal FileSource$, NewFileDateTime$())
On Error GoTo Handler

Dim hFile&
Dim FileDateTimeStr$
Dim NewFileTime(0 To 2) As FileTime
Dim ofst As OFSTRUCT
Dim Attribs$
Dim IsAttribsChanged As Boolean
Dim rv&
Dim i%

    SetFileDateTime& = FAILURE&
    
    For i% = 0 To 2
        FileDateTimeStr$ = NewFileDateTime$(i%)
        If IsDate(FileDateTimeStr$) Then
            FileDateTimeStr$ = Format(FileDateTimeStr$, "dd-MMM-yyyy hh:nn:ss")
            ConvertFromStrToDate& FileDateTimeStr$, NewFileTime(i%)
            'We now have time successfully converted time for writing back to file.
        Else
            Err.Raise vbObjectError, "SetFileDateTime", _
                "<" & FileDateTimeStr$ & "> is not a valid date!"
        End If
    Next i%
    
    'We can't change file info. if the file is read-only:
    'We must temporarily change attributes to remove read-only flag.
    GetFileAttribs& FileSource$, Attribs$
    If InStr(1, Attribs$, "R") Then
        SetFileAttribs FileSource$, "N"
        IsAttribsChanged = True
    End If
    
    hFile& = OpenFile(FileSource$, ofst, OF_WRITE)
    rv& = SetFileTime(hFile&, NewFileTime(0), NewFileTime(2), NewFileTime(1))
    
    CloseHandle hFile&
    
    'Change attributes back to original settings.
    If IsAttribsChanged Then SetFileAttribs FileSource$, Attribs$
        
    If rv <> SUCCESS Then Exit Function
    
    SetFileDateTime& = SUCCESS&
     
ExitProc:
    Exit Function

Handler:
    SetFileDateTime& = Err.Number
    'MsgBox "Error: " & CStr(Err.Number) & " - " & _
        Err.Description, vbExclamation, "SetFileDateTime Error"
    Resume ExitProc

End Function
'________________________________________________________________________

Public Function GetFileDateTime&(ByVal FileSource$, FileTime$())
On Error GoTo Handler

Dim hFile&
Dim CreateTime As FileTime
Dim LastAccTime As FileTime
Dim LastModTime As FileTime
Dim ft As FileTime
Dim SysTime As SYSTEMTIME
Dim ofst As OFSTRUCT
Dim i%

    'Initialise variables
    GetFileDateTime& = FAILURE&
    
    'Open and read file properties
    hFile& = OpenFile(FileSource$, ofst, OF_READ Or OF_SHARE_DENY_NONE)
    GetFileTime hFile&, CreateTime, LastAccTime, LastModTime
    CloseHandle hFile&

    For i% = 0 To 2
        FileTime$(i%) = "01-Jan-1601 00:00:00"
        Select Case i%
        Case 0
            ft = CreateTime
        Case 1
            ft = LastModTime
        Case 2
            ft = LastAccTime
        End Select
        
        'Take the locale time into effect i.e. regional time settings.
        FileTimeToLocalFileTime ft, ft
        'Convert to a meaningful strucure.
        FileTimeToSystemTime ft, SysTime
        
        'Reformat as a contiguous string.
        FileTime$(i%) = Format(SysTime.wDay, "00") _
            & "-" & MonthStr(SysTime.wMonth) _
            & "-" & SysTime.wYear _
            & " " & Format(SysTime.wHour, "00") _
            & ":" & Format(SysTime.wMinute, "00") _
            & ":" & Format(SysTime.wSecond, "00")
    Next i%
    
    GetFileDateTime& = SUCCESS&

ExitProc:
    Exit Function

Handler:
    GetFileDateTime& = Err.Number
    'MsgBox "Error: " & CStr(Err.Number) & " - " & _
        Err.Description, vbExclamation, "GetFileDateTime Error"
    Resume ExitProc
    
End Function
'________________________________________________________________________

Private Function SetFileAttribs&(ByVal FileSource$, ByVal Attribs$)
On Error GoTo Handler
Dim NewAttribs&
Dim IsFile As Byte
Dim rv&

    SetFileAttribs& = FAILURE&
        
    If InStr(1, Attribs$, "N") Or Attribs$ = vbNullString Then
         NewAttribs& = FILE_ATTRIBUTE_NORMAL
    Else
        If InStr(1, Attribs$, "R") Then
            NewAttribs& = NewAttribs& + FILE_ATTRIBUTE_READONLY
        End If
        If InStr(1, Attribs$, "A") Then
            NewAttribs& = NewAttribs& + FILE_ATTRIBUTE_ARCHIVE
        End If
        If InStr(1, Attribs$, "S") Then
            NewAttribs& = NewAttribs& + FILE_ATTRIBUTE_SYSTEM
        End If
        If InStr(1, Attribs$, "Handler") Then
            NewAttribs& = NewAttribs& + FILE_ATTRIBUTE_HIDDEN
        End If
    End If
    
    rv& = SetFileAttributes(FileSource$, NewAttribs&)
    If rv& = 1 Then SetFileAttribs& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    SetFileAttribs& = Err.Number
    'MsgBox "Error: " & CStr(Err.Number) & " - " & _
        Err.Description, vbExclamation, "SetFileAttribs Error"
    Resume ExitProc
    
End Function
'________________________________________________________________________

Private Function GetFileAttribs&(ByVal FileSource$, Attribs$)
On Error GoTo Handler

Dim CurrentAttribs&
Dim IsFile As Byte
Dim s$
Dim rv&

    GetFileAttribs& = FAILURE&
    s$ = vbNullString
    
    CurrentAttribs& = GetFileAttributes(FileSource$)
    
    'Bitwise comparison
    If CurrentAttribs& And FILE_ATTRIBUTE_NORMAL Then
        s$ = "N"
    Else
        If CurrentAttribs& And FILE_ATTRIBUTE_READONLY Then s$ = "R"
        If CurrentAttribs& And FILE_ATTRIBUTE_ARCHIVE Then s$ = s$ & "A"
        If CurrentAttribs& And FILE_ATTRIBUTE_SYSTEM Then s$ = s$ & "S"
        If CurrentAttribs& And FILE_ATTRIBUTE_HIDDEN Then s$ = s$ & "H"
    End If
    
    Attribs$ = s$
    GetFileAttribs& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    GetFileAttribs& = Err.Number
    'MsgBox "Error: " & CStr(Err.Number) & " - " & _
        Err.Description, vbExclamation, "GetFileAttribs Error"
    Resume ExitProc
    
End Function
'________________________________________________________________________

Private Function ConvertFromStrToDate&(ByVal str As String, NewDateTime As FileTime)
On Error GoTo Handler

Dim sdt As SYSTEMTIME
Dim NewFileTime As FileTime
Dim rv&

    ConvertFromStrToDate& = FAILURE&
    
    sdt.wDay = DatePart("d", str)
    sdt.wMonth = DatePart("m", str)
    sdt.wYear = DatePart("yyyy", str)
    sdt.wHour = DatePart("h", str)
    sdt.wMinute = DatePart("n", str)
    sdt.wSecond = DatePart("s", str)
    
    'Convert to FileTime struct.
    rv& = SystemTimeToFileTime(sdt, NewFileTime)
    'Take the locale time into effect i.e. regional time settings.
    rv& = LocalFileTimeToFileTime(NewFileTime, NewFileTime)
    NewDateTime = NewFileTime
    
    ConvertFromStrToDate& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    ConvertFromStrToDate& = Err.Number
    'MsgBox "Error: " & CStr(Err.Number) & " - " & _
        Err.Description, vbExclamation, "ConvertFromStrToDate Error"
    Resume ExitProc
        
End Function
