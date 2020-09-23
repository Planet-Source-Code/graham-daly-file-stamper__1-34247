VERSION 5.00
Begin VB.Form frmSetFileDateTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Stamper"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmSetFileDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   495
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Close and Exit."
      Top             =   6450
      WhatsThisHelpID =   4
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Date && Time"
      Height          =   2190
      Left            =   300
      TabIndex        =   13
      Top             =   4425
      Width           =   5565
      Begin VB.CommandButton cmdAuthor 
         Height          =   435
         Left            =   5100
         Picture         =   "frmSetFileDateTime.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "About this program"
         Top             =   1725
         Width           =   435
      End
      Begin VB.CheckBox chkAdvanced 
         Caption         =   "Advanced"
         Height          =   240
         Left            =   1425
         TabIndex        =   10
         Top             =   1725
         Width           =   1290
      End
      Begin VB.TextBox txtNewDate 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1425
         TabIndex        =   8
         Text            =   "01/01/2000"
         Top             =   1275
         Width           =   990
      End
      Begin VB.TextBox txtNewTime 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2550
         TabIndex        =   9
         Text            =   "00:00:00"
         Top             =   1275
         Width           =   765
      End
      Begin VB.TextBox txtNewDate 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   6
         Text            =   "01/01/2000"
         Top             =   825
         Width           =   990
      End
      Begin VB.TextBox txtNewTime 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2550
         TabIndex        =   7
         Text            =   "00:00:00"
         Top             =   825
         Width           =   765
      End
      Begin VB.CommandButton cmdSetTime 
         Caption         =   "&Set Date/Time"
         Height          =   795
         Left            =   3525
         TabIndex        =   11
         Top             =   375
         WhatsThisHelpID =   3
         Width           =   1740
      End
      Begin VB.TextBox txtNewTime 
         Height          =   315
         Index           =   0
         Left            =   2550
         TabIndex        =   5
         Text            =   "00:00:00"
         Top             =   375
         Width           =   765
      End
      Begin VB.TextBox txtNewDate 
         Height          =   315
         Index           =   0
         Left            =   1425
         TabIndex        =   4
         Text            =   "01/01/2000"
         Top             =   375
         Width           =   990
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         Caption         =   "Last Accessed:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   15
         Top             =   900
         Width           =   990
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   450
         Width           =   600
      End
   End
   Begin VB.Frame fraFrom 
      Caption         =   "Select Files"
      Height          =   3840
      Left            =   300
      TabIndex        =   12
      Top             =   300
      Width           =   5565
      Begin VB.DriveListBox drvMain 
         Height          =   315
         Left            =   225
         TabIndex        =   0
         ToolTipText     =   "Select Drive to transfer from."
         Top             =   465
         WhatsThisHelpID =   7
         Width           =   2715
      End
      Begin VB.FileListBox filMain 
         Height          =   2625
         Hidden          =   -1  'True
         Left            =   3150
         MultiSelect     =   2  'Extended
         System          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Double-click to view file Date/Time info"
         Top             =   450
         WhatsThisHelpID =   9
         Width           =   2190
      End
      Begin VB.DirListBox dirMain 
         Height          =   2565
         Left            =   225
         TabIndex        =   1
         Top             =   975
         WhatsThisHelpID =   8
         Width           =   2715
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All Files"
         Height          =   300
         Left            =   3150
         TabIndex        =   3
         Top             =   3150
         WhatsThisHelpID =   2
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmSetFileDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'________________________________________________________________________

Private Sub chkAdvanced_Click()
Dim IsAdvanced As Boolean

    IsAdvanced = (Me.chkAdvanced.Value = vbChecked)
    
    Me.txtNewDate(1).Enabled = IsAdvanced
    Me.txtNewDate(2).Enabled = IsAdvanced
    Me.txtNewTime(1).Enabled = IsAdvanced
    Me.txtNewTime(2).Enabled = IsAdvanced

    If IsAdvanced Then
        With Me.txtNewDate
            .Item(1).BackColor = vbWindowBackground
            .Item(2).BackColor = vbWindowBackground
        End With
        With Me.txtNewTime
            .Item(1).BackColor = vbWindowBackground
            .Item(2).BackColor = vbWindowBackground
        End With
    Else
        With Me.txtNewDate
            .Item(1).Text = .Item(0).Text
            .Item(1).BackColor = vbButtonFace
            .Item(2).Text = .Item(0).Text
            .Item(2).BackColor = vbButtonFace
        End With
        With Me.txtNewTime
            .Item(1).Text = .Item(0).Text
            .Item(1).BackColor = vbButtonFace
            .Item(2).Text = .Item(0).Text
            .Item(2).BackColor = vbButtonFace
        End With
    End If
    
End Sub

Private Sub Form_Load()

    Me.drvMain.Drive = "C:\"
    Me.dirMain.Path = "C:\"
    Me.cmdExit.Top = Me.Height
    
End Sub
'________________________________________________________________________

Private Sub chkSelectAll_Click()
Dim i%

    With Me.filMain
        For i% = 0 To Me.filMain.ListCount - 1
            .Selected(i%) = (Me.chkSelectAll.Value = 1)
        Next i%
        If .ListCount > 0 Then
            .Selected(0) = True
        End If
    End With
    
End Sub
'________________________________________________________________________

Private Sub cmdSetTime_Click()
On Error GoTo Handler

Dim FilePath$, FileSource$
ReDim TempDateTime$(0 To 2)
Dim IsInvalidData As Boolean
Dim i%, CountFiles&, CountStampedFiles&

    Me.MousePointer = vbHourglass
    
    'Initialise counters.
    CountFiles& = 0
    CountStampedFiles& = 0
    
    'Simple validation.
    IsInvalidData = Not IsDate(Me.txtNewDate(0).Text)
    IsInvalidData = IsInvalidData Or (Not IsDate(Me.txtNewTime(0).Text))
    
    If Me.chkAdvanced.Value = vbChecked Then
        IsInvalidData = IsInvalidData Or (Not IsDate(Me.txtNewDate(1).Text))
        IsInvalidData = IsInvalidData Or (Not IsDate(Me.txtNewTime(1).Text))
        IsInvalidData = IsInvalidData Or (Not IsDate(Me.txtNewDate(2).Text))
        IsInvalidData = IsInvalidData Or (Not IsDate(Me.txtNewTime(2).Text))
    End If
    
    If IsInvalidData Then
        Err.Raise vbObjectError, "cmdSetTime_Click", "Please enter valid date/time stamp info. for the selected file(s)!"
    End If
    
    'Validation was successful!
    FilePath$ = Me.filMain.Path
    AddRemSlash FilePath$, 1
    
    'Populate array.
    For i% = 0 To 2
        TempDateTime$(i%) = Format(Me.txtNewDate(i%).Text, "dd-MMM-yyyy")
        TempDateTime$(i%) = TempDateTime$(i%) & " " & Format(Me.txtNewTime(i%).Text, "hh:nn:ss")
    Next i%
    
    'Here's where the action happens!
    For i% = 0 To Me.filMain.ListCount - 1
        If Me.filMain.Selected(i%) Then
            FileSource$ = FilePath$ & Me.filMain.List(i%)
            If SetFileDateTime(FileSource$, TempDateTime$()) = SUCCESS& Then
                CountStampedFiles& = CountStampedFiles& + 1
            End If
            CountFiles& = CountFiles& + 1
        End If
    Next i%
    
    'Notify user when done...
    MsgBox CStr(CountStampedFiles&) & " of " _
        & CStr(CountFiles&) & " files successfully stamped!!!", vbInformation, "Success"
    
ExitProc:
    Me.MousePointer = vbDefault
    Exit Sub

Handler:
    MsgBox Err.Description, vbExclamation, "Error: " & CStr(Err.Number)
    Resume ExitProc
    
End Sub
'________________________________________________________________________

Private Sub filMain_DblClick()
Dim FileDateTime$(0 To 2)
Dim FilePath$, FileSource$

    With Me.filMain
        If .List(.ListIndex) <> vbNullString Then
            FilePath$ = Me.filMain.Path
            AddRemSlash FilePath$, 1
            FileSource$ = FilePath$ & .FileName
            GetFileDateTime& FileSource$, FileDateTime$()
            MsgBox "Details for <" & FileSource$ & "> as follows..." & vbCrLf & vbCrLf _
                & "  Created on: " & FileDateTime$(0) & vbCrLf _
                & "  Last modified on: " & FileDateTime$(1) & vbCrLf _
                & "  Last accessed on: " & FileDateTime$(2) & vbCrLf, _
                vbInformation, "File Date/Time Info"
        End If
    End With
    
End Sub
'________________________________________________________________________

Private Sub cmdExit_Click()

    Unload Me
    
End Sub
'________________________________________________________________________

Private Sub drvMain_Change()
On Error GoTo Handler
   
Dim rv%

    dirMain.Path = Me.drvMain.Drive
    chkSelectAll_Click
 
ExitProc:
    Exit Sub

Handler:
    If Err.Description = "Device unavailable" Then
        rv% = MsgBox(Me.drvMain.Drive & "\ is not accessible" & vbCrLf & vbCrLf & "The device is not ready", vbCritical + vbRetryCancel, "Error")
            If rv% = vbCancel Then
                Me.drvMain.Drive = Left(Me.dirMain.Path, 3)
                Exit Sub
            Else
                Resume
            End If
    Else
        MsgBox "Error " & str(Err.Number) & ": " & Err.Description, vbCritical, "Error"
        Resume ExitProc
    End If
    
End Sub
'________________________________________________________________________

Private Sub dirMain_Change()

    Me.filMain.Path = Me.dirMain.Path
    chkSelectAll_Click
    
End Sub
'________________________________________________________________________

Private Sub filMain_PathChange()

    Me.chkSelectAll.Enabled = (Me.filMain.ListCount > 0)
    Me.cmdSetTime.Enabled = Me.chkSelectAll.Enabled
    
End Sub
'________________________________________________________________________

Private Sub txtNewDate_Change(Index As Integer)

    If Index = 0 And Me.chkAdvanced.Value = vbUnchecked Then
        Me.txtNewDate(1).Text = Me.txtNewDate(0).Text
        Me.txtNewDate(2).Text = Me.txtNewDate(0).Text
    End If
    
End Sub
'________________________________________________________________________

Private Sub txtNewTime_Change(Index As Integer)
    
    If Index = 0 And Me.chkAdvanced.Value = vbUnchecked Then
        Me.txtNewTime(1).Text = Me.txtNewTime(0).Text
        Me.txtNewTime(2).Text = Me.txtNewTime(0).Text
    End If
    
End Sub
'________________________________________________________________________

Private Sub cmdAuthor_Click()

    MsgBox "This program is freeware. Enjoy!" & vbCrLf & vbCrLf _
        & "Copyright Graham Daly 2001" & vbCrLf _
        & "Email: g.daly@iol.ie", vbInformation, "About the author..."
        
End Sub

