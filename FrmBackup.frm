VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "www.MindWorkSoft.com"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   Icon            =   "FrmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Options"
      Height          =   1365
      Left            =   3180
      TabIndex        =   33
      Top             =   735
      Width           =   2580
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Left            =   2085
         Top             =   165
      End
      Begin VB.TextBox txtMin 
         Height          =   315
         Left            =   1395
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "30"
         Top             =   660
         Width           =   420
      End
      Begin VB.CheckBox chkAutoBackup 
         Caption         =   "Auto-backup "
         Height          =   315
         Left            =   225
         TabIndex        =   34
         Top             =   345
         Width           =   1245
      End
      Begin VB.Label lblElapse 
         Caption         =   "0"
         Height          =   225
         Left            =   795
         TabIndex        =   39
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Elapse"
         Height          =   225
         Left            =   225
         TabIndex        =   38
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label Label9 
         Caption         =   "Min."
         Height          =   225
         Left            =   1890
         TabIndex        =   37
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label8 
         Caption         =   "Backup every"
         Height          =   270
         Left            =   210
         TabIndex        =   36
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.PictureBox BoxRename 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   960
      ScaleHeight     =   1365
      ScaleWidth      =   3540
      TabIndex        =   22
      Top             =   3735
      Visible         =   0   'False
      Width           =   3540
      Begin VB.TextBox txtNewName 
         Height          =   345
         Left            =   1065
         TabIndex        =   26
         Top             =   465
         Width           =   2250
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   1830
         TabIndex        =   24
         Top             =   915
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   330
         Left            =   645
         TabIndex        =   23
         Top             =   915
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: do not put file extension"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   28
         Top             =   270
         Width           =   2235
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   3525
         X2              =   3525
         Y1              =   15
         Y2              =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   30
         X2              =   3525
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   15
         X2              =   3510
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "FILE RENAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   30
         TabIndex        =   27
         Top             =   45
         Width           =   3480
      End
      Begin VB.Label Label4 
         Caption         =   "New name?"
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   525
         Width           =   990
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Database Compact/Repair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   3165
      TabIndex        =   29
      Top             =   5010
      Width           =   2580
      Begin VB.CommandButton cmdCompactRepairDB 
         Caption         =   "Compact/Repair now"
         Height          =   390
         Left            =   285
         TabIndex        =   30
         Top             =   525
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Destination directory ----------->"
         Height          =   210
         Left            =   225
         TabIndex        =   31
         Top             =   255
         Width           =   2115
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      Picture         =   "FrmBackup.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   8940
      TabIndex        =   20
      Top             =   -15
      Width           =   8970
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   5865
         Top             =   -240
      End
      Begin VB.Label lblTimeDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6315
         TabIndex        =   32
         Top             =   30
         Width           =   2595
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MindWorkSoft 2004"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   7410
         TabIndex        =   21
         Top             =   435
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3180
      TabIndex        =   14
      Top             =   4170
      Width           =   2580
      Begin VB.CommandButton cmdRestore 
         Caption         =   "<-- Restore now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   15
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   3180
      TabIndex        =   11
      Top             =   2130
      Width           =   2580
      Begin VB.TextBox txtAppend 
         Height          =   330
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   1230
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Append text to filename"
         Height          =   315
         Left            =   255
         TabIndex        =   16
         Top             =   690
         Width           =   2070
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Backup now -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   13
         Top             =   1530
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Append date to filename"
         Height          =   315
         Left            =   255
         TabIndex        =   12
         Top             =   345
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "Append Text"
         Height          =   270
         Left            =   180
         TabIndex        =   18
         Top             =   1110
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7590
      TabIndex        =   8
      Top             =   5730
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Backup Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   5850
      TabIndex        =   4
      Top             =   720
      Width           =   2985
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   360
         Left            =   165
         TabIndex        =   10
         Top             =   4470
         Width           =   1140
      End
      Begin VB.FileListBox File2 
         Height          =   2040
         Left            =   135
         Pattern         =   "*.mdb"
         TabIndex        =   7
         Top             =   2400
         Width           =   2700
      End
      Begin VB.DirListBox Dir2 
         Height          =   1440
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   165
         TabIndex        =   5
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "List of backup files"
         Height          =   225
         Left            =   165
         TabIndex        =   9
         Top             =   2190
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   2985
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   165
         TabIndex        =   19
         Top             =   4920
         Width           =   1065
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   165
         TabIndex        =   3
         Top             =   330
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   135
         Pattern         =   "*.mdb"
         ReadOnly        =   0   'False
         System          =   -1  'True
         TabIndex        =   1
         Top             =   2460
         Width           =   2700
      End
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        Check3.Value = vbUnchecked
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = vbChecked Then
        Check1.Value = vbUnchecked
    End If
End Sub

Private Sub chkAutoBackup_Click()
    If chkAutoBackup.Value = vbChecked Then
        Timer2.Enabled = True
        Timer2.Interval = 60000
    Else
        Timer2.Enabled = False
        Counts = 0
    End If
End Sub

Private Sub cmdBackup_Click()
Dim SourceDir As String
Dim DestDir As String
Dim DatabaseFile As String
Dim DistanationFile As String
Dim Source As String
Dim Destination As String

'find out if there is file selected
If File1.FileName = "" Then
    MsgBox "please select file first", vbOKOnly, "No file selected"
    Exit Sub
End If



    DatabaseFile = File1.FileName
If Len(Dir1.Path) > 3 Then  'not a root directory
    SourceDir = Dir1.Path & "\"
    Source = SourceDir & DatabaseFile
Else
    SourceDir = Dir1.Path
    Source = SourceDir & DatabaseFile
End If

'if the want to append text to  filename
If Check3.Value = vbChecked Then
    DistanationFile = Left$(File1.FileName, Len(File1.FileName) - 4) & txtAppend.Text & Right$(File1.FileName, 4)
ElseIf Check1.Value = vbChecked Then
    DistanationFile = Left$(File1.FileName, Len(File1.FileName) - 4) & Format(Now, "dd-mm-yy-h-n") & Right$(File1.FileName, 4)
Else
    DistanationFile = DatabaseFile
End If



If Len(Dir2.Path) > 3 Then  'not a root directory
    DestDir = Dir2.Path & "\"
    Destination = DestDir & DistanationFile
Else
    DestDir = Dir2.Path
    Destination = DestDir & DistanationFile
End If

'start copying
    CopyFileWindowsWay Source, Destination
    File2.Refresh
End Sub

Private Sub cmdCancel_Click()
    BoxRename.Visible = False
End Sub

Private Sub cmdClose_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End
End Sub

Private Sub cmdCompactRepairDB_Click()
On Error GoTo ErrHandlerCompactRepair

Dim dbE As New DAO.DBEngine
Dim SourceDir As String
Dim DestDir As String
Dim DatabaseFile As String
Dim DistanationFile As String
Dim Source As String
Dim Destination As String

'find out if there is file selected
If File1.FileName = "" Then
    MsgBox "please select file first", vbOKOnly, "No file selected"
    Exit Sub
End If

    DatabaseFile = File1.FileName
If Len(Dir1.Path) > 3 Then  'not a root directory
    SourceDir = Dir1.Path & "\"
    Source = SourceDir & DatabaseFile
Else
    SourceDir = Dir1.Path
    Source = SourceDir & DatabaseFile
End If


'if the want to append text to  filename
    DistanationFile = Left$(File1.FileName, Len(File1.FileName) - 4) & "-REPAIRED" & Right$(File1.FileName, 4)

If Len(Dir2.Path) > 3 Then  'not a root directory
    DestDir = Dir2.Path & "\"
    Destination = DestDir & DistanationFile
Else
    DestDir = Dir2.Path
    Destination = DestDir & DistanationFile
End If

dbE.OpenDatabase Source, , , Password = "2003"

'Exit Sub
'start Compacting the datbase
    dbE.CompactDatabase Source, Destination  ' Compact and repair the DB
    File2.Refresh

Exit Sub
ErrHandlerCompactRepair:
MsgBox "Error: " & Err.Description, vbOKOnly, "Compact Error!!"
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandlerDelete
Dim CurrDir As String
Dim CurrFile As String

'ensure that file is selected
If File2.FileName = "" Then
    MsgBox "No file is selected", vbOKOnly, "Error"
    Exit Sub
End If

If Len(Dir2.Path) > 3 Then  'not a root directory
    CurrDir = Dir2.Path & "\"
    CurrFile = CurrDir & File2.FileName
Else
    CurrDir = Dir2.Path
    CurrFile = CurrDir & File2.FileName
End If

If MsgBox("Do you really delete file " & CurrFile & "?", vbYesNo, "Delete confirmation") = vbNo Then Exit Sub
    Kill CurrFile
    File2.Refresh
Exit Sub

ErrHandlerDelete:
    MsgBox "Error: " & Err.Description, vbOKOnly, "An error occured"
End Sub



Private Sub cmdOk_Click()
On erro GoTo ErrHandlerRename

    DatabaseFile = File1.FileName
If Len(Dir1.Path) > 3 Then  'not a root directory
    SourceDir = Dir1.Path & "\"
Else
    SourceDir = Dir1.Path
End If

    Source = SourceDir & DatabaseFile
    Name Source As SourceDir & txtNewName.Text & ".mdb"
    File1.Refresh
    BoxRename.Visible = False
Exit Sub
ErrHandlerRename:
MsgBox "Error: " & Err.Description, vbOKOnly, "File rename error!"
End Sub

Private Sub cmdRename_Click()
If File1.FileName = "" Then
    MsgBox "Please select file first to rename", vbOKOnly, "File not specified"
    Exit Sub
End If
    
    BoxRename.Visible = True
    txtNewName.SetFocus
End Sub

Private Sub cmdRestore_Click()
Dim RestoreSourceDir As String
Dim RestoreSourceFile As String
Dim RestoreSourse As String
Dim RestoreDistination As String

On Error GoTo ErrHandlerRestore

'find out if there is file selected
If File2.FileName = "" Then
    MsgBox "please select source file first", vbOKOnly, "No file selected"
    Exit Sub
End If

'combine directory and file for the source of the file to restore
RestoreSourceFile = File2.FileName
If Len(Dir2.Path) > 3 Then  'not a root directory
    RestoreSourceDir = Dir2.Path & "\"
Else
    RestoreSourceDir = Dir2.Path
End If
RestoreSourse = RestoreSourceDir & RestoreSourceFile

'find distination of file to restore
If Len(Dir1.Path) > 3 Then  'not a root directory
    RestoreDistination = Dir1.Path & "\" & RestoreSourceFile
Else
    RestoreDistination = Dir1.Path & RestoreSourceFile
End If

'start copying
CopyFileWindowsWay RestoreSourse, RestoreDistination
File1.Refresh

Exit Sub
ErrHandlerRestore:
MsgBox "Error: " & Err.Description, vbOKOnly, "Error restoring file"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
End Sub

Private Sub Dir2_Change()
File2.Path = Dir2.Path
File2.Refresh
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrHandlerDriveChange
Dir1.Path = Drive1.Drive
Dir1.Refresh
Exit Sub
ErrHandlerDriveChange:
MsgBox "Error: " & Err.Description, vbOKOnly, "Error"
End Sub

Private Sub Drive2_Change()
On Error GoTo ErrHandlerDriveChange
Dir2.Path = Drive2.Drive
Dir2.Refresh
Exit Sub
ErrHandlerDriveChange:
MsgBox "Error: " & Err.Description, vbOKOnly, "Error"

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
lblTimeDate.Caption = Now
End Sub

Private Sub Timer2_Timer()
    Counts = Counts + 1
    lblElapse.Caption = Counts
    
    If Counts >= Val(txtMin.Text) Then
        Counts = 0
        lblElapse.Caption = Counts
        cmdBackup_Click
    End If
End Sub
