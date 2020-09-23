VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataTransfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Transfer"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmDataTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox picBanner 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   5775
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   120
         Picture         =   "frmDataTransfer.frx":0442
         ScaleHeight     =   990
         ScaleWidth      =   1050
         TabIndex        =   4
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label lblInfo3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Start to begin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblInfo2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a database into a Text File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   3210
      End
      Begin VB.Label lblInfo1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This Application Transfers data from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   4365
      End
   End
   Begin MSComctlLib.ProgressBar pgTransfer 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdStartTransfer 
      Caption         =   "&Start"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox lstTables 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
   End
End
Attribute VB_Name = "frmDataTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuildFile(adRS As Recordset, sTable As String)
    
    On Error GoTo errBuildFile
    
    Dim lcv                                             As Long
    Dim sOutput                                         As String
    Dim iFreeHandle                                     As Integer
    Dim lCounter                                        As Long
    
    iFreeHandle = FreeFile
    
    pgTransfer.Value = 0
    pgTransfer.Max = adRS.RecordCount
    
    Do While Not adRS.EOF
            
            sOutput = "|"
            
            For lcv = 0 To adRS.Fields.Count - 1
               sOutput = sOutput & adRS.Fields(lcv) & "|"
            Next
                    
            lCounter = lCounter + 1
            pgTransfer.Value = lCounter
            
            Open "C:\" & sTable & ".txt" For Append As iFreeHandle

                Print #iFreeHandle, sOutput
                Close iFreeHandle
            
        adRS.MoveNext
    Loop
    
    Exit Sub
errBuildFile:
    errHandle Err.Number, Err.Description, Err.Source, "frmDataTransfer.BuildFile"
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdStartTransfer_Click()
    
    On Error GoTo errcmdStartTransfer_Click
    
    Dim sConn                                   As String
    Dim rst                                     As Recordset
    Dim sTableName                              As String
    Dim adRecordsetRS                           As Recordset
    Dim sSQL                                    As String
    Dim lcv                                     As Long
    
    Dim cn                                      As Connection
    
    
    Set cn = New Connection
    Set rst = New Recordset
    Set adRecordsetRS = New Recordset

    sConn = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Change"

    cn.Open sConn
    
    Set rst = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    
        Do While Not rst.EOF
            sTableName = rst("TABLE_NAME")
                lstTables.AddItem sTableName
            rst.MoveNext
        Loop
        
        DoEvents
        
        For lcv = 0 To lstTables.ListCount - 1
            lstTables.ListIndex = lcv
            sTableName = lstTables.Text
            
                If adRecordsetRS.State = adStateOpen Then
                    adRecordsetRS.Close
                End If
                
                sSQL = "Select * from [" & sTableName & "]"
                
                adRecordsetRS.Open sSQL, cn, adOpenStatic, adLockReadOnly, adCmdText
                    
                If Not adRecordsetRS.EOF Then
                    BuildFile adRecordsetRS, sTableName
                End If
        Next
        
    Set rst = Nothing
    Set cn = Nothing
    
    Exit Sub
errcmdStartTransfer_Click:
    errHandle Err.Number, Err.Description, Err.Source, "frmDataTransfer.cmdStartTransfer_Click"
End Sub
