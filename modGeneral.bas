Attribute VB_Name = "modGeneral"
Option Explicit
Public Sub errHandle(eNumber As Long, eDiscription As String, eSource As String, eProcedure As String)
    
    On Error GoTo errerrHandle
    
    Dim iRes                                                            As Integer
    Dim iFilenum                                                        As Integer
    Dim sErrString                                                      As String
    Dim dShell                                                          As Variant
    
    iRes = MsgBox("Error Number " & eNumber & vbCrLf & "Error Discription " & eDiscription & vbCrLf & _
            "Error Source " & eSource & vbCrLf & "Error in procedure " & eProcedure, vbCritical, "Error Occurred")
                        
    iFilenum = FreeFile()
    
    sErrString = "---------------------------" & vbCrLf & Now & vbCrLf & "Error Number " & eNumber & vbCrLf & "Error Discription " & eDiscription & vbCrLf & _
            "Error Source " & eSource & vbCrLf & "Error in procedure " & eProcedure & vbCrLf & _
                 "---------------------------"
                      
    Open "c:\ErrLog.txt" For Append As #iFilenum   ' Open file for output.
        Print #iFilenum, sErrString;
    Close iFilenum
    
    Exit Sub
errerrHandle:
    Resume Next

End Sub


