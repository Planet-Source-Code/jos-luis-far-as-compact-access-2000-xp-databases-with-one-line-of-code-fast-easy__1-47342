Attribute VB_Name = "ComprimirAccess2000"
'Compact Access Databases with one line of Code!
'"One line of Code", like this:  ;-)
'     CompactDB Database, ShowResume, Password
' or this example:
'     CompactDB App.Path & "\Datos.mdb", True, "Joselo.11·#"
'
'Private Sub CompactDB_Click()
'    If DB Is Nothing Then  'the database is Closed
'        CompactDB App.Path & "\Datos.mdb", True, "Joselo.11·#" 'compact
'    Else 'the database is open
'        DB.Close 'close
'        CompactDB App.Path & "\Datos.mdb", True, "Joselo.11·#" 'compact
'        [Call your Routine to load database] 'Open database
'    End If
'End Sub
'
'Works fine with Access 2000 and XP
'Actual Revision - Jan/20/2004
'Comments, sugestions, etc. are welcome.
'Written by José Luis Farías.
'Chile 1446 - Salto - Uruguay - CP 50.000
'JoseloFarias[at]adinet.com.uy
'¡¡¡Vamo' arriba Uruguay, carajo!!!
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'¡PLEASE!, if you use this Code sendme your Name and Country
'And if you like, emailme a program copy (source code if better)
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'
'
'            Jan/20/2004
' · some Form changes.
' · fixed some variables declarations. (Variants in place of Strings).
' · added Compact Password protected databases.
' · added some Comments.
' · added some Code examples.
' · added a compress level indicator.
'
'            Jul/25/2003
' · Initial Release

Public Ok As Boolean
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Function CompactDB(ByVal DataBase As String, Optional ShowResume As Boolean = True, Optional Password As String) As Boolean
' Assumes:Microsoft Jet and Replication objects 2.x Library (msjro.dll) reference.
'         Database is closed (see Sub CompactDB_Click() example on top).
'         If Database is password protected, the password passed are the correct.

    Dim JRO As New JRO.JetEngine
    Dim BD_Original As String, Dest_DB As String, TemporalFile As String, OriginalSize As String, FinalSize As String
    On Error GoTo CompactDB_Error
    DoEvents
        TemporalFile = App.Path & "\Temporal.mdb"
        If Password <> "" Then 'I try to use this as password to open the Database
            BD_Original = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBase & ";Jet OLEDB:" & "Database Password=" & Password & ";"
            Dest_DB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TemporalFile & ";Jet OLEDB:" & "Database Password=" & Password & ";"
        Else
            BD_Original = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBase & ";Jet OLEDB:Engine Type=5"
            Dest_DB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TemporalFile & ";Jet OLEDB:Engine Type=5"
        End If
        OriginalSize = GetFileSize(DataBase) 'get the original size
        If Dir(TemporalFile) <> "" Then Kill TemporalFile 'delete the temporal file if exist
        JRO.CompactDatabase BD_Original, Dest_DB 'Compact proccess
        FinalSize = GetFileSize(TemporalFile) 'get the final size
        If FinalSize > OriginalSize Then 'Compare the sizes
            Kill TemporalFile 'delete the temporal file
            MsgBox "The Final File if Bigger than the Original" 'Show result
            Exit Function
        End If
        If Dir(DataBase) <> "" Then Kill DataBase 'delete the source file
        Name TemporalFile As DataBase 'rename temporal as source
        Set JRO = Nothing
        CompactDB = True
        Ok = True
        If ShowResume Then 'in the other case, silent mode is enabled
            If OriginalSize = FinalSize Then 'Compare the sizes
                MsgBox "DataBase already Compresed", vbInformation, "Not Optimization Posible" 'Show result
            Else
                MsgBox "Original Size:   " & OriginalSize & vbCrLf & "Final Size:       " & FinalSize, vbInformation, "Optimization OK" 'Show result
            End If
        End If
        Exit Function
        
CompactDB_Error:
            CompactDB = False
            Ok = False
            Debug.Print Err.Number
            If Err.Number = -2147217843 Then 'Password Protected Database
'                MsgBox "Database Password protected." & vbCrLf & "Please enter the Correct Password"
            Else
                MsgBox Err.Description, vbExclamation
            End If
End Function
Public Function GetFileSize(File As Variant) As String
    On Error Resume Next
    Dim Bytes As Long
    Const Kb As Long = 1024
    Const Mb As Long = 1024 * Kb
    Const Gb As Long = 1024 * Mb
    Bytes = FileLen(File)
    If Bytes < Kb Then
        GetFileSize = Format(Bytes) & " bytes"
    ElseIf Bytes < Mb Then
        GetFileSize = Format(Bytes / Kb, "0.00") & " Kb"
    ElseIf Bytes < Gb Then
        GetFileSize = Format(Bytes / Mb, "0.00") & " Mb"
    Else
        GetFileSize = Format(Bytes / Gb, "0.00") & " Gb"
    End If
End Function
Private Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    strfileName.lpstrFilter = ""
    For intLoopCount = 1 To Len(WantedFilter)
        If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)
End Sub
Public Function OpenCommonDialog(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String
    Dim lngReturnValue As Long
    Dim intRest As Integer
    Dim i As Long
    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259)
    strfileName.nMaxFile = 260
    strfileName.flags = &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetOpenFileName(strfileName)
    strfileName.lpstrFile = Trim(strfileName.lpstrFile)
    i = Len(strfileName.lpstrFile)
    If i <> 1 Then
        OpenCommonDialog = Trim(strfileName.lpstrFile)
    Else
        OpenCommonDialog = ""
    End If
End Function
Private Function ShortPath(sPath As String, iMaxLen As Integer) As String
    Const DRIVE_LENGTH = 3         'Length of Drive, colon & slash in path
    Dim sLeft As String            'Left part of Path
    Dim sRight As String           'Right part of Path
    Dim iNextPos As Integer        'Position of Next "\"
    Dim iStart As Integer          'Position to start from
    If Len(sPath) <= iMaxLen Then
       ShortPath = sPath
       Exit Function
    End If
    iStart = DRIVE_LENGTH + 1                 'Start looking after Drive:\
    sLeft = Left$(sPath, DRIVE_LENGTH)        'Extract the drive from full path
    sRight = Right$(sPath, Len(sPath) - 3)    'Remove drive from right part
    Do While Len(sLeft & sRight) > iMaxLen    'Do until path shorter than Max Length
       iNextPos = InStr(iStart, sPath, "\")   'Find next "\" in path
       If iNextPos = 0 Then Exit Do           'Exit if no more "\" in path
       sLeft = sLeft & "...\"                 'Add another ...\ to short path
       sRight = Right$(sPath, Len(sPath) - iNextPos)   '
       iStart = iNextPos + 1
    Loop
    ShortPath = sLeft & sRight
End Function
