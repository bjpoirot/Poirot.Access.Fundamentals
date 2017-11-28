Option Compare Database
Option Explicit

' ************************************************************************
'
'   Poirot Software Engineering 2012
'   Auteur: B.J. Poirot
'
'   wijzigingen:
'   10-8-2012   - Initiele versie
'
'   1-9-2014 - aanpassing aan switchen van database
' ************************************************************************

Const ODBC_ADD_DSN As Long = 1
Const MAX_BUFFER_SIZE As Long = 1024
Const ODBCDriverDescription As String = "SQL Server"
'Contains all info for tables
Private mastODBCInfo() As tODBCInfo
Private Type tODBCInfo
    strTableName As String
    strNewName As String
    strConnectString As String
    strSourceTable As String
End Type
Declare Function SQLConfigDataSource Lib "odbccp32.dll" _
    (ByVal hwndParent As Long, _
    ByVal fRequest As Integer, _
    ByVal lpszDriver As String, _
    ByVal lpszAttributes As String) As Long

'--Windows API for screen metrics---
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd&, _
    ByVal hWndInsertAfter&, _
    ByVal X&, ByVal Y&, ByVal cX&, _
    ByVal cY&, ByVal wFlags&)
Global Const SWP_NOZORDER = &H4 'Ignores the hWndInsertAfter.
Global Const HWND_TOP = 0 'Moves MS Access window to top of Z-order.

' Set Window Position and Size
Function SizeAccess(cX As Long, cY As Long, cWidth As Long, cHeight As Long)

    Dim h As Long
    'Get handle to Microsoft Access.
    h = Application.hWndAccessApp
    
    'Position Microsoft Access.
    SetWindowPos h, 0, cX, cY, cWidth, cHeight, SWP_NOZORDER
    Application.RefreshDatabaseWindow
End Function

' DSN verbinding maken
Function PrepareDSN(ByVal strServerName As String, ByVal strDBName As String, _
                    ByVal strDSN As String) As Boolean
                    
On Error GoTo error_hdl
    Dim boolError As Boolean
    Dim strDSNString As String
    PrepareDSN = False

    
    strDSNString = Space(MAX_BUFFER_SIZE)
    strDSNString = ""
    strDSNString = strDSNString & "DSN=" & strDSN & Chr(0)
    strDSNString = strDSNString & "DESCRIPTION=" & "DSN Created Dynamically On " & CStr(Now) & Chr(0)
    strDSNString = strDSNString & "Server=" & strServerName & Chr(0)
    strDSNString = strDSNString & "Database=" & strDBName & Chr(0)
    strDSNString = strDSNString & "Trusted_Connection=Yes" & Chr(0)
    strDSNString = strDSNString & Chr(0)

   If Not CBool(SQLConfigDataSource(0, _
                        ODBC_ADD_DSN, _
                        ODBCDriverDescription, _
                        strDSNString)) Then
        boolError = True
        MsgBox ("Error in PrepareDSN::SQLConfigDataSource")
   End If

    If boolError Then
        Exit Function
    End If
    PrepareDSN = True
    Exit Function

error_hdl:
    MsgBox "PrepareDSN_ErrHandler::" & Err.Description
End Function


' BJP: function from http://access.mvps.org/access/tables/tbl0010.htm
'
' Functie welke alle linked tables met zelfde DSN opnieuw koppelt naar (andere) database
'
Function ReconnectLinkedTables(pDb As Database, pDSNname As String, pOldDatabaseName As String, pNewDatabaseName As String) As Boolean
    Dim db As Database, tdf As TableDef
    Dim varRet As Variant
    Dim strConnect As String
    Dim intTableCount As Integer
    Dim i As Integer
    Dim strTmp As String, strMsg As String
    Dim boolTablesPresent As Boolean
 
    On Error GoTo fReconnectODBC_Err
 
    If True Then
        Set db = pDb
        intTableCount = 0
        varRet = SysCmd(acSysCmdSetStatus, "Storing ODBC link info.....")
    
        boolTablesPresent = False
        For Each tdf In db.TableDefs
            strConnect = tdf.Connect
            If Len(strConnect) > 0 And Left$(tdf.Name, 1) <> "~" Then
                If Left$(strConnect, 4) = "ODBC" Then
                    If Nz(InStr(1, UCase(tdf.Connect), "DSN=" + UCase(pDSNname) + ";", vbTextCompare), 0) > 0 Then
                        ReDim Preserve mastODBCInfo(intTableCount)
                        With mastODBCInfo(intTableCount)
                            .strTableName = tdf.Name
                            .strSourceTable = tdf.SourceTableName
                            .strConnectString = tdf.Connect
                        End With
                        boolTablesPresent = True
                        intTableCount = intTableCount + 1
                    End If
                End If
            End If
        Next
        
        'now attempt relink
        If Not boolTablesPresent Then
            'No ODBC Tables present yet
            'Reconnect from the table info
            strMsg = "No ODBC tables were found in this database." & vbCrLf _
                            & "Do you wish to reconnect to all the ODBC sources ?"
            MsgBox strMsg, vbOKOnly, "ODBC Tables not present"
        Else
            For i = 0 To intTableCount - 1
                With mastODBCInfo(i)
                    varRet = SysCmd(acSysCmdSetStatus, "Attempting to relink '" _
                                        & .strTableName & "'.....")
                    strTmp = Format(Now(), "MMDDYY-hhmmss")
                    
                    db.TableDefs(.strTableName).Name = .strTableName & strTmp
                    'db.TableDefs(.strTableName).Connect = ChangeConnectionStringDB(.strConnectString, pOldDatabaseName, pNewDatabaseName)
                    db.TableDefs.Refresh
                    .strConnectString = ChangeConnectionStringDB(.strConnectString, pOldDatabaseName, pNewDatabaseName)
                    Set tdf = db.CreateTableDef(.strTableName, _
                                                                dbAttachSavePWD, _
                                                                .strSourceTable, _
                                                                .strConnectString)
                    
                    db.TableDefs.Append tdf
                    db.TableDefs.Delete .strTableName & strTmp
                End With
            Next
            db.TableDefs.Refresh
        End If
    End If
    varRet = SysCmd(acSysCmdClearStatus)
    ReconnectLinkedTables = True
  '  MsgBox "All ODBC tables were successfully reconnected.", _
   '                 vbInformation + vbOKOnly, "Success"

fReconnectODBC_Exit:
    Set tdf = Nothing
    Set db = Nothing
   ' Erase mastODBCInfo
Exit Function
fReconnectODBC_Err:
    Dim errX As Error

    If Errors.Count > 1 Then
        For Each errX In Errors
            strMsg = strMsg & "Error #: " & errX.Number & vbCrLf & errX.Description
        Next
        MsgBox strMsg, vbOKOnly + vbExclamation, "ODBC Errors in reconnect"
    Else
        strMsg = "Error #: " & Err.Number & vbCrLf & Err.Description
        MsgBox strMsg, vbOKOnly + vbExclamation, "VBA Errors in reconnect"
    End If
    ReconnectLinkedTables = False

    Resume fReconnectODBC_Exit
End Function



' Wijzig de database in een connectionstring
Public Function ChangeConnectionStringDB(pConnectionString As String, _
                                         pOldDatabase As String, _
                                         pNewDatabase As String) As String
    

    If Nz(InStr(1, UCase(pConnectionString), "DATABASE=" + UCase(pOldDatabase) + ";", vbTextCompare), 0) > 0 Then
            
        ChangeConnectionStringDB = Replace(pConnectionString, "DATABASE=" + UCase(pOldDatabase) + ";", "DATABASE=" + pNewDatabase + ";", , , vbTextCompare)
                
    ElseIf Nz(InStr(1, UCase(pConnectionString), "DATABASE=" + UCase(pNewDatabase) + ";", vbTextCompare), 0) > 0 Then
            
        ChangeConnectionStringDB = pConnectionString
        
    Else
        
        'when databsename is at last position the ; is not always used
        If Right(UCase(pConnectionString), Len(pOldDatabase)) = UCase(pOldDatabase) Then
                
            ChangeConnectionStringDB = Replace(pConnectionString, "DATABASE=" + UCase(pOldDatabase), "DATABASE=" + pNewDatabase, , , vbTextCompare)
                    
        ElseIf Right(UCase(pConnectionString), Len(pOldDatabase)) = UCase(pOldDatabase) Then
                
            ChangeConnectionStringDB = pConnectionString
            
        Else
        
            MsgBox "Database name not found in connectionstring", vbOKOnly, "Change Connectionstring"
            ChangeConnectionStringDB = pConnectionString
        
        End If
    
    End If


    
End Function



Function CheckLinkedTables(pDb As Database, pDSNname As String, pOldDatabaseName As String) As Boolean
    Dim db As Database, tdf As TableDef
    Dim varRet As Variant
    Dim strConnect As String
    Dim strTmp As String, strMsg As String
 
    On Error GoTo fReconnectODBC_Err
 
    If True Then
        Set db = pDb
        varRet = SysCmd(acSysCmdSetStatus, "Storing ODBC link info.....")
    
        For Each tdf In db.TableDefs
            strConnect = tdf.Connect
            If Len(strConnect) > 0 And Left$(tdf.Name, 1) <> "~" Then
                If Left$(strConnect, 4) = "ODBC" Then
                    If Nz(InStr(1, UCase(tdf.Connect), "DSN=" + UCase(pDSNname) + ";", vbTextCompare), 0) > 0 Then
                        If Nz(InStr(1, UCase(tdf.Connect), "DATABASE=" + UCase(pOldDatabaseName) + ";", vbTextCompare), 0) > 0 Then
                            
                            MsgBox tdf.Name + " still connected to old database, please reconnect manually."
                        
                        End If
                    End If
                End If
            End If
        Next
      
    End If
    varRet = SysCmd(acSysCmdClearStatus)
    CheckLinkedTables = True
 
fReconnectODBC_Exit:
    Set tdf = Nothing
    Set db = Nothing
   ' Erase mastODBCInfo
Exit Function
fReconnectODBC_Err:
    Dim errX As Error

    If Errors.Count > 1 Then
        For Each errX In Errors
            strMsg = strMsg & "Error #: " & errX.Number & vbCrLf & errX.Description
        Next
        MsgBox strMsg, vbOKOnly + vbExclamation, "ODBC Errors in reconnect"
    Else
        strMsg = "Error #: " & Err.Number & vbCrLf & Err.Description
        MsgBox strMsg, vbOKOnly + vbExclamation, "VBA Errors in reconnect"
    End If
    CheckLinkedTables = False

    Resume fReconnectODBC_Exit

End Function





' BJP: functie van http://www.access-programmers.co.uk/forums/showthread.php?t=99179
' Deze exporteert alle objecten naar textfiles
Public Sub ExportDatabaseObjects(pDatabase As Database, pSubFolder As String)
On Error Resume Next ' GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    'Dim db As DAO.Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = pDatabase
    
    sExportLocation = Application.CurrentProject.Path & "\" & pSubFolder & "\" 'Do not forget the closing back slash! ie: C:\Temp\
    
    FileSystem.MkDir sExportLocation
    
   ' tables exporteren is niet wenselijk omdat ook alle data wordt geëxporteerd.
   ' For Each td In db.TableDefs 'Tables
   '     If Left(td.Name, 4) <> "MSys" Then
   '          DoCmd.TransferText acExportDelim, , td.Name, sExportLocation & "Table_" & td.Name & ".txt", True
   '     End If
   ' Next td
    
   
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.Name, sExportLocation & "Form_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.Name, sExportLocation & "Report_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.Name, sExportLocation & "Macro_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.Name, sExportLocation & "Module_" & d.Name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
  '  MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportReport
' Author    : B.J. Poirot
' Date      : 18-9-2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ExportReport(pObjectType As AcObjectType, pDb As Database, pReportName As String)

    Dim db As Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = pDb
    
    sExportLocation = Application.CurrentProject.Path & "\" 'Do not forget the closing back slash! ie: C:\Temp\
    
    Application.SaveAsText pObjectType, pReportName, sExportLocation & "Report_" & pReportName & ".txt"
    
End Sub


' Procedure om tabellen te maken uit een ADO recorset.
' ( codevoorbeeld van internet aangepast)
Sub wMakeTableFromADO_Recordset(pRecordSet As ADODB.Recordset, pNewTableName As String)
    
    Dim con As ADODB.Connection
    Dim DAO_Type As Integer
    Dim db As Database
    Dim fld As DAO.Field
    Dim FldNo As Integer
    Dim fName As String
    Dim fSize As Long
    Dim fType As Integer
   
    Dim strQuery As String
    Dim tdfNew As TableDef
    Dim X As Long
    
    Set db = CurrentDb() ' This is just to get at the TableDefs
    
    Set con = CurrentProject.Connection
  '  Set rst.ActiveConnection = con
  '  rst.CursorType = adOpenKeyset
  '  rst.CursorLocation = adUseClient
  '  rst.LockType = adLockReadOnly
  '  rst.Open "Select * from DataTypesSamples" ' Just a test case to get some data into rst
    
    
    ' Create the new table structure & define all fields from the recordset properties
    On Error Resume Next
    DoCmd.Close acTable, pNewTableName
    DoCmd.DeleteObject AcObjectType.acTable, pNewTableName
    On Error GoTo 0
    
    Set tdfNew = db.CreateTableDef(pNewTableName) ' Create new TableDef object.
    
    Debug.Print
    
    For FldNo = 0 To pRecordSet.Fields.Count - 1
        fName = pRecordSet.Fields(FldNo).Name
        fSize = pRecordSet.Fields(FldNo).DefinedSize
        fType = pRecordSet.Fields(FldNo).Type
        
        If fType = 202 And fSize > 255 Then
            fSize = 255
        End If
        
        ' When creating the new fields in TableDef need to crossref
        ' the recordset's ADO field type with the DAO type
        
        '?? Note, we're not carrying over field Attributes like dbHyperlinkField
        '?? Also, not accounting for OLE Objects
        
        Select Case fType ' Convert the ADO type to DAO type
        Case adInteger: DAO_Type = 4 ' dbLong
        Case adUnsignedTinyInt: DAO_Type = 2 ' dbByte
        Case adCurrency: DAO_Type = 5 ' dbCurrency
        Case adDate: DAO_Type = 8 ' dbDate
        Case adNumeric: DAO_Type = 20 ' dbDecimal
        Case adSmallInt: DAO_Type = 3 ' dbInteger
        Case adSingle: DAO_Type = 6 ' dbSingle
        Case adDouble: DAO_Type = 7 ' dbDouble
        Case adLongVarWChar: DAO_Type = 12 ' dbMemo & Hyperlink
        Case adVarWChar: DAO_Type = 10 ' dbText
        Case adBoolean: DAO_Type = 1 ' dbBoolean
        End Select
    
        Debug.Print fName, fType, DAO_Type, fSize ' show the field info
    
        If DAO_Type <> 20 Then
            '?? it doesn't like dbDecimal - Run-time error '3259': "Invalid field data type"
            Set fld = tdfNew.CreateField(fName, DAO_Type, fSize)
            tdfNew.Fields.Append fld
        End If
        
    Next
    
    db.TableDefs.Append tdfNew ' Add this new table structure into TableDefs
    Set fld = Nothing
    Set tdfNew = Nothing
    Set db = Nothing
    
    Dim cntRecords As Integer
    
    If pRecordSet.RecordCount = -1 Then
        pRecordSet.MoveFirst
        Do While Not pRecordSet.EOF
            cntRecords = cntRecords + 1
            pRecordSet.MoveNext
        Loop
        pRecordSet.MoveFirst
    Else
        cntRecords = pRecordSet.RecordCount
    End If
    
    
    ' Load the recordset's data into new table
    For X = 1 To cntRecords
        ' loop on recordset getting each record and building up
        ' an INSERT statement for each. Seems like hard way.
        strQuery = "Insert Into " & pNewTableName & " Values("
        
        For FldNo = 0 To pRecordSet.Fields.Count - 1
            If IsNull(pRecordSet.Fields(FldNo).Value) Then
                strQuery = strQuery & "NULL,"
            Else
                Select Case pRecordSet.Fields(FldNo).Type
                Case adNumeric ' skip this (decimal), since we skipped it above
                    'strQuery = strQuery & pRecordSet.Fields(FldNo).Value & ","
                Case adVarWChar, adDate, adLongVarWChar ' Surround these with single quotes
                    
                    If Len(pRecordSet.Fields(FldNo).Value) > 255 Then
                        strQuery = strQuery & "'" & Left(pRecordSet.Fields(FldNo).Value, 255) & "',"
                    Else
                        If Len(pRecordSet.Fields(FldNo).Value) = 0 Then
                            strQuery = strQuery & "NULL,"
                        Else
                            strQuery = strQuery & "'" & pRecordSet.Fields(FldNo).Value & "',"
                        End If
                    End If
                Case Else
                    strQuery = strQuery & Replace(pRecordSet.Fields(FldNo).Value, ",", ".") & ","
                End Select
            End If
        Next
        
        ' Remove last comma and add ")" to close the Values() clause
        strQuery = Left$(strQuery, Len(strQuery) - 1) & ")"
        Debug.Print strQuery
        con.Execute strQuery ' add the record
        pRecordSet.MoveNext
    Next
    

End Sub ' wMakeTableFromADO_Recordset

'' Validate email address
Public Function ValidateEmailAddress(ByVal strEmailAddress As String) As Boolean
    On Error GoTo Catch
       
    Dim objRegExp As New RegExp
    Dim blnIsValidEmail As Boolean
    
    strEmailAddress = Trim(strEmailAddress)
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    blnIsValidEmail = objRegExp.Test(strEmailAddress)
    ValidateEmailAddress = blnIsValidEmail
    
    Exit Function
    
Catch:
    
    ValidateEmailAddress = False
    MsgBox "Module: wsbBasis - ValidateEmailAddress function" & vbCrLf & vbCrLf _
        & "Error#:  " & Err.Number & vbCrLf & vbCrLf & Err.Description

End Function


' Deze wordt direct op de SQL-server uitgevoerd
' LET OP!!! Er dient een view -object van type Pass-through te zij met de naam QRY_PASS_THROUGH
Public Sub CreateExcelFromSelect(pSelectStatement As String, pConnectionsting As String, pFilename As String)
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim qd As DAO.QueryDef
    Dim SqlStatement As String
    
    SqlStatement = pSelectStatement
    
    Set db = CurrentDb
    Set qd = db.QueryDefs("QRY_PASS_THROUGH")
    qd.Connect = pConnectionsting
    qd.SQL = SqlStatement
    
    Set rs = qd.OpenRecordset
       
    CreateExcelFromRecordset rs, pFilename
        
End Sub

Public Sub CreateExcelFromRecordset(pRS As DAO.Recordset, pFilename As String)
    
    Dim aExcel As Object
    Dim aBook As Object
    Dim aSheet As Object
    Dim aField As DAO.Field
    Dim r As Integer
    Dim f As Integer
    r = 1
    f = 1
    
    Set aExcel = CreateObject("Excel.Application")
    Set aBook = aExcel.Workbooks.Add
    Set aSheet = aBook.Worksheets(1)
    
    If Not pRS.EOF Then
              
        For Each aField In pRS.Fields
                  
            aSheet.Range("A1").Cells(r, f).Formula = aField.Name
                  
           If f < 255 Then
                f = f + 1
            End If
        Next
    End If
    
    f = 1
    
    Do While Not pRS.EOF
  
        For Each aField In pRS.Fields
                   
            aSheet.Range("A2").Cells(r, f).Formula = aField
                  
            ' Datumveld in Excel formateren
            If aField.Type = 8 Then
                aSheet.Range("A2").Cells(r, f).NumberFormat = "m/d/yyyy"
            End If
                  
           If f < 255 Then
                f = f + 1
            End If
        Next
        f = 1
        r = r + 1
        pRS.MoveNext
    Loop
   
    aExcel.Visible = True
    
End Sub


'********************************************************
' voegt een nieuw record in en blijft op huidige positie
'********************************************************
Public Sub InsertCloneRecord(pKeyField As Variant, pKeyValue As Variant, _
                              pCurrentFields As Fields, pCurrentForm As Form)
    
    Dim aRS As DAO.Recordset
    Dim i As Integer
    
    Set aRS = pCurrentForm.RecordsetClone
   
    With aRS
        
        .AddNew
         
         For i = 0 To pCurrentFields.Count - 1
            If .Fields(i).Name <> pKeyField And .Fields(i).DataUpdatable And .Fields(i).Attributes <> 33 Then
                .Fields(i).Value = pCurrentFields(i)
            End If
         Next
      
         aRS.Update
         aRS.Requery
         
         aRS.FindFirst pKeyField & " = " & pKeyValue
         
    End With

    pCurrentForm.Bookmark = aRS.Bookmark

End Sub

'-------------------------------------------------------------------
' Geef laatste 2 cijfers van jaar met weeknummer terug van een datum
' volgens nederlands model
'-------------------------------------------------------------------
Public Function GetYrWeek(Optional pDate)
    Dim jaar, weeknr As String
    
    If IsDate(pDate) Then
        ' jaar bepalen
        jaar = Format(pDate, "yy")
        ' weeknr bepalen & ev. aanvullen tot 2 characters
        weeknr = Format(pDate, "ww", vbMonday, vbFirstFourDays)
        weeknr = IIf(Len(weeknr) = 2, weeknr, "0" + weeknr)
        ' als einde van het jaar is komt het soms voor dat
        ' weeknummer 1 wordt teruggegeven van oude jaar ipv nieuwe
        If Month(pDate) = 12 And weeknr = "01" Then
            jaar = CStr(jaar + 1)
        End If
        If Month(pDate) = 1 And (weeknr = "52" Or weeknr = "53") Then
            jaar = CStr(jaar - 1)
        End If
        ' resultaat teruggeven
        GetYrWeek = jaar + weeknr
    Else
        GetYrWeek = ""
    End If
    
End Function

'-------------------------------------------------------------------
' Geef laatste 2 cijfers van jaar met weeknummer terug van een datum
' volgens nederlands model en (blank) als geen datum is ingevuld
'-------------------------------------------------------------------
Public Function GetYrWeekWithBlank(Optional pDate, Optional pReplyString As String)
    Dim jaar, weeknr As String
    
    If IsDate(pDate) Then
        ' jaar bepalen
        jaar = Format(pDate, "yy")
        ' weeknr bepalen & ev. aanvullen tot 2 characters
        weeknr = Format(pDate, "ww", vbMonday, vbFirstFourDays)
        weeknr = IIf(Len(weeknr) = 2, weeknr, "0" + weeknr)
        ' resultaat teruggeven
        GetYrWeekWithBlank = jaar + weeknr
    Else
        GetYrWeekWithBlank = pReplyString
    End If
    
End Function
'-------------------------------------------------------------------
' Geef laatste 2 cijfers van jaar met weeknummer terug van een datum
' volgens nederlands model
'-------------------------------------------------------------------
Public Function GetValueWithBlank(Optional pValue, Optional pReplyString As String)
    If Not IsMissing(pValue) Then
        ' resultaat teruggeven
        GetValueWithBlank = pValue
        If pValue = "" Or IsNull(pValue) Then
            GetValueWithBlank = pReplyString
        End If
    Else
        GetValueWithBlank = pReplyString
    End If
    
End Function



' Calculate the number of workdays between two dates
' BJP 15-2-2013 Refactored and Holidays added
Public Function CalcWorkDaysBetween(ByVal dt1 As Date, ByVal dt2 As Date, ByVal blnInclDay1 As Boolean, ByVal blnInclDay2 As Boolean) As Integer
    'Date 2 is bigger than date 1
    'Date 1 and 2 are both included in the nr of days
    'This procedure calcultates the number of working days between two dates.
    'Hollidays are taken into account if table tbl_Holiday is filled
    Dim nDay1, nDay2, nWeeks, nHolidays, nWorkdays As Integer
    Dim nRest As Double
    Dim rsHolidays As DAO.Recordset
        
    
    If dt1 > dt2 Then
        CalcWorkDaysBetween = 0
    Else
    
        ' First determine number of holidays
        Set rsHolidays = CurrentDb.OpenRecordset( _
             "SELECT Holiday FROM tbl_Holidays WHERE (((tbl_Holidays.Holiday)>=#" _
                + Format(dt1, "mm/dd/yyyy", vbMonday) + "# And (tbl_Holidays.Holiday)<=#" _
                + Format(dt2, "mm/dd/yyyy", vbMonday) + "#))")
        
        If rsHolidays.RecordCount > 0 Then
            rsHolidays.MoveFirst
            Do
                If DatePart("w", rsHolidays.Fields(0).Value, vbMonday, vbFirstFourDays) < 6 Then
                
                If rsHolidays.Fields(0).Value > dt1 And rsHolidays.Fields(0).Value < dt2 Then
                
                    nHolidays = nHolidays + 1
                End If
                If blnInclDay1 And rsHolidays.Fields(0).Value = dt1 Then
                    nHolidays = nHolidays + 1
                End If
                If blnInclDay2 And rsHolidays.Fields(0).Value = dt2 Then
                    nHolidays = nHolidays + 1
                End If
                
                End If
                rsHolidays.MoveNext
                
            Loop While rsHolidays.EOF = False
        End If
        
        ' Determine number of possible workdays
        nWeeks = DateDiff("ww", dt1, dt2, vbMonday, vbFirstFourDays)
        nWorkdays = nWeeks * 5
            
        ' first week and last week correction
        nDay1 = DatePart("w", dt1, vbMonday, vbFirstFourDays) 'day of the week, 1=monday, 2=tuesday, etc.
        nDay2 = DatePart("w", dt2, vbMonday, vbFirstFourDays)
        nWorkdays = nWorkdays - IIf(nDay1 + 1 > 5, 5, nDay1)
        nWorkdays = nWorkdays + IIf(nDay2 > 5, 5, nDay2 - 1)
                
        ' If first and last date are included
        If blnInclDay1 And nDay1 < 6 Then nRest = nRest + 1 'add start day
        If blnInclDay2 And nDay2 < 6 Then nRest = nRest + 1 'add end day
        
        
        ' Possible workdays minus holidays
        If nHolidays < nWorkdays + nRest Then
            CalcWorkDaysBetween = nWorkdays + nRest - nHolidays
        Else
            CalcWorkDaysBetween = 0
        End If
    End If

End Function

' Calculates the number of weeks between two dates
Public Function CalcWeeksBetween(ByVal dt1 As Date, ByVal dt2 As Date) As Integer
    'Date 2 can be bigger or smaller than date 1
    'If dt2 is bigger the result is positive; otherwise it is negative.
    'This procedure calcultates the number of weeks between two dates.
    'Hollidays are not taken into account
    Dim nDay1, nDay2 As Integer, nWeeks As Integer
    Dim nRest As Double
    
    nDay1 = DatePart("w", dt1, vbMonday, vbFirstFourDays) 'day of the week, 1=monday, 2=tuesday, etc.
    nDay2 = DatePart("w", dt2, vbMonday, vbFirstFourDays)
    
    nWeeks = Fix((Int(dt2) - Int(dt1)) / 7) 'Nr of whole weeks in interval
    nRest = Abs(Int(dt2) - Int(dt1) - nWeeks * 7) 'Remaining days

    If dt1 > dt2 Then
        If nRest > nDay1 Then
            CalcWeeksBetween = nWeeks - 1
        Else
            CalcWeeksBetween = nWeeks
        End If
    
    ElseIf dt2 > dt1 Then
        If nRest > nDay2 Then
            CalcWeeksBetween = nWeeks + 1
        Else
            CalcWeeksBetween = nWeeks
        End If
    ElseIf dt2 = dt1 Then
        CalcWeeksBetween = 0
    End If
    
End Function


    