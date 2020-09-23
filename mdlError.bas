Attribute VB_Name = "mdlError"
Option Explicit


Const sLocation As String = "mdlError"


Public Function API_Error(ByVal lError As Long) As String
On Error GoTo VB_Error

    Dim sDescription As String
    sDescription = String$(4096, 0)
    
    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lError, 0&, sDescription, Len(sDescription), 0&)
    sDescription = Str_CrLfTerm_Fix(Str_NullTerm_Fix(sDescription))
    
    If sDescription = vbNullString Then
        sDescription = "No description available."
    End If
    
    API_Error = sDescription
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\API_Error")
Resume Next
End Function

Public Function CommDlg_Error(ByVal lError As Long) As String
On Error GoTo VB_Error
    
    Select Case lError
        Case CDERR_DIALOGFAILURE: CommDlg_Error = "The common dialog box procedure's call to the DialogBox function failed."
        Case CDERR_FINDRESFAILURE: CommDlg_Error = "The common dialog box procedure failed to find a specified resource."
        Case CDERR_INITIALIZATION: CommDlg_Error = "The common dialog box procedure failed during initialization."
        Case CDERR_LOADRESFAILURE: CommDlg_Error = "The common dialog box procedure failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE: CommDlg_Error = "The common dialog box procedure failed to load a specified string."
        Case CDERR_LOCKRESFAILURE: CommDlg_Error = "The common dialog box procedure failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE: CommDlg_Error = "The common dialog box procedure was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE: CommDlg_Error = "The common dialog box procedure was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE: CommDlg_Error = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK: CommDlg_Error = "The ENABLEHOOK flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a pointer to a corresponding hook function."
        Case CDERR_NOTEMPLATE: CommDlg_Error = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding template."
        Case CDERR_REGISTERMSGFAIL: CommDlg_Error = "The RegisterWindowMessage function returned an error value when it was called by the common dialog box procedure."
        Case CDERR_STRUCTSIZE: CommDlg_Error = "The lStructSize member of a structure for the corresponding common dialog box is invalid."
        
        Case CFERR_MAXLESSTHANMIN: CommDlg_Error = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."
        Case CFERR_NOFONTS: CommDlg_Error = "No fonts exist."
        
        Case FNERR_BUFFERTOOSMALL: CommDlg_Error = "The buffer for a filename is too small."
        Case FNERR_INVALIDFILENAME: CommDlg_Error = "A filename is invalid."
        Case FNERR_SUBCLASSFAILURE: CommDlg_Error = "An attempt to subclass a list box failed because insufficient memory was available."
        
        Case FRERR_BUFFERLENGTHZERO: CommDlg_Error = "A member in a structure for the corresponding common dialog box points to an invalid buffer."
        
        Case PDERR_CREATEICFAILURE: CommDlg_Error = "The PrintDlg function failed when it attempted to create an information context."
        Case PDERR_DEFAULTDIFFERENT: CommDlg_Error = "An application called the PrintDlg function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer."
        Case PDERR_DNDMMISMATCH: CommDlg_Error = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
        Case PDERR_GETDEVMODEFAIL: CommDlg_Error = "The printer driver failed to initialize a DEVMODE structure."
        Case PDERR_INITFAILURE: CommDlg_Error = "The PrintDlg function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function."
        Case PDERR_LOADDRVFAILURE: CommDlg_Error = "The PrintDlg function failed to load the device driver for the specified printer."
        Case PDERR_NODEFAULTPRN: CommDlg_Error = "A default printer does not exist."
        Case PDERR_NODEVICES: CommDlg_Error = "No printer drivers were found."
        Case PDERR_PARSEFAILURE: CommDlg_Error = "The PrintDlg function failed to parse the strings in the [devices] section of the WIN.INI file."
        Case PDERR_PRINTERNOTFOUND: CommDlg_Error = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."
        Case PDERR_RETDEFFAILURE: CommDlg_Error = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was nonzero."
        Case PDERR_SETUPFAILURE: CommDlg_Error = "The PrintDlg function failed to load the required resources."
        Case Else: CommDlg_Error = "No description available."
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\CommDlg_Error")
Resume Next
End Function

Public Function Exception_Error(ByVal lError As Long) As String
On Error GoTo VB_Error
    
    Select Case lError
        Case EXCEPTION_ACCESS_VIOLATION: Exception_Error = "Access Violation"
        Case EXCEPTION_DATATYPE_MISALIGNMENT: Exception_Error = "Data Misalignment"
        Case EXCEPTION_BREAKPOINT: Exception_Error = "Breakpoint"
        Case EXCEPTION_SINGLE_STEP: Exception_Error = "Single Step"
        Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED: Exception_Error = "Array Bounds Exceeded"
        Case EXCEPTION_FLT_DENORMAL_OPERAND: Exception_Error = "Flt Denormal Operand"
        Case EXCEPTION_FLT_DIVIDE_BY_ZERO: Exception_Error = "Flt Divide By Zero"
        Case EXCEPTION_FLT_INEXACT_RESULT: Exception_Error = "Flt Inexact Result"
        Case EXCEPTION_FLT_INVALID_OPERATION: Exception_Error = "Flt Invalid Operation"
        Case EXCEPTION_FLT_OVERFLOW: Exception_Error = "Flt Overflow"
        Case EXCEPTION_FLT_STACK_CHECK: Exception_Error = "Flt Stack Check"
        Case EXCEPTION_FLT_UNDERFLOW: Exception_Error = "Flt Underflow"
        Case EXCEPTION_INT_DIVIDE_BY_ZERO: Exception_Error = "Int Divide By Zero"
        Case EXCEPTION_INT_OVERFLOW: Exception_Error = "Int Overflow"
        Case EXCEPTION_PRIV_INSTRUCTION: Exception_Error = "Privilaged Instruction"
        Case EXCEPTION_IN_PAGE_ERROR: Exception_Error = "In Page Error"
        Case EXCEPTION_ILLEGAL_INSTRUCTION: Exception_Error = "Illegal Instruction"
        Case EXCEPTION_NONCONTINUABLE_EXCEPTION: Exception_Error = "Non Continuable Exception"
        Case EXCEPTION_STACK_OVERFLOW: Exception_Error = "Stack Overflow"
        Case EXCEPTION_INVALID_DISPOSITION: Exception_Error = "Invalid Disposition"
        Case EXCEPTION_GUARD_PAGE: Exception_Error = "Gaurd Page Violation"
        Case EXCEPTION_INVALID_HANDLE: Exception_Error = "Invalid Handle"
        Case Else: Exception_Error = "No description available."
    End Select

Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Exception_Error")
Resume Next
End Function

Public Function MCI_Error(ByVal lError) As String
On Error GoTo VB_Error
    
    Dim sDescription As String
    sDescription = String$(128, 0)
    
    Call mciGetErrorString(lError, sDescription, Len(sDescription))
    sDescription = Str_NullTerm_Fix(sDescription)
    
    If sDescription = vbNullString Then
        sDescription = "No description available."
    End If
    
    MCI_Error = sDescription
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\MCI_Error")
Resume Next
End Function

Public Function PDH_Error(ByVal lError As Long) As String
On Error GoTo VB_Error
    
    Select Case lError
        Case PDH_CSTATUS_VALID_DATA: PDH_Error = "The returned data is valid."
        Case PDH_CSTATUS_NEW_DATA: PDH_Error = "The return data value is valid and different from the last sample."
        Case PDH_CSTATUS_NO_MACHINE: PDH_Error = "Unable to connect to specified machine or machine is off line."
        Case PDH_CSTATUS_NO_INSTANCE: PDH_Error = "The specified instance is not present."
        Case PDH_MORE_DATA: PDH_Error = "There is more data to return than would fit in the supplied buffer. Allocate a larger buffer and call the function again."
        Case PDH_CSTATUS_ITEM_NOT_VALIDATED: PDH_Error = "The data item has been added to the query, but has not been validated nor accessed. No other status information on this data item is available."
        Case PDH_RETRY: PDH_Error = "The selected operation should be retried."
        Case PDH_NO_DATA: PDH_Error = "No data to return."
        Case PDH_CALC_NEGATIVE_DENOMINATOR: PDH_Error = "A counter with a negative denominator value was detected."
        Case PDH_CALC_NEGATIVE_TIMEBASE: PDH_Error = "A counter with a negative timebase value was detected."
        Case PDH_CALC_NEGATIVE_VALUE: PDH_Error = "A counter with a negative value was detected."
        Case PDH_DIALOG_CANCELLED: PDH_Error = "The user cancelled the dialog box."
        Case PDH_END_OF_LOG_FILE: PDH_Error = "The end of the log file was reached."
        Case PDH_CSTATUS_NO_OBJECT: PDH_Error = "The specified object is not found on the system."
        Case PDH_CSTATUS_NO_COUNTER: PDH_Error = "The specified counter could not be found."
        Case PDH_CSTATUS_INVALID_DATA: PDH_Error = "The returned data is not valid."
        Case PDH_MEMORY_ALLOCATION_FAILURE: PDH_Error = "A PDH function could not allocate enough temporary memory to complete the operation. Close some applications or extend the pagefile and retry the function."
        Case PDH_INVALID_HANDLE: PDH_Error = "The handle is not a valid PDH object."
        Case PDH_INVALID_ARGUMENT: PDH_Error = "A required argument is missing or incorrect."
        Case PDH_FUNCTION_NOT_FOUND: PDH_Error = "Unable to find the specified function."
        Case PDH_CSTATUS_NO_COUNTERNAME: PDH_Error = "No counter was specified."
        Case PDH_CSTATUS_BAD_COUNTERNAME: PDH_Error = "Unable to parse the counter path. Check the format and syntax of the specified path."
        Case PDH_INVALID_BUFFER: PDH_Error = "The buffer passed by the caller is invalid."
        Case PDH_INSUFFICIENT_BUFFER: PDH_Error = "The requested data is larger than the buffer supplied. Unable to return the requested data."
        Case PDH_CANNOT_CONNECT_MACHINE: PDH_Error = "Unable to connect to the requested machine."
        Case PDH_INVALID_PATH: PDH_Error = "The specified counter path could not be interpreted."
        Case PDH_INVALID_INSTANCE: PDH_Error = "The instance name could not be read from the specified counter path."
        Case PDH_INVALID_DATA: PDH_Error = "The data is not valid."
        Case PDH_NO_DIALOG_DATA: PDH_Error = "The dialog box data block was missing or invalid."
        Case PDH_CANNOT_READ_NAME_STRINGS: PDH_Error = "Unable to read the counter and/or explain text from the specified machine."
        Case PDH_LOG_FILE_CREATE_ERROR: PDH_Error = "Unable to create the specified log file."
        Case PDH_LOG_FILE_OPEN_ERROR: PDH_Error = "Unable to open the specified log file."
        Case PDH_LOG_TYPE_NOT_FOUND: PDH_Error = "The specified log file type has not been installed on this system."
        Case PDH_NO_MORE_DATA: PDH_Error = "No more data is available."
        Case PDH_ENTRY_NOT_IN_LOG_FILE: PDH_Error = "The specified record was not found in the log file."
        Case PDH_DATA_SOURCE_IS_LOG_FILE: PDH_Error = "The specified data source is a log file."
        Case PDH_DATA_SOURCE_IS_REAL_TIME: PDH_Error = "The specified data source is the current activity."
        Case PDH_UNABLE_READ_LOG_HEADER: PDH_Error = "The log file header could not be read."
        Case PDH_FILE_NOT_FOUND: PDH_Error = "Unable to find the specified file."
        Case PDH_FILE_ALREADY_EXISTS: PDH_Error = "There is already a file with the specified file name."
        Case PDH_NOT_IMPLEMENTED: PDH_Error = "The function referenced has not been implemented."
        Case PDH_STRING_NOT_FOUND: PDH_Error = "Unable to find the specified string in the list of performance name and explain text strings."
        Case PDH_UNABLE_MAP_NAME_FILES: PDH_Error = "Unable to map to the performance counter name data files. The data will be read from the registry and stored locally."
        Case PDH_UNKNOWN_LOG_FORMAT: PDH_Error = "The format of the specified log file is not recognized by the PDH DLL."
        Case PDH_UNKNOWN_LOGSVC_COMMAND: PDH_Error = "The specified Log Service command value is not recognized."
        Case PDH_LOGSVC_QUERY_NOT_FOUND: PDH_Error = "The specified Query from the Log Service could not be found or could not be opened."
        Case PDH_LOGSVC_NOT_OPENED: PDH_Error = "The Performance Data Log Service key could not be opened. This may be due to insufficient privilege or because the service has not been installed."
        Case PDH_WBEM_ERROR: PDH_Error = "An error occured while accessing the WBEM data store.  The WBEM error code is contained in the LastError value."
        Case PDH_ACCESS_DENIED: PDH_Error = "Unable to access the desired machine or service. Check the permissions and authentication of the log service or the interactive user session against those on the machine or service being monitored."
        Case PDH_LOG_FILE_TOO_SMALL: PDH_Error = "The maximum log file size specified is too small to log the selected counters. No data will be recorded in this log file. Specify a smaller set of counters to log or a larger file size and retry this call."
        Case PDH_INVALID_DATASOURCE: PDH_Error = "Cannot connect to ODBC DataSource Name."
        Case PDH_INVALID_SQLDB: PDH_Error = "SQL Database does not contain a valid set of tables for Perfmon, use PdhCreateSQLTables."
        Case PDH_NO_COUNTERS: PDH_Error = "No counters were found for this Perfmon SQL Log Set."
        Case PDH_SQL_ALLOC_FAILED: PDH_Error = "Call to SQLAllocStmt failed with %1."
        Case PDH_SQL_ALLOCCON_FAILED: PDH_Error = "Call to SQLAllocConnect failed with %1."
        Case PDH_SQL_EXEC_DIRECT_FAILED: PDH_Error = "Call to SQLExecDirect failed with %1."
        Case PDH_SQL_FETCH_FAILED: PDH_Error = "Call to SQLFetch failed with %1."
        Case PDH_SQL_ROWCOUNT_FAILED: PDH_Error = "Call to SQLRowCount failed with %1."
        Case PDH_SQL_MORE_RESULTS_FAILED: PDH_Error = "Call to SQLMoreResults failed with %1."
        Case PDH_SQL_CONNECT_FAILED: PDH_Error = "Call to SQLConnect failed with %1."
        Case PDH_SQL_BIND_FAILED: PDH_Error = "Call to SQLBindCol failed with %1."
        Case PDH_CANNOT_CONNECT_WMI_SERVER: PDH_Error = "Unable to connect to the WMI server on requested machine."
        Case PDH_PLA_COLLECTION_ALREADY_RUNNING: PDH_Error = "Collection %1!s! is already running."
        Case PDH_PLA_ERROR_SCHEDULE_OVERLAP: PDH_Error = "The specified start time is after the end time."
        Case PDH_PLA_COLLECTION_NOT_FOUND: PDH_Error = "Collection %1!s! does not exist."
        Case PDH_PLA_ERROR_SCHEDULE_ELAPSED: PDH_Error = "The specified end time has already elapsed."
        Case PDH_PLA_ERROR_NOSTART: PDH_Error = "Collection %1!s! did not start, check the application event log for any errors."
        Case PDH_PLA_ERROR_ALREADY_EXISTS: PDH_Error = "Collection %1!s! already exists."
        Case PDH_PLA_ERROR_TYPE_MISMATCH: PDH_Error = "There is a mismatch in the settings type."
        Case PDH_PLA_ERROR_FILEPATH: PDH_Error = "The information specified does not resolve to a valid path name."
        Case PDH_PLA_SERVICE_ERROR: PDH_Error = "The Performance Logs & Alerts service did not repond."
        Case PDH_PLA_VALIDATION_ERROR: PDH_Error = "The information passed is not valid."
        Case PDH_PLA_VALIDATION_WARNING: PDH_Error = "The information passed is not valid."
        Case PDH_PLA_ERROR_NAME_TOO_LONG: PDH_Error = "The name supplied is too long."
        Case PDH_INVALID_SQL_LOG_FORMAT: PDH_Error = "SQL log format is incorrect. Correct format is SQL:<DSN-name>!<LogSet-Name>."
        Case PDH_COUNTER_ALREADY_IN_QUERY: PDH_Error = "Performance counter in PdhAddCounter() call has already been added in the performacne query. This counter is ignored."
        Case Else: PDH_Error = "No description available."
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\PDH_Error")
Resume Next
End Function


Public Sub Error_API(ByVal lError As Long, ByVal sSource As String, ByVal sFunction As String)
On Error GoTo VB_Error
    
    sErrorLogNum = sErrorLogNum + 1
    sErrorLog = sErrorLogNum & "\API\" & sSource & "\" & sFunction & "\" & lError & vbCrLf & sErrorLog
    
    Call Error_Write(sErrorLogNum & "\API\" & sSource & "\" & sFunction & "\" & lError & vbCrLf)
    
    If Forms_Loaded.bErrorLog = True Then
        frmErrorLog.txtLog.Text = sErrorLog
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Error_API")
Resume Next
End Sub

Public Sub Error_CommDlg(ByVal lError As Long, ByVal sSource As String, ByVal sFunction As String)
On Error GoTo VB_Error
    
    sErrorLogNum = sErrorLogNum + 1
    sErrorLog = sErrorLogNum & "\COMMDLG\" & sSource & "\" & sFunction & "\" & lError & vbCrLf & sErrorLog
    
    Call Error_Write(sErrorLogNum & "\COMMDLG\" & sSource & "\" & sFunction & "\" & lError & vbCrLf)
    
    If Forms_Loaded.bErrorLog = True Then
        frmErrorLog.txtLog.Text = sErrorLog
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Error_CommDlg")
Resume Next
End Sub

Public Sub Error_MCI(ByVal lError As Long, ByVal sSource As String, ByVal sFunction As String)
On Error GoTo VB_Error
    
    sErrorLogNum = sErrorLogNum + 1
    sErrorLog = sErrorLogNum & "\MCI\" & sSource & "\" & sFunction & "\" & lError & vbCrLf & sErrorLog
    
    Call Error_Write(sErrorLogNum & "\MCI\" & sSource & "\" & sFunction & "\" & lError & vbCrLf)
    
    If Forms_Loaded.bErrorLog = True Then
        frmErrorLog.txtLog.Text = sErrorLog
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Error_MCI")
Resume Next
End Sub

Public Sub Error_PDH(ByVal lError As Long, ByVal sSource As String, ByVal sFunction As String)
On Error GoTo VB_Error
    
    sErrorLogNum = sErrorLogNum + 1
    sErrorLog = sErrorLogNum & "\PDH\" & sSource & "\" & sFunction & "\" & lError & vbCrLf & sErrorLog
    
    Call Error_Write(sErrorLogNum & "\PDH\" & sSource & "\" & sFunction & "\" & lError & vbCrLf)
    
    If Forms_Loaded.bErrorLog = True Then
        frmErrorLog.txtLog.Text = sErrorLog
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Error_PDH")
Resume Next
End Sub

Public Sub Error_VB(ByVal oErr As ErrObject, ByVal sSource As String)
    sErrorLogNum = sErrorLogNum + 1
    sErrorLog = sErrorLogNum & "\VB\" & sSource & "\" & oErr.Number & vbCrLf & sErrorLog
    
    Call Error_Write(sErrorLogNum & "\VB\" & sSource & "\" & oErr.Number & vbCrLf)
    
    If Forms_Loaded.bErrorLog = True Then
        frmErrorLog.txtLog.Text = sErrorLog
    End If
End Sub

Public Sub Error_Write(ByVal sData As String)
    Dim hFile As Long
    Dim lStart As Long
    Dim lExtra As Long
    
    hFile = CreateFile(sAppPath & "\kira.log", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_ALWAYS, 0&, 0&)
    
    lStart = GetFileSize(hFile, lExtra)
    If lStart < 0 Then lStart = 0
    
    Call SetFilePointer(hFile, lStart, 0&, FILE_BEGIN)
    Call WriteFile(hFile, ByVal sData, Len(sData), lExtra, ByVal 0&)
    Call SetFilePointer(hFile, 0&, 0&, FILE_BEGIN)
    Call CloseHandle(hFile)
End Sub

Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
    Dim EXCEPTION_RECORD() As EXCEPTION_RECORD
    ReDim EXCEPTION_RECORD(0)
    EXCEPTION_RECORD(0) = ExceptionPtrs.ExceptionRecord
    
    Do
        If EXCEPTION_RECORD(UBound(EXCEPTION_RECORD)).ExceptionRecord = 0 Then
            Exit Do
        Else
            If IsBadReadPtr(EXCEPTION_RECORD(UBound(EXCEPTION_RECORD)).ExceptionRecord, Len(EXCEPTION_RECORD(0))) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\ExceptionHandler", "IsBadReadPtr")
            Else
                ReDim Preserve EXCEPTION_RECORD(UBound(EXCEPTION_RECORD) + 1)
                Call MoveMemory(EXCEPTION_RECORD(UBound(EXCEPTION_RECORD)), ByVal EXCEPTION_RECORD(UBound(EXCEPTION_RECORD) - 1).ExceptionRecord, Len(EXCEPTION_RECORD(0)))
            End If
        End If
        
        If bShutdown = True Then Exit Do
    Loop
    
    If EXCEPTION_RECORD(UBound(EXCEPTION_RECORD)).ExceptionFlags = EXCEPTION_NONCONTINUABLE Then
        Call Error_VB(Err, sLocation & "\ExceptionHandler")
        Call Main_Exit
    Else
        Call Err.Raise(1, sLocation & "\ExceptionHandler", Exception_Error(EXCEPTION_RECORD(UBound(EXCEPTION_RECORD)).ExceptionCode))
    End If
    
    ExceptionHandler = EXCEPTION_CONTINUE_SEARCH
End Function
