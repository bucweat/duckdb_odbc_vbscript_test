option explicit

REM From DUCKDB CI 
REM Reg Query "HKLM\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
REM Reg Query "HKLM\SOFTWARE\ODBC\ODBC.INI\DuckDB"
REM Reg Query "HKLM\SOFTWARE\ODBC\ODBCINST.INI\DuckDB Driver"
REM REG ADD "HKCU\SOFTWARE\ODBC\ODBC.INI\ODBC" //f
REM REG ADD "HKCU\SOFTWARE\ODBC\ODBC.INI\ODBC" //v Trace //t REG_SZ //d 1
REM REG ADD "HKCU\SOFTWARE\ODBC\ODBC.INI\ODBC" //v TraceDll //t REG_SZ //d "C:\Windows\system32\odbctrac.dll"
REM REG ADD "HKCU\SOFTWARE\ODBC\ODBC.INI\ODBC" //v TraceFile //t REG_SZ //d "D:\a\duckdb\duckdb\ODBC_TRACE.log"
REM echo "----------------------------------------------------------------"
REM Reg Query "HKCU\SOFTWARE\ODBC\ODBC.INI\ODBC"

class classTimer

    Private fStartTime
    Private fStopTime
    Private fCurrentTime
    Private lCounter
    Private t
    
    Private Sub Class_Initialize
        lCounter = 0
    end sub

    Private Sub Class_Terminate
    end sub

    Public Property Let StartTime(f) 
        fStartTime = f 
    End Property
    
    Public Property Get StartTime 
        StartTime = fStartTime 
    End Property 
    
    Public Property Let StopTime(f) 
        fStopTime = f 
    End Property
    
    Public Property Get StopTime 
        StopTime = fStopTime 
    End Property
    
    Public Property Get CurrentTime
        fCurrentTime = UnixEpoch +  (getMs/1000)
        CurrentTime = fCurrentTime 
    End Property 
    
    Public Property Get ElapsedTime
        fCurrentTime = UnixEpoch +  (getMs/1000)
        ElapsedTime = fCurrentTime - fStartTime
    End Property 
 
    public function StartTimer()
        StartTime = UnixEpoch + (getMs/1000)
        StartTimer = true
    end function
    
    public function StopTimer()
        StopTime = UnixEpoch +  (getMs/1000)
        StopTimer = true
    end function

    private function getMs
        t = Timer
        getMs = Int((t-Int(t)) * 1000)
    end function

    public function UnixEpoch
        Dim utc_now, t_diff, objWMIService, colItems, item
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        'Get UTC time string
        Set colItems = objWMIService.ExecQuery("Select * from Win32_UTCTime")
        For Each item In colItems
            If Not IsNull(item) Then
                utc_now = item.Month & "/" & item.Day & "/" & item.Year & " " _
                            & item.Hour & ":" & item.Minute & ":" & item.Second
            End If
        Next
        'Get UTC offset, not constant due to daylight savings
        t_diff = Abs(DateDiff("h", utc_now, Now()))
        'Calculate seconds since start of epoch
        UnixEpoch = DateDiff("s", "01/01/1970 00:00:00", DateAdd("h",t_diff,Now()))
    end function

end class

' constants for ADO connection/recordset
Const adOpenForwardOnly = 0
Const adOpenDynamic = 2
Const adOpenStatic = 3

Const adLockBatchOptimistic = 4
Const adLockOptimistic = 3
Const adLockPessimistic = 2
Const adLockReadOnly = 1

' constants for registry access
Const HKEY_CLASSES_ROOT     = &H80000000
Const HKEY_CURRENT_USER     = &H80000001
Const HKEY_LOCAL_MACHINE    = &H80000002
Const HKEY_USERS            = &H80000003
Const HKEY_CURRENT_CONFIG   = &H80000005

class classOdbcTests
    private objFSO
    private oWShell
    private objConn 
    private objRS 
    private strQuery 
    private iBitness
    private sBitPath
    private strFolder
    private separator
    private aQueryResults
    private dDataTypes
    private dbSqlite3
    private dbDuck
    private oTimer
    
    '*************************************************************************
    sub class_initialize()
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set oWShell = CreateObject("WScript.Shell")
        set oTimer = new classTimer
        
        separator = ","
        
        aQueryResults = Array()
        redim aQueryResults(5)
        aQueryResults(0) = ""
        aQueryResults(1) = ""
        set aQueryResults(2) = CreateObject("Scripting.Dictionary")
        aQueryResults(3) = ""
        aQueryResults(4) = ""

        on error resume next
        'try wscript first...if it works then move on
        strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
        'if variable not defined (error = 500), then we are running in an hta
        if err.number = 500 then
            strFolder = objFSO.GetParentFolderName(Replace(location.pathname,"%20"," "))
        end if
        on error goto 0
        
        log "strFolder [" & strFolder & "]"

        ' dictionary of ADO data types
        Set dDataTypes = CreateObject("Scripting.Dictionary")
        dDataTypes.add 20,"adBigInt" 'Indicates an eight-byte signed integer (DBTYPE_I8).
        dDataTypes.add 128,"adBinary" 'Indicates a binary value (DBTYPE_BYTES).
        dDataTypes.add 11,"adBoolean" 'Indicates a Boolean value (DBTYPE_BOOL).
        dDataTypes.add 8,"adBSTR" 'Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
        dDataTypes.add 136,"adChapter" 'Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
        dDataTypes.add 129,"adChar" 'Indicates a string value (DBTYPE_STR).
        dDataTypes.add 6,"adCurrency" 'Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
        dDataTypes.add 7,"adDate" 'Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        dDataTypes.add 133,"adDBDate" 'Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
        dDataTypes.add 134,"adDBTime" 'Indicates a time value (hhmmss) (DBTYPE_DBTIME).
        dDataTypes.add 135,"adDBTimeStamp" 'Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
        dDataTypes.add 14,"adDecimal" 'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        dDataTypes.add 5,"adDouble" 'Indicates a double-precision floating-point value (DBTYPE_R8).
        dDataTypes.add 0,"adEmpty" 'Specifies no value (DBTYPE_EMPTY).
        dDataTypes.add 10,"adError" 'Indicates a 32-bit error code (DBTYPE_ERROR).
        dDataTypes.add 64,"adFileTime" 'Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
        dDataTypes.add 72,"adGUID" 'Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
        dDataTypes.add 3,"adInteger" 'Indicates a four-byte signed integer (DBTYPE_I4).
        dDataTypes.add 205,"adLongVarBinary" 'Indicates a long binary value.
        dDataTypes.add 201,"adLongVarChar" 'Indicates a long string value.
        dDataTypes.add 203,"adLongVarWChar" 'Indicates a long null-terminated Unicode string value.
        dDataTypes.add 131,"adNumeric" 'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
        dDataTypes.add 138,"adPropVariant" 'Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
        dDataTypes.add 4,"adSingle" 'Indicates a single-precision floating-point value (DBTYPE_R4).
        dDataTypes.add 2,"adSmallInt" 'Indicates a two-byte signed integer (DBTYPE_I2).
        dDataTypes.add 16,"adTinyInt" 'Indicates a one-byte signed integer (DBTYPE_I1).
        dDataTypes.add 21,"adUnsignedBigInt" 'Indicates an eight-byte unsigned integer (DBTYPE_UI8).
        dDataTypes.add 19,"adUnsignedInt" 'Indicates a four-byte unsigned integer (DBTYPE_UI4).
        dDataTypes.add 18,"adUnsignedSmallInt" 'Indicates a two-byte unsigned integer (DBTYPE_UI2).
        dDataTypes.add 17,"adUnsignedTinyInt" 'Indicates a one-byte unsigned integer (DBTYPE_UI1).
        dDataTypes.add 132,"adUserDefined" 'Indicates a user-defined variable (DBTYPE_UDT).
        dDataTypes.add 204,"adVarBinary" 'Indicates a binary value.
        dDataTypes.add 200,"adVarChar" 'Indicates a string value.
        dDataTypes.add 12,"adVariant" 'Indicates an Automation Variant (DBTYPE_VARIANT).
        dDataTypes.add 139,"adVarNumeric" 'Indicates a numeric value.
        dDataTypes.add 202,"adVarWChar" 'Indicates a null-terminated Unicode character string.
        dDataTypes.add 130,"adWChar" 'Indicates a null-terminated Unicode character string (DBTYPE_WSTR).

        log "class_initialize"
    end sub
    
    '*************************************************************************
    sub class_terminate()
        set oTimer = nothing
        set objFSO = nothing
        set oWShell = nothing
        Set objRS = nothing
        Set objConn = nothing
        log "class_terminate"
    end sub
    
    '********************************************
    public function executeTests(i)
        select case i
            case 64
                iBitness = 64
                sBitPath = "64bit"
            case else
                log "script requires 64"
                exit function
        end select
        log "classOdbcTests configured to run " & sBitPath
        getArchitecture
        getInstalledOdbcDriverList ""
            
        if getInstalledOdbcDriverList("DuckDB Driver") then
            log "****************************************************" & vbCrLf & "DuckDB driver found!"
            
            if true then
                REM dbDuck = ":memory:"
                dbDuck = "test.duckdb"
                on error resume next : objFSO.deletefile dbDuck,true : on error goto 0
                if opendb("DUCKDB") then
                    
                    duckdbVersion
                    timestamp
                    pragmas
                    makeTheWeather
                    getTheWeather
                    duckdbGetSchema
                    duckdbGetSchemaSqliteMaster

                    closedb
                end if
            end if
            
        else
            log "****************************************************" & vbCrLf & "DuckDB driver NOT found!"
        end if
        
    end function
    
    '********************************************
    ' DuckDB
    '********************************************

    '********************************************
    ' DuckDB timestamp
    '********************************************
    public function timestamp() 
        REM note that none of these return partial seconds...here is a good reference
        REM https://stackoverflow.com/questions/48969397/adodb-unable-to-store-datetime-value-with-sub-second-precision
        
        logResult query2csv( "SELECT TIMESTAMP '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT TIMESTAMP '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT DATETIME '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT DATE '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT TIME '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT TIME '1992-09-20 11:30:00.123456' as t;" )
        logResult query2csv( "SELECT get_current_time() as t;" )
    end function

    '********************************************
    function duckdbGetSchema
        log "duckdbGetSchema same table reqeusted by GetSchema() .NET method"
        logresult query2csv( _
            "SELECT " & _
            "         table_catalog::VARCHAR ""TABLE_CAT"" " & _
            "       , table_schema ""TABLE_SCHEM"" " & _
            "       , table_name ""TABLE_NAME"" " & _
            "       , CASE " & _
            "                  WHEN table_type='BASE TABLE' " & _
            "                           THEN 'TABLE' " & _
            "                           ELSE table_type " & _
            "         END ""TABLE_TYPE"" " & _
            "       , '' ""REMARKS"" " & _
            "FROM " & _
            "         information_schema.tables " & _
            "WHERE " & _
            "         COALESCE(""TABLE_SCHEM"",'') LIKE  '%' ESCAPE '\' " & _
            "         AND COALESCE(""TABLE_NAME"",'') LIKE  '%' ESCAPE '\' " & _
            "         AND table_type IN ('BASE TABLE') " & _
            "ORDER BY " & _
            "         TABLE_TYPE " & _
            "       , TABLE_CATALOG " & _
            "       , TABLE_SCHEMA " & _
            "       , TABLE_NAME" & _
            ";" _
        )
    end function

    '********************************************
    function duckdbGetSchemaSqliteMaster
        log "sqlite_master table"
        logresult query2csv( "SELECT * from sqlite_master;" )
    end function
    
    '********************************************
    public function pragmas
        REM https://duckdb.org/docs/configuration/pragmas
        logResult query2csv( "SET log_query_path = 'C:\Users\charlie\Desktop\duck.log';" )
        
        logResult query2csv( "PRAGMA database_list;" )
        logResult query2csv( "PRAGMA database_size;" )
        
        logResult query2csv( "PRAGMA show_tables;" )
        logResult query2csv( "PRAGMA show_tables_expanded;" )
        logResult query2csv( "PRAGMA table_info('cities');" )
        logResult query2csv( "PRAGMA collations;" )
        logResult query2csv( "PRAGMA storage_info('cities');" )
        logResult query2csv( "PRAGMA functions;" )
        logResult query2csv( "PRAGMA version;" )
        logResult query2csv( "PRAGMA platform;" )
        logResult query2csv( "PRAGMA metadata_info;" )
        
        
        logResult query2csv( "PRAGMA enable_profiling;" )
        logResult query2csv( "SET enable_profiling = 'query_tree';" )
        logResult query2csv( "select * from cities;" )
        logResult query2csv( "PRAGMA disable_profiling;" )
        logResult query2csv( "select * from cities;" )
        
        
        logResult query2csv( "PRAGMA enable_verification;" )
        logResult query2csv( "select * from cities;" )
        logResult query2csv( "PRAGMA disable_verification;" )
        
        logResult query2csv( "SET log_query_path = '';" )
        
    end function
    
    '********************************************
    public function duckdbVersion
        logResult query2csv( "PRAGMA version;" )
    end function

    '********************************************
    ' Generic
    '********************************************
    
    '********************************************
    public function makeTheWeather
        logResult query2csv( "DROP TABLE IF EXISTS weather;" )
        
        ' from https://duckdb.org/docs/sql/introduction
        logResult query2csv( _
            "CREATE TABLE weather ( " & _
            "    city           VARCHAR, " & _
            "    temp_lo        INTEGER," & _
            "    temp_hi        INTEGER," & _
            "    prcp           REAL, " & _
            "    date           DATE " & _
            "); "_
        )

        logResult query2csv( "DROP TABLE IF EXISTS cities;" )
        logResult query2csv( _
            "CREATE TABLE cities ( " & _
            "    name            VARCHAR, " & _
            "    lat             DECIMAL, " & _
            "    lon             DECIMAL " & _
            "); "_
        )
        
        logResult query2csv("INSERT INTO weather VALUES ('San Francisco', 46, 50, 0.25, '1994-11-27');")
        logResult query2csv("INSERT INTO weather VALUES ('New York', 45, 50, 0.25, '1994-11-27');")
        logResult query2csv( _
            "INSERT INTO weather (city, temp_lo, temp_hi, prcp, date) " & _
            "VALUES('San Francisco', 43, 57, 0.0, '1994-11-29');" _
        )
        logResult query2csv( _
            "INSERT INTO weather (date, city, temp_hi, temp_lo) " & _
            "VALUES ('1994-11-29', 'Hayward', 54, 37); " _
        )
        logResult query2csv("INSERT INTO cities VALUES ('San Francisco', 1,1);")
        logResult query2csv("INSERT INTO cities VALUES ('New York',2,2);")
    end function
    
    '********************************************
    public function getTheWeather
        logResult query2csv("PRAGMA table_info('weather');")
        
        logResult query2csv("select * from weather;")
        
        logResult query2csv("select count(*) from weather;")
        
        logResult query2csv("SELECT city, (temp_hi+temp_lo)/2 AS temp_avg, date FROM weather;")

        logResult query2csv("SELECT * FROM weather WHERE city = 'San Francisco' AND prcp > 0.0;")
        
        logResult query2csv("SELECT * FROM weather ORDER BY city;")
        
        logResult query2csv("SELECT * FROM weather ORDER BY city, temp_lo;")
        
        logResult query2csv("SELECT DISTINCT city FROM weather;")
        
        logResult query2csv("SELECT DISTINCT city FROM weather ORDER BY city;")
        
        log "no location data for hayward"
        logResult query2csv("SELECT * FROM weather, cities WHERE city = 'hayward';")
        
        logResult query2csv("SELECT city, temp_lo, temp_hi, prcp, date, lon, lat FROM weather, cities WHERE city = 'hayward';")
        
        logResult query2csv("SELECT weather.city, weather.temp_lo, weather.temp_hi, weather.prcp, weather.date, cities.lon, cities.lat FROM weather, cities WHERE cities.name = weather.city;")
        
        logResult query2csv("SELECT * FROM weather INNER JOIN cities ON (weather.city = cities.name);")
        
        logResult query2csv("SELECT * FROM weather LEFT OUTER JOIN cities ON (weather.city = cities.name);")
        
        logResult query2csv("SELECT max(temp_lo) FROM weather;")
        
        logResult query2csv("SELECT city FROM weather WHERE temp_lo = (SELECT max(temp_lo) FROM weather);")
        
        logResult query2csv("SELECT city, max(temp_lo) FROM weather GROUP BY city;")
        
        logResult query2csv("SELECT city, max(temp_lo) FROM weather GROUP BY city HAVING max(temp_lo) < 40;")
        
        logResult query2csv("SELECT city, max(temp_lo) FROM weather WHERE city LIKE 'S%' GROUP BY city HAVING max(temp_lo) < 40;")
        
        logResult query2csv("UPDATE weather SET temp_hi = temp_hi - 2,  temp_lo = temp_lo - 2 WHERE date > '1994-11-28';")
        
        logResult query2csv("SELECT * FROM weather;")
        
        logResult query2csv("DELETE FROM weather WHERE city = 'Hayward';")
        
        logResult query2csv("SELECT * FROM weather;")
        
        ' logResult query2csv("DELETE FROM weather;")
        ' log "oopsie..delete removes all records from table!"
        ' logResult query2csv("SELECT * FROM weather;")
        ' logResult query2csv(DROP TABLE weather;")

    end function

    '********************************************
    ' Utility methods
    '********************************************
    
    '********************************************
    Public Function GetLogTime() 
        Dim strNow 
        strNow = Now() 
        GetLogTime = _
            Year(strNow) & "-" & _
            Pad(Month(strNow), 2, "0", True) & "-" & _
            Pad(Day(strNow), 2, "0", True) & "T" & _
            Pad(Hour(strNow), 2, "0", True) & ":" & _
            Pad(Minute(strNow), 2, "0", True) & ":" & _
            Pad(Second(strNow), 2, "0", True) 
    End Function
    
    '********************************************
    public function log (s)
        on error resume next
        ' try to log to cmd console
        wscript.echo s
        if err.number = 500 then
            ' must be in browser, so log to console
            console.log s
        end if
        on error goto 0
    end function
    
    '********************************************
    Public Function Pad(strText, nLen, strChar, bFront) 
        Dim nStartLen 
        If strChar = "" Then 
            strChar = "0" 
        End If 
        nStartLen = Len(strText) 
        If Len(strText) >= nLen Then 
            Pad = strText 
        Else 
            If bFront Then 
                Pad = String(nLen - Len(strText), strChar) & strText 
            Else 
                Pad = strText & String(nLen - Len(strText), strChar) 
            End If 
        End If 
    End Function 
    
    '********************************************
    public function logResult(r)
        ' log query2csv() result contained in aQueryResults to console
        if r >= 0 then
            log "QUERY  " & aQueryResults(0)
            log "Time   " & aQueryResults(4)
            log "return " & (r+1) & " rows"
            if len(aQueryResults(3)) > 0  then log "ERROR  " & aQueryResults(3)
            log aQueryResults(1)
            dim vKey: for each vKey in aQueryResults(2)
                log aQueryResults(2).item(vKey)
            next
            log ""
        else
            log "QUERY  " & aQueryResults(0)
            log "Time   " & aQueryResults(4)
            log "return " & r
            if len(aQueryResults(3)) > 0  then log "ERROR  " & aQueryResults(3)
            log ""
        end if
    end function
    
    '********************************************
    function opendb(p)
        opendb = false
        Set objConn = CreateObject("ADODB.Connection")
        Set ObjRS = CreateObject("ADODB.Recordset")
        dim sConnStr
        select case p
            case "DUCKDB"
                if dbDuck = ":memory:" then
                    ' no spaces allow (e.g. 'allow_unsigned_extensions = true;' will not work)
                    sConnStr = "Driver=DuckDB Driver;Database=:memory:;allow_unsigned_extensions=true;"
                    REM sConnStr = "Driver=DuckDB Driver;DataSource=:memory:;allow_unsigned_extensions=true;"
                    REM sConnStr = "Driver=DuckDB Driver;:memory:;"
                    REM sConnStr = "DRIVER=DuckDB Driver;"
                else
                    sConnStr = "Driver=DuckDB Driver;Database=" & dbDuck & ";allow_unsigned_extensions=true;"
                    REM sConnStr = "Driver=DuckDB Driver;DataSource=" & dbDuck & ";allow_unsigned_extensions=true;"
                    REM sConnStr = "Driver={DuckDB Driver};Database=" & dbDuck & ";"
                end if
        end select
        log "opendb(" & p & ") --> " & sConnStr & vbcrlf
        objConn.ConnectionString = sConnStr
        on error resume next
        objConn.open
        if err.number <> 0 then
            log err.number & " " & err.description & vbcrlf
            on error goto 0
            exit function
        end if
        on error goto 0
        ObjRS.CursorType = adOpenStatic
        Objrs.LockType = adLockOptimistic
        Set objRS.ActiveConnection = objConn
        opendb = true
    end function
  
    '********************************************
    sub closedb
        if objRS.State = 1 then objRS.Close
        if objConn.State = 1 then objConn.close
        Set objRS = nothing
        Set objConn = nothing
    end sub

    '********************************************
    function query(s)
        dim bOutputTextType: bOutputTextType = true
        dim ss: ss = "QUERY: " & s & vbcrlf
        on error resume next
        dim oRs: set oRs = objConn.execute(s)
        if err.number <> 0 then
            query = "QUERY ERROR --> " & s & ":" & err.description
            exit function
        else
            on error goto 0
            if oRs.state = 1 then
                if (not oRs.BOF) and (not oRs.EOF)  then
                    on error resume next
                    if err.number <> 0 then
                        log err.number & " " & err.description
                        err.clear
                    end if
                    ' if iBitness = 32 then oRs.MoveFirst
                    oRs.MoveFirst
                    on error goto 0
                    do while oRs.EOF = false
                        dim ff
                        on error resume next
                        dim fieldCount: fieldCount = oRs.fields.count
                        if err.number <> 0 then
                            log "0 ERROR " & err.number & "-->" & err.description
                            exit function
                        end if
                        on error goto 0
                        for each ff in oRs.Fields
                            REM log ff.Name & " " & ff.Type & " " & ff.Value
                            on error resume next
                            select case ff.type
                                case 1      ' vbNull
                                    ss = ss & ff.Name & ":NULL "
                                case 128    ' adBinary
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                    else
                                        ss = ss & ff.Name & "(adBinary):" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                    end if
                                case 2      ' vbInteger
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbInteger):" & ff.Value & separator
                                    end if
                                case 20      ' vbBigInteger
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbBigInteger):" & ff.Value & separator
                                    end if
                                case 3      ' vbLong
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbLong):" & ff.Value & separator
                                    end if
                                case 4      ' vbSingle
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbSingle):" & ff.Value & separator
                                    end if
                                case 5      ' vbDouble
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbDouble):" & ff.Value & separator
                                    end if
                                case 133    ' adDBDate
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(adDBDate):" & cdbl(ff.Value) & " " & ff.Value & separator
                                    end if
                                case 135    ' adDBTimeStamp
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(adDBTimeStamp):" & cdbl(ff.Value) & " " & ff.Value & separator
                                    end if
                                case 202    ' adVarWChar
                                    if typename(ff.Value) = "Null" then
                                        if not bOutputTextType then
                                            ss = ss & ff.Name & "(" & ff.type & "):Null" & separator
                                        else
                                            ss = ss & ff.Name & "(text):Null" & separator
                                        end if
                                    else
                                        if not bOutputTextType then
                                            ss = ss & ff.Name & "(" & ff.type & ")(" & len(ff.Value) & "):" & ff.Value & separator
                                        else
                                            ss = ss & ff.Name & "(text)(" & len(ff.Value) & "):" & ff.Value & separator
                                        end if
                                    end if
                                    REM wscript.echo typename(ff.value) & " " & lenb(ff.value) & " [" & BinaryToString(ff.value,false) & "]"
                                case 203    ' adLongVarWChar (memo)
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(adLongVarWChar):" & ff.Value & separator
                                    end if
                                case else
                                    REM ss = ss & "ELSE " & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    ss = ss & "ELSE " & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                            end select
                            if err.number <> 0 then
                                log "1 ERROR " & err.number & "-->" & err.description
                                log "BinaryToString " & BinaryToString(ff.value,true)
                                ss = ss & ff.Name  & "(" & ff.type & "):" &  "ERR "
                                err.clear
                            end if
                            on error goto 0
                        next
                        on error resume next
                        REM log "MoveNext BOF " & oRs.BOF & " EOF " & oRs.EOF
                        oRs.MoveNext
                        if err.number <> 0 then
                            log "2 ERROR " & err.number & "-->" & err.description
                            log "    query --> " & s
                        end if
                        on error goto 0
                        ' remove the last separator 
                        ss = left(ss,len(ss)-1)
                        ss = ss & vbcrlf
                    loop
                else
                    query = ss & "recordset contains no records" & vbcrlf
                    oRs.close
                    exit function
                end if
            else
                query = ss & "query did not return an open recordset" & vbcrlf
                exit function
            end if
        end if
        oRs.close
        set oRs = nothing
        query= ss
    end function

    '********************************************
    function query2csv(s)
        dim oTimerQuery2csv: set oTimerQuery2csv = new classTimer
        oTimerQuery2csv.StartTimer
        query2csv = -1
        aQueryResults(2).removeall
        aQueryResults(3) = ""
        dim rowCount: rowCount = 0
        dim bOutputTextType: bOutputTextType = true
        dim ss
        aQueryResults(0) = s
        on error resume next
        dim oRs: set oRs = objConn.execute(s)
        if err.number <> 0 then
            aQueryResults(3) = aQueryResults(3) & err.number & ":" & err.description & vbcrlf
            exit function
        else
            on error goto 0
            if oRs.state = 1 then
                dim ff
                ss = ""
                for each ff in oRs.Fields
                if bOutputTextType then
                    ss = ss & ff.Name & "(" & dDataTypes(ff.type) & " " & ff.type &" )" & separator
                else
                    ss = ss & ff.Name & "(" & ff.type & ")" & separator
                end if
                next
                ' remove the last separator 
                ss = left(ss,len(ss)-1)
                aQueryResults(1) = ss
                REM log ""
                REM log "QUERY  --> " & aQueryResults(0)
                REM log "HEADER --> " & aQueryResults(1)
                REM log ""
                if (not oRs.BOF) and (not oRs.EOF)  then
                    on error resume next
                    if err.number <> 0 then
                        log err.number & " " & err.description
                        err.clear
                    end if
                    oRs.MoveFirst
                    on error goto 0
                    do while oRs.EOF = false
                        ss = ""
                        on error resume next
                        dim fieldCount: fieldCount = oRs.fields.count
                        if err.number <> 0 then
                            aQueryResults(3) = aQueryResults(3) & "0: " & err.number & ":" & err.description & "|"
                            exit function
                        end if
                        on error goto 0
                        for each ff in oRs.Fields
                            REM log "[" & ff.name & "] " & dDataTypes(ff.type) & " (" & ff.type & ") " & (ff.value)
                            REM on error resume next
                            select case ff.type
                                case 1      ' vbNull
                                    ss = ss & ff.Name & ":NULL "
                                case 2      ' adSmallInt
                                    ss = ss & ff.Value & separator
                                case 3      ' adInteger
                                    ss = ss & ff.Value & separator
                                case 4      ' adSingle
                                    ss = ss & ff.Value & separator
                                case 5      ' adDouble
                                    ss = ss & ff.Value & separator
                                case 20     ' adBigInt
                                    ss = ss & ff.Value & separator
                                case 128    ' adBinary
                                    ss = ss & ff.Name & "(adBinary):" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                case 129    ' adChar
                                    ' specific handling for duckdb as we don't see this type in SQLite
                                    if typename(ff.Value) = "Null" then 
                                        ss = ss & "NULL" & separator
                                    else
                                        select case ff.Value
                                            case "T"
                                                ss = ss & "1" & separator
                                            case "F"
                                                ss = ss & "0" & separator
                                            case else
                                                ss = ss & ff.Value & separator
                                       end select
                                    end if
                                case 133    ' adDBDate
                                    ss = ss & ff.Name & "(adDBDate):" & ff.value & " bin2str:" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                    ' ss = ss & typename(ff.Value) & separator
                                case 135    ' adDBTimeStamp
                                    ss = ss & ff.Value & separator
                                case 202    ' adVarWChar
                                    if typename(ff.Value) = "Null" then
                                        ss = ss & "Null" & separator
                                    else
                                        ss = ss & chr(34) & replace(ff.Value,chr(34),chr(34)&chr(34)) & chr(34) & separator
                                    end if
                                case 203    ' adLongVarWChar (memo)
                                    ss = ss & chr(34) & replace(ff.Value,chr(34),chr(34)&chr(34)) & chr(34) & separator
                                case else
                                    ss = ss & ff.Value & separator
                            end select
                            if err.number <> 0 then
                                aQueryResults(3) = aQueryResults(3) & "1:" & err.number & ":" & err.description & "|"
                                err.clear
                            end if
                            on error goto 0
                        next
                        on error resume next
                        oRs.MoveNext
                        if err.number <> 0 then
                            aQueryResults(3) = aQueryResults(3) & "2 " & err.number & ":" & err.description & "|"
                        end if
                        on error goto 0
                        ' remove the last separator 
                        ss = left(ss,len(ss)-1)
                        aQueryResults(2).add rowCount, ss
                        query2csv = rowCount
                        rowCount = rowCount + 1
                    loop
                else
                    aQueryResults(3) = "recordset contains no records" & "|"
                    oRs.close
                    query2csv = -1
                    aQueryResults(4) = round(oTimerQuery2csv.ElapsedTime,3) & " seconds"
                    set oTimerQuery2csv = nothing
                    exit function
                end if
            else
                aQueryResults(3) = "query did not return an open recordset" & "|"
                query2csv = -1
                aQueryResults(4) = round(oTimerQuery2csv.ElapsedTime,3) & " seconds"
                set oTimerQuery2csv = nothing
                exit function
            end if
        end if
        oRs.close
        set oRs = nothing
        oTimerQuery2csv.StopTimer
        on error goto 0
        aQueryResults(4) = round(oTimerQuery2csv.ElapsedTime,3) & " seconds"
        set oTimerQuery2csv = nothing
    end function
    
    '********************************************
    function getArchitecture()
        ' get the system and process architecture info
        dim WshShell: Set WshShell = CreateObject("WScript.Shell")
        dim WshSysEnv: Set WshSysEnv = WshShell.Environment("SYSTEM")
        dim WshProcEnv: Set WshProcEnv = WshShell.Environment("PROCESS")
        log "architecture reports:"
        log "    SYSTEM " & WshSysEnv("PROCESSOR_ARCHITECTURE")
        log "    PROCESS " & WshProcEnv("PROCESSOR_ARCHITECTURE")
    end function
    
    '********************************************
    function getInstalledOdbcDriverList(s)
        getInstalledOdbcDriverList = false
        ' get ODBC installed drivers for selected architecture 
        ' (32 or 64 bit, depends on process running this code)
        dim odStart: odStart = now
        dim OutText: OutText = ""
        dim strValueName: strValueName = ""
        dim strValue: strValue = ""
        dim strKeyPath: strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
        dim arrValueNames,arrValueTypes
        dim objRegistry: Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
        objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
        dim i: For i = 0 to UBound(arrValueNames)
            strValueName = arrValueNames(i)
            objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
            OutText = OutText & "    " & arrValueNames(i) & " - " & strValue  & vbcrlf
            if s = arrValueNames(i) and strValue = "Installed" then getInstalledOdbcDriverList = true
        Next
        dim odFinish: odFinish = now
        if len(s) = 0 then
            log "ODBC Providers: (took " & round((odFinish-odStart)*86400,3) & " seconds)"
            log OutText
            log ""
        end if
        set objRegistry = nothing
    end function

    '********************************************
    function BinaryToString(Binary,bHex)
        'Antonin Foller, http://www.motobit.com
        'Optimized version of a simple BinaryToString algorithm.

        Dim cl1, cl2, cl3, pl1, pl2, pl3
        Dim L,v
        cl1 = 1
        cl2 = 1
        cl3 = 1
        L = LenB(Binary)

        Do While cl1<=L
            v = AscB(MidB(Binary,cl1,1))
            if bHex then
                if v <> 0 then
                    pl3 = pl3 & right("00" & Hex(v),2)
                else
                    pl3 = pl3 & "00"
                end if
            else
                if v <> 0 then
                    pl3 = pl3 & Chr(v)
                else
                    pl3 = pl3 & "."
                end if
            end if
            cl1 = cl1 + 1
            cl3 = cl3 + 1
            If cl3>300 Then
                pl2 = pl2 & pl3
                pl3 = ""
                cl3 = 1
                cl2 = cl2 + 1
                If cl2>200 Then
                    pl1 = pl1 & pl2
                    pl2 = ""
                    cl2 = 1
                End If
            End If
        Loop
        BinaryToString = pl1 & pl2 & pl3
    end function

end class
