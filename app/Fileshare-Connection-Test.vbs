Option Explicit
Dim strImRunningConnectionTestFilePath

' Uncomment next line, commented out for testing in an excel workbook IDE
FileShareConnectionTest(GetCurrentPath() & "\" & "Config.ini") ' Now we run our script

Function GetCurrentPath()
'    GetCurrentPath = ThisWorkbook.Path ' For testing in an excel workbook IDE
    GetCurrentPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1)
End Function

Function FileShareConnectionTest(strConfigIniFilePath)
    Dim strScriptPath
    strScriptPath = GetCurrentPath()
    If Not (CreateTmpImRunningFile(strScriptPath)) Then
        Exit Function
    End If
    Const vbOKOnly = 0
    Const vbCritical = 16
    'Configuration Settings for Network Stability tests
    If Not FileExists(strConfigIniFilePath) Then
        MsgBox "Error, failed to find file " & strConfigIniFilePath, vbCritical, "FileShareConnectionTest"
        FileShareConnectionTest = False
        Exit Function
    End If
    
    'Read variables in from the Config.ini file
    Dim strConnectionCheckVersion
    strConnectionCheckVersion = GetIniValue(strConfigIniFilePath, "Connection", "ConnectionCheckVersion")
    
    ' --- PINGS ---
    Dim intPingCount
    intPingCount = CInt(GetIniValue(strConfigIniFilePath, "Connection", "PingCount"))
    Dim intMaxPingResponseTime
    intMaxPingResponseTime = CInt(GetIniValue(strConfigIniFilePath, "Connection", "MaxPingResponseTime"))
    Dim dblSmallFileSizeKb
    
    ' --- Small file transfers ---
    dblSmallFileSizeKb = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "SmallFileSizeKb"))
    Dim intSmallFileTransferCountTotal
    intSmallFileTransferCountTotal = CInt(GetIniValue(strConfigIniFilePath, "Connection", "SmallFileTransferCountTotal"))
    Dim dblTimeTransferSmallMax
    dblTimeTransferSmallMax = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "TimeTransferSmallMax"))
    
    ' --- Medium file transfers ---
    Dim fPerformMediumFileTranfserTest 'As Boolean
    fPerformMediumFileTranfserTest = (GetIniValue(strConfigIniFilePath, "Connection", "fPerformMediumFileTranfserTest") = "True")
    Dim dblMediumFileSizeMb
    dblMediumFileSizeMb = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "MediumFileSizeMb"))
    Dim dblTimeTransferMediumMax
    dblTimeTransferMediumMax = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "TimeTransferMediumMax"))
    
    ' --- LOG Info ---
    Dim strNetworkSaveFolder
    strNetworkSaveFolder = GetIniValue(strConfigIniFilePath, "Connection", "NetworkSaveFolder")
    Dim strLogPathName
    strLogPathName = GetIniValue(strConfigIniFilePath, "Connection", "LogPathName")
    Dim strPrebuiltRandomFileName
    strPrebuiltRandomFileName = GetIniValue(strConfigIniFilePath, "Connection", "PrebuiltRandomFileName")
    Dim strCitrixUrl
    strCitrixUrl = GetIniValue(strConfigIniFilePath, "Connection", "CitrixUrl")

    'Build file path variables
    Dim fNetworkStable
    fNetworkStable = True
    Dim wshShell 'as object
    Dim fso 'as object
    Dim fil 'as object
    Dim strLogPath
    Dim strLogResults
    strLogPath = IIf( _
        FolderExists(strNetworkSaveFolder), _
        strNetworkSaveFolder & "\" & strLogPathName, _
        strScriptPath & "\" & strLogPathName _
    )
    Dim strPreBuiltRandomFilePath
    strPreBuiltRandomFilePath = strScriptPath & "\" & strPrebuiltRandomFileName
    Dim dblStartTimeSmall 'as double (really single)
    Dim dblStartTimeMedium 'as double (really single)
    Dim dblEndTimeSmall 'as double (really single)
    Dim dblEndTimeMedium 'as double (really single)
    Dim dblSmallFileSize 'as double
    Dim dblMediumFileSize 'as double
    Set wshShell = CreateObject("WScript.Shell") ' New IWshRuntimeLibrary.wshShell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strSmallFileName
    Dim strMediumFileName
    strSmallFileName = "RND_" & RemoveFileNameInvalidCharacters(RandomString(20)) & ".txt"
    strMediumFileName = "RND_" & RemoveFileNameInvalidCharacters(RandomString(21)) & ".txt"
    Dim strNetworkCheckFilePathSmall 'as string 100kb file
    strNetworkCheckFilePathSmall = strNetworkSaveFolder & "\" & strSmallFileName
    Dim strNetworkCheckFilePathMedium 'as string 2mb file
    strNetworkCheckFilePathMedium = strNetworkSaveFolder & "\" & strMediumFileName
    Dim strLocalSmallTempFilePath  'as string
    strLocalSmallTempFilePath = wshShell.ExpandEnvironmentStrings("%temp%") & "\" & strSmallFileName
    Dim strLocalMediumTempFilePath  'as string
    strLocalMediumTempFilePath = wshShell.ExpandEnvironmentStrings("%temp%") & "\" & strMediumFileName
    Dim lngSmallFileSizeActual
    dblStartTimeSmall = Timer()
    lngSmallFileSizeActual = BuildRandomFile(strLocalSmallTempFilePath, dblSmallFileSizeKb * (2 ^ 10))
    dblEndTimeSmall = Timer()
    Dim lngMediumFileSizeActual
    dblStartTimeMedium = Timer()
    If FileExists(strPreBuiltRandomFilePath) Then
        'Building a 2 MB, just takes too long so we try appending a line to the prebuilt file to make the file unique
        AppendRandomLineToFile strPreBuiltRandomFilePath, strLocalMediumTempFilePath, 500
        'strLocalMediumTempFilePath)
        If FileExists(strLocalMediumTempFilePath) Then
            lngMediumFileSizeActual = FileSize(strLocalMediumTempFilePath)
        Else
            lngMediumFileSizeActual = 0
        End If
    Else
        lngMediumFileSizeActual = 0
    End If
    If lngMediumFileSizeActual = 0 Then
        lngMediumFileSizeActual = BuildRandomFile(strLocalMediumTempFilePath, dblMediumFileSizeMb * (2 ^ 10) * (2 ^ 10))
        'If we had to make this file, we will copy it to the prebuilt location so we don't have to make it again
        CopyFile strLocalMediumTempFilePath , strPreBuiltRandomFilePath
    End If
    dblEndTimeMedium = Timer()
'    MsgBox "Build Random Files:" & dblEndTimeSmall - dblStartTimeSmall + dblEndTimeMedium - dblStartTimeMedium
    '=============================
    'Begin network stability tests
    '=============================
    'Perform pings of domain name controllers, check it's responce time
    If fNetworkStable Then
        Dim dblPingResponseTimeAverage
        Dim hostname
        hostname = GetAnyDcName()
        Dim strPingResults
        dblPingResponseTimeAverage = PingResponseTimeAverage(hostname, intPingCount)
        If dblPingResponseTimeAverage > intMaxPingResponseTime Then
            strPingResults = "Failed: Time exceeded limit " & intMaxPingResponseTime & ";"
            fNetworkStable = False
        Else
            strPingResults = "Success;"
        End If
        strLogResults = strPingResults & hostname & ";" & dblPingResponseTimeAverage & ";" & intPingCount & ";"
    End If
    If fNetworkStable Then
        'Transfer a small file to and from the network multiple times by deadline
        Dim intSmallFileTransferCount
        Dim strSmallTransferResults
        intSmallFileTransferCount = 0
        Dim dblStartSmallTransferTime
        dblStartSmallTransferTime = Timer()
        Do Until intSmallFileTransferCount >= intSmallFileTransferCountTotal Or Not (fNetworkStable)
            intSmallFileTransferCount = intSmallFileTransferCount + 1
            CopyFile strLocalSmallTempFilePath, strNetworkSaveFolder & "\" & strSmallFileName
            If FileExists(strLocalSmallTempFilePath) Then
                DeleteFile strLocalSmallTempFilePath
            Else
                strSmallTransferResults = strSmallTransferResults & "Failed: Network unavailable to write to;"
                fNetworkStable = False
            End If
            CopyFile strNetworkSaveFolder & "\" & strSmallFileName, strLocalSmallTempFilePath
            If FileExists(strNetworkSaveFolder & "\" & strSmallFileName) Then
                DeleteFile strNetworkSaveFolder & "\" & strSmallFileName
            Else
                strSmallTransferResults = strSmallTransferResults & "Failed: Network unavailable to read from:" & strNetworkSaveFolder & "\" & strSmallFileName & ";"
                fNetworkStable = False
            End If
            Dim dblTransferAverageTime
            dblTransferAverageTime = (Timer() - dblStartSmallTransferTime) / (intSmallFileTransferCount)
            If dblTransferAverageTime > dblTimeTransferSmallMax Then
                strSmallTransferResults = strSmallTransferResults & "Failed: Running average took too long;"
                fNetworkStable = False
            End If
        Loop
        'Some cleanup
        If FileExists(strLocalSmallTempFilePath) Then
            DeleteFile strLocalSmallTempFilePath
        End If
        If FileExists(strNetworkSaveFolder & "\" & strSmallFileName) Then
            DeleteFile strNetworkSaveFolder & "\" & strSmallFileName
        End If
        If fNetworkStable Then
            strSmallTransferResults = "Success;"
        End If
        strLogResults = strLogResults & strSmallTransferResults & dblTransferAverageTime & ";" & intSmallFileTransferCount & ";"
    Else
        strLogResults = strLogResults & ";;;"
    End If
    
    'Check speed of transfering a single medium sized file
    If fNetworkStable And fPerformMediumFileTranfserTest Then
        Dim strTransferSpeedResults
        Dim dblStartTimeMediumTransfer
        Dim dblUploadMediumTime
        If FileExists(strNetworkCheckFilePathMedium) Then
            DeleteFile strNetworkCheckFilePathMedium
        End If
        dblStartTimeMediumTransfer = Timer()
        'Test File Upload Time
        CopyFile strLocalMediumTempFilePath, strNetworkCheckFilePathMedium
        dblUploadMediumTime = Timer() - dblStartSmallTransferTime
        If FileExists(strNetworkCheckFilePathMedium) Then
            dblStartTimeMediumTransfer = Timer()
            If FileExists(strLocalMediumTempFilePath) Then
                DeleteFile strLocalMediumTempFilePath
            End If
            'Test File Download Time
            CopyFile strNetworkCheckFilePathMedium, strLocalMediumTempFilePath
            Dim dblDownloadMediumTime
            dblDownloadMediumTime = Timer() - dblStartSmallTransferTime
            If ((dblUploadMediumTime + dblDownloadMediumTime) / 2) > dblTimeTransferMediumMax Then
                strTransferSpeedResults = "Failed: transfer took too long, exceeded " & dblTimeTransferMediumMax & ";"
                strLogResults = strLogResults & strTransferSpeedResults & dblUploadMediumTime & ";" & dblDownloadMediumTime & ";" & lngMediumFileSizeActual
                fNetworkStable = False
            Else
                strTransferSpeedResults = "Success;"
            End If
            strLogResults = strLogResults & strTransferSpeedResults & dblUploadMediumTime & ";" & dblDownloadMediumTime & ";" & lngMediumFileSizeActual
        Else
            strTransferSpeedResults = "Failed to Upload file: " & strNetworkCheckFilePathMedium & ";"
            strLogResults = strLogResults & strTransferSpeedResults & ";;" & lngMediumFileSizeActual
            fNetworkStable = False
        End If
    Else
        strLogResults = strLogResults & ";;;" & lngMediumFileSizeActual
    End If
    WriteSpeedTestToLog _
        strConnectionCheckVersion & ";" & strLogResults, _
        strLogPath, _
        True, _
        False, _
        False, _
        True, _
        strScriptPath
    'Some cleanup
    DeleteFile strLocalMediumTempFilePath
    DeleteFile strNetworkCheckFilePathMedium
    If fNetworkStable Then
        FileShareConnectionTest = True
    Else
        FileShareConnectionTest = False
        Dim strCitrixGuidePath
        strCitrixGuidePath = strScriptPath & "\" & "documentation" & "\" & "Use CITRIX to Run Access Database.pdf"
        If FileExists(strCitrixGuidePath) Then
            OpenWithExplorer strCitrixGuidePath
            MsgBox _
                "Your network connection is not fast enough to connect to the database. " & _
                "If you are connected over a VPN you must use CITRIX to connect, a ""how to"" guide and Citrix will now open.", _
                vbCritical + vbOKOnly, _
                "MS Access: Fileshare Connection"
        Else
            MsgBox _
                "Your network connection is not fast enough to connect to the database. " & _
                "If you are connected over a VPN you must use CITRIX to connect, Citrix will now open.", _
                vbCritical + vbOKOnly, _
                "MS Access: Fileshare Connection"
        End If
        wshShell.Exec ( _
            """" & wshShell.ExpandEnvironmentStrings("%programfiles%") & "\" & _
            "Internet Explorer\iexplore.exe""" & _
            " " & strCitrixUrl _
        )
    End If
    DeleteFile strImRunningConnectionTestFilePath
End Function

Function CreateTmpImRunningFile(strWorkingFolder)
On Error Resume Next
    strImRunningConnectionTestFilePath = strWorkingFolder & "\" & "FileConnectionCheckRunning.txt"
    If FileExists(strImRunningConnectionTestFilePath) Then
        'If the I'm running file is greater than 90 seconds then attempt to delete it
        If ((Now() - FileCreated(strImRunningConnectionTestFilePath)) * 24 * 60 * 60) > 90 Then
            DeleteFile strImRunningConnectionTestFilePath
        Else
            CreateTmpImRunningFile = False
            Exit Function
        End If
    End If
    Touch strImRunningConnectionTestFilePath
    CreateTmpImRunningFile = (Err.Number = 0)
End Function

Sub Touch(strFilePath)
    Dim tf
    Set tf = CreateObject("Scripting.FileSystemObject").CreateTextFile( _
       strFilePath, _
       True _
    )
    tf.WriteLine vbNullString
    tf.Close
    'cleanup
    Set tf = Nothing
End Sub

' Return false if error occurs deleting file.
Function DeleteFile(strPath)
On Error Resume Next
Dim fso 'As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Err.Clear
    fso.DeleteFile strPath, True
    DeleteFile = (Err.Number = 0)
    ' Clean up
    Set fso = Nothing
End Function

Function RandomString(strLen)
'Modified from https://stackoverflow.com/a/7417797
    Dim str, min, max, i
    Const CHARACTERS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWQYZ0123456789`!@#$%^&*()_+~[]\;',./{}|:""<>?"
    min = 1
    max = Len(CHARACTERS)
    For i = 1 To strLen
        Randomize (Timer()*1000)
        str = str & Mid(CHARACTERS, Int((max - min + 1) * rnd(1) + min), 1)
    Next
    RandomString = str
End Function

' Return false if error occurs copying file.
Function CopyFile(strSourcePath, strDestinationPath) ' As Boolean
On Error Resume Next
Dim fso 'As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Err.Clear
    fso.CopyFile strSourcePath, strDestinationPath, True
    CopyFile = (Err.Number = 0)
    ' Clean up
    Set fso = Nothing
End Function

' Return true if file exists and false if file does not exist
Function FileExists(strPath)
Dim fso
    ' Note*: I used to use the vba.Dir function but using that function
    ' will lock the folder the file is in and prevents it from being deleted.
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(strPath)
    ' Clean up
    Set fso = Nothing
End Function

Function FolderExists(strPath)
Dim fso
    ' Note*: I used to use the vba.Dir function but using that function
    ' will lock the folder and prevent it from being deleted.
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(strPath)
    ' Clean up
    Set fso = Nothing
End Function

Function RemoveFileNameInvalidCharacters(strFileName) 'As String
    Const vbBinaryCompare = 0
    ' Function used to remove the following invalid characters from file names:
    ' < (less than)
    ' > (greater than)
    ' : (colon)
    ' " (double quote)
    ' / (forward slash)
    ' \ (backslash)
    ' | (vertical bar or pipe)
    ' ? (question mark)
    ' * (asterisk)
    strFileName = Replace(strFileName, "<", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, ">", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, ":", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, """", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, "/", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, "\", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, "|", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, "?", "_", 1, -1, vbBinaryCompare)
    strFileName = Replace(strFileName, "*", "_", 1, -1, vbBinaryCompare)
    RemoveFileNameInvalidCharacters = strFileName
End Function

'Builds a random file of minimum file size, will allways be slightly larger in size, returns completed file size
'Building a random file larger than 500KB with this method is simply too slow.
Function BuildRandomFile(strFilePath, dblFileSizeMinimum)
    Const ForAppending = 8
    'Length of Lines
    'Each character is a byte, and the cariage return line feed at the end of each line is two bytes.
    Const clngCharactersPerLine = 500
    'Number of lines to write between each file size check ~ 341 lines of 300 characters is 100KB
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dblCurrentSize 'As Double
    dblCurrentSize = 0
    
    Dim strOut 'as string
    strOut = ""
    Do
        strOut = strOut & RandomString(clngCharactersPerLine)
        dblCurrentSize = dblCurrentSize + clngCharactersPerLine
    Loop Until dblCurrentSize > dblFileSizeMinimum
    Dim tf
    'Write what's in memory to the file
    Set tf = fso.CreateTextFile( _
       strFilePath, _
       True _
    )
    tf.WriteLine (strOut)
    tf.Close
    
    BuildRandomFile = fso.GetFile(strFilePath).Size
End Function

Function FileSize(filePath)
On Error Resume Next
    FileSize = CreateObject("Scripting.FileSystemObject").GetFile(filePath).Size
End Function

Function FileCreated(filePath)
On Error Resume Next
    FileCreated = CreateObject("Scripting.FileSystemObject").GetFile(filePath).DateCreated
End Function

Function AppendRandomLineToFile(strPreBuiltRandomFilePath, strTmpDestination, lngCharactersPerLine)
On Error Resume Next
    CopyFile strPreBuiltRandomFilePath, strTmpDestination
    Const ForAppending = 8
    'Length of the line to add.
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dblCurrentSize 'As Double
    dblCurrentSize = 0
    Dim tf
    Set tf = fso.OpenTextFile( _
       strTmpDestination, _
       ForAppending, _
       True _
    )
    tf.WriteLine (RandomString(lngCharactersPerLine))
    tf.Close

    'Cleanup
    Set tf = Nothing
End Function

Function PingResponseTime(hostname) ' As Double
    'Modified from https://stackoverflow.com/a/25160410
    'You can use also name of computer
    ' Return TRUE, if ping was successful
    'Details of Win32_PingStatus at
    'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus
    Dim oPingResult ' As Variant
    For Each oPingResult In GetObject("winmgmts://./root/cimv2").ExecQuery _
        ("SELECT * FROM Win32_PingStatus WHERE Address = '" & hostname & "'")
        If IsObject(oPingResult) Then
            If oPingResult.StatusCode = 0 Then
                PingResponseTime = oPingResult.ResponseTime
                'Debug.Print "ResponseTime", oPingResult.ResponseTime 'You can also return ping time
               Exit Function
            Else
                PingResponseTime = 9999
            End If
        End If
    Next
End Function

Function PingResponseTimeAverage(hostname, intNumberOfPings)
'Can't use this ping to check what our ping to the file share would be, it's not using reachback, and allways takes one second + ping time?
    Dim dblStartTime
    dblStartTime = Timer()
    Dim ResponseTime
    ResponseTime = 0
    Dim ResponseTimeAverage
    Dim intPingCount
    intPingCount = intNumberOfPings
    Do Until intPingCount <= 0
        ResponseTime = ResponseTime + PingResponseTime(hostname)
        intPingCount = intPingCount - 1
    Loop
    ResponseTimeAverage = ResponseTime / intNumberOfPings
    PingResponseTimeAverage = ResponseTimeAverage
End Function

Function GetAnyDcName()
On Error Resume Next
    Dim oSysInfo 'As Object
    Set oSysInfo = CreateObject("AdSystemInfo")
    Dim strDomainControllerCommonName 'As String
    strDomainControllerCommonName = oSysInfo.GetAnyDcName
    GetAnyDcName = strDomainControllerCommonName
End Function

Function OpenWithExplorer(strFilePath)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    'Cleanup
    Set wshShell = Nothing
End Function

Sub WriteSpeedTestToLog( _
     strContent, _
     strFileName, _
     fRecordDescriptiveMachineInfo, _
     fOpenFile, _
     fOpenWithExplorerViceNotepad, _
     fAppendInsteadOfOverwriting, _
     strScriptName _
)
    Dim fso ' As Object
    Dim tf 'As Object
    Dim objNetwork 'As Object
    Dim aryContent
    Dim strHeaders
    If fRecordDescriptiveMachineInfo Then
        strHeaders = _
            "UserName,ComputerName,Time,Path,Version," & _
            "Ping Test Status,Host Pinged,Ping Response Time Average(ms),Ping Count," & _
            "Small File Transfer Test Status, Small File Transfer Average(ms),Small File Transfer Count," & _
            "Medium File Transfer Test Status, Medium File Upload Time(sec),Medium File Download Time(sec),Medium File Size(bytes)"
    Else
        strHeaders = _
            "Version,Ping Test Status,Host Pinged,Ping Response Time Average(ms),Ping Count," & _
            "Small File Transfer Test Status, Small File Transfer Average(ms),Small File Transfer Count," & _
            "Medium File Transfer Test Status, Medium File Upload Time(sec),Medium File Download Time(sec),Medium File Size(bytes)"
    End If
    aryContent = Split(strContent, ";")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Const dblLongWaitForLogMaxSeconds = 2
    Dim dblStartWriteToLog
    dblStartWriteToLog = Timer()
    Dim strLineToWrite

    If fRecordDescriptiveMachineInfo Then
        Set objNetwork = CreateObject("WScript.Network")
        strLineToWrite = _
            objNetwork.UserName & "," & _
            objNetwork.ComputerName & "," & _
            HandleCsvColumn(Now()) & "," & _
            HandleCsvColumn(strScriptName) & "," 
    End If 
    strLineToWrite = strLineToWrite & _ 
        HandleCsvColumn(Trim(aryContent(0))) & "," & _
        HandleCsvColumn(Trim(aryContent(1))) & "," & _
        HandleCsvColumn(Trim(aryContent(2))) & "," & _
        HandleCsvColumn(Trim(aryContent(3))) & "," & _
        HandleCsvColumn(Trim(aryContent(4))) & "," & _
        HandleCsvColumn(Trim(aryContent(5))) & "," & _
        HandleCsvColumn(Trim(aryContent(6))) & "," & _
        HandleCsvColumn(Trim(aryContent(7))) & "," & _
        HandleCsvColumn(Trim(aryContent(8))) & "," & _
        HandleCsvColumn(Trim(aryContent(9))) & "," & _
        HandleCsvColumn(Trim(aryContent(10))) & "," & _
        HandleCsvColumn(Trim(aryContent(11)))
    Do
        On Error Resume Next
        Err.Clear
        If fso.FileExists(strFileName) Then
            If fAppendInsteadOfOverwriting Then
                Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
            Else
                fso.DeleteFile strFileName
                Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
                tf.WriteLine strHeaders
            End If
        Else
            Set tf = fso.CreateTextFile(strFileName, True)
            tf.WriteLine strHeaders
        End If
    Loop While (Err.Number = 70 And (Timer() - dblStartWriteToLog) < dblLongWaitForLogMaxSeconds)
    If Err.Number = 70 Then ' After 2 seconds of trying the file is still locked, make a new filename
        Dim wshShell
        Set wshShell = CreateObject("WScript.Shell")
        Dim tmpPath
        tmpPath = strFileName & RemoveFileNameInvalidCharacters(RandomString(7)) & ".txt"
        If FileExists(tmpPath) Then
            If fAppendInsteadOfOverwriting Then
                Set tf = fso.OpenTextFile(tmpPath, 8, True) ' 8 = ForAppending
            Else
                fso.DeleteFile strFileName
                Set tf = fso.OpenTextFile(tmpPath, 8, True) ' 8 = ForAppending
                tf.WriteLine strHeaders
            End If
        Else
            Set tf = fso.CreateTextFile(tmpPath, True)
            tf.WriteLine strHeaders
        End If
        Set wshShell = Nothing
    End If
    On Error GoTo 0
    tf.WriteLine strLineToWrite
    tf.Close
    
    'Clean up
    Set tf = Nothing
    Set fso = Nothing
    Set objNetwork = Nothing
    ' Open file
    If fOpenFile Then
        If fOpenWithExplorerViceNotepad Then
            OpenWithExplorer strFileName
        Else
            Shell "Notepad.exe " & strFileName, vbNormalFocus
        End If
    End If
End Sub

Function HandleCsvColumn(strText)
    strText = IIf(Left(strText, 1) = "=", "`" & strText, strText)
    If Len(strText) > 0 Then
        HandleCsvColumn = """" & Replace(strText, """", """""") & """"
    End If
End Function

Function GetParentFolderName(strPath)
On Error Resume Next
Dim fso 'As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName = fso.GetParentFolderName(strPath)
    ' Clean up
    Set fso = Nothing
End Function

Function GetIniValue(strFilePath, strSection, strKey)
Const ForReading = 1
Dim fso ' As Scripting.FileSystemObject
Dim tf ' As Scripting.TextStream
Dim strLine
Dim nEqualPos
Dim strLeftString
    GetIniValue = ""
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If fso.FileExists(strFilePath) Then
        Set tf = fso.OpenTextFile(strFilePath, ForReading, False)
        Do While tf.AtEndOfStream = False
            ' Continue with next line
            strLine = Trim(tf.ReadLine)
            ' Check if section is found in the current line
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                Do While tf.AtEndOfStream = False
                    ' Continue with next line
                    strLine = Trim(tf.ReadLine)
                    ' Abort loop if next section is reached
                    If Left(strLine, 1) = "[" Then
                        Exit Do
                    End If
                    ' Find position of equal sign in the line
                    nEqualPos = InStr(1, strLine, "=", 1)
                    If nEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, nEqualPos - 1))
                        ' Check if item is found in the current line
                        If LCase(strLeftString) = LCase(strKey) Then
                            GetIniValue = Trim(Mid(strLine, nEqualPos + 1))
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If
                Loop
                Exit Do
            End If
        Loop
        tf.Close
    End If
End Function

'Using our own IIf to modify the error handling of built in IIf
'The Built in IIf always evaluates both the truepart and falsepart
'this can cause unnessesary errors
'Additionally the multil line method has been tested to perform faster that the buitin IIf
Function IIf( _
    fExpression, _
    strTruePart, _
    strFalsePart _
)
    If fExpression Then
        IIf = strTruePart
    Else
        IIf = strFalsePart
    End If

End Function
