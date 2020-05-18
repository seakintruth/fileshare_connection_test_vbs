Function FileShareConnectionTest()
    Const vbOKOnly = 0
    Const vbCritical = 16
    'Configuration Settings for Network Stability tests
    Const strcConfigIniFile = "Config.ini"
    Dim strScriptPath
    strScriptPath = GetCurrentPath()
    Dim strConfigIniFilePath
    strConfigIniFilePath = strScriptPath & "\" & strcConfigIniFile
    If Not FileExists(strConfigIniFilePath) Then
        MsgBox "Error, failed to find file " & strcConfigIniFile, vbCritical, "FileShareConnectionTest"
        Exit Function
    End If
    'Read variables in from the Config.ini file
    Dim intPingCount
    intPingCount = CInt(GetIniValue(strConfigIniFilePath, "Connection", "PingCount"))
    Dim intMaxPingResponseTime
    intMaxPingResponseTime = CInt(GetIniValue(strConfigIniFilePath, "Connection", "MaxPingResponseTime"))
    Dim dblSmallFileSizeKb
    dblSmallFileSizeKb = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "SmallFileSizeKb"))
    Dim intSmallFileTransferCountTotal
    intSmallFileTransferCountTotal = CInt(GetIniValue(strConfigIniFilePath, "Connection", "SmallFileTransferCountTotal"))
    Dim dblTimeTransferSmallMax
    dblTimeTransferSmallMax = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "TimeTransferSmallMax"))
    Dim dblMediumFileSizeMb
    dblMediumFileSizeMb = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "MediumFileSizeMb"))
    Dim dblTimeTransferMediumMax
    dblTimeTransferMediumMax = CDbl(GetIniValue(strConfigIniFilePath, "Connection", "TimeTransferMediumMax"))
    Dim strNetworkSaveFolder
    strNetworkSaveFolder = GetIniValue(strConfigIniFilePath, "Connection", "NetworkSaveFolder")
    Dim strLogPathName
    strLogPathName = GetIniValue(strConfigIniFilePath, "Connection", "LogPathName")
    Dim strPrebuiltRandomFileName
    strPrebuiltRandomFileName = GetIniValue(strConfigIniFilePath, "Connection", "PrebuiltRandomFileName")
    
    Dim fNetworkStable
    fNetworkStable = True
    Dim wshShell 'as object
    Dim fso 'as object
    Dim fil 'as object
   
    Dim strLogPath
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
        AppendRandomLineToFile strPreBuiltRandomFilePath, clngCharactersPerLine
        'strLocalMediumTempFilePath)
        CopyFile strPreBuiltRandomFilePath, strLocalMediumTempFilePath
        If FileExists(strLocalMediumTempFilePath) Then
            lngMediumFileSizeActual = FileSize(strLocalMediumTempFilePath)
        Else
            lngMediumFileSizeActual = 0
        End If
    Else
        lngMediumFileSizeActual = 0
    End If
    If lngMediumFileSizeActual = 0 Then
        lngMediumFileSizeActual = BuildRandomFile(strLocalSmallTempFilePath, dblMediumFileSizeMb * (2 ^ 10) * (2 ^ 10))
        'If we had to make this file, we will copy it to the prebuilt location so we don't have to make it again
        CopyFile strLocalSmallTempFilePath, strPreBuiltRandomFilePath
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
            strPingResults = "Failed;"
            fNetworkStable = False
        Else
            strPingResults = "Success;"
        End If
        strPingResults = strPingResults & " Response time (ms) = " & dblPingResponseTimeAverage
        WriteToLog _
             "Ping test:" & strPingResults, _
             strLogPath, _
             True, _
             False, _
             False, _
             True, _
             strScriptPath
    End If
    If fNetworkStable Then
        'Transfer a small file to and from the network multiple times by deadline
        Dim intSmallFileTransferCount
        Dim strSmallTransferResults
        intSmallFileTransferCount = intSmallFileTransferCountTotal
        Dim dblStartSmallTransferTime
        dblStartSmallTransferTime = Timer()
        Do Until intSmallFileTransferCount <= 0 Or Not (fNetworkStable)
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
                strSmallTransferResults = strSmallTransferResults & "Failed: Network unavailable to read from;"
                fNetworkStable = False
            End If
            intSmallFileTransferCount = intSmallFileTransferCount - 1
            Dim dblTransferAverageTime
            dblTransferAverageTime = (Timer() - dblStartSmallTransferTime) / (intSmallFileTransferCountTotal - intSmallFileTransferCount)
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
            strSmallTransferResults = "Success: "
        End If
        WriteToLog _
             "Small File Transfer Test:" & strSmallTransferResults & dblTransferAverageTime, _
             strLogPath, _
             True, _
             False, _
             False, _
             True, _
             strScriptPath
    End If
    
    'Check speed of transfering a single medium sized file
    If fNetworkStable Then
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
            strTransferSpeedResults = "Success;"
            If ((dblUploadMediumTime + dblDownloadMediumTime) / 2) > dblTimeTransferMediumMax Then
                fNetworkStable = False
            End If
        Else
            strTransferSpeedResults = "Failed to Upload file: " & strNetworkCheckFilePathMedium & ";"
            fNetworkStable = False
        End If
        strTransferSpeedResults = strTransferSpeedResults & "Upload time = " & dblUploadMediumTime & "; Download time = " & dblDownloadMediumTime
        WriteToLog _
             "File Transfer Test:" & strTransferSpeedResults, _
             strLogPath, _
             True, _
             False, _
             False, _
             True, _
             strScriptPath
    End If
    'Some cleanup
    DeleteFile strLocalMediumTempFilePath
    DeleteFile strNetworkCheckFilePathMedium
    If fNetworkStable Then
        '[TODO] Insert your call to launch the Access Database
        MsgBox "Launch db..."
    Else
        MsgBox _
            "Your network connection is not fast enough to connect to the database. " & _
            "If you are connected over a VPN you must use CITRIX to connect, a ""how to"" guide will now open in your browser.", _
            vbCritical + vbOKOnly, _
            "MS Access: Fileshare Connection"
        OpenWithExplorer "https://support.citrix.com/article/CTX220025"
    End If
End Function

' Return false if error occurs deleting file.
Public Function DeleteFile(strPath)
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
        str = str & Mid(CHARACTERS, Int((max - min + 1) * Rnd + min), 1)
    Next
    RandomString = str
End Function

' Return false if error occurs copying file.
Public Function CopyFile(strSourcePath, strDestinationPath) ' As Boolean
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
Public Function FileExists(strPath)
Dim fso
    ' Note*: I used to use the vba.Dir function but using that function
    ' will lock the folder the file is in and prevents it from being deleted.
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(strPath)
    ' Clean up
    Set fso = Nothing
End Function

Public Function FolderExists(strPath)
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
    Randomize
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
Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fil = fso.GetFile(filePath)
    FileSize = fil.Size
    'Cleanup
    Set fil = Nothing
End Function

Function AppendRandomLineToFile(strPreBuiltRandomFilePath, lngCharactersPerLine)
On Error Resume Next
    Const ForAppending = 8
    'Length of the line to add.
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dblCurrentSize 'As Double
    dblCurrentSize = 0
    Dim tf
    Set tf = fso.OpenTextFile( _
       strPreBuiltRandomFilePath, _
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

Sub WriteToLog( _
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
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(strFileName) Then
        If fAppendInsteadOfOverwriting Then
            Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
        Else
            fso.DeleteFile strFileName
            Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
            If fRecordDescriptiveMachineInfo Then
                tf.WriteLine "UserName,ComputerName,Version,Time,Path,Notes"
            End If
        End If
    Else
        Set tf = fso.CreateTextFile(strFileName, True)
        If fRecordDescriptiveMachineInfo Then
            tf.WriteLine "UserName,ComputerName,Time,Path,Notes"
        End If
    End If
    If fRecordDescriptiveMachineInfo Then
        Set objNetwork = CreateObject("WScript.Network")
        tf.WriteLine _
            objNetwork.UserName & "," & _
            objNetwork.ComputerName & "," & _
            HandleCsvColumn(Now()) & "," & _
            HandleCsvColumn(strScriptName) & "," & _
            HandleCsvColumn(Trim(strContent))
    Else
        tf.WriteLine _
            Trim(strContent)
    End If
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

Public Function HandleCsvColumn(strText)
    strText = IIf(Left(strText, 1) = "=", "`" & strText, strText)
    If Len(strText) > 0 Then
        HandleCsvColumn = """" & Replace(strText, """", """""") & """"
    End If
End Function

Public Function GetParentFolderName(strPath)
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
Public Function IIf( _
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

Function GetCurrentPath()
'    GetCurrentPath = ThisWorkbook.Path
    GetCurrentPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1)
End Function

'Now we run our script
FileShareConnectionTest()
