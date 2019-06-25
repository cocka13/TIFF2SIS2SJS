Attribute VB_Name = "modKontrola"
  DefLng A-Z

    Private Type STARTUPINFO
        cb              As Long
        lpReserved      As String
        lpDesktop       As String
        lpTitle         As String
        dwX             As Long
        dwY             As Long
        dwXSize         As Long
        dwYSize         As Long
        dwXCountChars   As Long
        dwYCountChars   As Long
        dwFillAttribute As Long
        dwFlags         As Long
        wShowWindow     As Integer
        cbReserved2     As Integer
        lpReserved2     As Long
        hStdInput       As Long
        hStdOutput      As Long
        hStdError       As Long
    End Type


Private Type PROCESS_INFORMATION
        hProcess        As Long
        hThread         As Long
        dwProcessID     As Long
        dwThreadID      As Long
End Type

Public Type udtShellAndWait         'pass information here
        sCommand As String              'command line for Shell
        bShellAndWaitRunning As Boolean 'shell and wait is running
        bNoTerminate  As Boolean        'no forced termination [no DoEvents]
        lMilliseconds As Long           'interrupt this often in milliseconds [1000 is 1 second]
        bLogFile As Boolean             'do a log file
        bTerminated As Boolean          'true if terminated by ShellAndWaitTerminate
        tStart   As STARTUPINFO
        tProcess As PROCESS_INFORMATION 'above structures
End Type



    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

    Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
        lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long

    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
        ByVal uExitCode As Long) As Long


    Private Const NORMAL_PRIORITY_CLASS As Long = &H20
    Private Const INFINITE              As Long = -1
    Private Const WAIT_TIMEOUT          As Long = &H102

    Public gsApp As String

Public Function ShellAndWait(tShellAndWait As udtShellAndWait) As Boolean

    Dim lRtn    As Long
    Dim iFN     As Integer
    Dim lMilliseconds As Long

    With tShellAndWait
        .bTerminated = False                'set not terminated yet into structure
        lMilliseconds = .lMilliseconds      'local variable
        If lMilliseconds <= 0 Then          'zero or negative then use 1 Second
            lMilliseconds = mclMillisecondsDefault  'default
        End If
        If .bNoTerminate Then               'don't allow terminate
            .lMilliseconds = INFINITE       'set to never return
        End If
        .bShellAndWaitRunning = True        'started
                ' Initialize the STARTUPINFO structure:
        .tStart.cb = Len(.tStart)


        ' Start the shelled application:
        lRtn = CreateProcessA(0&, .sCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, .tStart, .tProcess)

        ' Wait for the shelled application to finish:
        Do
            lRtn = WaitForSingleObject(.tProcess.hProcess, lMilliseconds)
'wait milliseconds
            If lRtn <> WAIT_TIMEOUT Then
                Exit Do
            End If
            DoEvents                                'allow other processes
        Loop While True

        lRtn = CloseHandle(.tProcess.hProcess)
        If lRtn <> 0 Then
            ShellAndWait = True                     'report success
        End If
        .bShellAndWaitRunning = False               'ended
    End With
End Function

Public Function ShellAndWaitTerminate(tShellAndWait As udtShellAndWait) As Boolean
    Dim lRtn As Long

    With tShellAndWait
        lRtn = TerminateProcess(.tProcess.hProcess, "0")
        If lRtn <> 0 Then                           'success
            lRtn = CloseHandle(.tProcess.hProcess)  'close handle, don't know if this is really needed!
            .bTerminated = True                     'set terminated
            ShellAndWaitTerminate = True            'report success
        End If
    End With
End Function


