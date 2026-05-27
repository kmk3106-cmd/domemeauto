' control_panel 자동 실행 (ASCII 파일명 — 한글경로/인코딩 이슈 회피).
' Windows 로그인 시 Startup 폴더의 shortcut 으로 호출됨.
'
' 디버깅:
'   - logs\autostart_panel_run.log 에 매 실행 시점 + 결정(skip/launch) 기록.
'   - logs\autostart_panel_stdout.log / _stderr.log 에 패널 실행 출력 캡처.
'
' 등록: %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\domemeauto-panel.lnk
'   Target    = wscript.exe
'   Arguments = "C:\Users\USER\domemeauto\autostart_panel.vbs"

Option Explicit

Dim sh, fs, projectDir, pythonExe, logsDir, runLog, stdoutLog, stderrLog, stamp
Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

projectDir = "C:\Users\USER\domemeauto"
pythonExe  = "C:\Users\USER\PycharmProjects\PythonProject\.venv\Scripts\python.exe"
If Not fs.FileExists(pythonExe) Then pythonExe = "python"

logsDir = projectDir & "\logs"
If Not fs.FolderExists(logsDir) Then fs.CreateFolder(logsDir)
runLog    = logsDir & "\autostart_panel_run.log"
stdoutLog = logsDir & "\autostart_panel_stdout.log"
stderrLog = logsDir & "\autostart_panel_stderr.log"

stamp = Now()

Sub Log(msg)
    Dim f
    On Error Resume Next
    Set f = fs.OpenTextFile(runLog, 8, True)  ' 8=append, True=create if missing
    f.WriteLine "[" & stamp & "] " & msg
    f.Close
    On Error GoTo 0
End Sub

Log "VBS 호출됨 (autostart_panel.vbs)"

' --- 중복 실행 방지: 8001 LISTEN 체크 ---
Dim exec, result
On Error Resume Next
Set exec = sh.Exec("cmd /c netstat -ano -p tcp | findstr :8001 | findstr LISTENING")
result = ""
If Err.Number = 0 Then result = exec.StdOut.ReadAll()
On Error GoTo 0

If InStr(result, ":8001") > 0 Then
    Log "  -> 8001 이미 LISTEN -> 종료(중복방지)"
    WScript.Quit 0
End If

' --- 패널 실행 (콘솔창 없이, stdout/stderr 로그파일로 캡처) ---
' cmd /c 로 감싸서 redirection 사용. 0=vbHide, False=비동기(블로킹X)
Dim cmdLine
cmdLine = "cmd /c """"" & pythonExe & """ -u """ & projectDir & "\control_panel.py"" 1>>""" & stdoutLog & """ 2>>""" & stderrLog & """"""
Log "  -> 실행: " & cmdLine
sh.CurrentDirectory = projectDir
sh.Run cmdLine, 0, False

Log "  -> 호출 완료 (백그라운드)"
WScript.Quit 0
