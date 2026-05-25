' 패널 자동 실행 (콘솔창 안 띄움 — Windows 시작프로그램 등록용)
' 로그인 시 자동으로 control_panel.py 를 백그라운드 실행.
' 이미 8001 포트가 떠 있으면 중복 실행 안 함.
'
' 등록: shell:startup 폴더에 이 .vbs 의 shortcut 또는 본 파일 복사.

Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

projectDir = "C:\Users\USER\domemeauto"
pythonExe = "C:\Users\USER\PycharmProjects\PythonProject\.venv\Scripts\python.exe"
If Not fs.FileExists(pythonExe) Then pythonExe = "python"

' 이미 떠 있는지 간단 체크: 8001 포트 LISTEN 확인 (netstat). 떠 있으면 종료.
On Error Resume Next
Set exec = sh.Exec("cmd /c netstat -ano -p tcp | findstr :8001 | findstr LISTENING")
result = exec.StdOut.ReadAll()
On Error GoTo 0
If InStr(result, ":8001") > 0 Then
    ' 이미 실행 중 — 종료 (중복 방지)
    WScript.Quit 0
End If

' 콘솔창 없이 (vbHide=0) 백그라운드 실행
cmd = """" & pythonExe & """ -u """ & projectDir & "\control_panel.py"""
sh.CurrentDirectory = projectDir
sh.Run cmd, 0, False

WScript.Quit 0
