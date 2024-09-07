Do While True
  dim r
  randomize
  r = int(rnd*100)
  If r < 10 Then
    Set WshShell = CreateObject("WScript.Shell")

    scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

    Dim HTM, wWidth, wHeight
    Set HTM = CreateObject("htmlfile")
    wWidth = HTM.parentWindow.screen.Width
    wHeight = HTM.parentWindow.screen.Height

    WshShell.Run scriptDir & "/ffplay-gif.exe " & scriptDir & "/freddy.gif -x " & wWidth & " -y " & wHeight & " -autoexit -alwaysontop -noborder -window_title funny", 5

    WshShell.AppActivate("funny")
  ElseIf r = 99 Then
    Exit Do
  Else
    WScript.Sleep 5000
  End If
Loop