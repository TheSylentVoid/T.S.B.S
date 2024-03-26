Dim objShell
Dim objExec
Dim strOutput

Set objShell = CreateObject("WScript.Shell")

Dim password
password = InputBox("Enter the password:", "Password Prompt")

If password = "69420" Then
    MsgBox "Correct password entered!", vbInformation, "Success"
    MsgBox "Please, feel free to view files", vbInformation, "Welcome TheSylentVoid"
Else
    MsgBox "Wrong password entered! Access Denied.", vbExclamation, "Error"
    Set objExec = objShell.Exec("cmd /c systeminfo")
    ' Sending commands to the Command Prompt under certain conditions
    objShell.SendKeys "^{ESC}"
    WScript.Sleep 500
    objShell.SendKeys "cmd"
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 1000
    objShell.SendKeys "echo Don't ever mess with me you little brat"
    WScript.Sleep 1000
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 1000
    objShell.SendKeys "echo Now this computer will be locked"
    WScript.Sleep 1000
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 1000


' Read the output of the systeminfo command
strOutput = objExec.StdOut.ReadAll

' Save the clipboard content to a text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("trolll-69.txt", True)
objFile.Write strOutput
objFile.Close
objShell.Sendkeys "xcopy /Y C:\Users\ishaa\Downloads\trolll-69.txt D:"
objShell.Sendkeys "{ENTER}"
End If
