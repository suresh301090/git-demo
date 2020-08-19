Set objWS = WScript.CreateObject("WScript.Shell")
   strLinkFile = "C:\Users\Me\Desktop\Git-Dir.LNK"
   Set objLink = objWS.CreateShortcut(strLinkFile)
   
   objLink.TargetPath = "C:\Program Files\Git\git-bash.EXE"
   '  objLink.Arguments = ""
   objLink.Description = "Current Git Directory"
   '  objLink.HotKey = "ALT+CTRL+F"
   '  objLink.IconLocation = "C:\Program Files\MyApp\MyProgram.EXE, 2"
   '  objLink.WindowStyle = "1"
   objLink.WorkingDirectory = "C:\Users\Me\Desktop\Git-demo\"
   objLink.Save