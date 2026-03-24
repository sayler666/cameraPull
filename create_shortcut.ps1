$shell = New-Object -COM WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcut = $shell.CreateShortcut("$desktop\CameraPull.lnk")
$shortcut.TargetPath = "uv"
$shortcut.Arguments = "run camera_pull.py"
$shortcut.WorkingDirectory = "d:\OneDrive\Projects\Code\python\cameraPull"
$shortcut.WindowStyle = 1
$shortcut.IconLocation = "$env:SystemRoot\System32\DDORes.dll,86"
$shortcut.Save()
