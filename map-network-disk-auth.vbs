On error resume next
ServerShare = "\\laptop\fileexchanger"
UserName = "accountant"
Password = "strongpassword"

Set NetworkObject = CreateObject("WScript.Network")
Set FSO = CreateObject("Scripting.FileSystemObject")

NetworkObject.MapNetworkDrive "X:", ServerShare, False, UserName, Password

