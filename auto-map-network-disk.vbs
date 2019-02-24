 
On Error Resume Next 
  
MapDrv "Z:", "\\LAPTOP\fileexchanger"  
 
Function MapDrv(DrvLet, UNCPath) 
 
    Dim WshNetwork      
    Dim Msg 
 
    Set WshNetwork = WScript.CreateObject("WScript.Network") 
 
    On Error Resume Next 
    WshNetwork.RemoveNetworkDrive DrvLet 
    WshNetwork.MapNetworkDrive DrvLet, UNCPath, "file", "file"  
     
    Select Case Err.Number 
        Case 0            
 
        Case -2147023694  
            WshNetwork.RemoveNetworkDrive DrvLet 
            WshNetwork.MapNetworkDrive DrvLet, UNCPath', '"file", "file" 
              
        Case -2147024811  
            WshNetwork.RemoveNetworkDrive DrvLet 
            WshNetwork.MapNetworkDrive DrvLet, UNCPath', '"file", "file"  
    End Select 
 
End Function 

