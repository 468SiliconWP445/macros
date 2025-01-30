Sub GetProcessorsName()
    Dim objWMIService As Object
    Dim objProcessor As Object
    Dim objItem As Object
    Dim cpuName As String

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    Set objProcessor = objWMIService.ExecQuery("Select * from Win32_Processor")

    For Each objItem In objProcessor
        cpuName = objItem.Name
    Next objItem

    MsgBox "CPU Name: " & cpuName
    
    Range("A1").Value = "CPU:"
    Range("B1").Value = cpuName
    
    Set objGPU = objWMIService.ExecQuery("Select * from Win32_VideoController")

    For Each objItem In objGPU
        gpuName = objItem.Name
    Next objItem

    MsgBox "GPU Name: " & gpuName
    
    Range("A2").Value = "GPU:"
    Range("B2").Value = gpuName
End Sub
