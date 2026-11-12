$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$wordProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if (-not $wordProcesses.Count) {
    '{"state":"closed"}'
    exit
}

$word = $null
$nativeWindow = $null
$document = $null
$body = $null
try {
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    }
    catch {
        # Word can expose its document window before it registers in the Running Object Table.
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class WordDocumentWindow {
    private delegate bool WindowCallback(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectID, ref Guid iid,
        [MarshalAs(UnmanagedType.IDispatch)] out object result);

    public static object Get(IntPtr parent) {
        object documentWindow = null;
        EnumChildWindows(parent, (window, parameter) => {
            var name = new StringBuilder(32);
            GetClassName(window, name, name.Capacity);
            if (name.ToString() != "_WwG") return true;
            var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
            return AccessibleObjectFromWindow(window, 0xfffffff0, ref dispatch, out documentWindow) != 0;
        }, IntPtr.Zero);
        return documentWindow;
    }
}
'@
        foreach ($process in $wordProcesses) {
            if ($process.MainWindowHandle -eq [IntPtr]::Zero) { continue }
            $nativeWindow = [WordDocumentWindow]::Get($process.MainWindowHandle)
            if ($null -ne $nativeWindow) { break }
        }
        if ($null -eq $nativeWindow) { throw 'Word document window unavailable' }
    }

    if ($null -ne $word) {
        if ($word.Documents.Count -eq 0) {
            '{"state":"closed"}'
            exit
        }
        $document = $word.ActiveDocument
    }
    else {
        $document = $nativeWindow.Document
    }
    $path = [string]$document.FullName
    if ($document.Saved -and $path.EndsWith('.docx', [StringComparison]::OrdinalIgnoreCase) -and
            [IO.File]::Exists($path)) {
        @{ state = 'saved'; path = $path } | ConvertTo-Json -Compress
        exit
    }

    $body = $document.Content
    $xml = [System.Xml.XmlDocument]::new()
    $xml.XmlResolver = $null
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($body.WordOpenXML)
    $namespaces = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $namespaces.AddNamespace('pkg', 'http://schemas.microsoft.com/office/2006/xmlPackage')
    $parts = @{}
    foreach ($name in @('document', 'footnotes', 'endnotes')) {
        $part = $xml.SelectSingleNode("/pkg:package/pkg:part[@pkg:name='/word/$name.xml']/pkg:xmlData/*", $namespaces)
        if ($null -ne $part) { $parts[$name] = $part.OuterXml }
    }
    if (-not $parts.ContainsKey('document')) { throw 'Word document XML unavailable' }
    @{ state = 'live'; path = $path; parts = $parts } | ConvertTo-Json -Compress -Depth 4
}
catch {
    '{"state":"busy"}'
}
finally {
    foreach ($object in @($body, $document, $nativeWindow, $word)) {
        if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
        }
    }
}
