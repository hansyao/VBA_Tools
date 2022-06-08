VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDLLAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA Side-By-Side COM/ActiveX tool
''' author: Hans Yao<hansyao@gmail.com>
''' -----------------------------------------------
''' https://github.com/hansyao/VBA_Tools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Hans Yao
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

Private Const NODE_ELEMENT = 1
Private Const ssfDESKTOP = 0
Private Const adTypeBinary = 1
Private Const adSaveCreateOverWrite = 2
Private Const INVALID_HANDLE_VALUE = -1

'code page
Private Const CP_ACP = 0 ' default to ANSI code page
Private Const CP_UTF8 = 65001 ' default to UTF-8 code page
Private Const CP_GB2312 = 936
Private Const CP_UNICODE = 1200

Private Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
Private Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Private Const IID_IStream As String = "{0000000C-0000-0000-C000-000000000046}"
Private Const IID_IClassFactory = "{00000001-0000-0000-C000-000000000046}"
Private Const IID_IClassFactory2 = "{B196B28F-BAB4-101A-B69C-00AA00341D07}"

#If VBA7 Then
Private Declare PtrSafe Function CLSIDFromString Lib "OLE32" (ByVal lpszCLSID As LongPtr, pclsid As GUID) As Long
Private Declare PtrSafe Function CLSIDFromProgID Lib "OLE32" (ByVal lpszProgID As LongPtr, pclsid As GUID) As Long
Private Declare PtrSafe Function CoCreateInstanceEx Lib "OLE32" (rclsid As GUID, ByVal pUnkOuter As Long, ByVal dwClsContent As CLSCTX, pServerInfo As LongPtr, ByVal cmq As Long, rgmqResults As LongPtr) As Long
Private Declare PtrSafe Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As LongPtr, ByVal dwFlags As moduleFlags) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function LoadPackagedLibrary Lib "kernel32" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long

Private Declare PtrSafe Function FindResourceA Lib "kernel32" (ByVal hInstance As LongPtr, ByVal lpName As String, ByVal lpType As String) As LongPtr
Private Declare PtrSafe Function FindResourceW Lib "kernel32" (ByVal hInstance As LongPtr, ByVal lpName As String, ByVal lpType As String) As LongPtr
Private Declare PtrSafe Function SizeofResource Lib "kernel32" (ByVal hInstance As LongPtr, ByVal hResInfo As LongPtr) As Long
Private Declare PtrSafe Function LoadResource Lib "kernel32" (ByVal hInstance As LongPtr, ByVal hResInfo As LongPtr) As LongPtr
Private Declare PtrSafe Function LockResource Lib "kernel32" (ByVal hResData As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeResource Lib "kernel32" (ByVal hResData As LongPtr) As Long
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Declare PtrSafe Function CreateActCtx Lib "kernel32" Alias "CreateActCtxA" (ByRef pActCtx As ACTCTX_) As LongPtr
Private Declare PtrSafe Sub ReleaseActCtx Lib "kernel32" (ByVal hActCtx As LongPtr)
Private Declare PtrSafe Function ActivateActCtx Lib "kernel32" (ByVal hActCtx As LongPtr, ByRef lpCookie As LongPtr) As Boolean
Private Declare PtrSafe Function DeactivateActCtx Lib "kernel32" (ByVal dwFlags As Long, ByVal ulCookie As LongPtr) As Boolean
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetCurrentActCtx Lib "kernel32" (ByRef lphActCtx As LongPtr) As Boolean
Private Declare PtrSafe Sub AddRefActCtx Lib "kernel32" (ByRef lphActCtx As LongPtr)

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32 " (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long

Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

#Else

Private Declare Function CLSIDFromString Lib "OLE32" (ByVal lpszCLSID As Long, pclsid As GUID) As Long
Private Declare Function CLSIDFromProgID Lib "OLE32" (ByVal lpszProgID As Long, pclsid As GUID) As Long
Private Declare Function CoCreateInstanceEx Lib "OLE32" (rclsid As GUID, ByVal pUnkOuter As Long, ByVal dwClsContent As Long, pServerInfo As COSERVERINFO, ByVal cmq As Long, rgmqResults As MULTI_QI) As Long

Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As moduleFlags) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadPackagedLibrary Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function FindResourceA Lib "kernel32" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceW Lib "kernel32" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function CreateActCtx Lib "kernel32" Alias "CreateActCtxW" (ByRef pActCtx As ACTCTX_) As Long
Private Declare Sub ReleaseActCtx Lib "kernel32" (ByVal hActCtx As Long)
Private Declare Function ActivateActCtx Lib "kernel32" (ByVal hActCtx As Long, ByRef lpCookie As Long) As Boolean
Private Declare Function DeactivateActCtx Lib "kernel32" (ByVal dwFlags As Long, ByVal ulCookie As Long) As Boolean
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function MultiByteToWideChar Lib "kernel32 " (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function GetCurrentActCtx Lib "kernel32" (ByRef lphActCtx As Long) As Boolean
Private Declare Sub AddRefActCtx Lib "kernel32" (ByRef lphActCtx As Long)

#End If

Private Type MULTI_QI
#If VBA7 Then
    pIID As LongPtr
    pItf As IUnknown
    hr As Long
#Else
    pIID As Long
    pItf As IUnknown
    hr As Long
#End If
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type COSERVERINFO
#If VBA7 Then
    dwReserved1 As Long
    pwszName As LongPtr
    pAuthInfo As String
    dwReserved2 As Long
#Else
    dwReserved1 As Long
    pwszName As Long
    pAuthInfo As String
    dwReserved2 As Long
#End If
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibraryexa
Private Enum moduleFlags
    DONT_RESOLVE_DLL_REFERENCES = &H1
    LOAD_IGNORE_CODE_AUTHZ_LEVEL = &H10
    LOAD_LIBRARY_AS_DATAFILE = &H2
    LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = &H40
    LOAD_LIBRARY_AS_IMAGE_RESOURCE = &H20
    LOAD_LIBRARY_SEARCH_APPLICATION_DIR = &H200
    LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = &H1000
    LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = &H100
    LOAD_LIBRARY_SEARCH_SYSTEM32 = &H800
    LOAD_LIBRARY_SEARCH_USER_DIRS = &H400
    LOAD_WITH_ALTERED_SEARCH_PATH = &H8
    LOAD_LIBRARY_REQUIRE_SIGNED_TARGET = &H80
    LOAD_LIBRARY_SAFE_CURRENT_DIRS = &H2000
End Enum

' // This represents the DLL used to hold a hosted service
Private Type PSVCHOST_DLL
    hModule As LongPtr
    pszDllPath As String
    pszManifestPath As String
    hActCtx As LongPtr
End Type


'https://docs.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-actctxa
Private Type ACTCTX_
    cbSize As Long
    dwFlags As ACTCTX_FLAG
    lpSource As String
    wProcessorArchitecture As Integer
    wLangId As Integer
    lpAssemblyDirectory As String
    lpResourceName As String
    lpApplicationName As String
    hModule As LongPtr
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-actctxa
Private Enum ACTCTX_FLAG
    ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID = 1
    ACTCTX_FLAG_LANGID_VALID = 2
    ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 4
    ACTCTX_FLAG_RESOURCE_NAME_VALID = 8
    ACTCTX_FLAG_SET_PROCESS_DEFAULT = 16
    ACTCTX_FLAG_APPLICATION_NAME_VALID = 32
    ACTCTX_FLAG_HMODULE_VALID = 128
End Enum

Public Enum CLSCTX
    CLSCTX_INPROC_SERVER = &H1
    CLSCTX_INPROC_HANDLER = &H2
    CLSCTX_LOCAL_SERVER = &H4
    CLSCTX_INPROC_SERVER16 = &H8
    CLSCTX_REMOTE_SERVER = &H10
    CLSCTX_INPROC_HANDLER16 = &H20
    CLSCTX_RESERVED1 = &H40
    CLSCTX_RESERVED2 = &H80
    CLSCTX_RESERVED3 = &H100
    CLSCTX_RESERVED4 = &H200
    CLSCTX_NO_CODE_DOWNLOAD = &O400
    CLSCTX_RESERVED5 = &H800
    CLSCTX_NO_CUSTOM_MARSHAL = &H1000
    CLSCTX_ENABLE_CODE_DOWNLOAD = &H2000
    CLSCTX_NO_FAILURE_LOG = &H4000
    CLSCTX_DISABLE_AAA = &H8000
    CLSCTX_ENABLE_AAA = &H10000
    CLSCTX_FROM_DEFAULT_CONTEXT = &H20000
    CLSCTX_INPROC = CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER
    CLSCTX_SERVER = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER
    CLSCTX_ALL = CLSCTX_SERVER Or CLSCTX_INPROC_HANDLER
End Enum

Private shell32 As Object
Private DOMDocument As Object
Private Stream As Object
Private RegRead As New clsRegRead
Private TypeLibInfo As New TLI.TypeLibInfo
Private dict As Object
Private XDoc As Object

Private mManifestXML As String
Private mManifestFile As String
Private mManifestDict As Object
Private mDllFile As String
Private actStatus As Boolean
#If VBA7 Then
Private hModule As LongPtr
Private actCookie As LongPtr
Private actHandle As LongPtr
#Else
Private hModule As Long
Private actCookie As Long
Private actHandle As Long
#End If

Private current_senario As Integer

Public Property Get dllFile() As String
    dllFile = mDllFile
End Property

Public Property Let dllFile(ByVal inputDllFile As String)
    mDllFile = inputDllFile
End Property

Public Property Get manifestFile() As String
    manifestFile = mManifestFile
End Property

Public Property Let manifestFile(ByVal nManifestFile As String)
    mManifestFile = nManifestFile
End Property

Public Property Get ManifestXML() As String
    ManifestXML = mManifestXML
End Property

Public Property Get ManifestDict() As Object
    ManifestDict = mManifestDict
End Property

Private Sub class_initialize()

    Set shell32 = VBA.CreateObject("Shell.Application")
    Set DOMDocument = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    Set Stream = VBA.CreateObject("ADODB.Stream")
    Set dict = CreateObject("Scripting.Dictionary")
    Set XDoc = CreateObject("MSXML2.DOMDocument.6.0")

End Sub

Private Sub Class_Terminate()
    Dim actCtxPtr As LongPtr

    hModule = GetModuleHandle(mDllFile)
    If hModule Then FreeLibrary (hModule)

    If actHandle <> 0& And _
        actHandle <> INVALID_HANDLE_VALUE And _
        current_senario <> 2 Then Call FinalizeActCtx

    Set shell32 = Nothing
    Set DOMDocument = Nothing
    Set Stream = Nothing
    Set TypeLibInfo = Nothing
    Set dict = Nothing
    Set XDoc = Nothing

End Sub

Public Sub FinalizeActCtx()
    Dim ret As Boolean

    If actHandle = 0& Or _
        actHandle = INVALID_HANDLE_VALUE Or _
        current_senario = 2 _
        Then actStatus = False: _
        Exit Sub

    If GetCurrentActCtx(actHandle) Then
        ret = DeactivateActCtx(0&, ByVal actCookie)
        If ret Then actCookie = 0&
        ReleaseActCtx actHandle: actHandle = 0&

        actStatus = False
    End If
End Sub

Private Function AssemblyName( _
    ByVal Company As String, _
    ByVal Description As String, _
    ByVal Name As String, _
    ByVal DllDateTime As Date) As String

    Company = AssemblyNamePart(Company)
    If Len(Description) Then
        Description = AssemblyNamePart(Description)
    Else
        Description = Name
    End If

    'Max length: 60.
    If Len(Company) Then
        AssemblyName = Left$(Company, 59 - Len(Description)) & "." & Description
    Else
        'These are supposed to be "strong" (unique) names so without a
        'value for Company we'll append a timestamp.
        AssemblyName = Left$(Description, 44) & Format$(DllDateTime, "\.YYYYMMDD\.HHNNSS")
    End If
End Function

Private Function AssemblyNamePart(ByVal Part As String) As String
    'Skips over runs of non-alphabetic characters, replacing each run by a "." character.
    Const MAX_LENGTH As Long = 45
    Dim S As String
    Dim i As Long
    Dim C As String
    Dim Skipping As Boolean
    Dim PastFirstAlpha As Boolean 'Avoiding outputting "." as first char of result.
    Dim SUsed As Long

    S = Space$(MAX_LENGTH)
    For i = 1 To Len(Part)
        C = Mid$(Part, i, 1)
        If C Like "[A-Za-z]" Then
            If Skipping Then
                If PastFirstAlpha Then
                    SUsed = SUsed + 1
                    Mid$(S, SUsed, 1) = "."
                End If
                Skipping = False
            End If
            SUsed = SUsed + 1
            Mid$(S, SUsed, 1) = C
            PastFirstAlpha = True 'We've copied at least one alpha character now.
            If SUsed >= MAX_LENGTH Then Exit For
        Else
            Skipping = True
        End If
    Next
    AssemblyNamePart = Left$(S, SUsed)
End Function

Private Function AttributeStrings(ByVal Mask As Integer) As String
    'TypeLibInfo.AttributeStrings is broken, returns nothing.
    'Instead call this passing TypeLibInfo.AttributeMask value:
    Dim S As String
    Dim Length As Long

    If Mask And LIBFLAG_FRESTRICTED Then S = S & "RESTRICTED,"
    If Mask And LIBFLAG_FCONTROL Then S = S & "CONTROL,"
    If Mask And LIBFLAG_FHIDDEN Then S = S & "HIDDEN,"
    If Mask And LIBFLAG_FHASDISKIMAGE Then S = S & "HASDISKIMAGE,"
    Length = Len(S)
    If Length Then AttributeStrings = Left$(S, Length - 1)
End Function

'create assembly Manifest - source:
'https://www.vbforums.com/showthread.php?796881-VB6-Reg-Free-COM-at-runtime-via-Microsoft-Windows-ActCtx
Public Sub CreateAssemblyManifest()
    'Returns True on success.
    Const XMLPI_ATTRIBS As String = "version=""1.0"" encoding=""UTF-8"" standalone=""yes"""
    Const ASM_V1_NAMESPACE As String = "urn:schemas-microsoft-com:asm.v1"
    Const PSOA_INTERFACE As String = "{00020424-0000-0000-C000-000000000046}"

    Dim ShellFolderItem As Object
    Dim ProcessingInstruction As Object
    Dim AssemblyNode As Object
    Dim Node As Object
    Dim FileVersion As String
    Dim Company As String
    Dim FirstCoClass As Boolean
    Dim Temp As String
    Dim Pos As Long
    Dim ThreadingModel As String
    Dim ResourceId As Integer
    Dim InterfaceInfo As TLI.InterfaceInfo
    Dim CoClassInfo As TLI.CoClassInfo

    With shell32
        Set ShellFolderItem = .NameSpace(ssfDESKTOP).ParseName(mDllFile)
        With ShellFolderItem
            FileVersion = .ExtendedProperty("System.FileVersion")
            Company = .ExtendedProperty("System.Company")
        End With
    End With
    Set ShellFolderItem = Nothing

    'Set TypeLibInfo = New TLI.TypeLibInfo
    'Aborts with exception if not a valid DLL:
    TypeLibInfo.ContainingFile = mDllFile

    If TypeLibInfo.CoClasses.Count = 0 Then
        MsgBox "DLL file contains no classes", vbOKOnly Or vbExclamation
        Exit Sub
    End If

    With DOMDocument
        Set ProcessingInstruction = .createProcessingInstruction("xml", XMLPI_ATTRIBS)
        .appendChild ProcessingInstruction
        Set AssemblyNode = .appendChild(.createNode(NODE_ELEMENT, "assembly", ASM_V1_NAMESPACE))
        With AssemblyNode
            .Text = vbNullString 'Hack to prevent serializing with collapsed tags, i.e.
                                 'XML output with <a></a> instead of <a/> form.
            .setAttribute "manifestVersion", "1.0"
            Set Node = .appendChild(DOMDocument.createNode(NODE_ELEMENT, "assemblyIdentity", ASM_V1_NAMESPACE))
            With Node
                .Text = vbNullString
                .setAttribute "name", AssemblyName(Company, _
                                                   TypeLibInfo.HelpString(0), _
                                                   TypeLibInfo.Name, _
                                                   FileDateTime(mDllFile))
                '.setAttribute "processorArchitecture", "x86"
                .setAttribute "type", "win32"
                .setAttribute "version", FileVersion
            End With
        End With
        With AssemblyNode
            Set Node = .appendChild(DOMDocument.createNode(NODE_ELEMENT, "file", ASM_V1_NAMESPACE))
            With Node
                .Text = vbNullString
                Temp = mDllFile
                Pos = InStrRev(Temp, "\")
                If Pos Then Temp = Mid$(Temp, Pos + 1)
                .setAttribute "name", Temp
                .setAttribute "hashalg", "SHA1"
                FirstCoClass = True
                For Each CoClassInfo In TypeLibInfo.CoClasses
                    Set Node = .appendChild(DOMDocument.createNode(NODE_ELEMENT, "comClass", ASM_V1_NAMESPACE))
                    With Node
                        .Text = vbNullString
                        .setAttribute "clsid", CoClassInfo.GUID
                        .setAttribute "tlbid", TypeLibInfo.GUID
                        Temp = CoClassInfo.HelpString(0)
                        If Len(Temp) Then
                            .setAttribute "description", Temp
                        End If
                        If FirstCoClass Then
                            'Attempt to retrieve threadingModel value, but this is only
                            'available if the library is registered:
                            With RegRead
                                On Error Resume Next
                                .OpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & CoClassInfo.GUID & "\InprocServer32"
                                If err.Number = 0 Then
                                    ThreadingModel = .QueryValue("ThreadingModel")
                                    err.Clear
                                    .CloseKey
                                End If
                                On Error GoTo 0
                            End With
                            FirstCoClass = False
                        End If
                        If Len(ThreadingModel) Then
                            .setAttribute "threadingModel", ThreadingModel
                        End If
                        .setAttribute "progid", TypeLibInfo.Name & "." & CoClassInfo.Name
                    End With
                Next
                Set Node = .appendChild(DOMDocument.createNode(NODE_ELEMENT, "typelib", ASM_V1_NAMESPACE))
                With Node
                    .Text = vbNullString
                    .setAttribute "tlbid", TypeLibInfo.GUID
                    ResourceId = TypeLibInfo.LCID And &HFFFF&
                    If ResourceId = 0 Then
                        ResourceId = 1
                    End If
                    .setAttribute "resourceid", Hex$(ResourceId)
                    .setAttribute "version", Hex$(TypeLibInfo.MajorVersion) & "." & Hex$(TypeLibInfo.MinorVersion)
                    Temp = TypeLibInfo.HelpFile
                    If Len(Temp) Then Temp = Left$(Temp, InStrRev(Temp, "\") - 1)
                    .setAttribute "helpdir", Temp
                    .setAttribute "flags", AttributeStrings(TypeLibInfo.AttributeMask)
                End With
            End With
            For Each InterfaceInfo In TypeLibInfo.Interfaces
                Set Node = .appendChild(DOMDocument.createNode(NODE_ELEMENT, "comInterfaceExternalProxyStub", ASM_V1_NAMESPACE))
                With Node
                    .Text = vbNullString
                    .setAttribute "name", InterfaceInfo.Name
                    .setAttribute "iid", InterfaceInfo.GUID
                    .setAttribute "tlbid", TypeLibInfo.GUID
                    .setAttribute "proxyStubClsid32", PSOA_INTERFACE
                End With
            Next
        End With
    End With

    mManifestXML = DOMDocument.XML
    embededManifestConv DOMDocument
    Call progidToDict(mManifestXML)

End Sub

Public Function pathNoExt()
    pathNoExt = Left(mDllFile, InStrRev(mDllFile, ".") - 1)
End Function

'conert manifest file from DOM object
Public Function embededManifestConv(ByVal DOMDocument As Object)
    Dim Temp As String
    Dim Bytes() As Byte
    Dim PartialDword As Long

    Temp = pathNoExt & ".manifest"
    With Stream
        .Open
        .Type = adTypeBinary
        DOMDocument.save Stream

        'Trim off unwanted trailing CRLF (or LF though MSXML generally uses CRLF):
        .Position = .Position - 2
        Bytes = .Read(2)
        If Bytes(0) = 13 And Bytes(1) = 10 Then
            .Position = .Position - 2
        ElseIf Bytes(1) = 10 Then
            .Position = .Position - 1
        End If
        .SetEOS

        'Pad to DWORD boundary if needed to create an embeddable manifest:
        PartialDword = .Position Mod 4
        If PartialDword Then
            Bytes = StrConv(Space$(4 - PartialDword), vbFromUnicode)
            .Write Bytes
        End If

        .SaveToFile Temp, adSaveCreateOverWrite
        .Close
    End With

    mManifestFile = Temp
End Function

'convert progid/clsid to dictionary from manifest xml
Public Function progidToDict(ByVal strManifest As String) As Object

    Dim lists As Object
    Dim listNode As Object
    Dim i As Long
    Dim iAtt As Object
    Dim sProgID As String
    Dim sCLSID As String
    Dim stlbid As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set XDoc = CreateObject("MSXML2.DOMDocument.6.0")

    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.loadXml strManifest

    'Get Document Elements
    Set lists = XDoc.SelectNodes("//*")

    For Each listNode In lists

        If VBA.StrComp(listNode.nodename, "comClass", vbTextCompare) <> 0 Then GoTo continue
        For Each iAtt In listNode.Attributes
            'Debug.Print listNode.nodename, iAtt.nodename, iAtt.nodeTypedValue
            Select Case iAtt.nodename
                Case "progid": sProgID = iAtt.nodeTypedValue
                Case "clsid": sCLSID = iAtt.nodeTypedValue
                Case "tlbid": stlbid = iAtt.nodeTypedValue
            End Select
        Next
        err.Clear
        'On Error GoTo continue
        dict.Add sProgID, sCLSID
continue:
    Next

    Set progidToDict = dict
    Set mManifestDict = dict

    Set XDoc = Nothing
    Set dict = Nothing

End Function

'apply method by Microsoft.Windows.ActCtx
Public Function oCtxTest(Optional ByVal clsid As String, Optional ByVal progid As String, Optional ByVal senario As Integer = 0) As Boolean
    Dim oCtx As Object
    Dim testObj As Object

    On Error GoTo err_handler
    InitObject testObj, clsid, progid, senario

    If VBA.StrComp(VBA.TypeName(testObj), "Nothing", vbTextCompare) = 0 Then
        oCtxTest = False: Exit Function
    End If

    oCtxTest = True

err_handler_exit:
    Set testObj = Nothing
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("oCtxTest", err.Number, err.Description, Erl)
        oCtxTest = False
    End If

    Resume err_handler_exit
End Function

Public Function getDllFromManifest(Optional ByVal manifestFile As String)
    Dim DOMDocument As Object
    Dim dllName As String
    Dim path As String

    If VBA.Len(manifestFile) = 0 Then manifestFile = mManifestFile

    Set DOMDocument = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    With DOMDocument
        .async = False
        .validateOnParse = False
        .Load manifestFile
    End With

    On Error GoTo err_handler
   dllName = DOMDocument.SelectSingleNode("//*[2]/@name").Text
    On Error GoTo err_handler
    If VBA.Len(dllName) = 0 Then err.Raise 100, , "not find dll filename "

    path = VBA.Left(manifestFile, VBA.InStrRev(manifestFile, "\"))

    dllName = path & dllName
    getDllFromManifest = dllName

    Set DOMDocument = Nothing

err_handler_exit:
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("getDllFromManifest", err.Number, err.Description, Erl)
    End If

    Resume err_handler_exit
End Function

Public Function MAKEWORD(ByVal wLow As Integer, ByVal wHigh As Integer) As Integer

    If wHigh And &H80 Then
        MAKEWORD = (((wHigh And &H7F) * 256) + wLow) Or &H8000
    Else
        MAKEWORD = (wHigh * 256) + wLow
    End If

End Function

Public Function HIWORD(ByVal dwValue As Long) As Long
    CopyMemory ByVal VarPtr(HIWORD), ByVal VarPtr(dwValue) + 2, 2
End Function

Public Function LOWORD(ByVal dwValue As Long) As Long
    CopyMemory ByVal VarPtr(LOWORD), ByVal VarPtr(dwValue), 2
End Function

Public Function MAKELONG(ByVal wLow As Long, ByVal wHi As Long) As Long

    If (wHi And &H8000&) Then
        MAKELONG = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
    Else
        MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHi))
    End If

End Function

'rewrote c launguage Micro: https://docs.microsoft.com/en-us/previous-versions/ms930224(v=msdn.10)
Public Function MAKEINTRESOURCE(ByVal lID As Long) As String
    MAKEINTRESOURCE = "#" & CStr(MAKELONG(lID, 0))
End Function

'get getManifest by resource API
Public Function getManifestByResourceApi(Optional ByVal codepage As Long = CP_UTF8) As String
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'author: Hans Yao (hansyow@gmail.com)
    'FindResourceA: https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-findresourcea
    'lpName: MAKEINTRESOURCE(resource_id)
    'lpType(resource type or RT_MANIFEST = MAKEINTRESOURCE(24)): https://docs.microsoft.com/en-us/windows/win32/menurc/resource-types

    'stackoverflow: concept initated by Jan 2009, implemented in VBA in Jun 2022 by Hans Yao<hansyow@gmail.com>
    'https://stackoverflow.com/questions/420852/reading-an-applications-manifest-file

    'resource id: https://docs.microsoft.com/en-us/windows/win32/sbscs/using-side-by-side-assemblies-as-a-resource
    '* Value of MANIFEST_RESOURCE_ID
    ' a. Manifest specifies the Process Default?
    ' b. Use for Static Imports?
    ' c. Use for an EXE?
    ' d. Use for a DLL?
    ' e. Uses Side-by-Side version of assemblies if compiled with -DISOLATION_AWARE_ENABLED?

    '                *      a    b      c   d     e
                    '1   Yes  Yes  Yes No Yes
                    '2   No  Yes  Yes Yes Yes
                    '3   No  No  Yes Yes Yes
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Dim resInfo As LongPtr
    Dim ret As LongPtr
    Dim RT_MANIFEST As String
    Dim lpName As String
    Dim resData As LongPtr
    Dim resSize As Long
    Dim manifestPtr As LongPtr
    Dim i As Integer
    Dim dllFile As String

    dllFile = mDllFile
    hModule = LoadLibraryEx(dllFile, 0&, LOAD_LIBRARY_AS_DATAFILE)

    If hModule = 0 Then Exit Function

    'loop all resouce_id(1 to 16) to get manifest
    For i = 1 To 16
        On Error Resume Next
        resInfo = FindResourceA(hModule, MAKEINTRESOURCE(i), MAKEINTRESOURCE(24))
        If resInfo <> 0 Then Exit For
    Next

    On Error GoTo err_handler
    If resInfo = 0 Then err.Raise 53

    resData = LoadResource(hModule, resInfo)
    resSize = SizeofResource(hModule, resInfo)
    If (resData And resSize) Then
        manifestPtr = LockResource(resData)

        'save xml manifest to class property
        mManifestXML = ptrToString(manifestPtr, codepage)
        getManifestByResourceApi = mManifestXML
    End If

err_handler_exit:
    FreeResource (resData)
    FreeLibrary hModule
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("getManifest", err.Number, err.Description, Erl)
    End If
    Resume err_handler_exit
End Function

#If VBA7 Then
Public Function ActiveActCtx(ByVal manifestFile As String, ByRef actHandle As LongPtr, ByRef actCookie As LongPtr) As Boolean
#Else
End Function
Public Function ActiveActCtx(ByVal manifestFile As String, ByRef actHandle As Long, ByRef actCookie As Long) As Boolean
#End If
    Dim ret As Boolean
    Dim actCTX As ACTCTX_
    Dim dllName As String

    If actStatus = True Then ActiveActCtx = True: Exit Function

    dllName = mDllFile
    '/* Check if we already have a loaded module for this service */
    hModule = GetModuleHandle(dllName)
    If hModule <> 0 Then FreeLibrary (hModule)

    'https://cpp.hotexamples.com/examples/-/-/ActivateActCtx/cpp-activateactctx-function-examples.html
    With actCTX
        .cbSize = Len(actCTX)
        .dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID
        .lpAssemblyDirectory = Application.path
        .lpSource = manifestFile
    End With

    '/* Activate the context */
    actHandle = CreateActCtx(actCTX)
    On Error GoTo err_handler
    If actHandle = INVALID_HANDLE_VALUE Or actHandle = 0& Then err.Raise 100, , "CreateActCtx: failed"

    '/* activate achandle */
    ret = ActivateActCtx(actHandle, actCookie)
    On Error GoTo err_handler
    If ret = False Then err.Raise 100, , "ActivateActCtx failed"

    actStatus = ret
    ActiveActCtx = ret

err_handler_exit:
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("ActiveActCtx", err.Number, err.Description, Erl)
        ReleaseActCtx actHandle: actHandle = 0&
        ret = False
        actStatus = False
        ActiveActCtx = False
    End If
    GoTo err_handler_exit

End Function

Public Function InitObject(ByRef myObj As Object, Optional ByVal clsid As String, Optional ByVal progid As String, Optional ByVal scenario As Integer = 0) As Boolean
    Dim ret As Boolean

    If VBA.Len(clsid) = 0 And VBA.Len(progid) = 0 Then
        On Error GoTo err_handler
        err.Raise 94, , "must have one of clsid and progid"
    End If

    current_senario = senario

    Select Case senario
        Case 0, 1:
            If actStatus = False Then
                ret = ActiveActCtx(Me.manifestFile, actHandle, actCookie)

                If ret = False Then
                    On Error GoTo err_handler
                    err.Raise 100, , "ActiveActCtx failed:" & vbNewLine & "clsid:" & clsid & vbNewLine & "progid:" & progid
                End If
            End If

            '/* now ready and load dll */
             '/* Check if we already have a loaded module for this service */
            hModule = GetModuleHandle(mDllFile)
            If hModule = 0 Then
                hModule = LoadLibraryEx(mDllFile, 0&, LOAD_WITH_ALTERED_SEARCH_PATH)
            End If

            On Error GoTo err_handler
            If hModule = 0 Then err.Raise 48

            Select Case senario
                Case 0:
                    Call CreateObjectExAPI(VBA.IIf(VBA.Len(clsid) = 0, progid, clsid), myObj, CLSCTX_INPROC_SERVER)
                Case 1:
                    On Error GoTo err_handler
                    If VBA.Len(progid) = 0 Then err.Raise 100, , "miss clsid,, clsid much be included in senario 1 "
                    Set myObj = VBA.GetObject("new:" & clsid)
            End Select

        Case 2:
            Dim oCtx As Object

            '/* now ready and load dll */
             '/* Check if we already have a loaded module for this service */
            hModule = GetModuleHandle(mDllFile)
            If hModule = 0 Then
                hModule = LoadLibraryEx(mDllFile, 0&, LOAD_WITH_ALTERED_SEARCH_PATH)
            End If

            actStatus = False
            Set oCtx = VBA.CreateObject("Microsoft.Windows.ActCtx")
            oCtx.manifest = mManifestFile
            On Error GoTo err_handler
            Set myObj = oCtx.CreateObject(progid)

    End Select

    If VBA.StrComp(VBA.TypeName(myObj), "Object", vbTextCompare) <> 0 Then
        On Error GoTo err_handler
        err.Raise 91
    End If

    InitObject = True

err_handler_exit:
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("InitObject", err.Number, err.Description, Erl)
        InitObject = False
        Resume err_handler_exit
    End If

End Function

'get function pointer from DLL
Public Function getDllFunction(ByVal procName As String, ByRef lpAddress) As Boolean
    Dim actCTX As ACTCTX_
    Dim ret As Boolean
    Dim hModule As LongPtr
    Dim myObj As Object

    Dim dllName As String

    dllName = mDllFile

    '/* Check if we already have a loaded module for this service */
    hModule = GetModuleHandle(dllName)

    '/* We don't -- load it */
    If hModule = 0 Then
        hModule = LoadLibraryEx(dllName, 0&, LOAD_WITH_ALTERED_SEARCH_PATH)
    End If

    If hModule = 0 Then err.Number = 100: err.Description = "failed to load " & dllName: GoTo err_handler

    '/* Next, get the address being looked up*/
    'OLE does not provide this function. DLLs that support the OLE Component Object Model (COM) must implement DllGetClassObject in OLE object handlers or DLL applications.
    'https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-dllgetclassobject
    lpAddress = GetProcAddress(hModule, procName)
    If lpAddress = 0 Then err.Number = 100: err.Description = "failed to get function address: " & procName: GoTo err_handler
    getDllFunction = True

err_handler_exit:
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("getDllFunction", err.Number, err.Description, Erl)
        getDllFunction = False
    End If
    GoTo err_handler_exit

End Function

'class: progid or clsid
Private Function CreateObjectExAPI(ByVal mClass As String, ByRef myObject As Object, Optional ByVal ServerName As String, Optional ByVal Context As CLSCTX = CLSCTX_INPROC_SERVER) As Boolean
    Dim rclsid As GUID
    Dim riid As GUID
    Dim hr As Long
    Dim ServerInfo As COSERVERINFO
    Dim mqi As MULTI_QI

    ' Convert IID_IUnknown string to binary IID
    hr = CLSIDFromString(StrPtr(IID_IUnknown), riid)
    If hr <> 0 Then err.Raise hr
    mqi.pIID = VarPtr(riid)
    ' Convert provided CLSID or ProgID string to binary CLSID
    If ((Left(mClass, 1) = "{") And (Right(mClass, 1) = "}") And _
        (Len(mClass) = 38)) Then
        hr = CLSIDFromString(StrPtr(mClass), rclsid)
        If hr <> 0 Then err.Raise hr
    Else
        ' Create CLSID from ProgID
        hr = CLSIDFromProgID(StrPtr(mClass), rclsid)
        If hr <> 0 Then err.Raise hr
    End If
    ' Set up appropriate server info
    If ServerName <> "" Then
        ServerInfo.pwszName = StrPtr(ServerName)
    End If

    'https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstanceex
    'HRESULT CoCreateInstanceEx(
    '  [in]      REFCLSID     Clsid,
    '  [in]      IUnknown     *punkOuter,
    '  [in]      DWORD        dwClsCtx,
    '  [in]      COSERVERINFO *pServerInfo,
    '  [in]      DWORD        dwCount,
    '  [in, out] MULTI_QI * pResults
    ');

    ' Create an instance of the object
    hr = CoCreateInstanceEx(rclsid, 0&, Context, ByVal VarPtr(ServerInfo), 1&, ByVal VarPtr(mqi))
    On Error GoTo err_handler
    If hr <> 0 Then err.Raise hr, , "createobject failed: " & mClass
    Set myObject = mqi.pItf
    CreateObjectExAPI = True

err_handler_exit:
    Exit Function

err_handler:
    If err.Number <> 0 Then
        Debug.Print err_message("CreateObjectExAPI", err.Number, err.Description, Erl)
        Resume err_handler_exit
    End If

End Function

#If VBA7 Then
    Private Function ptrToString(ByVal pString As LongPtr, Optional ByVal codepage As Long = CP_UTF8) As String
#Else
    Private Function ptrToString(ByVal pString As Long, Optional ByVal codepage As Long = CP_UTF8) As String
#End If
    Dim buf     As String
    Dim cSize   As Long
    Dim retVal  As Long

    cSize = MultiByteToWideChar(codepage, 0, pString, -1, 0, 0)

    If cSize <= 1 Then
        ptrToString = ""
        Exit Function
    End If

    ptrToString = String(cSize - 1, "*")

    retVal = MultiByteToWideChar(codepage, 0, pString, -1, StrPtr(ptrToString), cSize)
    If retVal = 0 Then
       err.Raise 1000, "ptrToString", "ptrToString error: " & err.LastDllError
       Exit Function
    End If

End Function
