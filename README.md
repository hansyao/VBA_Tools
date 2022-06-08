# VBA_Tools
vba utilities which compatible with 32bit and 64bits


## FreeReg COM/ActiveX
class: [clsDLLAsm.cls](https://github.com/hansyao/VBA_Tools/blob/main/vba-files/Class/clsDLLAsm.cls)

1. Idea: create a native solution to use COM/ActiveX control in VBA without registration (side-by-side Assembly)

2. **concept**: *call COM/ActiveX thru ActCtx Content*
>Microsoft provided The ActivateActCtx function activates the specified activation context. It does this by pushing the specified activation context to the top of the activation stack. The specified activation context is thus associated with the current thread and any appropriate side-by-side API functions.

3. **Implementation**:

  step 1. CreateAssemblyManifest from existing DLL/OCX controls

  step 2. Activate actCtx contents for specific DLL/OCX. there're three ways to implement it.

* *Solution 1*. win32 API [CoCreateInstanceEx](https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstanceex) to create object.
* *Solution 2*. VBA provided the native function [VBA.GetObject](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getobject-function) to create object.
* *Solution 3*. Microsoft also provided an ActiveX controls [Microsoft.Windows.ActCtx](https://docs.microsoft.com/en-us/windows/win32/sbscs/microsoft-windows-actctx-object) to create object

This class support all of above. and you can choose one of above three methods to implement it depends on your own decision.

4. **How to use**: code sample


```vb

'generally, you do only need one-time manifest generation. and then you would be able to directly set the property dll.manifestFile for future use.
'or you have already owned the manifestfile, just set dll.manifestFile but no more dll.CreateAssemblyManifest need.

function sampleObject()
  dim dll as new clsDLLAsm
  dim dllFile as string
  dim myObj as Object
  dim scenario as Integer
  dim clsid as string
  dim progid as string

  '0. CoCreateInstanceEx; 1. VBA.GetObject; 2. Microsoft.Windows.ActCtx
  scenario = 0

  'asign your .dll path to
  dll.dllFile = "your_dllfile_path.dll"

  'create .manifest file based on existing DLL/OCX
  dll.CreateAssemblyManifest
  If VBA.Len(dll.manifestFile) = 0 Then MsgBox "gather Manifest failed": Exit Sub

  'manifest file generated, then check the clsid & progid and put in InitObject

  '/* now ready and load dll and create object*/
  set myObj as object
  ret = dll.InitObject(myObj, clsid, progid, scenario)
  'myObj is byref which return the object you desired.

  if vba.TypeName(myObj) = "Nothing" then MsgBox "create object failed"

  MsgBox "object created successfully!"

  'then, do your action with myOjb.xxxx

  set dll = nothing '/* be sure release this class after finish all tasks.

end function

```

## Dependence

>need TypeLib Information support (TLBINFO32.DLL). As win10/11 64bit has removed this lib, you must install by your self follow the simple step [here](https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application). it's only for development purpose, no more need if you deploy it after generate manifest file.
