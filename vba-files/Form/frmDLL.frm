VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDLL 
   Caption         =   "DLL RegFree for VBA"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14895
   OleObjectBlob   =   "frmDLL.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit


Private Declare PtrSafe Sub ReleaseActCtx Lib "kernel32" (ByVal hActCtx As LongPtr)
Private Declare PtrSafe Function DeactivateActCtx Lib "kernel32" (ByVal dwFlags As LongPtr, ByVal ulCookie As LongPtr) As Boolean
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function GetCurrentActCtx Lib "kernel32" (ByRef lphActCtx As LongPtr) As Boolean

Private dll As New clsDLLAsm

Private Sub cmdSelectDll_Click()
    Dim dllFile As String
    Dim dict As Object
    Dim progid As Variant
    Dim iCount As Integer
    Dim manifestFile As String

    dllFile = Application.GetOpenFileName(FileFilter:="COM/OCX file (*.dll; *.ocx), *.dll;*.ocx", title:="select DLL/OCX file", MultiSelect:=False)
    If VBA.Len(dllFile) = 0 Or dllFile = "False" Then Exit Sub

    If dllFile = Me.lblDllFile Then Exit Sub

    manifestFile = Left(dllFile, VBA.InStrRev(dllFile, ".")) & "manifest"

    On Error Resume Next
    Set dll = Nothing
    With dll
        .dllFile = dllFile

        If VBA.Len(Dir(manifestFile)) = 0 Then
            On Error GoTo err_handler
            .CreateAssemblyManifest
            Set dict = .ManifestDict
        Else
            .manifestFile = manifestFile
            Set dict = .progidToDict(loadXmlStr(manifestFile))
        End If
        If VBA.Len(.manifestFile) = 0 Then MsgBox "gather Manifest failed": Exit Sub
        Me.lblFilePath.Caption = .manifestFile
    End With

    'show dll/ocx info in userform
    Me.lblDllFile.Caption = dllFile
    dictToListBox dict

err_handler_exit:

    Set dict = Nothing
    Set dll = Nothing
    Exit Sub
err_handler:
    If err.Number <> 0 Then
        MsgBox err_message("cmdSelectDll_Click", err.Number, err.Description, Erl), , "Warning"
    End If

    GoTo err_handler_exit

End Sub

Private Sub dictToListBox(ByVal dict As Object)
    Dim progid As Variant
    Dim iCount As Integer

    With Me.ListBox1
        .Clear
        .ColumnCount = 2
        For Each progid In dict
            iCount = .ListCount
            .AddItem
            .List(iCount, 0) = VBA.CStr(progid)
            .List(iCount, 1) = VBA.CStr(dict(progid))
        Next
    End With
End Sub

Private Sub cmdSelectManifest_Click()
    Dim manifestFile As String
    Dim strManifest As String
    Dim dict As Object

    On Error Resume Next
    Set dll = Nothing

    manifestFile = Application.GetOpenFileName(FileFilter:="COM/OCX file (*.manifest; *.xml), *.manifest; *.xml", title:="select manifest File", MultiSelect:=False)
    If VBA.Len(manifestFile) = 0 Or manifestFile = "False" Then Exit Sub
    If manifestFile = Me.lblFilePath Then Exit Sub

    dll.manifestFile = manifestFile
    dll.dllFile = dll.getDllFromManifest(manifestFile)

    If Me.lblDllFile.Caption = dll.dllFile Then MsgBox "previous dll filename exists!": Exit Sub

    Me.lblDllFile.Caption = dll.dllFile
    Me.lblFilePath.Caption = manifestFile
    'Set dll.ManifestDict = Nothing

    strManifest = loadXmlStr(manifestFile)

    'show dll/ocx info in userform
    dictToListBox dll.progidToDict(strManifest)

    Set dll = Nothing

End Sub

Private Sub cmdTest_Click()
    Dim manifestFile As String
    Dim progid As String
    Dim clsid As String
    Dim i As Integer
    Dim senario As Integer

    On Error Resume Next
    Set dll = Nothing

    senario = getSenario()

    Me.ListBox1.ColumnCount = 3
    If Not IsNull(Me.ListBox1.List(0, 2)) Then Exit Sub

    For i = 0 To Me.ListBox1.ListCount - 1
        With Me.ListBox1

            progid = .List(i, 0)
            clsid = .List(i, 1)
            dll.manifestFile = Me.lblFilePath.Caption
            dll.dllFile = Me.lblDllFile.Caption
            .List(i, 2) = VBA.IIf(dll.oCtxTest(clsid, progid, senario), "success", "failed")
        End With
    Next

    Set dll = Nothing
End Sub

Private Sub lblDllFile_Click()
    If VBA.Len(Me.lblDllFile) = 0 Then Exit Sub
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Me.lblDllFile.Caption
End Sub

Private Sub lblDllFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.lblDllFile
       If .ForeColor = rgb(0, 0, 0) Then .ForeColor = rgb(0, 0, 255) Else .ForeColor = rgb(0, 0, 0)
    End With
End Sub

Private Sub lblFilePath_Click()
    If VBA.Len(Me.lblFilePath) = 0 Then Exit Sub
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Me.lblFilePath.Caption
End Sub


Private Sub lblFilePath_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.lblFilePath
       If .ForeColor = rgb(0, 0, 0) Then .ForeColor = rgb(0, 0, 255) Else .ForeColor = rgb(0, 0, 0)
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me.lblFilePath
        .Caption = ""
        .AutoSize = False
        .WordWrap = True
    End With
    With Me.lblDllFile
        .Caption = ""
        .AutoSize = False
        .WordWrap = True
    End With
    Call frmLayout
End Sub


Private Function loadXmlStr(ByVal xmlFile As String) As String
    Dim DOMDocument As Object
    Set DOMDocument = VBA.CreateObject("MSXML2.DOMDocument.6.0")

    With DOMDocument
        .async = False
        .validateOnParse = False
        .Load xmlFile
    End With
    loadXmlStr = DOMDocument.XML
    Set DOMDocument = Nothing
End Function

Private Function frmLayout()
    Dim ctl As Variant
    Dim i As Integer

    Me.Controls("opt1").Caption = "方案1:CoCreateInstanceEx创建对象"
    Me.Controls("opt2").Caption = "方案2:GetObject(""new:"" & clsid)创建对象"
    Me.Controls("opt3").Caption = "方案3:createObject(""Microsoft.Windows.ActCtx"")创建对象"
    Me.Controls("opt1").Value = True
End Function

Private Function getSenario() As Integer
    Dim i As Integer
    Dim ret As Boolean

    getSenario = 0
    For i = 0 To 2
        ret = Me.Controls("opt" & i + 1).Value
        If ret Then getSenario = i: Exit Function
    Next

End Function

Private Sub UserForm_Terminate()
    Set dll = Nothing
End Sub
