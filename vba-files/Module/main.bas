Attribute VB_Name = "main"
Option Explicit


Public Function err_message(ByVal prog As String, ByVal errNo As Long, ByVal Description As String, ByVal errLine As Long)
    err_message = "The following error has occured..." & vbCrLf & _
           "Error Number: " & errNo & vbCrLf & _
           "Error Source: " & prog & vbCrLf & _
           "Error Description: " & Description & _
           VBA.Switch(errLine = 0, "", errLine <> 0, vbCrLf & "Line No: " & errLine)
End Function

public sub freeRegCom()
  on error goto err_handler
  frmDll.show

err_handler:
  if err.Number <> 0 then
    msgbox err_message("freeRegCom", err.Number, err.Description, erl)
  end if
end sub
