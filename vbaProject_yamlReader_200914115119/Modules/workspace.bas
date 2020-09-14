Attribute VB_Name = "workspace"
Option Explicit

Sub dev()

'Console.log "Hello!"
Call LoadJsonMain.Batch

'Call unittest

End Sub

Private Sub UnitTest()
'''' *************************************************
Console.info "-------------------- start !!"
Package.Include
''

''
Package.Terminate
Console.info "-------------------- end ...."
''
End Sub

Public Sub Export()
'''' *************************************************
Dim Exporter As Tool_ModuleExporter
Set Exporter = New Tool_ModuleExporter
Call Exporter.Export
''
End Sub
