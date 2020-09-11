VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tool_ModuleExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /****************************************
'''  * @file Tool_ModuleExporter.cls
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class T_ModuleExporter
'''  * @classdesc Module Exporter
'''  * @require C_File, C_Date
'''  */


Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  */
End Sub

Private Sub Class_Terminate()
''' /********************************************************
'''  * @event Class_Terminate
'''  */
'Destructor
End Sub

'//---------------------------------------
'// Instance method
'

Public Sub Export()
''' /********************************************************
'''  * @function Export
'''  *
'''  * @require Make VBA object trusted to access in security center
'''  * @require C_Date, C_File, C_Book
'''  */
'''
Dim C_Book As C_Book
Set C_Book = New C_Book
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
'' const
Const CONS_export_path As String = "vbaProject"
Const CONS_StdModule As Integer = 1
Const CONS_StdModuleFolder As String = "Modules"
Const CONS_ClassModule As Integer = 2
Const CONS_ClassModuleFolder As String = "ClassModules"
' variables
Dim objVbComponent As Object
Dim strExtention As String
Dim strFilePath As String
Dim strFolderPath As String
Dim strExportPath As String
Dim strWbName As String
Let strWbName = ThisWorkbook.Name
'
Dim strDtmformat As String
strDtmformat = format(Now(), "yymmddhhnnss")
Dim currentPath As String
currentPath = C_File.GetCurrentDirectory
'
strExportPath = C_File.BuildPath(currentPath, CONS_export_path) & "_" & C_File.GetBaseName(strWbName) & "_" & strDtmformat
C_File.CreateFolder (strExportPath)
C_File.CreateFolder (C_File.BuildPath(strExportPath, CONS_StdModuleFolder))
C_File.CreateFolder (C_File.BuildPath(strExportPath, CONS_ClassModuleFolder))
'
For Each objVbComponent In ThisWorkbook.VBProject.VBComponents
    ' Debug.Print objVbComponent.name
    If objVbComponent.Type <= 2 Then
        Select Case objVbComponent.Type
            Case CONS_StdModule
                strExtention = ".bas"
                strFolderPath = C_File.BuildPath(strExportPath, CONS_StdModuleFolder)
                
            Case CONS_ClassModule
                strExtention = ".cls"
                strFolderPath = C_File.BuildPath(strExportPath, CONS_ClassModuleFolder)
        End Select
        
        strFilePath = C_File.BuildPath(strFolderPath, objVbComponent.Name & strExtention)
        Console.info (strFilePath)
        Call objVbComponent.Export(strFilePath)
    End If
Next
'
Dim wbName As String
Let wbName = ThisWorkbook.Name
Dim wbUpdateDate As Date
Let wbUpdateDate = C_Book.GetBookProperties(wbName).Item("Last save time")
Dim dict As Dictionary
Set dict = New Dictionary
Call dict.Add("workbook", wbName)
Call dict.Add("updated", wbUpdateDate)
Dim json_str As String
Let json_str = JsonConverter.ConvertToJson(dict, Whitespace:=4)
Dim json_path As String
Let json_path = C_File.BuildPath(strExportPath, "book.json")
Call C_FileIO.WriteTextAllAsUTF8(json_path, json_str)
''
Set objVbComponent = Nothing
Set C_Book = Nothing
Set C_File = Nothing
Set C_FileIO = Nothing
'
End Sub

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Export
End Sub

