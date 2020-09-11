Attribute VB_Name = "DebugArea"
Option Explicit

Sub TestYaml2Json()
'''' *********************************************
''
Dim objYaml As O_YAML
Set objYaml = New O_YAML
''
Dim file_name As String
Let file_name = "input/yamlformat.yaml"
''
Dim jObj As cJobject
Set jObj = New cJobject
''
Set jObj = objYaml.YamlFileToJObject(file_name)
Console.info jObj.formatData
Console.info jObj.serialize
''
End Sub

Sub TestYaml2Range()
'''' *********************************************
''
Dim C_Range As C_Range
Set C_Range = New C_Range
''
Dim objYaml As O_YAML
Set objYaml = New O_YAML
''
Dim file_name As String
Let file_name = "input\yamlformat.yaml"
''
Dim jObj As cJobject
Set jObj = New cJobject
''
Dim aryLines As Variant
Let aryLines = objYaml.ReadYamlFile(file_name)
''
Console.Dir aryLines
''
Dim aryary As Variant
Let aryary = objYaml.YamlArrayToArrayArray(aryLines)
Console.Dir aryary
Call C_Range.PutArrayArray("A1", "Sheet1", aryary)
''
End Sub

Sub TestYaml2JsonFile()
'''' *********************************************
''
Dim C_Book As C_Book
Set C_Book = New C_Book
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
Dim objYaml As O_YAML
Set objYaml = New O_YAML
''
Dim file_name As String
Let file_name = "input\yamlformat.yaml"
Dim base_folder As String
Let base_folder = C_Book.GetThisWorkbookFolder
Dim outfile_name As String
Let outfile_name = "output\yamlformat.json"
Dim outfile_path As String
Let outfile_path = C_File.BuildPath(base_folder, outfile_name)
''
Dim jObj As cJobject
Set jObj = New cJobject
''
Set jObj = objYaml.YamlFileToJObject(file_name)
Console.info jObj.formatData
Console.info jObj.serialize
''
Dim json_str As String
Let json_str = usefulcJobject.JSONStringify(jObj)
Let json_str = JsonConverter.ConvertToJson(JsonConverter.ParseJson(json_str), 4)
Call C_FileIO.WriteTextAllAsUTF8NoneBOM(outfile_path, json_str)
''
End Sub

Sub TestJObjectToYaml()
'''' *********************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Book As C_Book
Set C_Book = New C_Book
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
Dim C_JObject As C_JObject
Set C_JObject = New C_JObject
Dim objYaml As O_YAML
Set objYaml = New O_YAML
''
Dim base_folder As String
Let base_folder = C_Book.GetThisWorkbookFolder
Dim outfile_name As String
Let outfile_name = "output\output.yaml"
Dim outfile_path As String
Let outfile_path = C_File.BuildPath(base_folder, outfile_name)
''
Dim jObj As cJobject
Set jObj = New cJobject
''
Call jObj.init(Nothing)
Call jObj.Add("tanbana")
Call jObj.Add("tanbana.name")
Call jObj.child("tanbana.name").setValue("Tanbana Jacobson")
Call jObj.Add("tanbana.job")
Call jObj.child("tanbana.job").setValue("Coder")
Call jObj.Add("tanbana.skills")
Call jObj.child("tanbana.skills").addArray
Call jObj.Add("tanbana.skills.1")
Call jObj.child("tanbana.skills.1").setValue("Python")
Call jObj.Add("tanbana.skills.2")
Call jObj.child("tanbana.skills.2").setValue("Peal")
Call jObj.Add("tanbana.skills.3")
Call jObj.child("tanbana.skills.3").setValue("PHP")
Console.info jObj.formatData
Console.info jObj.serialize
''
Dim lines As Variant
Let lines = C_JObject.ConverttoYamlArray(jObj)
''
Dim yaml_str As String
Let yaml_str = VBA.Join(lines, vbCrLf) + vbCrLf
Call C_FileIO.WriteTextAllAsUTF8NoneBOM(outfile_path, yaml_str)
''
End Sub

Sub TestYamlArrayArrayToYaml()
'''' *********************************************
''
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Book As C_Book
Set C_Book = New C_Book
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
Dim C_JObject As C_JObject
Set C_JObject = New C_JObject
Dim objYaml As O_YAML
Set objYaml = New O_YAML
''
''
Dim base_folder As String
Let base_folder = C_Book.GetThisWorkbookFolder
Dim outfile_name As String
Let outfile_name = "output\datadef.yaml"
Dim outfile_path As String
Let outfile_path = C_File.BuildPath(base_folder, outfile_name)
''
Dim targetParam As String
Let targetParam = "#  data design"
Dim shtName As String
Let shtName = "Sheet2"
Dim aryaryYaml As Variant
Let aryaryYaml = C_Range.GetCurrentRegionByKeyword(targetParam, shtName).Value
'Console.Dir aryaryYaml
Dim jObj As cJobject
Set jObj = New cJobject
Set jObj = objYaml.ConverttoJObject(targetParam, shtName, opt:=True)
'Console.info jObj.formatData
Dim lines As Variant
Let lines = C_JObject.ConverttoYamlArray(jObj)
'Console.Dir lines
Dim yaml_str As String
Let yaml_str = VBA.Join(lines, vbCrLf) & vbCrLf
'Console.info yaml_str
Call C_FileIO.WriteTextAllAsUTF8NoneBOM(outfile_path, yaml_str)
''
End Sub
