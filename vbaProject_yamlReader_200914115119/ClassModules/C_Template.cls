VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Template.cls
'''  *<pre> Template Utility Class
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/05/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Template
'''  * @classdesc XML Format Utility Class
'''  *
'''  */
'
Public Function GetFieldTemplateKey(ByVal val As String) As String
'''' ********************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim strRtn As String
''
If C_String.StartsWith(val, "{{") And C_String.EndsWith(val, "}}") Then
    Let strRtn = C_String.Chop(val, 2)
    Let strRtn = C_String.Chomp(strRtn, 2)
Else
    Let strRtn = ""
End If
''
Let GetFieldTemplateKey = strRtn
''
End Function


Public Function CreateXmlTemplate(recname As Variant, aryHeader As Variant) As String
''' /********************************************************
'''  *@function createXmlTemplate
'''  *<pre>XML形式Formatを、Header Arrayをもとに生成する。
'''  *  @param {Variant<String>} recname record name
'''  *  @param {Variant<String>()} aryHeader
'''  *  @return {String} createXmlTemplate
'''  */
    Dim str As String
    str = "<" & recname & " "
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        str = str & aryHeader(i) & "=" & """" & "${" & aryHeader(i) & "}" & """" & " "
    Next i
    str = str & "/>"
    CreateXmlTemplate = str
End Function

Public Function CreateXml2Template(recname As Variant, aryHeader As Variant) As String
''' /********************************************************
'''  *@function createXml2Template
'''  *<pre>XML形式Formatを、Header Arrayをもとに生成する。
'''  *  @param {Variant<String>} recname record name
'''  *  @param {Variant<String>()} aryHeader
'''  *  @return {String} createXml2Template
'''  */
    Dim str As String
    str = "<" & recname & ">" & vbCrLf
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        str = str & vbTab & "<" & aryHeader(i) & ">${" & aryHeader(i) & "}</" & aryHeader(i) & ">" & vbCrLf
    Next i
    str = str & "</" & recname & ">"
    CreateXml2Template = str
End Function


'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Function UnitTest()
    Dim aryHeader As Variant

''createXmlTemplate
    aryHeader = Array("no", "name", "address", "postNo", "note", "age")
    Console.info Me.CreateXmlTemplate("user", aryHeader)
    
''createXml2Template
    aryHeader = Array("no", "name", "address", "postNo", "note", "age")
    Console.info Me.CreateXml2Template("user", aryHeader)
    
End Function


