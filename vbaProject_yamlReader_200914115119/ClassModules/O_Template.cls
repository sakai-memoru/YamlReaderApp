VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "O_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file O_Template.cls
'''  *<br/> Template Utility
'''  *<br/> MiniTemplator.cls Wrapper
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  * @require MiniTemplator
'''  */
'''//--------------------------------------------------------
'''// Instance Property
'''//
''' /** @property @Public @name template_ @type {String} @desc */
Public template_ As String     ''Template Source
''' /** @property @Public @name dict_ @type {Scripting.Dictionary} @desc */
Public dict_ As Dictionary ''Merge Data
''' /** @property @private @name clsTemplator_ @type {MiniTemplator} @desc */
Public clsTemplator_ As MiniTemplator

Private Sub Class_Initialize()
''' ********************************************************
Set clsTemplator_ = New MiniTemplator
Set dict_ = New Dictionary
''
End Sub

Private Sub Class_Terminate()
''' ********************************************************
Set dict_ = Nothing
Set clsTemplator_ = Nothing
''
End Sub


Public Property Set dict(obj As Dictionary)
''' /********************************************************
'''  * @property dictionary setter(Primitve用）
'''  * @param {Object} obj
'''  */
Set Me.dict_ = obj
''
End Property

Public Property Let template(str)
''' /********************************************************
'''  * @property template setter(Primitve用）
'''  * @param {String} str
'''  */
Me.template_ = str
''
End Property

'''//--------------------------------------------------------
'''// Instance Method
'''//
Public Function SetTemplateByName(ByVal nameStr As String, Optional sheetName As Variant, Optional bookName As Variant) As Variant
''' /********************************************************
'''  *@function setTemplate
'''  *<br/>　Set template with a Range Address or Name
'''  * @param {String} rangeparam
'''  * @param {Variant} sheetname
'''  * @return {Variant} source of template
'''  */
If IsMissing(bookName) Then
    bookName = ThisWorkbook.Name
End If
If IsMissing(sheetName) Then
    Me.template_ = ThisWorkbook.Names(nameStr).RefersToRange.Value
Else
    Me.template_ = Workbooks(bookName).Names(nameStr).RefersToRange.Value
End If
''
SetTemplateByName = Me.template_
''
End Function

Public Function SetTemplateByRange(ByVal rng As Range) As Variant
''' /********************************************************
'''  *@function setTemplate
'''  *<br/>　Set template with a Range Address or Name
'''  * @param {String} rangeparam
'''  * @param {Variant} sheetname
'''  * @return {Variant} source of template
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "O_Template.setTemplateByRange"
If rng.Count = 1 Then
    Me.template_ = rng.Value
Else
    Err.Description = "A range size is not single."
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End If
''
SetTemplateByRange = Me.template_
''
End Function

Public Function SetTemplateByFile(ByVal templFile As String, Optional ByVal templFolder As Variant) As Variant
''' /********************************************************
'''  *@function SetTemplateByFile
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "O_Template.SetTemplateByFile"
''
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
If IsMissing(templFolder) Then
    Let templFolder = ThisWorkbook.path
End If
Dim templPath As String
Let templPath = C_File.BuildPath(templFolder, templFile)
Call clsTemplator_.ReadTemplateFromFile(templPath)
Let Me.template_ = clsTemplator_.TemplateString
''
SetTemplateByFile = Me.template_
''
End Function

Public Function merge(Optional ByVal escaped As Boolean = False) As Variant
'''' /********************************************************
''''  *@function merge
''''  *<br/>　merge key and value in dictionary into template
''''  * @param {Boolean} escaped Optional
''''  * @return {Variant} merge
''''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "O_Template.merge"
''
Dim i As Integer
Dim strRet As Variant
''
On Error GoTo EXCEPTION
    clsTemplator_.TemplateString = template_
    ''
    For i = 0 To dict_.Count - 1
        If clsTemplator_.VariableExists(dict_.keys(i)) Then
            If escaped Then
                clsTemplator_.SetVariableEsc dict_.keys(i), dict_.Items(i)
            Else
                clsTemplator_.SetVariable dict_.keys(i), dict_.Items(i)
            End If
        End If
    Next i
    ''
    strRet = clsTemplator_.GenerateOutputToString()
    ''
    merge = strRet
    Exit Function
    ''
EXCEPTION: 'Label
    Console.log ("Err.Source = " & Err.source)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function MergeWithBlock() As Variant
''' /********************************************************
'''  *@function mergeWithBlock
'''  *<br/>Blockを含むTemplateをマージする。
'''  *<br/>Dictionaryに、Block名でArray<Dictionary>を設定。
'''  * @return {Variant}
'''  */
    Dim i As Integer
    Dim j As Integer
    Dim strRet As Variant
    Const CONS_MODULE_NAME As String = "C_Template. mergeWithBlock"
    ''
On Error GoTo EXCEPTION
    clsTemplator_.TemplateString = template_
    ''
    For i = 0 To dict_.Count - 1
        If clsTemplator_.BlockExists(dict_.keys(i)) Then
            If IsArray(dict_(dict_.keys(i))) Then
                For j = LBound(dict_(dict_.keys(i))) To UBound(dict_(dict_.keys(i)))
                    If clsTemplator_.VariableExists(dict_(dict_.keys(i)).keys(j)) Then
                        clsTemplator_.SetVariable dict_(dict_.keys(i)).keys(j), dict_(dict_.keys(i)).Items(i)
                    End If
                Next j
            End If
            clsTemplator_.AddBlock dict_.keys(i)
        End If
        If clsTemplator_.VariableExists(dict_.keys(i)) Then
            clsTemplator_.SetVariable dict_.keys(i), dict_.Items(i)
        End If
    Next i
    ''
    strRet = clsTemplator_.GenerateOutputToString()
    ''
    MergeWithBlock = strRet
    Exit Function
    ''
EXCEPTION: 'Label
    Console.log ("Err.Source = " & Err.source)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function MergeWithBlock2() As Variant
''' /********************************************************
'''  *@function mergeWithBlock
'''  *<br/>Blockを含むTemplateをマージする。
'''  *<br/>Dictionaryに、Block名でArray<Dictionary>を設定。
'''  * @return {Variant}
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "O_Template.mergeWithBlock"
''
Dim i As Integer
Dim j As Integer
Dim strRet As Variant
''
On Error GoTo EXCEPTION
    Let clsTemplator_.TemplateString = template_
    ''
    Dim ary() As Variant
    Dim dictTemp As Dictionary
    Dim ky As Variant
    For i = 0 To dict_.Count - 1
        If clsTemplator_.BlockExists(dict_.keys(i)) Then
            If IsArray(dict_(dict_.keys(i))) Then
                Let ary = dict_(dict_.keys(i))
                For j = LBound(ary) To UBound(ary)
                    Set dictTemp = ary(j)
                    For Each ky In dictTemp
                        ''Console.info ky
                        If clsTemplator_.VariableExists(ky) Then
                            clsTemplator_.SetVariable ky, dictTemp.Item(ky)
                        End If
                    Next ky
                    clsTemplator_.AddBlock dict_.keys(i)
                Next j
            End If
        End If
        If clsTemplator_.VariableExists(dict_.keys(i)) Then
            clsTemplator_.SetVariable dict_.keys(i), dict_.Items(i)
        End If
    Next i
    ''
    Let strRet = clsTemplator_.GenerateOutputToString()
    ''
    Let MergeWithBlock2 = strRet
    Exit Function
    ''
EXCEPTION: 'Label
    Console.log ("Err.Source = " & Err.source)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function ExistsVariable(param As Variant) As Boolean
''' /********************************************************
''
Let ExistsVariable = clsTemplator_.VariableExists(param)
''
End Function

Public Sub Rest()
''' /********************************************************
''
Call clsTemplator_.ResetData
''
End Sub


'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim strHead As String
    Dim strBody As String
    Dim strFoot As String
    strHead = "/********************************************************"
    strFoot = " */"
    strBody = _
    " @functon ${function} " & vbCrLf & _
    "<pre>" & vbCrLf & _
    "${description}" & vbCrLf & _
    "</pre>" & vbCrLf & _
    "<!-- $BeginBlock ForEachParam --> " & vbCrLf & _
    "@param {${type}} ${name} ${description}" & vbCrLf & _
    "<!-- $BeginBlock ForEachParam --> " & vbCrLf & _
    "@return {${type}} ${name} ${description}" & vbCrLf
    
    Dim line As String
    line = "Public Function setTemplate(ByVal rangename As String, Optional sheetname As Variant) As Variant"

    Console.info strBody
    
End Sub

'/////////////////////////////////////////////////////////////
'//  Note        /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' ●MiniTemplator
'-------------------------------------------------------------
''■URL
''http://www.source-code.biz/MiniTemplator/
''
''■template syntax:
''■■Variables:
''       ${VariableName}
''
''■■Blocks:
''       <!-- $beginBlock blockName -->
''         ... block contents ...
''       <!-- $endBlock blockName -->
''
''    Conditional Blocks:
''       <!-- $if flag1 flag2 -->
''         ... included if flag1 or flag2 is set ...
''       <!-- $elseIf !flag3 flag4 -->
''         ... included if flag3 is not set or flag4 is set ...
''       <!-- $else -->
''         ... included if none of the above conditions is met ...
''       <!-- $endIf -->
''
''    Short form of conditional blocks:
''    (only recognized if TemplateSpecification.shortFormEnabled is true)
''       <$? flag1 flag2 >
''         ... included if flag1 or flag2 is set ...
''       <$: !flag3 flag4 >
''         ... included if flag3 is not set or flag4 is set ...
''       <$:>
''         ... included if none of the above conditions is met ...
''       <$/?>
''■■Example:
''       <$?de> Hallo Welt!
''       <$:fr> Bonjour tout le monde!
''       <$:  > Hello world!
''       <$/?>
''
''    Include a subtemplate:
''       <!-- $include relativeFileName -->
''
''■■General remarks:
''
''　・Variable names, block names, condition flags and commands (e.g. "$beginBlock") are case-insensitive.
''　・The same variable may be used multiple times within a template.
''　・Multiple blocks with the same name may occur within a template.
''　・Blocks can be nested.
''　・Conditional blocks ($if) and includes ($include) are resolved when the template is parsed. Parsing
''　　is done within the MiniTemplator constructor. Condition flags can be passed to the constructor
''　　using MiniTemplator.TemplateSpecification.
''　　Normal blocks ($beginBlock) must be added (and can be repeated) by the application program using
''　　addBlock().
''　　The MiniTemplatorCache class may be used to cache MiniTemplator objects with parsed templates.
''
''■Method Detail
''■■loadSubtemplate:
''  protected String loadSubtemplate(String subtemplateName) throws IOException
''■■reset
''  public void reset()
''
''■■setVariable
''  public void setVariable(String variableName, String variableValue, boolean isOptional) throws MiniTemplator.VariableNotDefinedException
''  public void setVariable(String variableName, String variableValue) throws MiniTemplator.VariableNotDefinedException
''  public void setVariable(String variableName, int variableValue) throws MiniTemplator.VariableNotDefinedException
''
''■■setVariableEsc
''  public void setVariableEsc(String variableName, String variableValue, boolean isOptional) throws MiniTemplator.VariableNotDefinedException
''  public void setVariableEsc(String variableName, String variableValue) throws MiniTemplator.VariableNotDefinedException
''
''■■variableExists
''  public boolean variableExists(String variableName)
''
''■■addBlock
''  public void addBlock(String blockName, boolean isOptional) throws MiniTemplator.BlockNotDefinedException
''  public void addBlock(String blockName) throws MiniTemplator.BlockNotDefinedException
''
''■■generateOutput
''  public void generateOutput(String outputFileName) throws IOException
''  public void generateOutput(Writer outputWriter)  throws IOException
''
''■■escapeHtml
''  public static String escapeHtml(String s)
''

