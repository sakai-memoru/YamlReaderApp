VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MiniTemplator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MiniTemplator is a compact template engine for HTML files.
'
' Template syntax:
'
'   Variables:
'     ${VariableName}
'
'   Blocks:
'     <!-- $BeginBlock BlockName -->
'     ... block content ...
'     <!-- $EndBlock BlockName -->
'
'   Include a subtemplate:
'     <!-- $Include RelativeFileName -->
'
' General remarks:
'  - Variable names and block names are case-insensitive.
'  - The same variable may be used multiple times within a template.
'  - Blocks can be nested.
'  - Multiple blocks with the same name may occur within a template.
'
' Public properties:
'   TemplateString          - Contains the content of the template file.
' Public methods:
'   ReadTemplateFromFile    - Reads the template from a file.
'   SetVariable             - Sets a template variable.
'   SetVariableEsc          - Sets a template variable to an escaped string value.
'   VariableExists          - Checks whether a variable exists.
'   AddBlock                - Adds an instance of a template block.
'   BlockExists             - Checks whether a block exists.
'   GenerateOutputToFile    - Generates the HTML page and writes it into a file.
'   GenerateOutputToString  - Generates the HTML page and returns it as a string.
'   ResetData               - Clears all variables and blocks.
'   EscapeHtml              - Escapes special HTML characters.
'
' Home page: www.source-code.biz/MiniTemplator
' License: This module is released under the GNU/LGPL license (www.gnu.org/licenses/lgpl.html).
' Copyright 2003: Christian d'Heureuse, Inventec Informatik AG, Switzerland. All rights reserved.
' This product is provided "as is" without warranty of any kind.
'
' Version history:
' 2001-10-24 Christian d'Heureuse (chdh): Module created.
' 2003-04-06 chdh: Minor changes to adjust to Java version.
' 2003-07-16 chdh: Method VariableExists added.
'   Method SetVariable changed to raise an error when the variable does not exist (instead of returning False).
' 2004-04-07 chdh: Parameter IsOptional added to method SetVariable.
'   Licensing changed from GPL to LGPL.
' 2004-04-18 chdh: Method BlockExists added.
' 2004-10-28 chdh:
'   Methods SetVariableEsc and EscapeHtml added.
'   Multiple blocks with the same name may now occur within a template.
'   No syntax error ("unknown command") is raised any more, if a HTML comment starts with "${".
' 2004-11-06 chdh:
'   "$Include" command implemented. Public property SubtemplateBasePath added.
' 2004-11-20 chdh:
'   "$Include" command changed so that the command text is not copied to the output file.
' 2005-09-27 chdh:
'   Speed of GenerateOutputToString() improved. Thanks to Dale Higgs for the hint.

Option Explicit
Option Compare Text

Private Const MaxNestingLevel = 50                             ' maximum number of block nestings
Private Const MaxInclTemplateSize = 1000000                    ' maximum length of template string when including subtemplates

Private template               As String                       ' content of the template file

Private Type VarTabRec                                         ' variables table record structure
   VarName                     As String                       ' variable name
   VarValue                    As String                       ' variable value
   End Type
Private Type VarRefTabRec                                      ' variable references table record structure
   VarNo                       As Integer                      ' variable no
   TPosBegin                   As Long                         ' template position of begin of variable reference
   TPosEnd                     As Long                         ' template position of end of variable reference
   BlockNo                     As Integer                      ' block no of the (innermost) block that contains this variable reference
   BlockVarNo                  As Integer                      ' block variable no. Index into BlockInstTab.BlockVarTab
   End Type
Private Type BlockTabRec                                       ' blocks table record structure
   BlockName                   As String                       ' block name
   NextWithSameName            As Integer                      ' block no of next block with same name or -1 (blocks are backward linked in relation to template position)
   TPosBegin                   As Long                         ' template position of begin of block
   TPosContentBegin            As Long                         ' template pos of begin of block content
   TPosContentEnd              As Long                         ' template pos of end of block content
   TPosEnd                     As Long                         ' template position of end of block
   NestingLevel                As Integer                      ' block nesting level
   ParentBlockNo               As Integer                      ' block no of parent block
   DefinitionIsOpen            As Boolean                      ' true while $BeginBlock processed but no $EndBlock
   Instances                   As Long                         ' number of instances of this block
   FirstBlockInstNo            As Long                         ' block instance no of first instance of this block or -1
   LastBlockInstNo             As Long                         ' block instance no of last instance of this block or -1
   CurrBlockInstNo             As Long                         ' current block instance no, used during generation of output file
   BlockVarCnt                 As Integer                      ' no of variables in block
   BlockVarNoToVarNoMap()      As Integer                      ' maps block variable numbers to variable numbers
   FirstVarRefNo               As Long                         ' variable reference no of first variable of this block or -1
   End Type
Private Type BlockInstTabRec                                   ' block instances table record structure
   BlockNo                     As Integer                      ' block number
   InstanceLevel               As Long                         ' instance level of this block
      ' InstanceLevel is an instance counter per block.
      ' (In contrast to BlockInstNo, which is an instance counter over the instances of all blocks)
   ParentInstLevel             As Long                         ' instance level of parent block
   NextBlockInstNo             As Long                         ' pointer to next instance of this block or -1
      ' Forward chain for instances of same block.
   BlockVarTab()               As String                       ' block instance variables
   End Type

Private VarTab()               As VarTabRec                    ' variables table, array index is variable no
Private VarTabCnt              As Integer                      ' no of entries used in VarTab
Private VarNameToNoMap         As Collection                   ' maps variable names to variable numbers
Private VarRefTab()            As VarRefTabRec                 ' variable references table
   ' Contains an entry for each variable reference in the template. Ordered by TemplatePos.
Private VarRefTabCnt           As Integer                      ' no of entries used in VarRefTab
Private BlockTab()             As BlockTabRec                  ' Blocks table, array index is block no
   ' Contains an entry for each block in the template. Ordered by TPosBegin.
Private BlockTabCnt            As Integer                      ' no of entries used in BlockTab
Private BlockNameToNoMap       As Collection                   ' maps block names to block numbers
Private OpenBlocksTab(0 To MaxNestingLevel) As Integer
   ' During parsing, this table contains the block numbers of the open parent blocks (nested outer blocks).
   ' Indexed by the block nesting level.
Private BlockInstTab()         As BlockInstTabRec              ' block instances table
   ' This table contains an entry for each block instance that has been added.
   ' Indexed by BlockInstNo.
Private BlockInstTabCnt        As Long                         ' no of entries used in BlockInstTab
Private TemplateValid          As Boolean                      ' true if a valid template is prepared

Private CurrentNestingLevel    As Integer                      ' current block nesting level during parsing
Private OutputMode             As Integer                      ' 0 = to file, 1 = to string
Private OutputFileHandle       As Integer                      ' file handle during writing of output file
Private OutputStringChunks()   As String                       ' string chunks of the generated HTML page
Private OutputStringChunksCnt  As Integer                      ' number of used entries in OutputStringChunks

'--- public properties -------------------------------------------------------------------------------------------------

' Base path for relative file names of subtemplates (for the $Include command).
' This path is prepended to the subtemplate file names. It must be set before
' ReadTemplateFromFile is called or the TemplateString property is set.
Public SubtemplateBasePath     As String

'--- template string handling ------------------------------------------------------------------------------------------

' Reads the template from a file.
Public Sub ReadTemplateFromFile(ByVal fileName As String)
   TemplateString = ReadFileIntoString(fileName)
   End Sub

' Assigns a new template string.
Public Property Let TemplateString(ByVal NewTemplate As String)
   TemplateValid = False
   template = NewTemplate
   ParseTemplate
   ResetData
   TemplateValid = True
   End Property

' Returns the template string.
Public Property Get TemplateString() As String
   TemplateString = template
   End Property

' Loads the template string for a subtemplate (used for the $Include command).
Private Function LoadSubtemplate(ByVal SubtemplateName)
   Dim SubtemplateFileName: SubtemplateFileName = CombineFileSystemPath(SubtemplateBasePath, SubtemplateName)
   LoadSubtemplate = ReadFileIntoString(SubtemplateFileName)
   End Function

'--- template parsing --------------------------------------------------------------------------------------------------

Private Sub ParseTemplate()
   InitParsing
   BeginMainBlock
   ParseTemplateCommands
   EndMainBlock
   CheckBlockDefinitionsComplete
   ParseTemplateVariables
   AssociateVariablesWithBlocks
   End Sub

Private Sub InitParsing()
   ReDim VarTab(0 To 255)
   VarTabCnt = 0
   Set VarNameToNoMap = New Collection
   ReDim VarRefTab(0 To 255)
   VarRefTabCnt = 0
   ReDim BlockTab(0 To 255)
   BlockTabCnt = 0
   Set BlockNameToNoMap = New Collection
   End Sub

' Registers the main block.
' The main block is an implicitly defined block that covers the whole template.
Private Sub BeginMainBlock()
   Dim BlockNo As Integer
   RegisterBlock "$InternalMainBlock$", BlockNo
   With BlockTab(BlockNo)
      .TPosBegin = 1
      .TPosContentBegin = 1
      .NestingLevel = 0
      .ParentBlockNo = -1
      .DefinitionIsOpen = True
      End With
   OpenBlocksTab(0) = BlockNo
   CurrentNestingLevel = 1
   End Sub

' Completes the main block registration.
Private Sub EndMainBlock()
   With BlockTab(0)
      .TPosContentEnd = Len(template) + 1
      .TPosEnd = Len(template) + 1
      .DefinitionIsOpen = False
      End With
   CurrentNestingLevel = CurrentNestingLevel - 1
   End Sub

' Parses commands within the template in the format "<!-- $command parameters -->".
Private Sub ParseTemplateCommands()
   Dim p As Long: p = 1
   Do
      Dim p0 As Long: p0 = InStr(p, template, "<!--", vbBinaryCompare)
      If p0 = 0 Then Exit Do
      p = InStr(p0, template, "-->", vbBinaryCompare)
      If p = 0 Then Err.Raise vbObjectError, , "Invalid HTML comment in template at offset " & p0 - 1 & "."
      p = p + 3
      Dim cmdl As String: cmdl = Mid$(template, p0 + 4, p - p0 - 7)
      Dim ResumeFromStart As Boolean
      ProcessTemplateCommand cmdl, p0, p, ResumeFromStart
      If ResumeFromStart Then p = p0
      Loop
   End Sub

Private Sub ProcessTemplateCommand(ByVal cmdl As String, ByVal CmdTPosBegin As Long, ByVal CmdTPosEnd As Long, ResumeFromStart As Boolean)
   ResumeFromStart = False
   Dim p As Integer: p = 1
   Dim cmd As String
   If Not ParseWord(cmdl, p, cmd) Then Exit Sub
   Dim parms As String: parms = Mid$(cmdl, p)
   Select Case cmd
      Case "$BeginBlock": ProcessBeginBlockCmd parms, CmdTPosBegin, CmdTPosEnd
      Case "$EndBlock":   ProcessEndBlockCmd parms, CmdTPosBegin, CmdTPosEnd
      Case "$Include":    ProcessincludeCmd parms, CmdTPosBegin, CmdTPosEnd: ResumeFromStart = True
      Case Else
         If left$(cmd, 1) = "$" And left$(cmd, 2) <> "${" Then _
            Err.Raise vbObjectError, , "Unknown command """ & cmd & """ in template at offset " & CmdTPosBegin - 1 & "."
      End Select
   End Sub

' Processes the $BeginBlock command.
Private Sub ProcessBeginBlockCmd(ByVal parms As String, ByVal CmdTPosBegin As Long, ByVal CmdTPosEnd As Long)
   Dim p As Integer: p = 1
   Dim BlockName As String
   If Not ParseWord(parms, p, BlockName) Then Err.Raise vbObjectError, , "Missing block name in $BeginBlock command in template at offset " & CmdTPosBegin - 1 & "."
   If Trim$(Mid$(parms, p)) <> "" Then Err.Raise vbObjectError, , "Extra parameter in $BeginBlock command in template at offset " & CmdTPosBegin - 1 & "."
   Dim BlockNo As Integer
   RegisterBlock BlockName, BlockNo
   With BlockTab(BlockNo)
      .TPosBegin = CmdTPosBegin
      .TPosContentBegin = CmdTPosEnd
      .NestingLevel = CurrentNestingLevel
      .ParentBlockNo = OpenBlocksTab(CurrentNestingLevel - 1)
      End With
   OpenBlocksTab(CurrentNestingLevel) = BlockNo
   CurrentNestingLevel = CurrentNestingLevel + 1
   If CurrentNestingLevel > MaxNestingLevel Then Err.Raise vbObjectError, , "Block nesting overflow in template at offset " & CmdTPosBegin - 1 & "."
   End Sub

' Processes the $EndBlock command.
Private Sub ProcessEndBlockCmd(ByVal parms As String, ByVal CmdTPosBegin As Long, ByVal CmdTPosEnd As Long)
   Dim p As Integer: p = 1
   Dim BlockName As String
   If Not ParseWord(parms, p, BlockName) Then Err.Raise vbObjectError, , "Missing block name in $EndBlock command in template at offset " & CmdTPosBegin - 1 & "."
   If Trim$(Mid$(parms, p)) <> "" Then Err.Raise vbObjectError, , "Extra parameter in $EndBlock command in template at offset " & CmdTPosBegin - 1 & "."
   Dim BlockNo As Integer
   If Not LookupBlockName(BlockName, BlockNo) Then Err.Raise vbObjectError, , "Undefined block name """ & BlockName & """ in $EndBlock command in template at offset " & CmdTPosBegin - 1 & "."
   CurrentNestingLevel = CurrentNestingLevel - 1
   With BlockTab(BlockNo)
      If Not .DefinitionIsOpen Then Err.Raise vbObjectError, , "Multiple $EndBlock command for block """ & BlockName & """ in template at offset " & CmdTPosBegin - 1 & "."
      If .NestingLevel <> CurrentNestingLevel Then Err.Raise vbObjectError, , "Block nesting level mismatch at $EndBlock command for block """ & BlockName & """ in template at offset " & CmdTPosBegin - 1 & "."
      .TPosContentEnd = CmdTPosBegin
      .TPosEnd = CmdTPosEnd
      .DefinitionIsOpen = False
      End With
   End Sub

Private Sub RegisterBlock(ByVal BlockName As String, BlockNo As Integer)
   If BlockTabCnt > UBound(BlockTab) Then _
      ReDim Preserve BlockTab(0 To 2 * BlockTabCnt)
   BlockNo = BlockTabCnt
   Dim BlockWithSameName As Integer: If Not LookupBlockName(BlockName, BlockWithSameName) Then BlockWithSameName = -1
   With BlockTab(BlockNo)
      .BlockName = BlockName
      .NextWithSameName = BlockWithSameName
      .DefinitionIsOpen = True
      .Instances = 0
      .FirstBlockInstNo = -1
      .LastBlockInstNo = -1
      .BlockVarCnt = 0
      .FirstVarRefNo = -1
      ReDim .BlockVarNoToVarNoMap(0 To 31)
      End With
   If BlockWithSameName <> -1 Then BlockNameToNoMap.remove BlockName
   BlockNameToNoMap.Add BlockNo, BlockName
   BlockTabCnt = BlockTabCnt + 1
   End Sub

' Checks that all block definitions are closed.
Private Sub CheckBlockDefinitionsComplete()
   Dim BlockNo As Integer
   For BlockNo = 0 To BlockTabCnt - 1
      With BlockTab(BlockNo)
         If .DefinitionIsOpen Then _
            Err.Raise vbObjectError, , "Missing $EndBlock command in template for block """ & .BlockName & """."
         End With
      Next
   If CurrentNestingLevel <> 0 Then _
      Err.Raise vbObjectError, , "Block nesting level error at end of template."
   End Sub

' Processes the $Include command.
Private Sub ProcessincludeCmd(ByVal parms As String, ByVal CmdTPosBegin As Long, ByVal CmdTPosEnd As Long)
   Dim p As Integer: p = 1
   Dim SubtemplateName As String
   If Not ParseQuotedStringOrWord(parms, p, SubtemplateName) Then Err.Raise vbObjectError, , "Missing or invalid subtemplate name in $Include command in template at offset " & CmdTPosBegin - 1 & "."
   If Trim$(Mid$(parms, p)) <> "" Then Err.Raise vbObjectError, , "Extra parameter in $Include command in template at offset " & CmdTPosBegin - 1 & "."
   InsertSubtemplate SubtemplateName, CmdTPosBegin, CmdTPosEnd
   End Sub

Private Sub InsertSubtemplate(ByVal SubtemplateName As String, ByVal TPos1 As Long, ByVal TPos2 As Long)
   If Len(template) > MaxInclTemplateSize Then _
      Err.Raise vbObjectError, , "Subtemplate include aborted because the internal template string is longer than " & MaxInclTemplateSize & " characters."
   Dim Subtemplate As String
   'Subtemplate = LoadSubtempla@te(SubtemplateName) ''FIXME
   ' (Copying the template to insert a subtemplate is a bit slow. In a future implementation of MiniTemplator,
   ' a table could be used that contains references to the string fragments.)
   template = left$(template, TPos1 - 1) & Subtemplate & Mid$(template, TPos2)
   End Sub

' Parses variable references within the template in the format "${VarName}" .
Private Sub ParseTemplateVariables()
   Dim p As Long: p = 1
   Do
      p = InStr(p, template, "${", vbBinaryCompare)
      If p = 0 Then Exit Do
      Dim p0 As Long: p0 = p
      p = InStr(p, template, "}", vbBinaryCompare)
      If p = 0 Then Err.Raise vbObjectError, , "Invalid variable reference in template at offset " & p0 - 1 & "."
      p = p + 1
      Dim VarName As String: VarName = Trim$(Mid$(template, p0 + 2, p - p0 - 3))
      If Len(VarName) = 0 Then Err.Raise vbObjectError, , "Empty variable name in template at offset " & p0 - 1 & "."
      RegisterVariableReference VarName, p0, p
      Loop
   End Sub

Private Sub RegisterVariableReference(ByVal VarName As String, ByVal TPosBegin As Long, ByVal TPosEnd As Long)
   Dim VarNo As Integer
   If Not LookupVariableName(VarName, VarNo) Then _
      RegisterVariable VarName, VarNo
   If VarRefTabCnt > UBound(VarRefTab) Then _
      ReDim Preserve VarRefTab(0 To 2 * VarRefTabCnt)
   Dim VarRefNo As Integer: VarRefNo = VarRefTabCnt
   With VarRefTab(VarRefNo)
      .TPosBegin = TPosBegin
      .TPosEnd = TPosEnd
      .VarNo = VarNo
      End With
   VarRefTabCnt = VarRefTabCnt + 1
   End Sub

Private Sub RegisterVariable(ByVal VarName As String, VarNo As Integer)
   If VarTabCnt > UBound(VarTab) Then _
      ReDim Preserve VarTab(0 To 2 * VarTabCnt)
   VarNo = VarTabCnt
   With VarTab(VarNo)
      .VarName = VarName
      .VarValue = vbNullString
      End With
   VarNameToNoMap.Add VarNo, VarName
   VarTabCnt = VarTabCnt + 1
   End Sub

' Associates variable references with blocks.
Private Sub AssociateVariablesWithBlocks()
   Dim VarRefNo As Integer: VarRefNo = 0
   Dim ActiveBlockNo As Integer: ActiveBlockNo = 0
   Dim NextBlockNo As Integer: NextBlockNo = 1
   Do While VarRefNo < VarRefTabCnt
      With VarRefTab(VarRefNo)
         Dim VarRefTPos As Long: VarRefTPos = .TPosBegin
         Dim VarNo As Integer: VarNo = .VarNo
         End With
      If VarRefTPos >= BlockTab(ActiveBlockNo).TPosEnd Then
         ActiveBlockNo = BlockTab(ActiveBlockNo).ParentBlockNo
         GoTo AssocVarsWithBlocksIterate
         End If
      If NextBlockNo < BlockTabCnt Then
         If VarRefTPos >= BlockTab(NextBlockNo).TPosBegin Then
            ActiveBlockNo = NextBlockNo
            NextBlockNo = NextBlockNo + 1
            GoTo AssocVarsWithBlocksIterate
            End If
         End If
      With BlockTab(ActiveBlockNo)
         If VarRefTPos < .TPosBegin Then ProgramLogicError 1
         If .BlockVarCnt > UBound(.BlockVarNoToVarNoMap) Then _
            ReDim Preserve .BlockVarNoToVarNoMap(0 To 2 * .BlockVarCnt)
         Dim BlockVarNo As Integer: BlockVarNo = .BlockVarCnt
         .BlockVarNoToVarNoMap(BlockVarNo) = VarNo
         .BlockVarCnt = .BlockVarCnt + 1
         If .FirstVarRefNo = -1 Then .FirstVarRefNo = VarRefNo
         End With
      With VarRefTab(VarRefNo)
         .BlockNo = ActiveBlockNo
         .BlockVarNo = BlockVarNo
         End With
      VarRefNo = VarRefNo + 1
AssocVarsWithBlocksIterate:
      Loop
   End Sub

'--- build up (template variables and blocks) --------------------------------------------------------------------------

' Clears all variables and blocks.
' This method can be used to produce another HTML page with the same
' template. It is faster than creating another MiniTemplator object,
' because the template does not have to be parsed again.
' All variable values are cleared and all added block instances are deleted.
Public Sub ResetData()
   Dim VarNo As Integer
   For VarNo = 0 To VarTabCnt - 1
      VarTab(VarNo).VarValue = ""
      Next
   Dim BlockNo As Integer
   For BlockNo = 0 To BlockTabCnt - 1
      With BlockTab(BlockNo)
         .Instances = 0
         .FirstBlockInstNo = -1
         .LastBlockInstNo = -1
         End With
      Next
   ReDim BlockInstTab(0 To 255)
   BlockInstTabCnt = 0
   End Sub

' Sets a template variable.
' For variables that are used in blocks, the variable value
' must be set before addBlock is called.
' If the variable does not exist within the template and IsOptional
' is False, a runtime-error is generated.
Public Sub SetVariable(ByVal VariableName As String, ByVal VariableValue As String, Optional ByVal IsOptional As Boolean = False)
   If Not TemplateValid Then Err.Raise vbObjectError, , "Template not valid."
   Dim VarNo As Integer
   If Not LookupVariableName(VariableName, VarNo) Then
      If IsOptional Then Exit Sub
      Err.Raise vbObjectError, , "Variable """ & VariableName & """ not defined in template."
      End If
   VarTab(VarNo).VarValue = VariableValue
   End Sub

' Sets a template variable to an escaped string value.
' This method is identical to SetVariable, except that the characters <, >,
' &, ' and " of variableValue are replaced by their corresponding HTML/XML
' character entity codes.
' For variables that are used in blocks, the variable value
' must be set before addBlock is called.
' If the variable does not exist within the template and IsOptional
' is False, a runtime-error is generated.
Public Sub SetVariableEsc(ByVal VariableName As String, ByVal VariableValue As String, Optional ByVal IsOptional As Boolean = False)
   SetVariable VariableName, EscapeHtml(VariableValue), IsOptional
   End Sub

' Checks whether a variable with the specified name exists within the template.
' Returns True if the variable exists, or False if no variable with the
' specified name exists in the template.
Public Function VariableExists(ByVal VariableName As String) As Boolean
   If Not TemplateValid Then Err.Raise vbObjectError, , "Template not valid."
   Dim VarNo As Integer
   VariableExists = LookupVariableName(VariableName, VarNo)
   End Function

' Adds an instance of a template block.
' If the block contains variables, these variables must be set
' before the block is added.
' If the block contains subblocks (nested blocks), the subblocks
' must be added before this block is added.
' If multiple blocks exist with the specified name, an instance
' is added for each block occurence.
Public Sub AddBlock(ByVal BlockName As String)
   If Not TemplateValid Then Err.Raise vbObjectError, , "Template not valid."
   Dim BlockNo As Integer
   If Not LookupBlockName(BlockName, BlockNo) Then _
      Err.Raise vbObjectError, , "Block """ & BlockName & """ not defined in template."
   Do While BlockNo <> -1
      AddBlockByNo BlockNo
      BlockNo = BlockTab(BlockNo).NextWithSameName
      Loop
   End Sub

Private Sub AddBlockByNo(ByVal BlockNo As Integer)
   Dim BlockInstNo As Long
   RegisterBlockInstance BlockInstNo
   With BlockTab(BlockNo)
      Dim InstanceLevel As Long: InstanceLevel = .Instances
      .Instances = .Instances + 1
      If .FirstBlockInstNo = -1 Then .FirstBlockInstNo = BlockInstNo
      If .LastBlockInstNo <> -1 Then _
         BlockInstTab(.LastBlockInstNo).NextBlockInstNo = BlockInstNo        ' set forward pointer of chain
      .LastBlockInstNo = BlockInstNo
      Dim ParentBlockNo As Integer: ParentBlockNo = .ParentBlockNo
      Dim BlockVarCnt As Integer: BlockVarCnt = .BlockVarCnt
      End With
   With BlockInstTab(BlockInstNo)
      .BlockNo = BlockNo
      .InstanceLevel = InstanceLevel
      If ParentBlockNo = -1 Then
         .ParentInstLevel = -1
       Else
         .ParentInstLevel = BlockTab(ParentBlockNo).Instances
         End If
      .NextBlockInstNo = -1
      If BlockVarCnt > 0 Then _
         ReDim .BlockVarTab(0 To BlockVarCnt - 1)
      Dim BlockVarNo As Integer
      For BlockVarNo = 0 To BlockVarCnt - 1                                   ' copy instance variables for this block
         Dim VarNo As Integer: VarNo = BlockTab(BlockNo).BlockVarNoToVarNoMap(BlockVarNo)
         .BlockVarTab(BlockVarNo) = VarTab(VarNo).VarValue
         Next
      End With
   End Sub

Private Sub RegisterBlockInstance(BlockInstNo As Long)
   If BlockInstTabCnt > UBound(BlockInstTab) Then _
      ReDim Preserve BlockInstTab(0 To 2 * BlockInstTabCnt)
   BlockInstNo = BlockInstTabCnt
   BlockInstTabCnt = BlockInstTabCnt + 1
   End Sub

' Checks whether a block with the specified name exists within the template.
' Returns True if the block exists, or False if no block with the
' specified name exists in the template.
Public Function BlockExists(ByVal BlockName As String) As Boolean
   If Not TemplateValid Then Err.Raise vbObjectError, , "Template not valid."
   Dim BlockNo As Integer
   BlockExists = LookupBlockName(BlockName, BlockNo)
   End Function

'--- output generation -------------------------------------------------------------------------------------------------

' Generates the HTML page and writes it into a file.
Public Sub GenerateOutputToFile(ByVal fileName As String)
   On Error Resume Next
      Kill fileName
      Dim err_c, err_d: err_c = Err.Number: err_d = Err.Description
      On Error GoTo 0
   If err_c <> 0 And err_c <> 53 Then Err.Raise vbObjectError, , "Error while deleting old version of " & fileName & ". " & err_d
   OutputFileHandle = FreeFile
   On Error Resume Next
      Open fileName For Binary Access Write Lock Read Write As OutputFileHandle
      err_c = Err.Number: err_d = Err.Description
      On Error GoTo 0
   If err_c <> 0 Then Err.Raise vbObjectError, , "Unable to open file " & fileName & ". " & err_d
   OutputMode = 0
   On Error Resume Next
      GenerateOutput
      Close OutputFileHandle
      err_c = Err.Number: err_d = Err.Description
      On Error GoTo 0
   If err_c <> 0 Then Err.Raise err_c, , "Error while writing file " & fileName & ". " & err_d
   End Sub

' Generates the HTML page and returns it as a string.
Public Function GenerateOutputToString() As String
   OutputMode = 1
   ReDim OutputStringChunks(0 To 31)
   OutputStringChunksCnt = 0
   GenerateOutput
   GenerateOutputToString = Join(OutputStringChunks, "")
   Erase OutputStringChunks
   End Function

Private Sub GenerateOutput()
   If Not TemplateValid Then Err.Raise vbObjectError, , "Template not valid."
   If BlockTab(0).Instances = 0 Then _
      AddBlockByNo 0                                           ' add main block
   Dim BlockNo As Integer
   For BlockNo = 0 To BlockTabCnt - 1
      With BlockTab(BlockNo)
         .CurrBlockInstNo = .FirstBlockInstNo
         End With
      Next
   WriteBlockInstances 0, -1
   End Sub

' Writes all instances of a block that are contained within a specific
' parent block instance.
' Called recursively.
Private Sub WriteBlockInstances(ByVal BlockNo As Integer, ByVal ParentInstLevel As Integer)
   Do
      Dim BlockInstNo As Long: BlockInstNo = BlockTab(BlockNo).CurrBlockInstNo
      If BlockInstNo = -1 Then Exit Do
      With BlockInstTab(BlockInstNo)
         If .ParentInstLevel < ParentInstLevel Then ProgramLogicError 2
         If .ParentInstLevel > ParentInstLevel Then Exit Do
         WriteBlockInstance BlockInstNo
         BlockTab(BlockNo).CurrBlockInstNo = .NextBlockInstNo
         End With
      Loop
   End Sub

Private Sub WriteBlockInstance(ByVal BlockInstNo As Integer)
   Dim BlockNo As Integer: BlockNo = BlockInstTab(BlockInstNo).BlockNo
   Dim TPos As Long: TPos = BlockTab(BlockNo).TPosContentBegin
   Dim SubBlockNo As Integer: SubBlockNo = BlockNo + 1
   Dim VarRefNo As Integer: VarRefNo = BlockTab(BlockNo).FirstVarRefNo
   Do
      Dim TPos2 As Long, kind As Integer
      TPos2 = BlockTab(BlockNo).TPosContentEnd: kind = 0       ' assume end-of-block
      If VarRefNo <> -1 And VarRefNo < VarRefTabCnt Then       ' check for variable reference
         With VarRefTab(VarRefNo)
            If .TPosBegin < TPos Then
               VarRefNo = VarRefNo + 1
               GoTo WriteBlockInstanceIterate
               End If
            If .TPosBegin < TPos2 Then
               TPos2 = .TPosBegin
               kind = 1
               End If
            End With
         End If
      If SubBlockNo < BlockTabCnt Then                         ' check for subblock
         With BlockTab(SubBlockNo)
            If .TPosBegin < TPos Then
               SubBlockNo = SubBlockNo + 1
               GoTo WriteBlockInstanceIterate
               End If
            If .TPosBegin < TPos2 Then
               TPos2 = .TPosBegin
               kind = 2
               End If
            End With
         End If
      If TPos2 > TPos Then _
         WriteString Mid$(template, TPos, TPos2 - TPos)
      Select Case kind
         Case 0   ' end of block
            Exit Do
         Case 1   ' variable
            With VarRefTab(VarRefNo)
               If .BlockNo <> BlockNo Then ProgramLogicError 4
               WriteString BlockInstTab(BlockInstNo).BlockVarTab(.BlockVarNo)
               TPos = .TPosEnd
               End With
            VarRefNo = VarRefNo + 1
         Case 2   ' sub block
            With BlockTab(SubBlockNo)
               If .ParentBlockNo <> BlockNo Then ProgramLogicError 3
               WriteBlockInstances SubBlockNo, BlockInstTab(BlockInstNo).InstanceLevel
               TPos = .TPosEnd
               End With
            SubBlockNo = SubBlockNo + 1
         End Select
WriteBlockInstanceIterate:
      Loop
   End Sub

Private Sub WriteString(ByVal s As String)
   Select Case OutputMode
      Case 0                                                   ' output to file
         Put OutputFileHandle, , s
      Case 1                                                   ' output to string
         If OutputStringChunksCnt > UBound(OutputStringChunks) Then _
            ReDim Preserve OutputStringChunks(0 To 2 * UBound(OutputStringChunks) + 1)
         OutputStringChunks(OutputStringChunksCnt) = s
         OutputStringChunksCnt = OutputStringChunksCnt + 1
      End Select
   End Sub

'--- name lookup routines ----------------------------------------------------------------------------------------------

' Maps variable name to variable number.
Private Function LookupVariableName(ByVal VarName As String, VarNo As Integer) As Boolean
   On Error Resume Next
      VarNo = VarNameToNoMap.Item(VarName)
      LookupVariableName = Err.Number = 0
      On Error GoTo 0
   End Function

' Maps block name to block number.
' If there are multiple blocks with the same name, the block number of the last
' registered block with that name is returned.
Private Function LookupBlockName(ByVal BlockName As String, BlockNo As Integer) As Boolean
   On Error Resume Next
      BlockNo = BlockNameToNoMap.Item(BlockName)
      LookupBlockName = Err.Number = 0
      On Error GoTo 0
   End Function

'--- general utility routines ------------------------------------------------------------------------------------------

' Reads a file into a string.
' Returns the string with the file content.
Private Function ReadFileIntoString(ByVal fileName As String) As String
   On Error Resume Next
      ' First we have to check whether the file exists, because "Open For Binary" would
      ' create a file if it does not exist.
      GetAttr fileName
      Dim ErrNo, ErrDescr: ErrNo = Err.Number: ErrDescr = Err.Description
      On Error GoTo 0
   If ErrNo <> 0 Then Err.Raise vbObjectError, , "Error while reading " & fileName & ". " & ErrDescr
   Dim fh: fh = FreeFile
   On Error Resume Next
      Open fileName For Binary Access Read Lock Write As fh
      ErrNo = Err.Number: ErrDescr = Err.Description
      On Error GoTo 0
   If ErrNo <> 0 Then Err.Raise vbObjectError, , "Error while reading " & fileName & ". " & ErrDescr
   On Error Resume Next
      Dim s As String: s = Input$(LOF(fh), fh)
      Close fh
      ErrNo = Err.Number: ErrDescr = Err.Description
      On Error GoTo 0
   If ErrNo <> 0 Then Err.Raise vbObjectError, , "Error while reading " & fileName & ". " & ErrDescr
   ReadFileIntoString = s
   End Function

' Parses the next word w within the string s starting at position p.
' Returns False when the end of the string is reached.
Private Function ParseWord(ByVal s As String, p As Integer, w As String) As Boolean
   p = SkipBlanks(s, p)
   If p > Len(s) Then Exit Function
   Dim p0 As Integer: p0 = p
   p = SkipNonBlanks(s, p)
   w = Mid$(s, p0, p - p0)
   ParseWord = True
   End Function

' Parses a quoted string within the string s starting at position p.
' Returns False when no valid quoted string is available.
Private Function ParseQuotedString(ByVal s As String, p As Integer, w As String) As Boolean
   p = SkipBlanks(s, p)
   If p > Len(s) Then Exit Function
   If Mid$(s, p, 1) <> """" Then Exit Function
   p = p + 1
   Dim p0 As Integer: p0 = p
   p = InStr(p, s, """", vbBinaryCompare)
   If p = 0 Then Exit Function
   w = Mid$(s, p0, p - p0)
   p = p + 1
   ParseQuotedString = True
   End Function

' If the next non-blank character in string s is a quote, a quoted string is parsed.
' Otherwise a word is parsed.
' Returns False when no valid quoted string or word is available.
Private Function ParseQuotedStringOrWord(ByVal s As String, p As Integer, w As String) As Boolean
   p = SkipBlanks(s, p)
   If p > Len(s) Then Exit Function
   If Mid$(s, p, 1) = """" Then
      ParseQuotedStringOrWord = ParseQuotedString(s, p, w)
    Else
      ParseQuotedStringOrWord = ParseWord(s, p, w)
      End If
   End Function

Private Function SkipBlanks(ByVal s As String, ByVal p As Integer) As Integer
   Dim s_len As Integer: s_len = Len(s)
   Do While p <= s_len
      Dim c As Long: c = AscW(Mid$(s, p, 1))
      If c > 32 Or c < 0 Then Exit Do
      p = p + 1
      Loop
   SkipBlanks = p
   End Function

Private Function SkipNonBlanks(ByVal s As String, ByVal p As Integer) As Integer
   Dim s_len As Integer: s_len = Len(s)
   Do While p <= s_len
      Dim c As Long: c = AscW(Mid$(s, p, 1))
      If c <= 32 And c >= 0 Then Exit Do
      p = p + 1
      Loop
   SkipNonBlanks = p
   End Function

Private Sub ProgramLogicError(ByVal ErrorId)
   ' Should never happen.
   Err.Raise vbObjectError, , "MiniTemplator: Program logic error " & ErrorId & "."
   End Sub

' Escapes special HTML characters.
' Replaces the characters <, >, &, ' and " by their corresponding HTML/XML character entity codes.
Public Function EscapeHtml(ByVal s As String) As String
   Dim p As Integer, s2 As String
   For p = 1 To Len(s)
      Dim c As String: c = Mid$(s, p, 1)
      Select Case c
         Case "<":  s2 = s2 & "&lt;"
         Case ">":  s2 = s2 & "&gt;"
         Case "&":  s2 = s2 & "&amp;"
         Case "'":  s2 = s2 & "&#39;"
         Case """": s2 = s2 & "&#34;"
         Case Else: s2 = s2 & c
         End Select
      Next
   EscapeHtml = s2
   End Function

' Combines two file system path names.
Private Function CombineFileSystemPath(ByVal Path1, ByVal Path2)
   If Path1 = "" Or Path2 = "" Then CombineFileSystemPath = Path2: Exit Function
   If left(Path2, 1) = "\" Then Path2 = Mid(Path2, 2)
   If Right(Path1, 1) = "\" Or Right(Path1, 1) = ":" Then
      CombineFileSystemPath = Path1 & Path2
    Else
      CombineFileSystemPath = Path1 & "\" & Path2
      End If
   End Function
