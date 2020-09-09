VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cregXLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:08 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414836/raw/cregXLib.cls
Option Explicit
' v2.02
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
' for building up a library of useful regex expressions
Private pName As String
Private pRegex As RegExp
Public Property Get Pattern() As String
    Pattern = pRegex.Pattern
End Property
Public Property Let Pattern(p As String)
    pRegex.Pattern = p
End Property
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(p As String)
   pName = p
End Property
Public Property Get ignorecase() As Boolean
    ignorecase = pRegex.ignorecase
End Property
Public Property Let ignorecase(p As Boolean)
     pRegex.ignorecase = p
End Property
Public Property Get rGlobal() As Boolean
    rGlobal = pRegex.Global
End Property
Public Property Let rGlobal(p As Boolean)
   pRegex.Global = p
End Property

Public Sub init(sName As String, _
            Optional spat As String = "", _
            Optional bIgnoreSpaces As Boolean = True, _
            Optional bIgnoreCase As Boolean = True, _
            Optional bGlobal As Boolean = True)
    
    Dim s As String
    s = spat
    If bIgnoreSpaces Then
        s = Replace(s, " ", "")
    End If
    Set pRegex = New RegExp
    With pRegex
        .Pattern = s
        .ignorecase = bIgnoreCase
        .Global = bGlobal
    End With
    pName = sName

End Sub
Public Function getString(sFrom As String) As String
    Dim mc As MatchCollection, am As match, rs As String
    Set mc = pRegex.Execute(sFrom)
    rs = ""
    For Each am In mc
        rs = rs & am.Value
    Next am
    getString = rs
End Function
Public Function getGroup(sFrom As String, groupNumber As Long) As String
    Dim mc As MatchCollection, am As match, bm As SubMatches, rs As String
    Set mc = pRegex.Execute(sFrom)
    rs = ""
    
    
    If mc.Count > 1 And mc.Count >= groupNumber Then
        rs = mc.Item(groupNumber - 1).Value
    ElseIf mc.Count = 1 Then
        If mc.Item(0).SubMatches.Count >= groupNumber Then
    ' dont really understand this yet
            rs = mc.Item(0).SubMatches(groupNumber - 1)
        End If
    End If
    
    getGroup = rs
End Function
Function getReplace(sFrom As String, sTo As String) As String
    getReplace = pRegex.Replace(sFrom, sTo)

End Function
Function getTest(sFrom As String) As Boolean
    getTest = pRegex.Test(sFrom)

End Function




