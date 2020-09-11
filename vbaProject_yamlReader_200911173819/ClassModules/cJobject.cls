VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJobject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:45 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414365/raw
' this is used for object serliazation. Its just basic JSON with only string data types catered for
Option Explicit
' v2.25 3414365
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Private pParent As cJobject
Private pValue As Variant
Private pKey As String
Private pChildren As Collection
Private pValid As Boolean
Private pIndex As Long
Const cNull = "_null"
Const croot = "_deserialization"
Private pFake As Boolean            ' not a real key
Private pisArrayRoot                ' this is the root of an array
Private pPointer As Long            ' this one is used for deserializing string
Private pJstring As String          ' so is this
Private pWhatNext As String
Private pActive As Boolean
Private pJtype As eDeserializeType
Private pBacktrack As cJobject        ' used in parsing
Public Enum eDeserializeType
    eDeserializeNormal
    eDeserializeGoogleWire
End Enum
' this is for treeview - i couldnt find it anywhere
Public Enum tvw
    tvwFirst = 0
    tvwLast = 1
    tvwNext = 2
    tvwPrevious = 3
    tvwChild = 4
End Enum
Public Enum eSort
    eSortNone
    eSortAscending
    eSortDescending
End Enum
Public Enum eTypeofColumn
    eTCdate
    eTCnumeric
    eTCtext
    eTCmixed
    eTCboolean
    eTCunknown
End Enum
Private pbyKey As String
Public Property Get backtrack() As cJobject
    Set backtrack = pBacktrack
End Property
Public Property Set backtrack(back As cJobject)
    Set pBacktrack = back
End Property

Public Property Get Self() As cJobject
    Set Self = Me
End Property
Public Property Get isValid() As Boolean
    isValid = pValid
End Property
Public Property Let setValid(good As Boolean)
    pValid = good
End Property
Public Property Get jString() As String
    jString = pJstring
End Property
Public Property Get fake() As Boolean
    fake = pFake
    If Not pParent Is Nothing Then
        fake = fake And pParent.isArrayRoot
    End If
End Property
Public Property Get childIndex() As Long
    childIndex = pIndex
End Property
Public Property Let childIndex(p As Long)
    pIndex = p
End Property
Public Property Get isArrayRoot() As Boolean
    isArrayRoot = pisArrayRoot
End Property
Public Property Get isArrayMember() As Boolean
    If Not pParent Is Nothing Then
        isArrayMember = pParent.isArrayRoot
    Else
        isArrayMember = False
    End If
End Property
Public Property Let isArrayRoot(p As Boolean)
    pisArrayRoot = p
End Property

Public Property Get Parent() As cJobject
    Set Parent = pParent
End Property
Public Property Set Parent(p As cJobject)
    Set pParent = p
End Property
Public Property Get isRoot() As Boolean
    isRoot = (root Is Me)
End Property
Public Sub clearParent()
    Set pParent = Nothing
End Sub
Public Property Get root() As cJobject
    Dim jo As cJobject
    ' the root is the object with no parent
    Set jo = Me
    While Not jo.Parent Is Nothing
        Set jo = jo.Parent
    Wend
    Set root = jo
End Property
Public Property Get key() As String
    key = pKey
End Property
Public Property Get Value() As Variant
    If IsObject(pValue) Then
        Set Value = pValue
    Else
        Value = pValue
    End If
End Property
Public Function setValue(p As Variant) As cJobject
    If IsObject(p) Then
        Set pValue = p
    Else
        pValue = p
    End If
    Set setValue = Me
End Function
Public Function isVoid(Optional childName As String = vbNullString) As Boolean
    If childName = vbNullString Then
        If isObjValue Then
            isVoid = pValue Is Nothing
        Else
            isVoid = IsEmpty(pValue)
        End If
    Else
        isVoid = child(childName).isVoid
    End If
End Function
Public Function isObjValue(Optional childName As String = vbNullString) As Boolean
    If childName = vbNullString Then
        isObjValue = IsObject(Value)
    Else
        isObjValue = child(childName).isObjValue
    End If
End Function
Public Function cValue(Optional childName As String = vbNullString) As Variant
    If childName = vbNullString Then
        cValue = Value
    Else
        cValue = child(childName).Value
    End If
End Function
Public Function kids(Optional childName As String = vbNullString) As Collection
    If (childName = vbNullString) Then
        Set kids = children
    Else
        Set kids = child(childName).children
    End If
End Function
Public Function toString(Optional childName As String = vbNullString) As String
    If isObjValue Then
        toString = "object:" + TypeName(getObject())
    Else
        toString = CStr(cValue(childName))
    End If
End Function
Public Property Get getObject(Optional childName As String = vbNullString) As Object
    Dim t As cJobject
    If childName <> vbNullString Then
        Set t = Me.child(childName)
    Else
        Set t = Me
    End If
    
    Set getObject = t.Value
   
End Property
Public Property Let Value(p As Variant)
    pValue = p
End Property

Public Property Get children() As Collection
    Set children = pChildren
End Property
Public Property Set children(p As Collection)
    Set pChildren = p
End Property
Public Property Get hasChildren() As Boolean
    hasChildren = False
    If Not pChildren Is Nothing Then
        hasChildren = (pChildren.Count > 0)
    End If
End Property
Public Function deleteChild(childName As String) As cJobject
    ' this deletes a child from the children collection
    Dim jOb As cJobject, target As cJobject
    Set target = childExists(childName)
    
    If (Not target Is Nothing) Then
        children.remove target.childIndex
        For Each jOb In children
            If jOb.childIndex > target.childIndex Then
                jOb.childIndex = jOb.childIndex - 1
            End If
        Next jOb
        target.tearDown
       
    End If
    Set deleteChild = Me
End Function
Public Function valueIndex(v As Variant) As Long
    ' check to see if h is in the cj array
    Dim cj As cJobject
    valueIndex = 0
    For Each cj In children
        If Not cj.isObjValue Then
            If cj.Value = v Then
                valueIndex = cj.childIndex
                Exit Function
            End If
        End If
    Next cj
    
End Function

Public Function toTreeView(tr As Object, Optional bEnableCheckBoxes As Boolean = False) As Object
    ' this populates a treeview with a cJobject
    tr.CheckBoxes = bEnableCheckBoxes
    Set toTreeView = treeViewPopulate(tr, Me)
    
End Function

Private Function treeViewPopulate(tr As Object, cj As cJobject, Optional Parent As cJobject = Nothing)
    Dim c As cJobject, s As String
    s = vbNullString
    If cj.hasChildren Then
        s = cj.key
    Else
        s = cj.key + " : " & cj.toString
    End If
    If (Not Parent Is Nothing) Then
        tr.nodes.Add Parent.fullkey, tvwChild, cj.fullkey, s
    Else
        tr.nodes.Add(, , cj.fullkey, cj.key).Expanded = True
    End If
    For Each c In cj.children
        treeViewPopulate tr, c, cj
    Next c
    Set treeViewPopulate = tr
    
End Function
Public Function init(p As cJobject, Optional k As String = cNull, Optional v As Variant = Empty) As cJobject
    Set pParent = p
    Set pBacktrack = p
    pbyKey = "sortKey"
    
    pFake = (k = cNull)
    If pFake Then
        pKey = CStr(pIndex)
    Else
        pKey = k
    End If

    If Not pParent Is Nothing Then
        If Not child(pKey) Is Nothing Then
            MsgBox ("Programming error " & pKey & " is a duplicate object")
            pValid = False
        Else
            pIndex = pParent.children.Count + 1
            If pFake Then
                pKey = CStr(pIndex)
            End If
            pParent.children.Add Me, pKey
        End If
    End If

    Set init = setValue(v)

End Function

Public Function child(s As String) As cJobject
    Dim aString As Variant, n As Long, jo As cJobject, jc  As cJobject
    
    If Len(s) > 0 Then
        aString = Split(s, ".")
        Set jo = Me
        ' we take something x.y.z and find the child
        For n = LBound(aString) To UBound(aString)
            Set jc = jo.childExists(CStr(aString(n)))
            Set jo = jc
            If jo Is Nothing Then Exit For
        Next n
    End If
    Set child = jo

End Function
Public Function insert(Optional s As String = cNull, Optional v As Variant = Empty) As cJobject
    Dim joNew As cJobject, sk As String
    Set joNew = childExists(s)

    If joNew Is Nothing Then
        ' if its an array, use the child index as the name if there is no name given
        If pisArrayRoot And s = cNull Then
            sk = cNull

        Else
            sk = s
        End If
        
        Set joNew = New cJobject
        joNew.init Me, sk, v
    Else
        If Not IsEmpty(v) Then joNew.setValue v
    End If
    Set insert = joNew
End Function
Public Function Add(Optional k As String = cNull, Optional v As Variant = Empty) As cJobject
    Dim aString As Variant, n As Long, jo As cJobject, jc  As cJobject
    aString = Split(k, ".")
    Set jo = Me
    ' we take something x.y.z and add z with parent of y
    For n = LBound(aString) To UBound(aString)
        Set jc = jo.insert(CStr(aString(n)), v)
        Set jo = jc
    Next n
    Set Add = jo
End Function
Public Function addArray() As cJobject
    pisArrayRoot = True
    Set addArray = Me
End Function
' check if this childExists in current children
Public Function childExists(s As String) As cJobject
    On Error GoTo handle
    Set childExists = pChildren(s)
    Exit Function
handle:
    Set childExists = Nothing
End Function
Private Function unSplitToString(A As Variant, delim As String, _
    Optional startAt As Long = -999, Optional howMany As Long = -999, _
    Optional startAtEnd As Boolean = False) As String
    Dim s As String, c As cStringChunker, i As Long

    ' sort out possible boundaries
    If startAt = -999 Then startAt = LBound(A)
    If howMany = -999 Then howMany = UBound(A) - startAt + 1
    If startAtEnd Then startAt = UBound(A) - howMany + 1
    ' will return nullstring on outside bounds
    If startAt < LBound(A) Or howMany + startAt - 1 > UBound(A) Then
        unSplitToString = vbNullString
    Else
        Set c = New cStringChunker
        ' combine and convert to string
        For i = startAt To startAt + howMany - 1
            c.Add(CStr(A(i))).Add delim
        Next i
        unSplitToString = c.chopIf(delim).content
        Set c = Nothing
    End If
    End Function
Public Function findInArray(prop As String, val As Variant, Optional asKey As Boolean = True) As cJobject
    ' find a particular prop value in an array
    Dim cj As cJobject, p As cJobject
    For Each cj In children
        Set p = cj.childExists(prop)
        If (isSomething(p)) Then
            If (Not p.isObjValue() And Not p.hasChildren) Then
                If compareAsKey(p.Value, val, asKey) Then
                    Set findInArray = p
                    Exit Function
                End If
            End If
        End If
    Next cj
    Set findInArray = Nothing
End Function

Public Function find(s As String) As cJobject
    Dim jo As cJobject, f As cJobject, k As String, fk As String, possible As Boolean
    k = makeKey(s)
    fk = makeKey(fullkey(False))
    
    ' need to deal with find("x.y.z") as well as simple find("x")
    Dim kk As String, A As Variant, B As Variant
    B = Split(fk, ".")
    A = Split(k, ".")
    kk = unSplitToString(B, ".", , arrayLength(A), True)
    
    'now the fullkey is the same number of items as the key to compare it against
    If kk = k Then
        Set f = Me
    ElseIf hasChildren Then
        For Each jo In pChildren
            Set f = jo.find(s)
            If Not f Is Nothing Then Exit For
        Next jo
    End If
    Set find = f
End Function
Public Function convertToArray() As cJobject
    ' here's where have something like {x:{a:'x',b:'y'}} and we need to make {x:[{a:'x',b:'y'}]}
    Dim kids As Collection, newParent As cJobject, jOb As cJobject, newRoot As cJobject, i As Long
    
    ' if its got no kids but has a value then we need to assign that value
    
    If Not hasChildren Then
        addArray
        If Not IsEmpty(Value) Then
            ' make a space for the value
            Add , Value
        Else
            ' do nothing
        End If
        Set convertToArray = Me
    Else
        ' we need to make a space for the object and for each child
        Set kids = children
        ' remove current item
        Parent.children.remove (key)
        ' reset child indices
        i = 0
        For Each jOb In Parent.children
            i = i + 1
            jOb.childIndex = i
        Next jOb
        
        ' add a new version of me
        Set newRoot = Parent.Add(key).addArray

        ' move over contents
        With newRoot.Add
            For Each jOb In kids
                .Add jOb.key, jOb.Value
            Next jOb
        End With
        
        Set convertToArray = newRoot
    End If

    
End Function
Public Function fullkey(Optional includeRoot As Boolean = True) As String
    ' reconstruct full key to parent
    Dim s As String, jo As cJobject
    Set jo = Me
    While Not jo Is Nothing
        If (Not jo.isRoot) Or includeRoot Then s = jo.key & "." & s
        Set jo = jo.Parent
    Wend
    If Len(s) > 0 Then s = left(s, Len(s) - 1)
    fullkey = s
    
End Function

Public Function findByValue(x As Variant) As cJobject
    Dim jOb As cJobject, Result As cJobject
    
    If Not isObjValue() And Value = x Then
        Set findByValue = Me
        Exit Function
    
    Else
        For Each jOb In children
            Set Result = jOb.findByValue(x)
            If Not Result Is Nothing Then
                Set findByValue = Result
                Exit Function
            End If
        Next jOb
    End If
    
End Function
Public Function hasKey() As Boolean
    hasKey = pKey <> vbNullString And _
        pKey <> cNull And _
        (hasChildren Or Not isArrayMember) And Not pFake
End Function
Public Function needsCurly() As Boolean
    needsCurly = hasKey And Not isArrayRoot
    If hasChildren Then
        needsCurly = pChildren(1).hasKey
    End If
    
End Function

Public Function needsSquare() As Boolean

    needsSquare = isArrayRoot

End Function
Public Function stringify(Optional blf As Boolean) As String
    stringify = serialize(blf)
End Function
Public Function serialize(Optional blf As Boolean = False) As String
' make a JSON string of this structure
  Dim t As cStringChunker
  
  Set t = New cStringChunker
  If Not fake Then t.Add "{"
  recurseSerialize Me, t, blf
  If Not fake Then t.Add "}"
 
  serialize = t.content
End Function
Public Property Get needsIndent() As Boolean
    needsIndent = needsCurly Or needsSquare
End Property
Public Function recurseSerialize(jOb As cJobject, Optional soFar As cStringChunker = Nothing, _
                Optional blf As Boolean = False) As cStringChunker
  Dim s As String, jo As cJobject, t As cStringChunker
  Static indent As Long
  If indent = 0 Then indent = 3
  If soFar Is Nothing Then
    Set t = New cStringChunker
  Else
    Set t = soFar
  End If

  If blf And (jOb.hasKey Or jOb.needsCurly) Then t.Add Space(indent)
  
  If jOb.hasKey Then
    t.Add(Quote(jOb.key)).Add (":")
  End If
  
  If Not (jOb.hasChildren Or jOb.isArrayRoot) Then
    If blf And Not jOb.hasKey Then s = s & Space(indent)
    
    Dim vt As VbVarType
    vt = VarType(jOb.Value)
    
    ' an embedded object
    If (jOb.isObjValue()) Then
        t.Add Quote(jOb.toString)
        
    ' translate to null
    ElseIf jOb.isVoid() Then
        t.Add "null"
    
    ' take care to avoid localized version of true/false
    ElseIf vt = vbBoolean Then

        If (jOb.Value) Then
            t.Add ("true")
        Else
            t.Add ("true")
        End If
    
    ' take care to avoid localized version of decimal points
    ElseIf (vt = vbDouble Or _
        vt = vbSingle Or _
        vt = vbCurrency) Then
        
        t.Add Replace(LCase(jOb.toString), ",", ".")
    
    ' numbers shouldnt be quoted
    ElseIf (vt = vbLong Or _
            vt = vbInteger Or _
            vt = vbByte) Then
            
        t.Add (jOb.toString)
   
    ' turn dates to ISO
    ElseIf (vt = vbDate) Then
        t.Add (toISODateTime(jOb.Value))
        
    ' quote and escapeify everything else
    Else
        t.Add Quote(CStr(escapeify(jOb.Value)))
    End If
        
  Else
    ' arrays need squares
    
    If jOb.needsSquare Then t.Add "["
    If jOb.needsCurly Then t.Add "{"
    If blf And Not jOb.isArrayRoot Then t.Add vbLf
    If jOb.needsIndent Then
        indent = indent + 3
    End If
    
    For Each jo In jOb.children
      recurseSerialize(jo, t, blf).Add (",")
      If blf Then t.Add (vbLf)
    Next jo
    
    ' get rid of trailing comma
    t.chopWhile(" ").chopIf(vbLf).chopIf (",")

    
    If jOb.needsIndent Then
        indent = indent - 3
        If blf Then t.Add vbLf
    End If
    If blf Then t.Add Space(indent)
    If jOb.needsCurly Then t.Add "}"
    If jOb.needsSquare Then t.Add " ]"
    
  End If
  Set recurseSerialize = t
End Function

Public Property Get longestFullKey() As Long
    longestFullKey = clongestFullKey(root)
End Property
Public Function clone() As cJobject
    Dim cj As cJobject
    Set cj = New cJobject
    Set cj = cj.init(Nothing).append(Me).children(1)
    cj.clearParent
    Set clone = cj
End Function
Public Function arrayAppend(arrJob As cJobject) As cJobject
    Dim j As cJobject
    For Each j In arrJob.children
        Me.append j
    Next j
    Set arrayAppend = Me
End Function
Public Function mergeArrayOfJobs() As cJobject
    Dim jOb As cJobject, head As cJobject, jo As cJobject
    
    ' Me should be an array of Jobjects, which are themsleves arrays
    ' that needs to be merged into the first object
    Debug.Assert isArrayRoot
    
    For Each jOb In children
        Debug.Assert jOb.isObjValue And jOb.isArrayMember
        If (jOb.childIndex > 1) Then
            For Each jo In jOb.getObject().children
                Set jo.Parent = head
                jo.childIndex = head.children.Count + 1
                head.children.Add jo, CStr(jo.childIndex)
            Next jo
        Else
            Set head = jOb.getObject
        End If
    Next jOb
    
    Set mergeArrayOfJobs = head
End Function
Public Function extractAndBlowArrayOfJobs() As cJobject
    Dim jOb As cJobject, newArray As cJobject, jo As cJobject, head As cJobject, joc As cJobject
    
    ' Me should be an array of Jobjects, which are themsleves objects
    ' will blow out the objects ffor stringifiyinh
    Debug.Assert isArrayRoot
    
    Set newArray = New cJobject
    With newArray.init(Nothing).addArray
        '// select out any objects and add them
        For Each jOb In children
            Set head = .Add
            For Each jo In jOb.children
                If (jo.isObjValue) Then
                    For Each joc In jo.getObject.children
                        Set joc.Parent = head
                        joc.childIndex = head.children.Count
                        head.children.Add joc, joc.key
                    Next joc
                End If
            Next jo
        Next jOb
    End With

    
    Set extractAndBlowArrayOfJobs = newArray
End Function
Public Function replaceBranch(byThis As cJobject) As cJobject
    
    Debug.Assert Not Parent Is Nothing
    Debug.Assert Parent.hasChildren
    
    ' set the parent as me
    Set byThis.Parent = Parent
    
    ' remove me
    remove
    
    ' add me and set my index
    Dim k As String
    k = byThis.key
    byThis.childIndex = byThis.Parent.children.Count + 1
    If byThis.fake Then
        k = CStr(byThis.childIndex)
    End If
    byThis.Parent.children.Add byThis, k
    Set replaceBranch = byThis
    
    ' kill the previous
    Me.tearDown


End Function
Public Function merge(mergeThisIntoMe As cJobject) As cJobject
    ' merge this cjobject with another
    ' items in merged with are replaced with items in Me
    Dim cj As cJobject, p As cJobject
    
    Set p = Me.find(mergeThisIntoMe.fullkey(False))
    
    If p Is Nothing Then
    ' i dont have it yet
        Set p = Me.append(mergeThisIntoMe)
    Else
    ' actually i do have it already
        If p.isArrayRoot Then
            ' but its an array - i need to get rid of it
            Set p = p.remove
            Set p = p.append(mergeThisIntoMe)
        Else
            p.setValue (mergeThisIntoMe.Value)
        End If
    End If
    ' now the other childreb tio merge in
    For Each cj In mergeThisIntoMe.children
       p.merge cj
    Next cj
    Set merge = Me

End Function
Public Function remove() As cJobject
    ' removes a branch
    Dim cj As cJobject, p As cJobject, i As Long
    
    Debug.Assert Not Parent Is Nothing
    Debug.Assert Parent.hasChildren
    
    Parent.children.remove childIndex
    ' fix the childindices
    i = 0
    For Each cj In Parent.children
        i = i + 1
        cj.childIndex = i
    Next cj
    Set remove = Parent

End Function
Public Function append(appendThisToMe As cJobject) As cJobject
    ' append another object to me
    Dim cj As cJobject, p As cJobject

    If appendThisToMe.Parent Is Nothing Then
        Set p = Me.Add(appendThisToMe.key, appendThisToMe.Value)
    
    ElseIf Not appendThisToMe.fake Then
        Set p = Me.Add(appendThisToMe.key, appendThisToMe.Value)
    
    Else
        Set p = Me.Add(, appendThisToMe.Value)
    End If
    
    If appendThisToMe.isArrayRoot Then p.addArray
    For Each cj In appendThisToMe.children
       p.append cj
    Next cj
    Set append = Me
End Function
Public Property Get depth(Optional l As Long = 0) As Long
    Dim jo As cJobject
    l = l + 1
    For Each jo In pChildren
        l = jo.depth(l)
    Next jo
    depth = l
End Property
Private Function clongestFullKey(jOb As cJobject, Optional soFar As Long = 0) As Long
    Dim jo As cJobject
    Dim l As Long
    l = Len(jOb.fullkey)
    If l < soFar Then l = soFar
    If Not jOb.children Is Nothing Then
        For Each jo In jOb.children
            l = clongestFullKey(jo, l)
        Next jo
    End If
    clongestFullKey = l
End Function
Public Property Get formatData(Optional bDebug As Boolean = False) As String
    formatData = cformatdata(root, , bDebug)
End Property
Private Function cformatdata(jOb As cJobject, Optional soFar As String = "", Optional bDebug As Boolean = False) As String
    Dim jo As cJobject, ji As cJobject
    Dim s As String
    s = soFar

        s = s & itemFormat(jOb, bDebug)
        If jOb.hasChildren Then
            For Each ji In jOb.children
                s = cformatdata(ji, s, bDebug)
            Next ji
        End If


    cformatdata = s
End Function
Private Function itemFormat(jo As cJobject, Optional bDebug As Boolean = False) As String
    Dim s As String
    s = jo.fullkey & Space(longestFullKey + 4 - Len(jo.fullkey)) _
            & jo.toString
    If bDebug Then
        s = s + "("
        s = s & "debug: Haskey :" & jo.hasKey & " NeedsCurly :" & jo.needsCurly & " NeedsSquare:" & jo.needsSquare
        s = s + " isArrayMember:" & jo.isArrayMember & " isArrayRoot:" & jo.isArrayRoot & " Fake:" & jo.fake
        s = s & ")"
    
    End If
    itemFormat = s + vbCrLf
End Function
Public Sub jdebug()
    Debug.Print formatData(True)
End Sub
Private Function Quote(s As String) As String
    Quote = q & s & q
End Function
'/**
' *this will take a JSON string and parse it into a cJobject branch under ME
' *it behaves as the JavaScript JSON.parse except it allows for single quoted props & values too
' *@param {string} s the string to be parsed
' *@param {edeserializetype} jtype type of serialization
' *@param {boolean} complain whether to complain in the event of an error
' *@param {boolean} dealWithLikelyTimestamps whether to convert things that look like JS or unix timestamps to excel dates
' *@return {cJobject} the branch header
'*/
Public Function parse(s As String, _
    Optional jtype As eDeserializeType, _
    Optional complain As Boolean = True, _
    Optional dealWithLikelyTimestamps As Boolean = True) As cJobject
    
    Dim j As cJobject
    Set j = deSerialize(s, jtype, complain, dealWithLikelyTimestamps)
    If j.key = croot Then
        ' drop fake header
        j.sever
    End If
    Set parse = j
End Function
'/**
' *this will take a JSON string and deserialize into a cJobject branch under ME
' *@param {string} s the string to be parsed
' *@param {edeserializetype} jtype type of serialization
' *@param {boolean} complain whether to complain in the event of an error
' *@param {boolean} dealWithLikelyTimestamps whether to convert things that look like JS or unix timestamps to excel dates
' *@return {cJobject} the branch header
'*/
Public Function deSerialize(s As String, _
    Optional jtype As eDeserializeType = eDeserializeNormal, _
    Optional complain As Boolean = True, _
    Optional dealWithLikelyTimestamps As Boolean = True) As cJobject
    
    ' this will take a JSON string and deserialize into a cJobject branch starting at ME
    ' prepare string for processing
    Dim jo As cJobject

    pPointer = 1
    pJstring = noisyTrim(s)
    Set jo = New cJobject
    jo.init Nothing, croot
    pJtype = jtype
    Set jo = dsLoop(jo, complain, dealWithLikelyTimestamps)
    ' already has its own root
    If jtype = eDeserializeGoogleWire Then
        Set jo = jo.children(1)
        jo.clearParent
    End If
    jo.setValid = pValid
    Set deSerialize = jo
End Function
Public Function sever() As cJobject
    pKey = cNull
    Set pParent = Nothing
    Set sever = Me
    pFake = True
End Function
Private Function noisyTrim(s As String) As String
    Dim ns As String
    ns = Trim(s)
    If Len(ns) > 0 Then
        While (isNoisy(Right(ns, 1)))
            ns = left(ns, Len(ns) - 1)
        Wend
    End If
    noisyTrim = ns
End Function
Private Function nullItem(jOb As cJobject) As cJobject
    Set nullItem = Nothing
    
    If peek() = "," Then
    ' need an array element
    ' simulate a { 'x':'x}
        If pJtype = eDeserializeGoogleWire Then
            Set nullItem = jOb.Add.Add("v")
        Else
            Set nullItem = jOb.Add
        End If
    End If
    

End Function

Private Function dsLoop(jOb As cJobject, complain As Boolean, _
        dealWithLikelyTimestamps As Boolean) As cJobject
    
    Dim cj As cJobject, jo As cJobject, ws As String
    Set jo = jOb
    pActive = True
    pWhatNext = "{["
    While pPointer <= Len(pJstring) And pActive
        Set jo = dsProcess(jo, complain, dealWithLikelyTimestamps)
    Wend
    Set dsLoop = jOb
End Function
Private Function okWhat(what As String) As Boolean

    okWhat = (InStr(pWhatNext, nOk) <> 0 And _
                (what = "." Or what = "-" Or IsNumeric(what))) Or _
                (InStr(pWhatNext, what) <> 0)

            
End Function
Private Function peekNextToken() As String
    ' this is in case the next token is a special
    Dim k As Long
    peekNextToken = vbNullString

    ignoreNoise
    k = pPointer
    While Not (isQuote(pointedAt(k)) Or isNoisy(pointedAt(k)) Or _
        IsNumeric(pointedAt(k)) Or 0 <> InStr("[]{},.:", pointedAt(k)))
        k = k + 1
    Wend
    If (k > pPointer) Then peekNextToken = LCase(pointedAt(, k - pPointer))
    
End Function

Private Function doNextToken() As String
    Dim nextToken As String
    ' poke around to fix exceptions like null, false, true
    nextToken = peekNextToken
    If nextToken = "null" Then
        pPointer = pPointer + Len(nextToken)
        ignoreNoise
        doNextToken = pointedAt
    ElseIf nextToken = "false" Then
        doNextToken = "0"
        pPointer = pPointer + Len(nextToken)
    ElseIf nextToken = "true" Then
        doNextToken = "1"
        pPointer = pPointer + Len(nextToken)
    End If
End Function
Private Function dsProcess(jOb As cJobject, complain As Boolean, _
        dealWithLikelyTimestamps As Boolean) As cJobject
        
    Dim k As Long, jo As cJobject, s As String, what As String, jd As cJobject, v As Variant
    Dim nextToken As String, nt As String, A As Variant, av As String, jt As cJobject
    'are we done?
    Set dsProcess = jOb
    If pPointer > Len(pJstring) Then Exit Function

    Set jo = jOb
    ignoreNoise

    nextToken = doNextToken
    If nextToken <> vbNullString Then
        what = nextToken
    Else
        what = pointedAt
    End If
    
    ' is it what was expected
    
    If Not okWhat(what) Then
        badJSON pWhatNext, , complain
        Exit Function
    End If
    ' process next token
    Select Case what
    ' start of key:value pair- do nothing except set up to get the key name
        Case "{"
            pPointer = pPointer + 1
            If jo.isArrayRoot Then Set jo = jo.Add
            Set dsProcess = jo
            pWhatNext = anyQ & ",}"
            
    ' its the beginning of an array - need to kick off a new array
        Case "["
            pPointer = pPointer + 1
            If jo.isArrayRoot Then
                ' this is a double [[
                Set jo = jo.Add
            End If
            If nullItem(jo.addArray) Is Nothing Then
                pWhatNext = nOk & anyQ & "{],["
            Else
                pWhatNext = ","
            End If
            Set dsProcess = jo

            
     ' could be a key or an array value
        Case q, qs, "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "."
            v = getvItem(dealWithLikelyTimestamps, , nextToken)
            If IsEmpty(v) Then
                badJSON pWhatNext, , complain
            Else
                ' start of key/value pair
                If peek() = ":" Then
                ' add as a new key, and set up for getting the value
                    Set jt = jo
                    Set jo = jo.Add(CStr(v))
                    Set jo.backtrack = jt
                    
                    pWhatNext = ":"
                ElseIf jo.isArrayRoot Then
                ' an array value is allowed without a key
                    jo.Add , v
                    pWhatNext = ",]"
                Else
                    badJSON pWhatNext, , complain
                End If
                Set dsProcess = jo

            End If

    ' its the value of a pair
        Case ":"
            pPointer = pPointer + 1
            nt = peekNextToken
            v = getvItem(dealWithLikelyTimestamps, , doNextToken)

            If IsEmpty(v) And nt <> "null" Then
                ' about to start an array rather than get a value
                pWhatNext = "{["
            Else
                ' store the value, come back for the next
                ' boolean hack
                If (v = 1 And nt = "true") Then
                    v = True
                ElseIf (v = 0 And nt = "false") Then
                    v = False
                End If
                jo.Value = v
                Set jo = jo.backtrack
                pWhatNext = ",}"
            End If
            Set dsProcess = jo
            
        Case ","
    ' another value - same set
            pPointer = pPointer + 1
            If nullItem(jo) Is Nothing Then
                pWhatNext = nOk & anyQ & "{}],["
            Else
                pWhatNext = ","
            End If
            Set dsProcess = jo
            
        Case "}"
    ' backup a level
            pPointer = pPointer + 1
            pWhatNext = ",]}"
            Set dsProcess = jo.backtrack
               
        Case "]"
    ' backup a level
            pPointer = pPointer + 1
            pWhatNext = ",}]"
            Set dsProcess = jo.backtrack
            
        Case Else
    ' unexpected thing happened
            badJSON pWhatNext, , complain
    
    End Select

    
End Function
Private Function nOk() As String
    ' some character to say that a numeric is ok
    nOk = chr(254)
End Function
Private Function getvItem(dealWithLikelyTimestamps As Boolean, _
    Optional whichQ As String = "", _
    Optional nextToken As String = vbNullString) As Variant
    
    Dim s As String
    ' is it a string?
    getvItem = Empty
    ignoreNoise
    Select Case nextToken
        Case "1"
            getvItem = 1
        Case "0"
            getvItem = 0
        Case Else
            If isQuote(pointedAt) Then
                getvItem = getQuotedItem(whichQ)
            Else
    ' maybe its a number
                s = getNumericItem
                If Len(s) > 0 Then getvItem = toNumber(s, dealWithLikelyTimestamps)
            End If
    End Select
    
End Function
Private Function peek() As String
    Dim k As Long
    ' peek ahead to next non noisy character
    k = pPointer
    ignoreNoise
    peek = pointedAt
    pPointer = k
End Function
Private Function peekBehind() As String
    Dim k As Long
    k = pPointer - 1
    While k > 0 And isNoisy(pointedAt(k))
        k = k - 1
    Wend
    If k > 0 Then
        peekBehind = pointedAt(k)
    End If
End Function
Private Function toNumber(sIn As String, dealWithLikelyTimestamps As Boolean) As Variant
    ' convert string to numeric , either double or long
    Dim ts As String, s As String, x As Date
 ' find out the '.' separator for this locale
    ts = Mid(CStr(1.1), 2, 1)
 ' and use it so that cdbl works properly
    s = Replace(sIn, ".", ts)
    On Error GoTo overflow
   

    If InStr(1, s, ts) Then
        toNumber = CDbl(s)
    Else
        toNumber = CLng(s)
    End If
    Exit Function
    
overflow:
    'perhaps this is a javascript date
    On Error GoTo overflowAgain
    If (Len(s) = 13 And dealWithLikelyTimestamps And Not InStr(1, s, ts)) Then
        x = DateAdd("s", CDbl(left(s, 10)), DateSerial(1970, 1, 1))
    Else
        ' we're not going to convert to date, so try to overflow to a double
        x = CDbl(s)
    End If
    toNumber = x
    Resume Next
    Exit Function
    
overflowAgain:
    'this wasnt a javascript date or a convertable number
    toNumber = 0
    Resume Next
    Exit Function
    
End Function
Private Function pointedAt(Optional pos As Long = 0, Optional sLen As Long = 1) As String
    ' return what ever the currently quoted character is
    Dim k As Long
    If pos = 0 Then
        k = pPointer
    Else
        k = pos
    End If
    pointedAt = Mid(pJstring, k, sLen)
End Function

Private Function getQuotedItem(Optional whichQ As String = "") As String
    Dim s As String, k As Long, wq As String
    ignoreNoise
    s = ""

    If isQuote(pointedAt, whichQ) Then
        wq = pointedAt
    ' extract until the next matching quote
        k = pPointer + 1

        While Not isQuote(pointedAt(k), wq)
          If isUnicode(pointedAt(k, 2)) Then
            s = s & ChrW(CLng("&H" & pointedAt(k + 2, 4)))
            'S = S & StrConv(Hex2Dec(pointedAt(k + 2, 4)), vbFromUnicode)
            k = k + 6
            
          ElseIf isEscape(pointedAt(k)) Then
            Select Case LCase(pointedAt(k + 1))
                Case "t"
                    s = s & vbTab
                Case "n"
                    s = s & vbLf
                Case "r"
                    s = s & vbCr
                Case Else
                    s = s & pointedAt(k + 1)
            End Select
            k = k + 2
          Else
            s = s & pointedAt(k)
            k = k + 1
          End If
        Wend
        pPointer = k + 1
    End If
    getQuotedItem = s

End Function

Private Function getNumericItem() As String
    Dim s As String, k As Long, eAllowed As Boolean
    ignoreNoise
    s = vbNullString
    eAllowed = False
    k = pPointer
    While IsNumeric(pointedAt(k)) Or pointedAt(k) = "." Or pointedAt(k) = "-" Or (eAllowed And pointedAt(k) = "E")
        s = s & pointedAt(k)
        eAllowed = InStr(1, s, "E") < 1
        k = k + 1
    Wend
    pPointer = pPointer + Len(s)

    getNumericItem = s
    
End Function


Private Function isQuote(s As String, Optional whichQ As String = "") As Boolean
    If Len(whichQ) = 0 Then
        ' any quote
        isQuote = (s = q Or s = qs)
    Else
        isQuote = (s = whichQ)
    End If
End Function
Private Sub badJSON(pWhatNext As String, Optional Add As String = "", Optional complain As Boolean = True)
    If (complain) Then
        MsgBox Add & "got " & pointedAt & " expected --(" & pWhatNext & _
            ")-- Bad JSON at character " & CStr(pPointer) & " starting at " & _
            Mid(pJstring, pPointer)
    End If
    pValid = False
    pActive = False
    
End Sub

Private Sub ignoreNoise(Optional pos As Long = 0, Optional extraNoise As String = "")
    Dim k As Long, t As Long
    If pos = 0 Then
        t = pPointer
    Else
        t = pos
    End If
    For k = t To Len(pJstring)
        If Not isNoisy(Mid(pJstring, k, 1), extraNoise) Then Exit For
    Next k
    pPointer = k
End Sub
Private Function isNoisy(s As String, Optional extraNoise As String = "") As Boolean
    isNoisy = InStr(vbTab & " " & vbCrLf & vbCr & vbLf & extraNoise, s)
End Function
Private Function isEscape(s As String) As Boolean
    isEscape = (s = "\")
End Function
Private Function isUnicode(s As String) As Boolean
    isUnicode = LCase(s) = "\u"
End Function
Private Function q() As String
    q = chr(34)
End Function
Private Function qs() As String
    qs = chr(39)
End Function
Private Function anyQ() As String
    anyQ = q & qs
End Function

Public Function sortByValue() As cJobject
    ' sorts an objects children by their keys
    Dim jOb As cJobject, i As Long
    
    ' this will be accessed by needswap
    
    SortColl children, eSortAscending
    
    ' now need to sort out the indexes
    i = 0
    For Each jOb In children
        i = i + 1
        jOb.childIndex = i
    Next jOb
    Set sortByValue = Me
End Function
Public Function needSwap(cc As cJobject, e As eSort) As Boolean
    ' this can be used from a sorting alogirthm
    
    Dim A As Variant, B As Variant

    Select Case e
        Case eSortAscending
            needSwap = sortNorm(cValue(pbyKey)) > sortNorm(cc.cValue(pbyKey))
            
        Case eSortDescending
            needSwap = sortNorm(cValue(pbyKey)) < sortNorm(cc.cValue(pbyKey))
    
        Case Else
            needSwap = False
    End Select
End Function
Private Function sortNorm(v As Variant) As Variant

    If (VarType(v) = vbString) Then
        sortNorm = LCase(v)
    Else
        sortNorm = v
    End If
    
End Function
Private Sub Class_Initialize()
    pisArrayRoot = False
    pValid = True
    pIndex = 1
    Set pChildren = New Collection
End Sub
Public Sub tearDown()
    Dim cj As cJobject
    If Not pChildren Is Nothing Then
        For Each cj In pChildren
            cj.tearDown
        Next cj
    End If
    Set pParent = Nothing
    Set pBacktrack = Nothing
    Set pChildren = Nothing
End Sub












