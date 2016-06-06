VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NpJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum NpJsonError
  npjUnknownError
  npjInvalidSyntax
  npjDoubledQuoteExpected
  npjCharacterExpected
  npjDigitExpected
  npjNumberOverflow
  npjNameExpected
  npjUnclosedString
  npjRepeatingKey
  npjInvalidEscapechar
  npjInvalidUnicode
  npjOutOfMemory
End Enum

Public Enum NpJsonType
  json_Null
  json_Object
  json_Array
  json_Boolean
  json_Number
  json_String
End Enum

Private Type NpJsonRec
  json_type As NpJsonType
  Key As String
  Value As Variant
End Type

Private Type NpJsonPage
  List() As NpJsonRec
End Type

Private mBuffer() As NpJsonPage
Private Const mPageLimit As Long = 4096
Private Const mPageSize As Long = 4096
Private Const mPageMask As Long = mPageSize - 1
Private mPageCount As Long
Private mFreeList As Long

Private mText As String
Private mErrorPos As Long
Private mErrorLineNo As Long
Private mErrorLinePos As Long
Private mErrorCode As NpJsonError
Private mParam As String
Private mLineNo As Long
Private mLinePos As Long
Private mPos As Long
Private mTabSize As Long
Private mIndentSize As Long
Private mExtendedSyntax As Boolean

Property Get ExtendedSyntax() As Boolean
  ExtendedSyntax = mExtendedSyntax
End Property

Property Let ExtendedSyntax(ByVal vNewValue As Boolean)
  mExtendedSyntax = vNewValue
End Property

Property Get ErrorLineNo() As Long
  ErrorLineNo = mErrorLineNo
End Property

Property Get ErrorLinePos() As Long
  ErrorLinePos = mErrorLinePos
End Property

Property Get ErrorCode() As NpJsonError
  ErrorCode = mErrorCode
End Property

Property Get TabSize() As Long
  If mTabSize <= 0 Then
    TabSize = 8
  Else
    TabSize = mTabSize
  End If
End Property

Property Let TabSize(ByVal vNewValue As Long)
  mTabSize = vNewValue
  If mTabSize < 0 Then mTabSize = 0
End Property

Property Get IndentSize() As Long
  If mTabSize <= 0 Then
    IndentSize = 2
  Else
    IndentSize = mIndentSize
  End If
End Property

Property Let IndentSize(ByVal vNewValue As Long)
  mIndentSize = vNewValue
  If mIndentSize < 0 Then mIndentSize = 0
End Property

Property Get ErrorMessage() As String
  ErrorMessage = mErrorMessage(mErrorCode)
End Property

Private Function mErrorMessage(ByVal vErrorCode As NpJsonError) As String
  Select Case vErrorCode
    Case npjUnknownError
      mErrorMessage = "Unknwon error"
    Case npjInvalidSyntax
      mErrorMessage = "Invalid syntax"
    Case npjDoubledQuoteExpected
      mErrorMessage = "Doubled quote expected"
    Case npjCharacterExpected
      mErrorMessage = "character or string """ & mParam & """ expected"
    Case npjDigitExpected
      mErrorMessage = "Digit expected"
    Case npjNumberOverflow
      mErrorMessage = "Numeric overflow " & mParam
    Case npjNameExpected
      mErrorMessage = "Name expected"
    Case npjUnclosedString
      mErrorMessage = "Unclosed string (doubled quote missing)"
    Case npjRepeatingKey
      mErrorMessage = "Reapiting key " & mParam
    Case npjInvalidEscapechar
      mErrorMessage = "Invalid escape char " & mParam
    Case npjInvalidUnicode
      mErrorMessage = "Invalid unicode character " & mParam
    Case npjOutOfMemory
      mErrorMessage = "Out of JSON memory"
  End Select
End Function

Function NewJson(ByVal json_type As NpJsonType, Optional ByVal Key As String, Optional ByVal Value As Variant) As Long
  Dim i As Long
  Dim PageNo As Long
  If mPageCount = 0 Then
    mFreeList = -1
    ReDim mBuffer(mPageLimit - 1)
  End If
  If mFreeList = -1 Then
    If mPageCount = mPageLimit Then
      NewJson = -1
      SetError npjOutOfMemory
    End If
    ReDim mBuffer(mPageCount).List(mPageSize - 1)
    mFreeList = mPageCount * mPageSize
    For i = 0 To mPageSize - 2
      mBuffer(mPageCount).List(i).Value = mFreeList + i + 1
    Next
    mBuffer(mPageCount).List(mPageSize - 1).Value = -1
    mPageCount = mPageCount + 1
  End If
  PageNo = mFreeList \ mPageSize
  i = mFreeList And mPageMask
  With mBuffer(PageNo).List(i)
    NewJson = mFreeList
    mFreeList = .Value
    .json_type = json_type
    .Key = Key
    If json_type = json_Object Or json_type = json_Array Then
      Set .Value = New Collection
    Else
      .Value = Value
    End If
  End With
End Function

Sub JsonAddToArray(ByVal owner_json As Long, ByVal child_json As Long)
  mBuffer(owner_json \ mPageSize).List(owner_json And mPageMask).Value.Add child_json
End Sub

Sub JsonAddToObject(ByVal owner_json As Long, ByVal Key As String, ByVal child_json As Long)
  mBuffer(owner_json \ mPageSize).List(owner_json And mPageMask).Value.Add child_json, StrKey(Key)
  mBuffer(child_json \ mPageSize).List(child_json And mPageMask).Key = Key
End Sub

Property Get JsonKey(ByVal json_index As Long) As String
  JsonKey = mBuffer(json_index \ mPageSize).List(json_index And mPageMask).Key
End Property

Property Get JsonValue(ByVal json_index As Long) As Variant
  With mBuffer(json_index \ mPageSize).List(json_index And mPageMask)
    If IsObject(.Value) Then
      Set JsonValue = .Value
    Else
      JsonValue = .Value
    End If
  End With
End Property

Property Get JsonType(ByVal json_index As Long) As NpJsonType
  With mBuffer(json_index \ mPageSize).List(json_index And mPageMask)
    JsonType = .json_type
  End With
End Property

Sub RemoveJson(ByVal json_index As Long)
  Dim f
  If json_index < 0 Then Exit Sub
  With mBuffer(json_index \ mPageSize).List(json_index And mPageMask)
    Select Case .json_type
      Case json_Array, json_Object
      For Each f In .Value
        RemoveJson f
      Next
    End Select
    .Value = mFreeList
  End With
  mFreeList = json_index
End Sub

Function Parse(ByVal s As String, ByRef Index As Long) As Boolean
  Index = -1
  mPos = 1
  mText = s
  ParseClear
  If mParseValue(Index) Then
    SkipChars ""
    If mPos > Len(mText) Then
      Parse = True
    Else
      SetError npjInvalidSyntax
    End If
  End If
End Function

Function ParseScript(ByVal s As String, ByRef Index As Long) As Boolean
  Dim sub_index As Long
  Dim VarName As String
  Index = -1
  mPos = 1
  mText = s
  ParseClear
  Do
    If Not SkipWord("var") Then Exit Do
    SkipChars ""
    If Not mParseName(VarName) Then Exit Function
    If KeyInObject(Index, VarName) Then Exit Function
    If Not SkipChars("=") Then
      SetError npjCharacterExpected, "="
      Exit Function
    End If
    If Not mParseValue(sub_index) Then
      RemoveJson sub_index
      Exit Function
    End If
    If Index = -1 Then
      Index = NewJson(json_Object)
    End If
    JsonAddToObject Index, VarName, sub_index
    SkipChars ";"
  Loop
  SkipChars ""
  If mPos > Len(mText) Then
    ParseScript = True
  Else
    SetError npjCharacterExpected, "var"
  End If
End Function

Private Function mParseValue(ByRef Index As Long) As Boolean
  Dim s As String
  Index = -1
  
  If SkipChars("""") Then
    If Not mParseString(s) Then Exit Function
    Index = NewJson(json_String, , s)
  ElseIf SkipWord("true") Then
    Index = NewJson(json_Boolean, , True)
  ElseIf SkipWord("false") Then
    Index = NewJson(json_Boolean, , False)
  ElseIf SkipWord("null") Then
    Index = NewJson(json_Null)
  ElseIf SkipChars("{") Then
    If Not mParseObject(Index) Then Exit Function
  ElseIf SkipChars("[") Then
    If Not mParseArray(Index) Then Exit Function
  Else
    If SkipNumber(s) Then
      Index = NewNumberJson(s)
      If Index = -1 Then Exit Function
    Else
      SetError npjInvalidSyntax
      Exit Function
    End If
  End If
  mParseValue = True
End Function

Private Function NewNumberJson(ByVal s As String) As Long
  Dim v As Double
  On Error GoTo Final
  v = Val(s)
  NewNumberJson = NewJson(json_Number, , v)
  Exit Function
Final:
  SetError npjNumberOverflow, s
  NewNumberJson = -1
End Function

Private Function mParseObject(ByRef Index As Long) As Boolean
  Dim IsFirst As Boolean
  Index = NewJson(json_Object)
  IsFirst = True
  Do While Not SkipChars("}")
    If Not IsFirst Then
      If Not SkipChars(",") Then
        SetError npjCharacterExpected, ","
        Exit Function
      End If
    End If
    If Not ParseObjectPair(Index) Then Exit Function
    IsFirst = False
  Loop
  SkipChars ";"
  mParseObject = True
End Function

Private Function mParseArray(ByRef Index As Long) As Boolean
  Dim sub_index As Long
  Dim IsFirst As Boolean
  Index = NewJson(json_Array)
  IsFirst = True
  Do While Not SkipChars("]")
    If Not IsFirst Then
      If Not SkipChars(",") Then
        SetError npjCharacterExpected, ","
        Exit Function
      End If
    End If
    If Not mParseValue(sub_index) Then
      RemoveJson sub_index
      Exit Function
    End If
    JsonAddToArray Index, sub_index
    IsFirst = False
  Loop
  mParseArray = True
End Function

Private Function SkipChars(ByVal Chr As String) As Boolean
  Dim s As String
  Do While mPos <= Len(mText)
    s = CurrentChar ' Mid(mText, mPos, 1)
    If InStr(1, vbCrLf & vbTab & " ", s) > 0 Then
      mPos = mPos + 1
      If s = vbCr Then
        mLineNo = mLineNo + 1
        mLinePos = 1
        If CurrentChar = vbLf Then
          mPos = mPos + 1
        End If
      ElseIf s = vbLf Then
        mLinePos = 1
        mLineNo = mLineNo + 1
      ElseIf s = vbTab Then
        mLinePos = 1 + Int((mLinePos + TabSize - 1) / TabSize) * TabSize
      Else
        mLinePos = mLinePos + 1
      End If
    Else
      If InStr(1, Chr, s) > 0 Then
        mPos = mPos + 1
        mLinePos = mLinePos + 1
        SkipChars = True
      End If
      Exit Function
    End If
  Loop
End Function

Private Function mParseString(ByRef s As String) As Boolean
  Dim i As Long
  Dim c As String
  i = mPos
  Do While mPos <= Len(mText)
    SkipChars ""
    c = CurrentChar
    If c = "\" Then
      mPos = mPos + 1
      mLinePos = mLinePos + 1
    ElseIf c = """" Then
      If Not InvEscapeStr(Mid(mText, i, mPos - i), s) Then Exit Function
      mPos = mPos + 1
      mLinePos = mLinePos + 1
      mParseString = True
      Exit Function
    End If
    mPos = mPos + 1
    mLinePos = mLinePos + 1
  Loop
  SetError npjUnclosedString
End Function

Private Function mParseName(ByRef s As String) As Boolean
  Dim i As Long
  Dim c As String
  SkipChars ""
  i = mPos
  Do While mPos <= Len(mText)
    c = CurrentChar
    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM1234567890_qwertyuiopasdfghjklzxcvbnm", c) = 0 Then Exit Do
    mPos = mPos + 1
    mLinePos = mLinePos + 1
  Loop
  If mPos = i Then
    s = ""
    SetError npjNameExpected
  Else
    s = Mid(mText, i, mPos - i)
    mParseName = True
  End If
End Function

Private Function SkipWord(ByVal s As String) As Boolean
  Dim c As String
  SkipChars ""
  If mPos + Len(s) > Len(mText) Then Exit Function
  If Mid(mText, mPos, Len(s)) <> s Then Exit Function
  c = Mid(mText, mPos + Len(s), 1)
  If LCase(c) <> UCase(c) Then Exit Function
  mPos = mPos + Len(s)
  mLinePos = mLinePos + Len(s)
  SkipWord = True
End Function

Private Function SkipNumber(ByRef s As String) As Boolean
  Dim i As Long
  Dim c As String
  
  SkipChars ""
  i = mPos
  c = CurrentChar
  If c = "-" Then
    mPos = mPos + 1
    c = CurrentChar
    mLinePos = mLinePos + 1
  End If
  Select Case c
    Case "0"
      mPos = mPos + 1
      mLinePos = mLinePos + 1
    Case "1" To "9"
      Do
        mPos = mPos + 1
        mLinePos = mLinePos + 1
        If InStr(1, "0123456789", CurrentChar) = 0 Then Exit Do
      Loop
    Case "."
      If mExtendedSyntax Then
        If InStr(1, "0123456789", Mid(mText, mPos + 1, 1)) = 0 Then Exit Function
      Else
        SetError npjDoubledQuoteExpected
        Exit Function
      End If
    Case Else
      Exit Function
  End Select
  If CurrentChar = "." Then
    Do
      mPos = mPos + 1
      mLinePos = mLinePos + 1
      If InStr(1, "0123456789", CurrentChar) = 0 Then Exit Do
    Loop
  End If
  If UCase(CurrentChar) = "E" Then
    mPos = mPos + 1
    mLinePos = mLinePos + 1
    Select Case CurrentChar
      Case "+", "-"
        mPos = mPos + 1
        mLinePos = mLinePos + 1
    End Select
    Select Case CurrentChar
      Case "0" To "9"
      Case Else
        SetError npjDigitExpected
        Exit Function
    End Select
    Do
      mPos = mPos + 1
      mLinePos = mLinePos + 1
      If InStr(1, "0123456789", CurrentChar) = 0 Then Exit Do
    Loop
  End If
  s = Mid(mText, i, mPos - i)
  SkipNumber = True
End Function

Private Function KeyInObject(ByVal Index As Long, ByVal Key As String) As Boolean
  Dim i
  On Error GoTo Final
  i = mBuffer(Index \ mPageSize).List(Index And mPageMask).Value(StrKey(Key))
  SetError npjRepeatingKey, Key
  KeyInObject = True
Final:
End Function

Private Function ParseObjectPair(ByRef Index As Long) As Boolean
  Dim s As String
  Dim sub_index As Long
  If Not SkipChars("""") Then
    If mExtendedSyntax Then
      If Not mParseName(s) Then Exit Function
    Else
      SetError npjDoubledQuoteExpected
      Exit Function
    End If
  Else
    If Not mParseString(s) Then Exit Function
  End If
  If KeyInObject(Index, s) Then Exit Function
  If Not SkipChars(":") Then
    SetError npjCharacterExpected, ":"
    Exit Function
  End If
  If Not mParseValue(sub_index) Then
    RemoveJson sub_index
    Exit Function
  End If
  JsonAddToObject Index, s, sub_index
  ParseObjectPair = True
End Function

Property Get CurrentChar() As String
  CurrentChar = Mid(mText, mPos, 1)
End Property

Private Property Get InvEscapeLen(ByVal s As String) As Long
  Dim i As Long
  Dim c As Long
  i = 1
  Do
    i = InStr(i, s, "\")
    If i = 0 Then Exit Do
    i = i + 1
    Select Case Mid(s, i, 1)
      Case """", "\", "/", "b", "f", "n", "r", "t"
        i = i + 1
        c = c + 1
      Case "u"
        i = i + 4
        c = c + 5
    End Select
  Loop
  InvEscapeLen = Len(s) - c
End Property

Private Property Get EscapeLen(ByVal s As String) As Long
  Const EscapeChars As String = vbCrLf & vbTab & """\/" & vbBack & vbFormFeed
  Dim i As Long
  Dim c As Long
  For i = 1 To Len(s)
    If InStr(1, EscapeChars, Mid(s, i, 1)) > 0 Then
      c = c + 1
    End If
  Next
  EscapeLen = Len(s) + c
End Property

Private Property Get EscapeStr(ByVal s As String) As String
  Dim o As String
  Dim i As Long
  Dim j As Long
  Dim c As String
  o = Space(EscapeLen(s))
  
  j = 1
  For i = 1 To Len(s)
    c = Mid(s, i, 1)
    Select Case c
      Case vbCr
        Mid(o, j, 2) = "\r"
        j = j + 2
      Case vbLf
        Mid(o, j, 2) = "\n"
        j = j + 2
      Case vbTab
        Mid(o, j, 2) = "\t"
        j = j + 2
      Case """"
        Mid(o, j, 2) = "\"""
        j = j + 2
      Case "\"
        Mid(o, j, 2) = "\\"
        j = j + 2
      Case "/"
        Mid(o, j, 2) = "\/"
        j = j + 2
      Case vbBack
        Mid(o, j, 2) = "\b"
        j = j + 2
      Case vbFormFeed
        Mid(o, j, 2) = "\f"
        j = j + 2
      Case Else
        'If mExtendedSyntax Then
          Mid(o, j, 1) = c
          j = j + 1
        'Else
        '  SetError npjInvalidEscapechar, c
        'End If
    End Select
  Next
  EscapeStr = o
End Property

Function SetCharUnicode(ByRef o As String, ByVal j As Long, ByVal s As String) As Boolean
  On Error GoTo Final
  Mid(o, j, 1) = ChrW(CLng("&H" & s))
  SetCharUnicode = True
  Exit Function
Final:
  SetError npjInvalidUnicode, "\u" & s
End Function

Private Property Get InvEscapeStr(ByVal s As String, ByRef o As String) As Boolean
  Dim i As Long
  Dim j As Long
  Dim c As String
  Dim PrevIsSlash As Boolean
  o = Space(InvEscapeLen(s))
  
  j = 1
  i = 1
  Do While i <= Len(s)
    c = Mid(s, i, 2)
    Select Case c
      Case "\r"
        Mid(o, j, 1) = vbCr
        i = i + 2
      Case "\n"
        Mid(o, j, 1) = vbLf
        i = i + 2
      Case "\t"
        Mid(o, j, 1) = vbTab
        i = i + 2
      Case "\"""
        Mid(o, j, 1) = """"
        i = i + 2
      Case "\\"
        Mid(o, j, 1) = "\"
        i = i + 2
      Case "\/"
        Mid(o, j, 1) = "/"
        i = i + 2
      Case "\b"
        Mid(o, j, 1) = vbBack
        i = i + 2
      Case "\f"
        Mid(o, j, 1) = vbFormFeed
        i = i + 2
      Case "\u"
        If Not SetCharUnicode(o, j, Mid(s, i + 2, 4)) Then Exit Function
        i = i + 6
      Case Else
        If mExtendedSyntax Or Left(c, 1) <> "\" Then
          Mid(o, j, 1) = Mid(c, 1, 1)
          i = i + 1
        Else
          mPos = mPos - 2
          mLinePos = mLinePos - 2
          SetError npjInvalidEscapechar, c
          Exit Property
        End If
    End Select
    j = j + 1
  Loop
  InvEscapeStr = True
End Property

Function JsonText(ByVal Index As Long) As String
  Dim List As Collection
  Dim l As Long
  Dim f
  Dim s As String
  Set List = New Collection
  
  mJsonText List, Index, 0
  
  For Each f In List
    l = l + Len(f)
  Next
  
  s = Space(l)
  l = 1
  For Each f In List
    Mid(s, l, Len(f)) = f
    l = l + Len(f)
  Next
  
  JsonText = s
End Function

Private Function ValStr(ByVal Value As Double) As String
  Dim s As String
  s = Trim(Str(Value))
  If Left(s, 1) = "." Then
    s = "0" & s
  ElseIf Left(s, 2) = "-." Then
    s = "-0" & Mid(s, 2)
  End If
  ValStr = s
End Function

Private Property Get ExistsObjectsOrArrays(ByVal Index As Long) As Boolean
  Dim f
  For Each f In mBuffer(Index \ mPageSize).List(Index And mPageMask).Value
    Select Case JsonType(f)
      Case json_Array, json_Object
        ExistsObjectsOrArrays = True
        Exit Function
    End Select
  Next
End Property

Private Sub NewLine(ByVal List As Collection, ByVal Indent As Long)
  List.Add vbCrLf & Space(Indent)
End Sub

Private Sub mJsonText(ByVal List As Collection, ByVal Index As Long, ByVal Indent As Long)
  Dim PageNo As Long
  Dim PageOfs As Long
  Dim f
  Dim IsFirst As Boolean
  Dim HasObject As Boolean
  
  PageNo = Index \ mPageSize
  PageOfs = Index And mPageMask
  
  Select Case mBuffer(PageNo).List(PageOfs).json_type
    Case json_Array
      List.Add "["
      IsFirst = True
      HasObject = ExistsObjectsOrArrays(Index)
      For Each f In mBuffer(PageNo).List(PageOfs).Value
        If IsFirst Then
          IsFirst = False
        Else
          List.Add ","
        End If
        If HasObject Then NewLine List, Indent + IndentSize
        mJsonText List, f, Indent + IndentSize
      Next
      List.Add "]"
    Case json_Boolean
      List.Add IIf(mBuffer(PageNo).List(PageOfs).Value, "true", "false")
    Case json_Null
      List.Add "null"
    Case json_Number
      List.Add ValStr(mBuffer(PageNo).List(PageOfs).Value)
    Case json_Object
      List.Add "{"
      IsFirst = True
      HasObject = ExistsObjectsOrArrays(Index)
      For Each f In mBuffer(PageNo).List(PageOfs).Value
        If IsFirst Then
          IsFirst = False
        Else
          List.Add ","
        End If
        If HasObject Then NewLine List, Indent + IndentSize
        List.Add """" & mBuffer(f \ mPageSize).List(f And mPageMask).Key & """:"
        mJsonText List, f, Indent + IndentSize
      Next
      List.Add "}"
    Case json_String
      List.Add """"
      List.Add EscapeStr(mBuffer(PageNo).List(PageOfs).Value)
      List.Add """"
  End Select
End Sub

Private Sub SetError(ByVal ErrCode As NpJsonError, Optional ByVal Param As String = "")
  mErrorCode = ErrCode
  mParam = Param
  mErrorLineNo = mLineNo
  mErrorLinePos = mLinePos
  mErrorPos = mPos
End Sub

Private Sub ParseClear()
  mErrorCode = npjUnknownError
  mParam = ""
  mErrorLineNo = 0
  mErrorLinePos = 0
  mErrorPos = 0
  mLineNo = 1
  mLinePos = 1
End Sub

Private Function StrKey(ByVal s As String) As String
  Dim u As String
  Dim l As String
  Dim Key As String
  Dim i As Long
  Dim j As Long
  
  u = UCase(s)
  
  Key = Space(2 * Len(s))
  j = 1
  For i = 1 To Len(s)
    Mid(Key, j, 1) = Mid(u, i, 1)
    If Mid(u, i, 1) <> Mid(s, i, 1) Then
      Mid(Key, j + 1, 1) = "_"
    End If
    j = j + 2
  Next
  StrKey = Key
End Function

