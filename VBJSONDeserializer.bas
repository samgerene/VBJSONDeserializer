Attribute VB_Name = "VBJSONDeserializer"
' VBJSONDeserializer is a VB6 adaptation of the VB-JSON project @
' http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
' BSD Licensed

Option Explicit

Private Const A_CURLY_BRACKET_OPEN As Integer = 123     ' AscW("{")
Private Const A_CURLY_BRACKET_CLOSE As Integer = 125    ' AscW("}")
Private Const A_SQUARE_BRACKET_OPEN As Integer = 91     ' AscW("[")
Private Const A_SQUARE_BRACKET_CLOSE As Integer = 93    ' AscW("]")
Private Const A_BRACKET_OPEN As Integer = 40            ' AscW("(")
Private Const A_BRACKET_CLOSE As Integer = 41           ' AscW(")")
Private Const A_COMMA As Integer = 44                   ' AscW(",")
Private Const A_DOUBLE_QUOTE As Integer = 34            ' AscW("""")
Private Const A_SINGLE_QUOTE As Integer = 39            ' AscW("'")
Private Const A_BACKSLASH As Integer = 92               ' AscW("\")
Private Const A_FORWARDSLASH As Integer = 47            ' AscW("/")
Private Const A_COLON As Integer = 58                   ' AscW(":")
Private Const A_SPACE As Integer = 32                   ' AscW(" ")
Private Const A_ASTERIX As Integer = 42                 ' AscW("*")
Private Const A_VBCR As Integer = 13                    ' AscW("vbcr")
Private Const A_VBLF As Integer = 10                    ' AscW("vblf")
Private Const A_VBTAB As Integer = 9                    ' AscW("vbTab")
Private Const A_VBCRLF As Integer = 13                  ' AscW("vbcrlf")

Private Const A_b As Integer = 98                       ' AscW("b")
Private Const A_f As Integer = 102                      ' AscW("f")
Private Const A_n As Integer = 110                      ' AscW("n")
Private Const A_r As Integer = 114                      ' AscW("r"
Private Const A_t As Integer = 116                      ' AscW("t"))
Private Const A_u As Integer = 117                      ' AscW("u")

Private m_parserrors As String
Private m_str() As Integer
Private m_length As Long

Public Function GetParserErrors() As String
   GetParserErrors = m_parserrors
End Function

Public Function parse(ByRef str As String) As Object

Dim index As Long
index = 1

GenerateStringArray str

m_parserrors = vbNullString
On Error Resume Next

Call skipChar(index)

Select Case m_str(index)
Case A_SQUARE_BRACKET_OPEN
    Set parse = parseArray(str, index)
Case A_CURLY_BRACKET_OPEN
    Set parse = parseObject(str, index)
Case Else
    m_parserrors = "Invalid JSON"
End Select

'clean array
Erase m_str

End Function

Private Sub GenerateStringArray(ByRef str As String)

Dim i As Long
Dim Length As Long
Dim s As String

m_length = Len(str)
ReDim m_str(1 To m_length)

For i = 1 To m_length
    m_str(i) = AscW(Mid$(str, i, 1))
Next i

End Sub

Private Function parseObject(ByRef str As String, ByRef index As Long) As Dictionary

Set parseObject = New Dictionary
Dim sKey As String
Dim charint As Integer

Call skipChar(index)

If m_str(index) <> A_CURLY_BRACKET_OPEN Then
   m_parserrors = m_parserrors & "Invalid Object at position " & index & " : " & Mid$(str, index) & vbCrLf
   Exit Function
End If

index = index + 1

Do
    Call skipChar(index)
    
    charint = m_str(index)
    
    If charint = A_COMMA Then
        index = index + 1
        Call skipChar(index)
    ElseIf charint = A_CURLY_BRACKET_CLOSE Then
        index = index + 1
        Exit Do
    ElseIf index > m_length Then
         m_parserrors = m_parserrors & "Missing '}': " & Right(str, 20) & vbCrLf
         Exit Do
    End If

    ' add key/value pair
    sKey = parseKey(index)
    On Error Resume Next

    parseObject.Add sKey, parseValue(str, index)
    If Err.Number <> 0 Then
        m_parserrors = m_parserrors & Err.Description & ": " & sKey & vbCrLf
        Exit Do
    End If
Loop

End Function

Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection

Dim charint As Integer

Set parseArray = New Collection

Call skipChar(index)

If Mid$(str, index, 1) <> "[" Then
    m_parserrors = m_parserrors & "Invalid Array at position " & index & " : " + Mid$(str, index, 20) & vbCrLf
    Exit Function
End If
   
index = index + 1

Do
    Call skipChar(index)
    
    charint = m_str(index)
    
    If charint = A_SQUARE_BRACKET_CLOSE Then
        index = index + 1
        Exit Do
    ElseIf charint = A_COMMA Then
        index = index + 1
        Call skipChar(index)
    ElseIf index > m_length Then
         m_parserrors = m_parserrors & "Missing ']': " & Right(str, 20) & vbCrLf
         Exit Do
    End If
    
    'add value
    On Error Resume Next
    parseArray.Add parseValue(str, index)
    If Err.Number <> 0 Then
        m_parserrors = m_parserrors & Err.Description & ": " & Mid$(str, index, 20) & vbCrLf
        Exit Do
    End If
Loop

End Function

Private Function parseValue(ByRef str As String, ByRef index As Long)

   Call skipChar(index)

    Select Case m_str(index)
    Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
        parseValue = parseString(str, index)
        Exit Function
    Case A_SQUARE_BRACKET_OPEN
        Set parseValue = parseArray(str, index)
        Exit Function
    Case A_t, A_f
        parseValue = parseBoolean(str, index)
        Exit Function
    Case A_n
        parseValue = parseNull(str, index)
        Exit Function
    Case A_CURLY_BRACKET_OPEN
        Set parseValue = parseObject(str, index)
        Exit Function
    Case Else
        parseValue = parseNumber(str, index)
        Exit Function
    End Select

End Function

Private Function parseString(ByRef str As String, ByRef index As Long) As String

   Dim quoteint As Integer
   Dim charint As Integer
   Dim Code    As String
   
   Call skipChar(index)
   
   quoteint = m_str(index)
   
   index = index + 1
   
   Do While index > 0 And index <= m_length
   
      charint = m_str(index)
      
      Select Case charint
        Case A_BACKSLASH

            index = index + 1
            charint = m_str(index)

            Select Case charint
            Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
                parseString = parseString & ChrW$(charint)
                index = index + 1
            Case A_b
                parseString = parseString & vbBack
                index = index + 1
            Case A_f
                parseString = parseString & vbFormFeed
                index = index + 1
            Case A_n
                    parseString = parseString & vbLf
                  index = index + 1
            Case A_r
                parseString = parseString & vbCr
                index = index + 1
            Case A_t
                parseString = parseString & vbTab
                  index = index + 1
            Case A_u
                index = index + 1
                Code = Mid$(str, index, 4)

                parseString = parseString & ChrW$(Val("&h" + Code))
                index = index + 4
            End Select

        Case quoteint
        
            index = index + 1
            Exit Function

         Case Else
            parseString = parseString & ChrW$(charint)
            index = index + 1
      End Select
   Loop
   
End Function

Private Function parseNumber(ByRef str As String, ByRef index As Long)

   Dim Value   As String
   Dim Char    As String

   Call skipChar(index)
   Do While index > 0 And index <= m_length
      Char = Mid$(str, index, 1)
      If InStr("+-0123456789.eE", Char) Then
         Value = Value & Char
         index = index + 1
      Else
         parseNumber = CDec(Value)
         Exit Function
      End If
   Loop
End Function

Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

   Call skipChar(index)
   
   If Mid$(str, index, 4) = "true" Then
      parseBoolean = True
      index = index + 4
   ElseIf Mid$(str, index, 5) = "false" Then
      parseBoolean = False
      index = index + 5
   Else
      m_parserrors = m_parserrors & "Invalid Boolean at position " & index & " : " & Mid$(str, index) & vbCrLf
   End If

End Function

Private Function parseNull(ByRef str As String, ByRef index As Long)

   Call skipChar(index)
   
   If Mid$(str, index, 4) = "null" Then
      parseNull = Null
      index = index + 4
   Else
      m_parserrors = m_parserrors & "Invalid null value at position " & index & " : " & Mid$(str, index) & vbCrLf
   End If

End Function

Private Function parseKey(ByRef index As Long) As String

   Dim dquote  As Boolean
   Dim squote  As Boolean
   Dim charint As Integer
   
   Call skipChar(index)
   
    Do While index > 0 And index <= m_length
    
        charint = m_str(index)
        
        Select Case charint
        Case A_DOUBLE_QUOTE
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
            
                Call skipChar(index)
                
                If m_str(index) <> A_COLON Then
                    m_parserrors = m_parserrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
                

        Case A_SINGLE_QUOTE
            squote = Not squote
            index = index + 1
            If Not squote Then
                Call skipChar(index)
                
                If m_str(index) <> A_COLON Then
                    m_parserrors = m_parserrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
                
            End If
        
        Case A_COLON
            index = index + 1
            If Not dquote And Not squote Then
                Exit Do
            Else
                parseKey = parseKey & ChrW$(charint)
            End If
        Case Else
            
            If A_VBCRLF = m_str(charint) Then
            ElseIf A_VBCR = m_str(charint) Then
            ElseIf A_VBLF = m_str(charint) Then
            ElseIf A_VBTAB = m_str(charint) Then
            ElseIf A_SPACE = m_str(charint) Then
            Else
                parseKey = parseKey & ChrW$(charint)
            End If

            index = index + 1
        End Select
    Loop

End Function

Private Sub skipChar(ByRef index As Long)

Dim bComment As Boolean
Dim bStartComment As Boolean
Dim bLongComment As Boolean

Do While index > 0 And index <= m_length
    
    Select Case m_str(index)
    Case A_VBCR, A_VBLF
        If Not bLongComment Then
            bStartComment = False
            bComment = False
        End If
    
    Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
        'do nothing
        
    Case A_FORWARDSLASH
        If Not bLongComment Then
            If bStartComment Then
                bStartComment = False
                bComment = True
            Else
                bStartComment = True
                bComment = False
                bLongComment = False
            End If
        Else
            If bStartComment Then
                bLongComment = False
                bStartComment = False
                bComment = False
            End If
        End If
    Case A_ASTERIX
        If bStartComment Then
            bStartComment = False
            bComment = True
            bLongComment = True
        Else
            bStartComment = True
        End If
    Case Else
        If Not bComment Then
            Exit Do
        End If
    End Select

    index = index + 1
Loop

End Sub
