Attribute VB_Name = "jsonParser"
Option Explicit
Private i As Long, token As Variant, dic As Object

Function parseJSON(JSON As String, Optional key As String = "root") As Object
    ''' Parses a JSON string into a dictionary object with each JSON path stored as the name of a dictionary
    ''' key, and the corresponding JSON value as the value.



    Dim j As Long
    Dim Matches As Object
    Dim match As Object
    Const Pattern As String = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    i = 0
    j = 1
    
    'Breaks the JSON string into individual tokens and stores them in an array
    With CreateObject("vbscript.regexp")
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = Pattern
        
        If .test(JSON) Then
            Set Matches = .Execute(JSON)
            ReDim token(1 To Matches.Count) As String
            
            For Each match In Matches
                token(j) = match.Value
                If Len(match.submatches(0)) Or match.Value = """""" Then token(j) = match.submatches(0)
                j = j + 1
            Next
        End If
    End With
    
    'Create dictionary object
    Set dic = CreateObject("Scripting.Dictionary")
    
    'Write values of token array into dictionary
    If token(1) = "{" Then
        parseObj key
    Else
        parseArr key
    End If
    
    'Return populated dictionary object
    Set parseJSON = dic

End Function

Sub parseObj(key As String)
    ''' Loops through token array and writes JSON objects to dictionary key



    Do: i = i + 1
        Select Case token(i)
            Case "]"
            
            Case "["
                parseArr key
                
            Case "{"
                If token(i + 1) = "}" Then
                    i = i + 1
                    dic.Add key, "null"
                Else
                    parseObj key
                    If i = UBound(token) Then Exit Do
                End If
                
            Case "}"
                If InStr(key, ".") Then key = Left(key, InStrRev(key, ".") - 1)
                Exit Do
                
            Case ":"
                key = key & "." & token(i - 1)
                
            Case ","
                If InStr(key, ".") Then key = Left(key, InStrRev(key, ".") - 1)
                
            Case Else: If token(i + 1) <> ":" Then dic.Add key, token(i)
            
        End Select
    Loop
    
End Sub

Sub parseArr(key As String)
    ''' Loops through JSON arrays and writes JSON array entries to dictionary key



    Dim j As Long
    
    Do: i = i + 1
        Select Case token(i)
            Case "}"
            
            Case "{"
                parseObj key & "[" & j & "]"
                
            Case "["
                parseArr key
                If i = UBound(token) Then Exit Do
                
            Case "]"
                Exit Do
                
            Case ":"
                key = key & "[" & j & "]"
                
            Case ","
                j = j + 1
                
            Case Else
                dic.Add key & "[" & j & "]", token(i)
                
        End Select
    Loop
    
End Sub
