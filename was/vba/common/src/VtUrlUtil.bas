Attribute VB_Name = "VtUrlUtil"
Option Explicit

Public Function VtUrlEncode(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Integer
    Dim result As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = Asc(ch)

        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122   ' 0-9 A-Z a-z
                result = result & ch
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Hex(code)
        End Select
    Next

    VtUrlEncode = result
End Function

Public Function VtUrlDecode(ByVal s As String) As String
    Dim i As Long
    Dim result As String
    Dim hexVal As String

    i = 1
    Do While i <= Len(s)
        Select Case Mid$(s, i, 1)
            Case "+"
                result = result & " "
                i = i + 1

            Case "%"
                If i + 2 <= Len(s) Then
                    hexVal = Mid$(s, i + 1, 2)
                    result = result & Chr$(CLng("&H" & hexVal))
                    i = i + 3
                Else
                    result = result & "%"
                    i = i + 1
                End If

            Case Else
                result = result & Mid$(s, i, 1)
                i = i + 1
        End Select
    Loop

    VtUrlDecode = result
End Function
