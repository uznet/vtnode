Attribute VB_Name = "VtEnv"
Option Explicit

Public Function VtGetEnv(ByVal variableName As String, _
                         Optional ByVal defaultValue As String = "") As String
    Dim value As String

    value = Environ$(variableName)

    If Len(value) = 0 Then
        VtGetEnv = defaultValue
    Else
        VtGetEnv = value
    End If
End Function

