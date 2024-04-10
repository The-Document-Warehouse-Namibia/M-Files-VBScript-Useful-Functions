Class StringBuilder_
    Private buffer

    Private Sub Class_Initialize()
        buffer = ""
    End Sub

    Public Sub Append(text)
        buffer = buffer & text
    End Sub

    Public Sub AppendLine(text)
        buffer = buffer & vbCrLf & text
    End Sub

    Public Function ToString()
        ToString = buffer
    End Function

    Public Sub Clear()
        buffer = ""
    End Sub
End Class