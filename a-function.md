# a-function

```
Function HelloWorldNTimes(rng As Range) As String
    Dim n As Integer
    Dim result As String
    n = rng.Value
    result = ""
    
    Dim i As Integer
    For i = 1 To n
        result = result & "Hello World!  " & vbCrLf
    Next i

    HelloWorldNTimes = result
End Function
```
