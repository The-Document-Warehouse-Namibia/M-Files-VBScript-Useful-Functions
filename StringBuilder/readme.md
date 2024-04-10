# StringBuilder Class

A simple implementation of a StringBuilder class in VBScript.

## Usage

1. Create an instance of the StringBuilder class.
2. Append strings or lines using the `Append` or `AppendLine` methods respectively.
3. Get the final string using the `ToString` method.
4. Optionally, clear the content using the `Clear` method.

## Example

```vbscript
' Create a new instance of StringBuilder
Dim sb
Set sb = New StringBuilder_

' Append some text
sb.Append "Hello"
sb.AppendLine " World!"

' Get the final string
Dim result
result = sb.ToString()

' Output the result
WScript.Echo result

' Clear the content
sb.Clear()
