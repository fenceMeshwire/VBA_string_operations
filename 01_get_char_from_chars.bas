Option Explicit

' ----------------------------------------------------------------------
Sub get_single_characters()

Dim intCounter As Integer
Dim strChar, strCharacters As String
Dim varCharacters As Variant

strCharacters = "XFC32KT0XA2F39213"

varCharacters = separate_char_in_characters(strCharacters)

For intCounter = LBound(varCharacters) To UBound(varCharacters)
  strChar = varCharacters(intCounter)
  Debug.Print (strChar)
Next intCounter

End Sub

' ----------------------------------------------------------------------
Function separate_char_in_characters(strCharacters As String) As Variant

Dim intCounter As Integer
Dim strChar As String
Dim varCharacters As Variant

ReDim varCharacters(Len(strCharacters) - 1)
For intCounter = 1 To Len(strCharacters)
    varCharacters(intCounter - 1) = Mid$(strCharacters, intCounter, 1)
Next

separate_char_in_characters = varCharacters

End Function

' ----------------------------------------------------------------------
' Note: Python3 has a much simpler approach for the same result:

' #!/usr/bin/env python3
' # Python 3.9.5

' strCharacters = "XFC32KT0XA2F39213"
' for strChar in strCharacters:
'   print(strChar)
