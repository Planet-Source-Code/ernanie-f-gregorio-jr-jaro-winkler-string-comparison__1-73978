<div align="center">

## Jaro\-Winkler String Comparison


</div>

### Description

I think this is the first Jaro-Winkler Algorithm here on PSC.

Description:

The Jaroâ€“Winkler distance (Winkler, 1990) is a measure of similarity between two strings. It is a variant of the Jaro distance metric (Jaro, 1989, 1995) and mainly used in the area of record linkage (duplicate detection). The higher the Jaroâ€“Winkler distance for two strings is, the more similar the strings are. The Jaroâ€“Winkler distance metric is designed and best suited for short strings such as person names. The score is normalized such that 0 equates to no similarity and 1 is an exact match.

References:

http://en.wikipedia.org/wiki/Jaro%E2%80%93Winkler_distance

<B>lingpipe </B>

http://lingpipe-blog.com/2006/12/13/code-spelunking-jaro-winkler-string-comparison/
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ernanie F\. Gregorio Jr\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ernanie-f-gregorio-jr.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ernanie-f-gregorio-jr-jaro-winkler-string-comparison__1-73978/archive/master.zip)





### Source Code

```

Public Function JaroWrinkler(ByVal prmKeyword As String, prmCompareTo As String) As Double
  Dim iProximity As Integer ' set the number of adjacent characters to compare to
  Dim i As Integer
  Dim x As Integer
  Dim iFrom As Integer
  Dim iTo As Integer
  Dim iMatchCharacters As Integer
  Dim iTransposeCount As Integer
  Dim iJaro As Double
  prmCompareTo = UCase$(Trim$(prmCompareTo))
  prmKeyword = UCase$(Trim$(prmKeyword))
  If prmCompareTo <> prmKeyword Then ' check if the two words are the same
    If InStr(1, prmCompareTo, prmKeyword) <= 0 Then
      ' compute for the proximity of character checking
      ' allows matching characters to be up to X number of characters away.
      If Len(prmCompareTo) >= Len(prmKeyword) Then
        iProximity = (Len(prmCompareTo) / 2) - 1
      Else
        iProximity = (Len(prmKeyword) / 2) - 1
      End If
      For i = 1 To Len(prmKeyword)
        ' this is the index of the character to be compared to
        iTo = (i + iProximity) - 1
        ' get the left most side character based on the iProximity
        If i <= iProximity Then
          iFrom = 1
        Else
          iFrom = i - iProximity + 1
        End If
        ' start the letter by letter comparison
        For x = iFrom To iTo
          If Mid$(prmKeyword, i, 1) = Mid$(prmCompareTo, x, 1) Then
            If i = x Then
              iMatchCharacters = iMatchCharacters + 1
              GoTo exitfor
            End If
            iMatchCharacters = iMatchCharacters + 1
            iTransposeCount = iTransposeCount + 1
            Exit For
          End If
        Next
exitfor:
      Next
      iTransposeCount = iTransposeCount \ 2
      If iMatchCharacters > 0 Then
        x = 0
        For i = 1 To 4
          If Mid$(prmKeyword, i, 1) = Mid$(prmCompareTo, i, 1) Then
            x = x + 1
          Else
            Exit For
          End If
        Next
        iJaro = ((iMatchCharacters / Len(prmKeyword)) + _
              (iMatchCharacters / Len(prmCompareTo)) + _
              ((iMatchCharacters - iTransposeCount) / iMatchCharacters)) / 3
        If x > 0 Then
          JaroWrinkler = iJaro + 0.1 * x * (1 - iJaro)
        Else
          JaroWrinkler = iJaro
        End If
      Else
        JaroWrinkler = 0
      End If
    Else ' return 1 result if the keyword is within the search string
      JaroWrinkler = 1
    End If
  Else ' return a 1 result if the string are the same
    JaroWrinkler = 1
  End If
End Function
```

