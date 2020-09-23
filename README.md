<div align="center">

## ArrPtr/StrArrPtr function for vb5 and vb6 \(FIXED\!\)


</div>

### Description

This function returns a pointer to the SAFEARRAY header of any Visual Basic array, including a Visual Basic string array. Substitutes both ArrPtr and StrArrPtr. This function will work with vb5 or vb6 without modification. Normally you need to declare a VarPtr alias into msvbvm50.dll or msvbvm60.dll depending on the vb version, but this function will work with vb5 or vb6.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-arrptr-strarrptr-function-for-vb5-and-vb6-fixed__1-69970/archive/master.zip)





### Source Code


<br />
<font face="Times"><h1>ArrayPtr function for vb5 and vb6</h1></font>
<pre><font face="Courier New" color="#0000a0">Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (pDest As Any, pSrc As Any, ByVal lByteLen As Long)</font>
<font color="#008000">' + ArrayPtr ++++++++++++++++++++++++Rd+
' This function returns a pointer to the
' SAFEARRAY header of any Visual Basic
' array, including a Visual Basic string
' array.
' Substitutes both ArrPtr and StrArrPtr.
' This function will work with vb5 or
' vb6 without modification.</font>
<font color="#0000a0">Public Function ArrayPtr(Arr) As Long</font>
<font color="#008000">  ' Thanks to Francesco Balena and Monte Hansen</font>
<font color="#0000a0">  Dim iDataType As Integer
  On Error GoTo UnInit
  CopyMemory iDataType, Arr, 2&</font>            <font color="#008000">' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit</font>
<font color="#0000a0">  If (iDataType And vbArray) = vbArray Then</font>      <font color="#008000">' if a valid array was passed</font>
<font color="#0000a0">    CopyMemory ArrayPtr, ByVal VarPtr(Arr) + 8&, 4&</font> <font color="#008000">' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.</font>
<font color="#0000a0">  End If
UnInit:
End Function</font>
<font color="#008000">' ++++++++++++++++++++++++++++++++++++++</font></pre>

