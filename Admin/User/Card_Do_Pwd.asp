<%
Function Encrypt(theNumber) 
''加密后
On Error Resume Next 
Dim n, szEnc, t, HiN, LoN, i 
n = CDbl((theNumber + 1570) ^ 2 - 7 * (theNumber + 1570) - 450) 
If n < 0 Then szEnc = "R" Else szEnc = "J" 
n = CStr(abs(n)) 
For i = 1 To Len(n) step 2 
 t = Mid(n, i, 2) 
 If Len(t) = 1 Then 
 szEnc = szEnc & t 
 Exit For 
 End If 
 HiN = (CInt(t) And 240) / 16 
 LoN = CInt(t) And 15 
 szEnc = szEnc & Chr(Asc("M") + HiN) & Chr(Asc("C") + LoN) 
Next 
Encrypt = szEnc 
End Function 

Function Decrypt(theNumber) 
''解密后
On Error Resume Next 
Dim e, n, sign, t, HiN, LoN, NewN, i 
e = theNumber 
If Left(e, 1) = "R" Then sign = -1 Else sign = 1 
e = Mid(e, 2) 
NewN = "" 
For i = 1 To Len(e) step 2 
 t = Mid(e, i, 2) 
 If Asc(t) >= Asc("0") And Asc(t) <= Asc("9") Then 
 NewN = NewN & t 
 Exit For 
 End If 
 HiN = Mid(t, 1, 1) 
 LoN = Mid(t, 2, 1) 
 HiN = (Asc(HiN) - Asc("M")) * 16 
 LoN = Asc(LoN) - Asc("C") 
 t = CStr(HiN Or LoN) 
 If Len(t) = 1 Then t = "0" & t 
 NewN = NewN & t 
Next 
e = CDbl(NewN) * sign 
Decrypt = CLng((7 + sqr(49 - 4 * (-450 - e))) / 2 - 1570) 
End Function 

'Dim Need_Do_Pwd_Str
'Need_Do_Str = "123456"
'response.Write("加密后:"&Encrypt(Need_Do_Pwd_Str))
'response.Write("<br>解密后:"&Decrypt(  Encrypt(Need_Do_Pwd_Str)  ))
%>





