<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Server.ScriptTimeout=9999999
Dim Conn
MF_Default_Conn
Dim RS
Set RS = Server.CreateObject(G_FS_RS)
RS.Open "Select Content from FS_NS_News",Conn,1,3
do while Not RS.Eof
	RS("Content") = Replace_Str(RS("Content") & "")
	RS.Update
	RS.MoveNext
Loop
RS.Close
Set RS = Nothing
Response.Write("Ö´ÐÐÍê³É")
Function Replace_Str(f_Str)
	Replace_Str = f_Str
	Dim Replace_Special_Str,Arr_Replace_Special_Str,i_Replace_Special_Str,Temp_Arr_Replace_Special_Str
	Replace_Special_Str = "&ldquo;:¡°|&rdquo;:¡±|&lsquo;:¡®"
	Arr_Replace_Special_Str = Split(Replace_Special_Str,"|")
	For i_Replace_Special_Str = LBound(Arr_Replace_Special_Str) to UBound(Arr_Replace_Special_Str)
		Temp_Arr_Replace_Special_Str = Split(Arr_Replace_Special_Str(i_Replace_Special_Str),":")
		if UBound(Temp_Arr_Replace_Special_Str) >= 1 then
			Replace_Str = Replace(Replace_Str,Temp_Arr_Replace_Special_Str(0),Temp_Arr_Replace_Special_Str(1))
		end if
	Next
end Function
%>