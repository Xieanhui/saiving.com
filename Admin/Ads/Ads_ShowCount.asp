<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,AddrConn
MF_Default_Conn
MF_IP_Conn
MF_Session_TF 
'Ȩ���ж�
'Call MF_Check_Pop_TF("NS_Class_000001")
Dim str_Show_Sql,o_Show_Rs
set o_Show_Rs = Conn.execute("select ID,AdIPAdress from FS_AD_Source where Address is null or Address=''")
do while not o_Show_Rs.eof 
	Conn.execute("Update FS_AD_Source set Address='"&Get_Address(o_Show_Rs("AdIPAdress"))&"' where ID="&o_Show_Rs("ID"))
	o_Show_Rs.movenext
loop
set o_Show_Rs = nothing
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo

int_RPP=5 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ͳ��___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	Dim int_ID,str_AdName,strShowErr,TmpStr
	int_ID=Request.QueryString("ID")
	If int_ID="" or IsNull(int_ID) Then
		strShowErr = "<li>��������!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		If Isnumeric(int_ID)=False Then
			strShowErr = "<li>��������!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			int_ID=Clng(int_ID)
		End If
	End If
	str_AdName=Request.QueryString("AdName")
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td>������[<font color="red"><%=str_AdName%></font>]����Դͳ��  |  <a href="javascript:history.back();">������һ��</a></td>
  </tr>
</table>
<%
	str_Show_Sql="Select AdID,Address,Count(Address) as Addrnum from FS_AD_Source group by AdID,Address having AdID="&CintStr(int_ID)&" order by Count(Address) desc"
	Set o_Show_Rs= CreateObject(G_FS_RS)
	o_Show_Rs.Open str_Show_Sql,Conn,1,1
	If Not o_Show_Rs.Eof Then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="33%" align="center" class="xingmu">����</td>
    <td width="10%" align="center" class="xingmu">�������</td>
    <td align="center" class="xingmu">ռ�ܵ�������ٷֱ�</td>
  </tr>
<%
		o_Show_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = CintStr(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>o_Show_Rs.PageCount Then cPageNo=o_Show_Rs.PageCount 
		o_Show_Rs.AbsolutePage=cPageNo
	
		For int_Start=1 TO int_RPP  
		Tmpstr = formatnumber(o_Show_Rs("Addrnum")/Conn.execute("Select Count(*) from FS_AD_Source where AdID="&CintStr(int_ID))(0)*100,3)&"%"
%>
  <tr class="hback">
    <td height="20" align="center" class="hback"><%=o_Show_Rs("Address")%></td>
    <td align="center" class="hback"><%=o_Show_Rs("Addrnum")%></td>
    <td class="hback"><img src="../Images/bar2.gif" height="12" width="<%=Tmpstr%>"><%=Tmpstr%></td>
  </tr>
<%
			o_Show_Rs.MoveNext
			If o_Show_Rs.Eof or o_Show_Rs.Bof Then Exit For
		Next
		Response.Write "<tr><td class=""hback"" colspan=""4"" align=""left"">"&fPageCount(o_Show_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td></tr>"

%>
</table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>��ǰû��������Ϣ!</td></tr></table>"
	End If
%>
</body>
</html>
<%
Function Get_Address(LoginIP)'ת��IP��ַ���бȽ�
	dim str1,str2,str3,str4 ,num,db_address,Get_Address_Rs
	num = "" : db_address = ""	
	str1=left(LoginIP,instr(LoginIP,".")-1)
	LoginIP=mid(LoginIP,instr(LoginIP,".")+1)
	str2=left(LoginIP,instr(LoginIP,".")-1)
	LoginIP=mid(LoginIP,instr(LoginIP,".")+1)
	str3=left(LoginIP,instr(LoginIP,".")-1)
	str4=mid(LoginIP,instr(LoginIP,".")+1)
	num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
	Set Get_Address_Rs=AddrConn.execute("select Country from Address where StarIP <="&num&" and EndIP >="&num&"")
	'����IP�õ���ַ
	if not Get_Address_Rs.eof then
		db_address = Get_Address_Rs(0)
	end if	
	Get_Address_Rs.close
	set Get_Address_Rs = Nothing
	Get_Address = db_address
End Function
AddrConn.close
set AddrConn = Nothing
Set Conn=nothing
%>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





