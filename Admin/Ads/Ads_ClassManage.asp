<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS003") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo,Temp_C_ID

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
<form name="AdsClass" action="" method="post">
<%
	Dim str_Ad_OP,lng_ClassID,strShowErr,CheckAllID,i,TempID
	str_Ad_OP=Request.QueryString("Ad_OP")
	lng_ClassID=Clng(Request.QueryString("ID"))
	CheckAllID=Request.Form("Checkallbox")
	Select Case str_Ad_OP
		Case "Lock"
			Conn.execute("Update FS_AD_Class Set Lock=1 where AdClassID="&lng_ClassID&"")
			Conn.execute("Update FS_AD_Info Set AdLock=1 where AdClassID="&lng_ClassID&"")
			strShowErr = "<li>�����ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "Unlock"
			Conn.execute("Update FS_AD_Class Set Lock=0 where AdClassID="&lng_ClassID&"")
			Conn.execute("Update FS_AD_Info Set AdLock=0 where AdClassID="&lng_ClassID&"")
			strShowErr = "<li>��������ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "DelOne"
			Conn.execute("delete  from FS_AD_TxtInfo where AdID In(select AdId from FS_AD_TxtInfo where AdId In(select AdID from FS_AD_Info where AdClassID="&lng_ClassID&"))")
			Conn.execute("delete  from FS_AD_Class where AdClassID="&lng_ClassID&"")
			Conn.execute("delete  from FS_AD_Info where AdClassID="&lng_ClassID&"")
			strShowErr = "<li>ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "P_Lock"
			CheckAllIDFLag("����")
			Conn.execute("update FS_AD_Info set AdLock=1 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			Conn.execute("update FS_AD_Class set Lock=1 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			strShowErr = "<li>���������ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "P_UnLock"
			CheckAllIDFLag("�������")
			Conn.execute("update FS_AD_Info set AdLock=0 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			Conn.execute("update FS_AD_Class set Lock=0 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			strShowErr = "<li>������������ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "P_Del"
			CheckAllIDFLag("ɾ��")
			Temp_C_ID=CheckAllID
			CheckAllID=split(CheckAllID,",")
			For i=0 to Ubound(CheckAllID)
				TempID=TempID&CheckAllID(i)&","
			Next
			Conn.execute("delete  from FS_AD_TxtInfo where AdID In(select AdId from FS_AD_TxtInfo where AdId In(select AdID from FS_AD_Info where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")))")
			Conn.execute("delete  from FS_AD_Class where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")")
			Conn.execute("delete  from FS_AD_Info where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")")
			strShowErr = "<li>����ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "AdClassDelAll"
			Conn.execute("delete  from FS_AD_Info where AdClassID in (select AdClassID from FS_AD_Class)")
			Conn.execute("delete  from FS_AD_Class")
			strShowErr = "<li>ȫ��ɾ���ɹ�!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
	End Select
	Sub CheckAllIDFLag(Showstr)
		If CheckAllID="" or IsNull(CheckAllID) Then
			strShowErr = "<li>��ѡ��Ҫ"&Showstr&"����Ŀ!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End Sub
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">���������</td>
  </tr>
   <tr class="hback"> 
    <td class="hback"><a href="Ads_AddClass.asp?OpCType=Add">��ӷ���</a> | <a href="javascript:P_Lock();">��������</a> | <a href="javascript:P_UnLock();">��������</a> | <a href="javascript:P_Del();">����ɾ��</a> | <a href="javascript:AdClassDelAll();">ɾ��ȫ��</a> | <a href="javascript:history.back();">������һ��</a></td>
  </tr>
</table>
<%
	Dim str_Adclass_Sql,o_Adclass_Rs,str_ClassType
	str_Adclass_Sql="Select AdClassID,AdClassName,AddDate,Lock from FS_AD_Class"
	Set o_Adclass_Rs= CreateObject(G_FS_RS)
	o_Adclass_Rs.Open str_Adclass_Sql,Conn,1,1
	If Not o_Adclass_Rs.Eof Then 
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="36%" align="center" class="xingmu">��Ŀ����</td>
    <td width="15%" align="center" class="xingmu">���ʱ��</td>
    <td width="15%" align="center" class="xingmu">��ǰ��Ŀ�����</td>
    <td width="5%" align="center" class="xingmu">״̬</td>
    <td width="27%" align="center" class="xingmu">����</td>
  </tr>
<%
		o_Adclass_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo>o_Adclass_Rs.PageCount Then cPageNo=o_Adclass_Rs.PageCount 
		If cPageNo<=0 Then cPageNo=1		
		o_Adclass_Rs.AbsolutePage=cPageNo
	
		For int_Start=1 TO int_RPP  
%>
  <tr class="hback">
    <td class="hback" align="Center"><%=o_Adclass_Rs("AdClassName")%></td>
    <td class="hback" align="Center"><%=o_Adclass_Rs("AddDate")%></td>
    <td class="hback" align="Center"><%=Conn.execute("Select Count(*) from FS_AD_Info Where AdClassID="&o_Adclass_Rs("AdClassID")&"")(0)%></td>
<%
			str_ClassType=Cint(o_Adclass_Rs("Lock"))
			Select Case str_ClassType
				Case 0
					str_ClassType="����"
				Case 1
					str_ClassType="<font color=""red"">����</font>"
			End Select
%>
    <td class="hback" align="Center"><%=str_ClassType%></td>
    <td class="hback" align="Center"><a href="javascript:Lock('<%=o_Adclass_Rs("AdClassID")%>');">����</a> | <a href="javascript:Unlock('<%=o_Adclass_Rs("AdClassID")%>');">����</a> | <a href="javascript:DelOne('<%=o_Adclass_Rs("AdClassID")%>');">ɾ��</a> | <a href="javascript:Update('<%=o_Adclass_Rs("AdClassID")%>');">�޸�</a> | <input type="checkbox" name="Checkallbox" value="<%=o_Adclass_Rs("AdClassID")%>"></td>
  </tr>
<%
			o_Adclass_Rs.MoveNext
		If o_Adclass_Rs.Eof or o_Adclass_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""5"" align=""left"">"&fPageCount(o_Adclass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ȫѡ<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');"" value=""0""></td></tr>"
%>
</table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>��ǰû�й����Ŀ!</td></tr></table>"
	End If
%>
</form>
</body>
</html>
<script language="javascript">
function P_Lock()
{
	if(confirm('�˲���������ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ��������'))
	{
		document.AdsClass.action="?Ad_OP=P_Lock";
		document.AdsClass.submit();
	}
}
function P_UnLock()
{
	if(confirm('�˲�������������ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ�����������'))
	{
		document.AdsClass.action="?Ad_OP=P_UnLock";
		document.AdsClass.submit();
	}
}
function P_Del()
{
	if(confirm('�˲�����ɾ��ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ��ɾ����'))
	{
		document.AdsClass.action="?Ad_OP=P_Del";
		document.AdsClass.submit();
	}
}
function AdClassDelAll()
{
	if(confirm('�˲�����ɾ��������Ŀ�Լ�������Ŀ�Ĺ�棿\n��ȷ��ɾ����'))
	{
		location='?Ad_OP=AdClassDelAll';
	}
}
function CheckAll(CheckType)
{
	var checkBoxArray=document.all(CheckType)
	if(checkBoxArray[checkBoxArray.length-1].checked)
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=true;
			
		}
	}else
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=false;
		}
	}
}
function Lock(ID)
{	
	if(confirm('�˲���������ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ��������'))
	{
		location='?Ad_OP=Lock&ID='+ID;
	}
}
function Unlock(ID)
{
	if(confirm('�˲�������������ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ�����������'))
	{
		location='?Ad_OP=Unlock&ID='+ID;
	}
}
function DelOne(ID)
{
	if(confirm('�˲�����ɾ��ѡ����Ŀ�Լ�����Ŀ�����й�棿\n��ȷ��ɾ����'))
	{
		location='?Ad_OP=DelOne&ID='+ID;
	}
}
function Update(ID)
{
	location='Ads_AddClass.asp?ID='+ID;
}
</script>
<%
Set Conn=Nothing
%><!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





