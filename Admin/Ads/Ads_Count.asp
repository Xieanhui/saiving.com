<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS002") then Err_Show
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

Dim str_Action,int_ID,strShowErr,CheckAllID
str_Action=Request.QueryString("Action")
int_ID=Request.QueryString("ID")
CheckAllID=Request.Form("Checkallbox")

Select Case str_Action
	Case "DelOne"
		If int_ID="" Or IsNull(int_ID) Then
			strShowErr = "<li>��������!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			If Isnumeric(int_ID)=False Then
				strShowErr = "<li>��������!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("delete From FS_AD_Source Where AdID="&Clng(int_ID)&"")
				strShowErr = "<li>���ͳ����Ϣ�ɹ�!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp?Page="&Request.QueryString("Page")&"")
				Response.end
			End If
		End If
	Case "P_Del"
		If CheckAllID="" or IsNull(CheckAllID) Then
			strShowErr = "<li>��ѡ��Ҫ���ͳ����Ϣ�Ĺ��!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		CheckAllID=split(CheckAllID,",")
		For i=0 to Ubound(CheckAllID)
			TempID=TempID&CheckAllID(i)&","
		Next
		Conn.execute("delete  From FS_AD_Source Where AdID in (" & FormatIntArr(TempID) & ")")
		strShowErr = "<li>�������ͳ����Ϣ�ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp?Page="&Request.QueryString("Page")&"")
		Response.end
	Case "DelAll"
		Conn.execute("Delete From FS_AD_Source")
		strShowErr = "<li>���ͳ����Ϣ�ɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp")
		Response.end
End Select
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ͳ��___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">�����ϸͳ��</td>
  </tr>
</table>
<%
	Dim o_Count_Rs,str_Count_Sql
	str_Count_Sql= "Select AdID,AdName,AdClickNum,AdShowNum from FS_AD_Info order by AdID Desc"
	Set o_Count_Rs= CreateObject(G_FS_RS)
	o_Count_Rs.Open str_Count_Sql,Conn,1,1
	If Not o_Count_Rs.Eof Then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
<form name="AdsCount" method="post" action="">
  <tr class="hback">
    <td width="279" align="center" class="xingmu">�������</td>
    <td width="90" align="center" class="xingmu">���������</td>
    <td width="82" align="center" class="xingmu">����ʾ����</td>
    <td width="101" align="center" class="xingmu">��Դͳ��</td>
    <td width="84" align="center" class="xingmu">ʱͳ��</td>
    <td width="73" align="center" class="xingmu">��ͳ��</td>
    <td width="68" align="center" class="xingmu">��ͳ��</td>
    <td width="99" align="center" class="xingmu">���ͳ����Ϣ</td>
  </tr>
<%
		o_Count_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>o_Count_Rs.PageCount Then cPageNo=o_Count_Rs.PageCount 
		o_Count_Rs.AbsolutePage=cPageNo
	
		For int_Start=1 TO int_RPP  
%>
  <tr class="hback">
    <td class="hback"><%=o_Count_Rs("AdName")%></td>
    <td align="center" class="hback"><%=o_Count_Rs("AdClickNum")%></td>
    <td align="center" class="hback"><%=o_Count_Rs("AdShowNum")%></td>
    <td align="center" class="hback"><a href="javascript:ShowSource('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>');">����鿴</a></td>
    <td align="center" class="hback"><a href="javascript:ShowHour('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">����鿴</a></td>
    <td align="center" class="hback"><a href="javascript:ShowDay('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">����鿴</a></td>
    <td align="center" class="hback" width="68"><a href="javascript:ShowMonth('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">����鿴</a></td>
    <td align="center" class="hback" width="99"><a href="javascript:DelCount('<%=o_Count_Rs("AdID")%>');">��� | <input type="checkbox" name="Checkallbox" value="<%=o_Count_Rs("AdID")%>"></a></td>
  </tr>
<%
		o_Count_Rs.MoveNext
		If o_Count_Rs.Eof or o_Count_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""8"" align=""left"">"&fPageCount(o_Count_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;ȫѡ<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');""> <input type=""button"" name=""NamePDel"" onclick=""javascript:P_Del();"" value="" ������� ""> <input type=""button"" name=""NameDelAll"" onclick=""javascript:DelAll();"" value="" ȫ����� ""></td></tr>"
%>
</form>
</table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>��ǰû�й��!</td></tr></table>"
	End If
	o_Count_Rs.Close
	Set o_Count_Rs=Nothing
%>
</body>
</html>
<script language="javascript">
function ShowSource(ID,Name)
{
	window.location='Ads_ShowCount.asp?AdName='+Name+'&ID='+ID;
}
function ShowDay(ID,Name)
{
	window.location='Visit_DaysStatistic.asp?AdName='+Name+'&ID='+ID;
}
function ShowHour(ID,Name)
{
	window.location='Visit_HoursStatistic.asp?AdName='+Name+'&ID='+ID;
}
function ShowMonth(ID,Name)
{
	window.location='Visit_MonthsStatistic.asp?AdName='+Name+'&ID='+ID;
}
function DelCount(ID)
{
	if (confirm('�˲�������մ˹�����е�ͳ����Ϣ\n��ȷ�������'))
	{
		window.location='?Action=DelOne&ID='+ID;
	}
}
function P_Del()
{
	if (confirm('�˲����������ѡ������е�ͳ����Ϣ\n��ȷ�������'))
	{
		document.AdsCount.action='?Action=P_Del';
		document.AdsCount.submit();
	}
}
function DelAll()
{
	if (confirm('�˲�����������й��ͳ����Ϣ\n��ȷ�������'))
	{
		window.location='?Action=DelAll';
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
</script>
<%
Set Conn=nothing
%><!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





