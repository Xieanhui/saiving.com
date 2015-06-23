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

int_RPP=5 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页

Dim str_Action,int_ID,strShowErr,CheckAllID
str_Action=Request.QueryString("Action")
int_ID=Request.QueryString("ID")
CheckAllID=Request.Form("Checkallbox")

Select Case str_Action
	Case "DelOne"
		If int_ID="" Or IsNull(int_ID) Then
			strShowErr = "<li>参数错误!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			If Isnumeric(int_ID)=False Then
				strShowErr = "<li>参数错误!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("delete From FS_AD_Source Where AdID="&Clng(int_ID)&"")
				strShowErr = "<li>清空统计信息成功!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp?Page="&Request.QueryString("Page")&"")
				Response.end
			End If
		End If
	Case "P_Del"
		If CheckAllID="" or IsNull(CheckAllID) Then
			strShowErr = "<li>请选择要清空统计信息的广告!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
		CheckAllID=split(CheckAllID,",")
		For i=0 to Ubound(CheckAllID)
			TempID=TempID&CheckAllID(i)&","
		Next
		Conn.execute("delete  From FS_AD_Source Where AdID in (" & FormatIntArr(TempID) & ")")
		strShowErr = "<li>批量清空统计信息成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp?Page="&Request.QueryString("Page")&"")
		Response.end
	Case "DelAll"
		Conn.execute("Delete From FS_AD_Source")
		strShowErr = "<li>清空统计信息成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Count.asp")
		Response.end
End Select
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告统计___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">广告详细统计</td>
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
    <td width="279" align="center" class="xingmu">广告名称</td>
    <td width="90" align="center" class="xingmu">被点击次数</td>
    <td width="82" align="center" class="xingmu">已显示次数</td>
    <td width="101" align="center" class="xingmu">来源统计</td>
    <td width="84" align="center" class="xingmu">时统计</td>
    <td width="73" align="center" class="xingmu">日统计</td>
    <td width="68" align="center" class="xingmu">月统计</td>
    <td width="99" align="center" class="xingmu">清空统计信息</td>
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
    <td align="center" class="hback"><a href="javascript:ShowSource('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>');">点击查看</a></td>
    <td align="center" class="hback"><a href="javascript:ShowHour('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">点击查看</a></td>
    <td align="center" class="hback"><a href="javascript:ShowDay('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">点击查看</a></td>
    <td align="center" class="hback" width="68"><a href="javascript:ShowMonth('<%=o_Count_Rs("AdID")%>','<%=o_Count_Rs("AdName")%>')">点击查看</a></td>
    <td align="center" class="hback" width="99"><a href="javascript:DelCount('<%=o_Count_Rs("AdID")%>');">清空 | <input type="checkbox" name="Checkallbox" value="<%=o_Count_Rs("AdID")%>"></a></td>
  </tr>
<%
		o_Count_Rs.MoveNext
		If o_Count_Rs.Eof or o_Count_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""8"" align=""left"">"&fPageCount(o_Count_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;全选<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');""> <input type=""button"" name=""NamePDel"" onclick=""javascript:P_Del();"" value="" 批量清空 ""> <input type=""button"" name=""NameDelAll"" onclick=""javascript:DelAll();"" value="" 全部清空 ""></td></tr>"
%>
</form>
</table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>当前没有广告!</td></tr></table>"
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
	if (confirm('此操作将清空此广告所有的统计信息\n你确定清空吗？'))
	{
		window.location='?Action=DelOne&ID='+ID;
	}
}
function P_Del()
{
	if (confirm('此操作将清空所选广告所有的统计信息\n你确定清空吗？'))
	{
		document.AdsCount.action='?Action=P_Del';
		document.AdsCount.submit();
	}
}
function DelAll()
{
	if (confirm('此操作将清空所有广告统计信息\n你确定清空吗？'))
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
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





