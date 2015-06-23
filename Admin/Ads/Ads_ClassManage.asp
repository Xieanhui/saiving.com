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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告统计___Powered by foosun Inc.</title>
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
			strShowErr = "<li>锁定成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "Unlock"
			Conn.execute("Update FS_AD_Class Set Lock=0 where AdClassID="&lng_ClassID&"")
			Conn.execute("Update FS_AD_Info Set AdLock=0 where AdClassID="&lng_ClassID&"")
			strShowErr = "<li>解除锁定成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "DelOne"
			Conn.execute("delete  from FS_AD_TxtInfo where AdID In(select AdId from FS_AD_TxtInfo where AdId In(select AdID from FS_AD_Info where AdClassID="&lng_ClassID&"))")
			Conn.execute("delete  from FS_AD_Class where AdClassID="&lng_ClassID&"")
			Conn.execute("delete  from FS_AD_Info where AdClassID="&lng_ClassID&"")
			strShowErr = "<li>删除成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&Error	Url=../Ads_ClassManage.asp")
			Response.end
		Case "P_Lock"
			CheckAllIDFLag("锁定")
			Conn.execute("update FS_AD_Info set AdLock=1 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			Conn.execute("update FS_AD_Class set Lock=1 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			strShowErr = "<li>批量锁定成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "P_UnLock"
			CheckAllIDFLag("解除锁定")
			Conn.execute("update FS_AD_Info set AdLock=0 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			Conn.execute("update FS_AD_Class set Lock=0 where AdClassID in (" & FormatIntArr(CheckAllID) & ")")
			strShowErr = "<li>批量解除锁定成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "P_Del"
			CheckAllIDFLag("删除")
			Temp_C_ID=CheckAllID
			CheckAllID=split(CheckAllID,",")
			For i=0 to Ubound(CheckAllID)
				TempID=TempID&CheckAllID(i)&","
			Next
			Conn.execute("delete  from FS_AD_TxtInfo where AdID In(select AdId from FS_AD_TxtInfo where AdId In(select AdID from FS_AD_Info where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")))")
			Conn.execute("delete  from FS_AD_Class where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")")
			Conn.execute("delete  from FS_AD_Info where AdClassID in (" & FormatIntArr(Temp_C_ID) & ")")
			strShowErr = "<li>批量删除成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Case "AdClassDelAll"
			Conn.execute("delete  from FS_AD_Info where AdClassID in (select AdClassID from FS_AD_Class)")
			Conn.execute("delete  from FS_AD_Class")
			strShowErr = "<li>全部删除成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
	End Select
	Sub CheckAllIDFLag(Showstr)
		If CheckAllID="" or IsNull(CheckAllID) Then
			strShowErr = "<li>请选择要"&Showstr&"的栏目!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End Sub
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">广告分类管理</td>
  </tr>
   <tr class="hback"> 
    <td class="hback"><a href="Ads_AddClass.asp?OpCType=Add">添加分类</a> | <a href="javascript:P_Lock();">批量锁定</a> | <a href="javascript:P_UnLock();">批量解锁</a> | <a href="javascript:P_Del();">批量删除</a> | <a href="javascript:AdClassDelAll();">删除全部</a> | <a href="javascript:history.back();">返回上一级</a></td>
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
    <td width="36%" align="center" class="xingmu">栏目名称</td>
    <td width="15%" align="center" class="xingmu">添加时间</td>
    <td width="15%" align="center" class="xingmu">当前栏目广告数</td>
    <td width="5%" align="center" class="xingmu">状态</td>
    <td width="27%" align="center" class="xingmu">操作</td>
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
					str_ClassType="正常"
				Case 1
					str_ClassType="<font color=""red"">锁定</font>"
			End Select
%>
    <td class="hback" align="Center"><%=str_ClassType%></td>
    <td class="hback" align="Center"><a href="javascript:Lock('<%=o_Adclass_Rs("AdClassID")%>');">锁定</a> | <a href="javascript:Unlock('<%=o_Adclass_Rs("AdClassID")%>');">解锁</a> | <a href="javascript:DelOne('<%=o_Adclass_Rs("AdClassID")%>');">删除</a> | <a href="javascript:Update('<%=o_Adclass_Rs("AdClassID")%>');">修改</a> | <input type="checkbox" name="Checkallbox" value="<%=o_Adclass_Rs("AdClassID")%>"></td>
  </tr>
<%
			o_Adclass_Rs.MoveNext
		If o_Adclass_Rs.Eof or o_Adclass_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""5"" align=""left"">"&fPageCount(o_Adclass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;全选<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');"" value=""0""></td></tr>"
%>
</table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>当前没有广告栏目!</td></tr></table>"
	End If
%>
</form>
</body>
</html>
<script language="javascript">
function P_Lock()
{
	if(confirm('此操作将锁定选中栏目以及此栏目下所有广告？\n你确定锁定吗？'))
	{
		document.AdsClass.action="?Ad_OP=P_Lock";
		document.AdsClass.submit();
	}
}
function P_UnLock()
{
	if(confirm('此操作将解锁锁定选中栏目以及此栏目下所有广告？\n你确定解除锁定吗？'))
	{
		document.AdsClass.action="?Ad_OP=P_UnLock";
		document.AdsClass.submit();
	}
}
function P_Del()
{
	if(confirm('此操作将删除选中栏目以及此栏目下所有广告？\n你确定删除吗？'))
	{
		document.AdsClass.action="?Ad_OP=P_Del";
		document.AdsClass.submit();
	}
}
function AdClassDelAll()
{
	if(confirm('此操作将删除所有栏目以及所有栏目的广告？\n你确定删除吗？'))
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
	if(confirm('此操作将锁定选中栏目以及此栏目下所有广告？\n你确定锁定吗？'))
	{
		location='?Ad_OP=Lock&ID='+ID;
	}
}
function Unlock(ID)
{
	if(confirm('此操作将解锁锁定选中栏目以及此栏目下所有广告？\n你确定解除锁定吗？'))
	{
		location='?Ad_OP=Unlock&ID='+ID;
	}
}
function DelOne(ID)
{
	if(confirm('此操作将删除选中栏目以及此栏目下所有广告？\n你确定删除吗？'))
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
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





