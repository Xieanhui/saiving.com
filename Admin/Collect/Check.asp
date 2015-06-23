<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp"-->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
Dim SiteID,Action,NewsSql,RsNewsObj,CurrPage,AllPageNum,RecordNum,i,SiteName,RsTempObj,AttributeStr,DelID,DelIDArray,str_History
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
if not MF_Check_Pop_TF("CS003") then Err_Show
int_RPP=30 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"

Action = Request("Action")
SiteID = Request("SiteID")
if Action = "Del" then
	DelID = Request("ID")
	if DelID <> "" then
		if DelID = "record" then
			CollectConn.Execute("Delete from FS_News where History=1")
		else
			DelIDArray = Split(DelID,"***")
			for i = LBound(DelIDArray) to UBound(DelIDArray)
				if DelIDArray(i) <> "" then
					CollectConn.Execute("Delete from FS_News where ID=" & CintStr(DelIDArray(i)))
				end if
			Next
		end if
	end if	
end if
CurrPage = NoSqlHack(Request("Page"))
if Request("check")="1" then
	str_History = " and History=1"
elseif Request("check")="0" then
	str_History = " and History=0"
elseif Request("check") = "all" Then
	str_History = ""
Else
	str_History = " and History=0"
end if
NewsSql = "Select * from FS_News where 0=0"&NoSqlHack(str_History)&" Order by ID Desc"
Set RsNewsObj = Server.CreateObject(G_FS_RS)
RsNewsObj.Open NewsSql,CollectConn,1,1
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>[site] 管理后台 -- 风讯内容管理系统 FoosunCMS V5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<BODY topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="hback">选择状态：<a href="Check.asp?check=0">采集数据</a> -- <a href="Check.asp?check=1">历史记录</a> -- <a href="Check.asp?check=all">全部采集记录</a> </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="26" nowrap class="xingmu">
      <div align="center"> 标题</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">状态</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">采集站点</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">添加日期</div></td>
    <td width="15%" height="20" nowrap class="xingmu"> 
      <div align="center">操作</div></td>
  </tr>
  <%
if Not RsNewsObj.Eof then
	if CurrPage = "" then
		CurrPage = 1
	else
		CurrPage = CInt(CurrPage)
	end if
	RsNewsObj.PageSize = int_RPP
	RecordNum = RsNewsObj.RecordCount
	AllPageNum = RsNewsObj.PageCount
	if CurrPage > AllPageNum then CurrPage = AllPageNum
	RsNewsObj.AbsolutePage = Cint(CurrPage)
	for i = 1 to RsNewsObj.PageSize
		if RsNewsObj.Eof then Exit For
		Set RsTempObj = CollectConn.Execute("Select SiteName from FS_Site where ID=" & RsNewsObj("SiteID"))
		if Not RsTempObj.Eof then
			SiteName = RsTempObj("SiteName")
		else
			SiteName = "未知"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
%>
  <tr class="hback"> 
    <td height="26" nowrap> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><input type="checkbox" CName="<% = RsNewsObj("Title") %>" value="<% = RsNewsObj("ID") %>" name="NewsID"></td>
          <td><% = Left(RsNewsObj("Title"),20) %></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> 
        <%
		AttributeStr = ""
		if RsNewsObj("History") = True then
			AttributeStr = "<font color=""red"">已入库</fonr>"
		else
			AttributeStr = "未入库"
		end if
		Response.Write(AttributeStr)
		%>
        </div></td>
    <td nowrap><div align="center">
        <% = SiteName %>
      </div></td>
    <td nowrap><div align="center">
        <% = RsNewsObj("AddDate") %>
      </div></td>
    <td nowrap class="SpanStyle"> <div align="center"><span style="cursor:hand;" onClick="if (confirm('确定要修改吗？')) location='EditNews.asp?NewsIDStr=<% = RsNewsObj("ID") %>';">修改</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要删除吗？')) location='?Action=Del&ID=<% = RsNewsObj("ID") %>';">删除</span></div></td>
  </tr>
  <%
		RsNewsObj.MoveNext
	next
%>
  <tr  class="hback"> 
    <td height="30" colspan="5" nowrap>
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <div align="right">　　　　　
              <input type="checkbox" name="checkAll" value="" onClick="selectAll(document.all.NewsID,this.checked)">
              全选&nbsp;&nbsp;
			  <input onClick="DeleteNews('record');" type="button" name="Submit" value=" 删除全部已入库新闻 ">
			  &nbsp;&nbsp;
			  <input onClick="MoveNews('all');" type="button" name="Submit" value=" 全部入库 ">
			  &nbsp;&nbsp;
              <input onClick="MoveNews('');" type="button" name="Submit" value=" 入 库 ">
			  &nbsp;&nbsp;<input onClick="DeleteNews('');" type="button" name="Submit" value=" 删 除 ">
              <%
			  Response.Write"<br>"
			Response.Write(fPageCount(RsNewsObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,CurrPage))
			Response.Write"<br>"
		%>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <%
end if
%>
</table>
</BODY>
</HTML>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsNewsObj = Nothing
%>
<script language="JavaScript">
function MoveNews(f_All_ID)
{
	var ID_Str='',CName_Str='';
	if (f_All_ID=='')
	{
		if (document.all.NewsID.length)
		{
			for (var i=0;i<document.all.NewsID.length;i++)
			{	
				if(document.all.NewsID(i).checked)
				{
					if (ID_Str!=''){ID_Str=ID_Str+'***'+document.all.NewsID(i).value;}else{ID_Str=document.all.NewsID(i).value;}
					if (CName_Str!=''){CName_Str=CName_Str+'***'+document.all.NewsID(i).CName;}else{CName_Str=document.all.NewsID(i).CName;}
				}
			}
		}
		else{if(document.all.NewsID.checked){ID_Str=document.all.NewsID.value;CName_Str=document.all.NewsID.CName;}}
	}
	else {ID_Str='all';CName_Str='全部新闻';}
	if (ID_Str!='')
	{
		if(confirm('确定要入库吗？'))location='MoveNews.asp?ID='+ID_Str+'&CName='+CName_Str;
	}
	else alert('请选择要入库的新闻');
}

function DeleteNews(f_ID_Type)
{
	var ID_Str='';
	if (f_ID_Type=='')
	{
		if (document.all.NewsID.length)
		{
			for (var i=0;i<document.all.NewsID.length;i++)
			{	
				if(document.all.NewsID(i).checked)
				{
					if (ID_Str!=''){ID_Str=ID_Str+'***'+document.all.NewsID(i).value;}else{ID_Str=document.all.NewsID(i).value;}
				}
			}
		}
		else{if(document.all.NewsID.checked)ID_Str=document.all.NewsID.value;}
	}
	else
	{
		ID_Str=f_ID_Type
	}
	if (ID_Str!='')
	{
		if(confirm('确定要删除吗？')) location='?Action=Del&ID='+ID_Str;
	}
	else alert('请选择要删除的新闻');
}

function selectAll(f_OBJ,f_Flag)
{
	if (f_OBJ.length)
	{
		for (var i=0;i<f_OBJ.length;i++)
		{	
			f_OBJ(i).checked=f_Flag;
		}
	}
	else{f_OBJ.checked=f_Flag;}
}
</script>





