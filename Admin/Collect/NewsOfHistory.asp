<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

'判断权限结束
Dim Action
Action = Request("Action")
if Action = "DelAll" then
	if Not JudgePopedomTF(Session("Name"),"P080400") then Call ReturnError1()
	CollectConn.Execute("Delete from FS_News where History=1 and SiteID in(select id from fs_site where siteid="&Session("SiteID")&")")
end if
Dim NewsSql,RsNewsObj,CurrPage,AllPageNum,RecordNum,i,SysClassCName,SiteName,RsTempObj
CurrPage = Request("CurrPage")
NewsSql = "Select * from FS_News where History=1 and SiteID in(select id from fs_site where siteid="&Session("SiteID")&") Order by ID Desc"
Set RsNewsObj = Server.CreateObject(G_FS_RS)
RsNewsObj.Open NewsSql,CollectConn,1,1
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>新闻采集</TITLE>
</HEAD>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<BODY topmargin="2" leftmargin="2" onClick="SelectNews();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" align="center" alt="删除" onClick="DelNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
			<td width=2 class="Gray">|</td>
          <td width="70" align="center" alt="删除全部" onClick="DelAll();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除全部</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="入库" onClick="MoveNewsToSystem();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">入库</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;</td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td height="26" nowrap class="ButtonListLeft"> <div align="center">标题</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">新闻长度</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">目标栏目</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">采集站点</div></td>
    <td width="15%" height="24" nowrap class="ButtonList"> 
      <div align="center">添加日期</div></td>
  </tr>
  <%
if Not RsNewsObj.Eof then
	if CurrPage = "" then
		CurrPage = 1
	else
		CurrPage = CInt(CurrPage)
	end if
	RsNewsObj.PageSize = 18
	RecordNum = RsNewsObj.RecordCount
	AllPageNum = RsNewsObj.PageCount
	if CurrPage > AllPageNum then CurrPage = AllPageNum
	RsNewsObj.AbsolutePage = Cint(CurrPage)
	for i = 1 to RsNewsObj.PageSize
		if RsNewsObj.Eof then Exit For
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & NoSqlHack(RsNewsObj("ClassID")) & "' and SiteID ="&NoSqlHack(Session("SiteID")))
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "栏目不存在"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
		Set RsTempObj = CollectConn.Execute("Select SiteName from FS_Site where ID=" & CintStr(RsNewsObj("SiteID"))&" and SiteID ="&NoSqlHack(Session("SiteID")))
		if Not RsTempObj.Eof then
			SiteName = RsTempObj("SiteName")
		else
			SiteName = "未知"
		end if
		RsTempObj.Close
		Set RsTempObj = Nothing
%>
  <tr> 
    <td height="26" nowrap>
<table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../../Images/Info/WordNews.gif" width="24" height="22"></td>
          <td><span class="TempletItem" NewsID=<% = RsNewsObj("ID") %>>
<% = Left(RsNewsObj("Title"),20) %></span></td>
        </tr>
      </table></td>
    <td nowrap><div align="center"> 
        <% = Len(RsNewsObj("Content")) %>
        字符</div></td>
    <td nowrap><div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap><div align="center"> 
        <% = SiteName %>
      </div></td>
    <td nowrap><div align="center"> 
        <% = RsNewsObj("AddDate") %>
      </div></td>
  </tr>
  <%
		RsNewsObj.MoveNext
	next
%>
  <tr> 
    <td height="30" colspan="5" nowrap><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <div align="right"> 
              <%
			Response.Write"&nbsp;共<b>"& AllPageNum & "</b>页<b>" & RecordNum & "</b>条记录，每页<b>" & RsNewsObj.pagesize & "</b>条，本页是第<b>"& CurrPage &"</b>页"
			if Int(CurrPage) > 1 then
				Response.Write"&nbsp;<a href=?CurrPage=1>首页</a>&nbsp;"
				Response.Write"&nbsp;<a href=?CurrPage=" & Cstr(CInt(CurrPage)-1) & ">上页</a>&nbsp;"
			end if
			if Int(CurrPage) < AllPageNum then
				Response.Write"&nbsp;<a href=?CurrPage=" & Cstr(Cint(CurrPage)+1) & ">下页</a>"
				Response.Write"&nbsp;<a href=?CurrPage=" & AllPageNum & ">末页</a>&nbsp;"
			end if
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
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelNews();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MoveNewsToSystem();','入库','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('location.reload();','刷新','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContent='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',删除,入库,';
	else DisabledContentMenuStr='';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.NewsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectNews()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
				for (i=MaxIndex-1;i<EndIndex;i++)
				{
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
			else
			{
				for (i=EndIndex;i<MaxIndex-1;i++)
				{	
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
				ListObjArray[ElIndex].Obj.className='TempletSelectItem';
				ListObjArray[ElIndex].Selected=true;
			}
		}
	}
	else
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	}
}
function EditNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) window.location='EditNews.asp?NewsIDStr='+SelectedNews;
		else alert('请选择一条新闻');
	}
	else alert('请选择新闻');
}
function DelNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=DelNews.asp&PageTitle=删除新闻&NewsIDStr='+SelectedNews,200,120,window);
	else alert('请选择新闻');
}
function DelAll()
{
	if (confirm('确定要删除吗？')) location='?Action=DelAll'
}
function MoveNewsToSystem()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=MoveNewsToSystem.asp&PageTitle=新闻入库&DelNews=true&NewsIDStr='+SelectedNews,200,120,window);
	else alert('请选择新闻');
}
</script>





