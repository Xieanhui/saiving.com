<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

if Request("action") = "Del" then
	if Request("id") <> "" then CollectConn.Execute("Delete from FS_Rule where id in (" & FormatIntArr(Replace(Request("id"),"***",",")) & ")")
	Response.Redirect("Rule.asp")
	Response.End
end if

if Request.Form("Result") = "add" then
	if Request.Form("RuleName")="" then 
		Response.Write("<script>alert('请填写规则名称');history.back();</script>")
		Response.End
	end if
    Dim Sql,RsEditObj
	Set RsEditObj = Server.CreateObject(G_FS_RS)
	Sql = "Select * from FS_Rule"
	RsEditObj.Open Sql,CollectConn,1,3
	RsEditObj.AddNew
	RsEditObj("RuleName") = Request.Form("RuleName")
	Dim KeywordSetting
	If InStr(Request.Form("KeywordSetting"),"[过滤字符串]")<>0 then
		KeywordSetting = Split(Request.Form("KeywordSetting"),"[过滤字符串]",-1,1)
		RsEditObj("HeadSeting") = NoSqlHack(KeywordSetting(0))
		RsEditObj("FootSeting") = NoSqlHack(KeywordSetting(1))
	Else
		RsEditObj("HeadSeting") = ""
		RsEditObj("FootSeting") = ""
	End If
	RsEditObj("ReContent") = NoSqlHack(Request.Form("ReContent"))
	RsEditObj("AddDate") = Now()
	RsEditObj.update
	RsEditObj.close
	Set RsEditObj = Nothing
	Response.Redirect("Rule.asp")
	Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改新闻</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2" onselectstart="return false;">
<%
if Request("action") = "AddRule" then
	Call Add()
else
	Call Main()
end if
Sub Main()
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" align="center" alt="添加" style="cursor:hand;" onClick="location='?Action=AddRule';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
			<td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="后退" style="cursor:hand;" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td width="40%" height="26" class="xingmu"> <div align="center">规则名称</div></td>
    <td width="20%" class="xingmu"><div align="center">时间</div></td>
    <td width="20%" height="20" class="xingmu"> <div align="center">操作</div></td>
  </tr>
  <%
Dim RsSite,Sitesql,CheckInfo,StrPage,Select_Count,Select_PageCount,i,ApplyStation,RsTempObj
Set RsSite = Server.CreateObject(G_FS_RS)
SiteSql = "select * from FS_Rule order by id desc"
RsSite.Open SiteSql,CollectConn,1,1
if Not RsSite.Eof then
	StrPage = Request.QueryString("Page")
	if StrPage <= 1 or StrPage = "" then 
		StrPage = 1
	else 
		StrPage = CInt(StrPage)
	end if
	RsSite.PageSize = 12
	Select_Count = RsSite.RecordCount
	Select_PageCount = RsSite.PageCount
	if StrPage > Select_PageCount then StrPage = Select_PageCount
	RsSite.AbsolutePage = CInt(StrPage)
	for i=1 to RsSite.PageSize
		IF RsSite.Eof Then Exit For
%>
  <tr class="hback"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/folder.gif" width="16" height="16"></td>
          <td><span class="TempletItem" RuleID="<% = RsSite("ID") %>"> 
            <% = RsSite("RuleName") %>
            </span></td>
        </tr>
      </table></td>
    <td> <div align="center">
        <% = RsSite("AddDate") %>
      </div></td>
    <td> <div align="center"><span style="cursor:hand;" onClick="if (confirm('确定要修改吗?')){location='Rulemodify.asp?RuleId=<% = RsSite("ID") %>';}">属性</span>&nbsp;&nbsp;<span style="cursor:hand;" onClick="if (confirm('确定要删除吗?')){location='?action=Del&Id=<% = RsSite("ID") %>';}">删除</span></div></td>
  </tr>
  <%
		RsSite.MoveNext
	next
  %>
  <tr class="hback"> 
    <td colspan="4"> <table  width="100%" border="0" cellpadding="5" cellspacing="0">
        <tr> 
          <td height="30"> <div align="right"> 
              <%
				Response.Write"&nbsp;共<b>" & Select_PageCount & "</b>页<b>" & Select_Count & "</b>条记录，每页<b>" & RsSite.pagesize & "</b>条，本页是第<b>" & StrPage &"</b>页"
				if Int(StrPage)>1 then
					Response.Write "&nbsp;<a href=?Page=1>第一页</a>&nbsp;"
					Response.Write "&nbsp;<a href=?Page=" & CStr(CInt(StrPage) - 1) & ">上一页</a>&nbsp;"
				end if
				if Int(StrPage) < Select_PageCount then
					Response.Write "&nbsp;<a href=?Page=" & CStr(CInt(StrPage) + 1 ) & ">下一页</a>"
					Response.Write "&nbsp;<a href=?Page="& Select_PageCount &">最后一页</a>&nbsp;"
				end if
				Response.Write"<br>"
				RsSite.close
				Set RsSite = Nothing
				%>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <% 
end if
%>
</table>
<%End Sub%>
<%
Sub Add()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="Form1" method="post" action="" id="Form1">
    <tr class="hback"> 
      <td height="25" colspan="5" valign="middle"> 
        <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="35" align="center" alt="保存" style="cursor:hand;" onClick="document.Form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="后退" style="cursor:hand;" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="add"></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback"> 
      <td width="100" height="34"> 
        <div align="center">规则名称</div></td>
      <td> 
        <input style="width:100%;" name="RuleName" type="text" id="RuleName"> 
        <div align="right"></div></td>
    </tr>
    <tr  class="hback"> 
      <td height="110"> 
        <div align="center">过滤字符串</div></td>
      <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.KeywordSetting.rows>2)document.Form1.KeywordSetting.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.KeywordSetting.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        &nbsp;&nbsp;可用标签:<font onClick="addTag('[过滤字符串]')" style="CURSOR: hand"><b>[过滤字符串]</b></font> 
        &nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
        <br>
	  <textarea name="KeywordSetting" id="KeywordSetting" onFocus="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" style="width:100%;"></textarea> 
        <div align="right"></div></td>
    </tr>
    <tr class="hback"> 
      <td> 
        <div align="center">替换为</div></td>
      <td> 
        <textarea name="ReContent" rows="5" style="width:100%;"></textarea></td>
    </tr>
  </form>
</table>
<%End Sub%>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
currObj = "uuuu";
function getActiveText(obj)
{
	currObj = obj;
}

function addTag(code)
{
	addText(code);
}

function addText(ibTag)
{
	var isClose = false;
	var obj_ta = currObj;
	if (obj_ta.isTextEdit)
	{
		obj_ta.focus();
		var sel = document.selection;
		var rng = sel.createRange();
		rng.colapse;

		if((sel.type == "Text" || sel.type == "None") && rng != null)
		{
			rng.text = ibTag;
		}
		obj_ta.focus();
		return isClose;
	}
	else return false;
}	
-->
</script>






