<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

Dim RuleID
RuleID = Request("RuleID")
if Request.Form("Result")="Edit" then
    Dim Sql,RsEditObj
	if RuleID <> "" then
		Set RsEditObj = Server.CreateObject(G_FS_RS)
		Sql = "Select * from FS_Rule where id=" & CintStr(RuleID) & ""
		RsEditObj.Open Sql,CollectConn,1,3
		if RsEditObj.Eof then
			Response.Write"<script>alert(""没有修改规则"");location.href=""javascript:history.back()"";</script>"
			Response.End
		end if
		RsEditObj("RuleName") = NoSqlHack(Request.Form("RuleName"))
		Dim KeywordSetting
		If InStr(Request.Form("KeywordSetting"),"[过滤字符串]")<>0 then
			KeywordSetting = Split(NoSqlHack(Request.Form("KeywordSetting")),"[过滤字符串]",-1,1)
			RsEditObj("HeadSeting") = NoSqlHack(KeywordSetting(0))
			RsEditObj("FootSeting") = NoSqlHack(KeywordSetting(1))
		End If
		RsEditObj("ReContent") = NoSqlHack(Request.Form("ReContent"))
		RsEditObj.UpDate
		RsEditObj.Close
		Set RsEditObj = Nothing
	else
		Response.Write"<script>alert(""参数传递错误"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
	Response.Redirect("Rule.asp")
	Response.End
end if

Dim RsRuleObj
if RuleID <> "" then
	Set RsRuleObj = Server.CreateObject(G_FS_RS)
	Sql = "Select * from FS_Rule where id=" & CintStr(RuleID)
	RsRuleObj.Open Sql,CollectConn,1,3
	if RsRuleObj.Eof then
		Response.Write"<script>alert(""没有修改规则"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
else
	Response.Write"<script>alert(""参数传递错误"");location.href=""javascript:history.back()"";</script>"
	Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改新闻</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<link href="FS_css.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2" onselectstart="return false;">
<form name="form1" id="form1" method="post" action="">
  <div align="center">
    <table width="98%" border="0" cellpadding="1" cellspacing="1" class="table">
      <tr class="hback"> 
        <td colspan="5" valign="middle"> <table width="100%" height="20" border="0" cellpadding="5" cellspacing="1">
            <tr> 
              <td width="35" align="center" class="BtnMouseOut" style="cursor:hand;" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" alt="保存">保存</td>
              <td width=2 class="Gray">|</td>
              <td width="35" align="center" alt="后退" style="cursor:hand;" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
              <td>&nbsp; <input name="Result" type="hidden" id="Result4" value="Edit"> 
                <input name="id" type="hidden" id="id2" value="<% = RuleID %>"></td>
            </tr>
          </table></td>
      </tr>
    </table>
    <table width="98%" border="0" cellpadding="3" cellspacing="1" class="table">
      <tr class="hback"> 
        <td width="100"> <div align="center">规则名称</div></td>
        <td> <input name="RuleName" style="width:100%;" type="text" id="RuleName" value="<% = RsRuleObj("RuleName") %>"> 
          <div align="right"></div></td>
      </tr>
      <tr class="hback"> 
        <td> <div align="center">过滤字符串</div></td>
        <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.form1.KeywordSetting.rows>2)document.form1.KeywordSetting.rows-=1;return false;" style='cursor:hand'><b>缩小</b></span> 
          <span onClick="document.form1.KeywordSetting.rows+=1;return false;" style='cursor:hand'><b>扩大</b></span> 
          &nbsp;&nbsp;可用标签:<font onClick="addTag('[过滤字符串]')" style="CURSOR: hand"><b>[过滤字符串]</b></font> 
          &nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br> 
          <br> <textarea name="KeywordSetting"  onfocus="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" id="textarea2" style="width:100%;"><% = RsRuleObj("HeadSeting") %>[过滤字符串]<% = RsRuleObj("FootSeting") %></textarea> 
        </td>
      </tr>
      <tr class="hback"> 
        <td> <div align="center"> 替换为</div></td>
        <td colspan="3"><textarea style="width:100%;" name="ReContent" cols="30" rows="5" id="ReContent"><% = RsRuleObj("ReContent") %></textarea></td>
      </tr>
    </table>
  </div>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
Set RsRuleObj = Nothing
%>

<script language="javaScript">
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

</script>






