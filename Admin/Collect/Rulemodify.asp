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
			Response.Write"<script>alert(""û���޸Ĺ���"");location.href=""javascript:history.back()"";</script>"
			Response.End
		end if
		RsEditObj("RuleName") = NoSqlHack(Request.Form("RuleName"))
		Dim KeywordSetting
		If InStr(Request.Form("KeywordSetting"),"[�����ַ���]")<>0 then
			KeywordSetting = Split(NoSqlHack(Request.Form("KeywordSetting")),"[�����ַ���]",-1,1)
			RsEditObj("HeadSeting") = NoSqlHack(KeywordSetting(0))
			RsEditObj("FootSeting") = NoSqlHack(KeywordSetting(1))
		End If
		RsEditObj("ReContent") = NoSqlHack(Request.Form("ReContent"))
		RsEditObj.UpDate
		RsEditObj.Close
		Set RsEditObj = Nothing
	else
		Response.Write"<script>alert(""�������ݴ���"");location.href=""javascript:history.back()"";</script>"
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
		Response.Write"<script>alert(""û���޸Ĺ���"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
else
	Response.Write"<script>alert(""�������ݴ���"");location.href=""javascript:history.back()"";</script>"
	Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸�����</title>
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
              <td width="35" align="center" class="BtnMouseOut" style="cursor:hand;" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" alt="����">����</td>
              <td width=2 class="Gray">|</td>
              <td width="35" align="center" alt="����" style="cursor:hand;" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
              <td>&nbsp; <input name="Result" type="hidden" id="Result4" value="Edit"> 
                <input name="id" type="hidden" id="id2" value="<% = RuleID %>"></td>
            </tr>
          </table></td>
      </tr>
    </table>
    <table width="98%" border="0" cellpadding="3" cellspacing="1" class="table">
      <tr class="hback"> 
        <td width="100"> <div align="center">��������</div></td>
        <td> <input name="RuleName" style="width:100%;" type="text" id="RuleName" value="<% = RsRuleObj("RuleName") %>"> 
          <div align="right"></div></td>
      </tr>
      <tr class="hback"> 
        <td> <div align="center">�����ַ���</div></td>
        <td> &nbsp;&nbsp;�������� <span onClick="if(document.form1.KeywordSetting.rows>2)document.form1.KeywordSetting.rows-=1;return false;" style='cursor:hand'><b>��С</b></span> 
          <span onClick="document.form1.KeywordSetting.rows+=1;return false;" style='cursor:hand'><b>����</b></span> 
          &nbsp;&nbsp;���ñ�ǩ:<font onClick="addTag('[�����ַ���]')" style="CURSOR: hand"><b>[�����ַ���]</b></font> 
          &nbsp;&nbsp;&nbsp;<font onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br> 
          <br> <textarea name="KeywordSetting"  onfocus="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" id="textarea2" style="width:100%;"><% = RsRuleObj("HeadSeting") %>[�����ַ���]<% = RsRuleObj("FootSeting") %></textarea> 
        </td>
      </tr>
      <tr class="hback"> 
        <td> <div align="center"> �滻Ϊ</div></td>
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






