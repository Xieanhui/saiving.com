<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF

Dim RsEditObj,EditSql,SiteID
Dim LinkHeadSetting,LinkFootSetting
Dim ObjUrl,ListHeadSetting,ListFootSetting,NewsLinkStr
Dim HandSetAuthor,HandSetSource,HandSetAddDate,WebCharset
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write("û���޸ĵ�վ��")
	else
		ObjUrl = RsEditObj("ObjUrl")
		ListHeadSetting = RsEditObj("ListHeadSetting")
		ListFootSetting = RsEditObj("ListFootSetting")
		HandSetAuthor = RsEditObj("HandSetAuthor")
		HandSetSource = RsEditObj("HandSetSource")
		HandSetAddDate = RsEditObj("HandSetAddDate")
		WebCharset = RsEditObj("WebCharset")
	end if
else
	Response.write("û���޸ĵ�վ��")
end if
Dim ListSetting
If InStr(Request.Form("LinkSetting"),"[�б�URL]") = 0 Then
	Response.Write "<script>alert('�б�URLû�����û����ò���ȷ��');history.back();</script>"
	Response.End 
End if
ListSetting = Split(Request.Form("LinkSetting"),"[�б�URL]",-1,1)
LinkHeadSetting = ListSetting(0)
LinkFootSetting = ListSetting(1)

if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject(G_FS_RS)
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("LinkHeadSetting") = LinkHeadSetting
	RsAddObj("LinkFootSetting") = LinkFootSetting
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if

Dim ResponseAllStr,NewsListStr
ResponseAllStr = GetPageContent(ObjURL,WebCharset)
NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
NewsLinkStr = FormatUrl(GetOtherContent(NewsListStr,LinkHeadSetting,LinkFootSetting),ObjUrl)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" method="post" action="SiteFiveStep.asp" id="Form1">
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr  class="hback">
            <td width="50" style="cursor:hand" align="center" alt="���Ĳ�" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
			<td width=2 class="Gray">|</td>
            <td width="50" style="cursor:hand" align="center" alt="���岽" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" style="cursor:hand" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"> <input type="hidden" name="NewsLinkStr" value="<% = NewsLinkStr %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="20%"> <div align="center">����</div></td>
      <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.PageTitleSetting.rows>2)document.Form1.PageTitleSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.PageTitleSetting.rows+=1" style='cursor:hand'><b>����</b></span>
	  &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="5"></td>
          </tr>
        </table>
        <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="PageTitleSetting" cols="50" rows="3" id="textarea4" style="width:100%;"><%=RsEditObj("PageTitleHeadSetting")%>[����]<%=RsEditObj("PageTitleFootSetting")%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td> <div align="center">����</div></td>
      <td> &nbsp;&nbsp;�������� <span onClick="if(document.Form1.PagebodySetting.rows>2)document.Form1.PagebodySetting.rows-=1" style='cursor:hand'><b>��С</b></span> 
        <span onClick="document.Form1.PagebodySetting.rows+=1" style='cursor:hand'><b>����</b></span> 
        &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
        <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onChange="getActiveText(this)" name="PagebodySetting" cols="50" rows="6" id="textarea" style="width:100%;"><%=RsEditObj("PagebodyHeadSetting")%>[����]<%=RsEditObj("PagebodyFootSetting")%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td height="26" colspan="4"> <div align="left"> ����������������������������������
<input name="OtherSetType" type="radio" onClick="ChangeSetOption(0);" value="0" checked>
          �������� 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(1);" value="1">
          ������Դ 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(2);" value="2">
          ����ʱ�� 
        </div></td>
    </tr>
    <tr class="hback" id="SetAuthor" style="display:;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAuthor" value="<% = HandSetAuthor %>"></td>
          </tr>
          <tr> 
            <td width="20%"> <div align="center">����</div></td>
            <td colspan="3">&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.AuthorSetting.rows>2)document.Form1.AuthorSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.AuthorSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			 &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="AuthorSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AuthorHeadSetting")%>[����]<%=RsEditObj("AuthorFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" id="SetSource" style="display:none;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetSource" value="<% = HandSetSource %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">��Դ</div></td>
            <td colspan="3">&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.SourceSetting.rows>2)document.Form1.SourceSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.SourceSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[��Դ]')" style="CURSOR: hand"><b>[��Դ]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="SourceSetting" cols="50" rows="3" id="textarea9a" style="width:100%;"><%=RsEditObj("SourceHeadSetting")%>[��Դ]<%=RsEditObj("SourceFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" id="SetAddTime" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAddDate" value="<% = HandSetAddDate %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">����ʱ��</div></td>
            <td>&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.AddDateSetting.rows>2)document.Form1.AddDateSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.AddDateSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[����ʱ��]')" style="CURSOR: hand"><b>[����ʱ��]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="AddDateSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AddDateHeadSetting")%>[����ʱ��]<%=RsEditObj("AddDateFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" > 
      <td height="26" colspan="4"> <div align="center">
          <input name="OtherNewsType" <% if RsEditObj("OtherNewsType") = 0 then Response.Write("checked") %> type="radio" onClick="ChangeNewsSetOption(0);" value="0" checked>
          ���������ŷ�ҳ 
          <input type="radio" <% if RsEditObj("OtherNewsType") = 1 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(1);" value="1">
          ��Ƿ�ҳ
          <input type="radio" <% if RsEditObj("OtherNewsType") = 2 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(2);" value="2">
          ҳ���ҳ
          <!--input type="radio" <% if RsEditObj("OtherNewsType") = 3 then Response.Write("checked") %> name="OtherNewsType" onClick="ChangeNewsSetOption(3);" value="3">
          �����ֶ���ҳ-->
		  </div></td>
    </tr>
    <tr class="hback" id="SetCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="20%"> 
              <div align="center">��ҳ����<br>(��һҳ)</div></td>
      <td>&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.OtherNewsPageSetting.rows>2)document.Form1.OtherNewsPageSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.OtherNewsPageSetting.rows+=1" style='cursor:hand'><b>����</b></span>&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.OtherNewsPageSetting);" onClick="addTag('[��ҳ����]')" style="CURSOR: hand"><b>[��ҳ����]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageSetting" cols="50" rows="3" id="textarea5" style="width:100%;"><%=RsEditObj("OtherNewsPageHeadSetting")%>[��ҳ����]<%=RsEditObj("OtherNewsPageFootSetting")%></textarea><br /><span style="color:red;">��:<% = Server.HTMLEncode("<a href=") %>"[��ҳ����]"<% = Server.HTMLEncode(">") %>��һҳ Ҫ�� ��һҳ ����Ϊ����ҳ����Ψһ�ַ�</span></td>
    </tr>
        </table></td>
    </tr>
    <tr class="hback"  id="SetIndexCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback" > 
            <td width="20%"> 
            <div align="center">��ҳ����</div></td>
            <td width="80%" >
			&nbsp;&nbsp;��������
			<span onClick="if(document.Form.IndexRule.rows>2)document.Form.IndexRule.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.IndexRule.rows+=1" style='cursor:hand'><b>����</b></span>
			 &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.form1.OtherNewsPageIndexSetting);" onClick="addTag('[��ҳ����]')" style="CURSOR: hand"><b>[��ҳ����]</b></font> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;<font onmouseover="getActiveText(document.form1.OtherNewsPageIndexSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
			<textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageIndexSetting" cols="50" rows="3" id="OtherNewsPageIndexSetting" style="width:100%;"><%=RsEditObj("OtherNewsPageIndexSetting")%></textarea><br /><span style="color:red;">��:<% = Server.HTMLEncode("&nbsp;</font><a href=") %>"[��ҳ����]"<% = Server.HTMLEncode(">") %>[����]<% = Server.HTMLEncode("</a>") %>Ҫ�� [��ҳ����] ǰ�ַ�������Ϊ����ҳ����Ψһ����</span></td>
          </tr>
        </table></td>
    </tr>
	<!--tr class="hback"  id="SetHandCutPage" style="display:none;"> 
      <td width="10%"> <div align="center">��ҳ����</div></td>
      <td height="26">&nbsp;&nbsp;��������<span onClick="if(document.Form.HandPageContent.rows>2)document.Form.HandPageContent.rows-=1" style='cursor:hand'><b>��С</b></span><span onClick="document.Form.HandPageContent.rows+=1" style='cursor:hand'><b>����</b></span><textarea  name="OtherNewsPageIndexSettingHandPageContent" rows="6" id="OtherNewsPageIndexSettingHandPageContent" style="width:100%;"><% = RsEditObj("OtherNewsPageIndexSettingHandPageContent") %></textarea>
	  </td>
	</tr-->	
</table>
</form>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr class="hback">
    <td colspan="2" height="28" class="xingmu"> <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ԥ�����:<a href="<% = NewsLinkStr %>" target="_blank"> 
        <% = NewsLinkStr %>
        </a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set RsEditObj = Nothing
Set CollectConn = Nothing
%>
<script language="JavaScript">
function ChangeSetOption(Flag)
{
	switch (Flag)
	{
		case 0 :
			document.all.SetAuthor.style.display='';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
		case 1 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='';
			document.all.SetAddTime.style.display='none';
			break;
		case 2 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='';
			break;
		case 999 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
		default :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			break;
	}
}

function ChangeNewsSetOption(f_Flag)
{
	switch (f_Flag)
	{
		case 0 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='none';
			break;
		case 1 :
			document.all.SetCutPage.style.display='';
			document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='none';
			break;
		case 2 :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='';
			//document.all.SetHandCutPage.style.display='none';
			break;
		//case 3 :
			//document.all.SetCutPage.style.display='none';
			//document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='';
			//break;
		default :
			document.all.SetCutPage.style.display='none';
			document.all.SetIndexCutPage.style.display='none';
			//document.all.SetHandCutPage.style.display='none';
			break;
	}
}		  

currObj = "uuuu";
function getActiveText(obj)
{
	obj.focus();
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





