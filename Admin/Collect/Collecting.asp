<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/WaterPrint_Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="CS_Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
Response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
if not MF_Check_Pop_TF("CS_collect") then Err_Show
Dim p_SYS_ROOT_DIR,SiteID,ErrorInfoStr,Action,SaveIMGPath,ListHeadSetting,ListFootSetting,LinkHeadSetting,LinkFootSetting
Dim PagebodyHeadSetting,PagebodyFootSetting,PageTitleHeadSetting,PageTitleFootSetting,OtherPageFootSetting,OtherPageHeadSetting
Dim OtherNewsPageHeadSetting,OtherNewsPageFootSetting,AuthorHeadSetting,AuthorFootSetting,SourceHeadSetting,SourceFootSetting
Dim AddDateHeadSetting,AddDateFootSetting,IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim IsStyle,IsDiv,IsA,IsClass,IsFont,IsSpan,IsObjectTF,IsIFrame,IsScript,HandSetAuthor,HandSetSource,HandSetAddDate,TextTF,SaveRemotePic,IsReverse
Dim ObjURL,ReturnValue,CollectStartLocation,CollectEndFlag,CollectObjURL,CollectedPageURL,p_DoMain_Str
Dim SiteName,CollectingSiteID,CollectSiteIndex,AllNewsNumber,CollectOKNumber,CollectPageNumber,Num,CollectType
Dim OtherNewsType,OtherNewsPageIndexSetting,OtherNewsPageIndexSettingStartPageNum,OtherNewsPageIndexSettingEndPageNum,OtherNewsPageIndexSettingHandPageContent
Dim WebCharset,WaterPrintTF,CS_SiteReKeyID,Temp_picPath,AuditTF
Dim AutoCollect
AutoCollect=False
if G_VIRTUAL_ROOT_DIR = "" then
	p_SYS_ROOT_DIR = ""
else
	p_SYS_ROOT_DIR = "/" & G_VIRTUAL_ROOT_DIR
end if

p_DoMain_Str = "http://"&Request.Cookies("FoosunMFCookies")("FoosunMFDomain")
Action = Request("Action")
SiteID = Request("SiteID")
ErrorInfoStr = ""
CollectEndFlag = False
CollectedPageURL = Request("CollectedPageURL")
AllNewsNumber = Request("AllNewsNumber")
if AllNewsNumber = "" then
	AllNewsNumber = 0
else
	AllNewsNumber = CLng(AllNewsNumber)
end if
CollectOKNumber = Request("CollectOKNumber")
if CollectOKNumber = "" then
	CollectOKNumber = 0
else
	CollectOKNumber = CLng(CollectOKNumber)
end if
CollectSiteIndex = Request("CollectSiteIndex")
if CollectSiteIndex = "" then
	CollectSiteIndex = 0
else
	CollectSiteIndex = CInt(CollectSiteIndex)
end if
CollectPageNumber = Request("CollectPageNumber")
if CollectPageNumber = "" then
	CollectPageNumber = 0
else
	CollectPageNumber = CInt(CollectPageNumber)
end if
CollectStartLocation = Request("CollectStartLocation")
if CollectStartLocation = "" then CollectStartLocation = 0
Num = Request("Num")
If Num = "allNews" Or Num="" Then 
	Num = 10
Else
	if Not IsNumeric(Num) then
		Num = 10
	else
		Num = CInt(Num)
	end if
End If
CollectType = Request("CollectType")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>[site] �����̨ -- ��Ѷ���ݹ���ϵͳ FoosunCMS V5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
  <tr bgcolor="xingmu"> 
    <td height="26" colspan="5" valign="middle" class="hback">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td style="cursor:hand;" width="35" id="StopCollect" align="center" alt="ֹͣ�ɼ�" onClick="location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="xingmu">ȡ��</td>
		  <td width=2 class="Gray">|</td>
          <td style="cursor:hand;" width="35" id="SaveCollect" align="center" alt="����ɼ����Ȳ�����" onClick="location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="xingmu">����</td>
		  <td width=2 class="Gray">|</td>
		  <td style="cursor:hand;" width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="xingmu">����</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="tabble">
  <tr class="hback_1">
    <td height="20"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
		<%If CollectType="ResumeCollect" then%>
			<td width="50%;" align="right"><font color="#FF0000" id="CollectEndArea">��������</font></td>
		<%else%>
			<td width="50%;" align="right"><font color="#FF0000" id="CollectEndArea">���ڲɼ�</font></td>
		<%End if%>
			<td width="50%;">&nbsp;<font color="#FF0000" id="ShowInfoArea" size="+1">&nbsp;</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td valign="middle" class="hback">
<%
if Action = "Submit" then
	if SiteID <> "" then
		GetCollectPara
		If AllNewsNumber>=Num And Num<>0 Then 
			CollectEndFlag = True
		End If
		if CollectEndFlag then
			if ErrorInfoStr <> "" then
				Response.Write(ErrorInfoStr)
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>�ɼ�����</strong>�� ����ȡ" & AllNewsNumber & "�����ţ��ɼ��ɹ�" & CollectOKNumber & "�����š�"
				Response.Write(ReturnValue)
				Response.Write("<script language=""JavaScript"">setTimeout('SetCollectEndStr();',100);</script>")
			end if
		elseif CollectType<>"ResumeCollect" Then
			GetNewsPageContent()
			if CollectStartLocation = 0 then
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ���ҳ" & CollectPageNumber & "</font></strong>��" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ���ҳ" & CollectPageNumber + 1 & "</font></strong>��" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			end if
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ�վ��</font></strong>��" & SiteName & "<br>" & ReturnValue
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ����</font></strong>���Ѿ���ȡ" & AllNewsNumber & "�����ţ�����" & CollectOKNumber & "������<br>" & ReturnValue
			Response.Write(ReturnValue & "<meta http-equiv=""refresh"" content=""2;url=Collecting.asp?Action=Submit&CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&Num=" & Num & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber & """>")
		else
			ResumeGetNewsPageContent()
			if CollectStartLocation = 0 then
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ���ҳ" & CollectPageNumber & "</font></strong>��" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ���ҳ" & CollectPageNumber + 1 & "</font></strong>��" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			end if
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ�վ��</font></strong>��" & SiteName & "<br>" & ReturnValue
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>�ɼ����</font></strong>���Ѿ���ȡ" & AllNewsNumber & "�����ţ�������" & CollectOKNumber & "������<br>" & ReturnValue
			Response.Write(ReturnValue & "<meta http-equiv=""refresh"" content=""2;url=Collecting.asp?Action=Submit&CollectType=ResumeCollect&CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber & """>")
		end if
	end if
end if
%>
	</td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var ForwardShow=true;
function ShowPromptInfo()
{
	var TempStr=document.all.ShowInfoArea.innerText;
	if (ForwardShow==true)
	{
		if (TempStr.length>4) ForwardShow=false;
		document.all.ShowInfoArea.innerText=TempStr+'.';
	}
	else
	{
		if (TempStr.length==2) ForwardShow=true;
		document.all.ShowInfoArea.innerText=TempStr.substr(0,TempStr.length-1);
	}
}
function SetCollectEndStr()
{
	document.all.CollectEndArea.innerText='�ɼ�����,3���Ӻ󷵻���ҳ��';
	setTimeout("location='Site.asp';",3000);
}
window.setInterval('ShowPromptInfo()',500);</script>
<% if Action = "" then %>
<script language="JavaScript">
setTimeout("location='?SiteID=<% = SiteID %>&CollectType=<%= CollectType %>&Action=Submit&Num=<%= Num %>';",10);
</script>
<% end if %>
<%
Set Conn = Nothing
Set CollectConn = Nothing

%>





