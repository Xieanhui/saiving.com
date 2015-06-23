<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
'on error resume next
Dim Conn,User_Conn,PrizeID,ChangePrizeRs,prizeDesc,PrizeName,needPoint,PrizePic,storage,startDate,endDate,provider,perUserNum
Dim i,str_CurrPath
'************************************Update
if Request.QueryString("Act")="editprize" then
	PrizeID=NoSqlHack(Request.QueryString("PrizeID"))
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_award") then Err_Show 

	Set ChangePrizeRs=Server.CreateObject(G_FS_RS)
	ChangePrizeRs.open "select prizeID,PrizeName,prizeDesc,PrizePic,NeedPoint,storage,StartDate,EndDate,provider,perUserNum from FS_ME_Prize where prizeID="&PrizeID,User_Conn,1,1
	if not ChangePrizeRs.eof then
		PrizeName=ChangePrizeRs("PrizeName")
		needPoint=ChangePrizeRs("needPoint")
		PrizePic=ChangePrizeRs("PrizePic")
		startDate=ChangePrizeRs("startDate")
		endDate=ChangePrizeRs("endDate")
		prizeDesc=ChangePrizeRs("prizeDesc")
		storage=ChangePrizeRs("storage")
		provider=ChangePrizeRs("provider")
		perUserNum=ChangePrizeRs("perUserNum")
	end if
elseif Request.QueryString("Act")="add" then
 startDate=datevalue(Now())
 needPoint=0
end if
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;"> 
<%
	if NoSqlHack(Request.QueryString("act"))="editprize" then
		Response.Write("<form name='PrizePanel' id='PrizePanel' method='post' action='awardAction.asp?act=editPrizeaction&Prizeid="&NoSqlHack(Request.QueryString("Prizeid"))&"'>")
	else
		Response.Write("<form name='PrizePanel' id='PrizePanel' method='post' action='awardAction.asp?act=addPrizeaction'>")
	end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <tr class="hback">
    <td align="right" class="xingmu" colspan="2"><div align="left">奖品项目设置&nbsp;&nbsp;| &nbsp;<a href="#" onClick="history.back()">后退</a></div></td></tr> 
        <tr class="hback"> 
          <td align="right">兑换物品：</td> 
          <td> <input name="PrizeName" type="text" id="PrizeName" value="<%=PrizeName%>" size="50" maxlength="20"/>
          <font color="#FF0000">*</font><span id="PrizeName_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">需要积分： </td>
    <td><input name="needpoint" type="text" id="needpoint"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%=needPoint%>" size="50" maxlength="4"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
    <font color="#FF0000">*</font><span id="needpoint_Alert"></span>
    </td>
  </tr>
<tr class="hback">
  <td align="right">物品说明：</td>
  <td><textarea name="PrizeDesc" cols="60" rows="10" id="PrizeDesc"><%=prizeDesc%></textarea></td>
</tr>
<tr class="hback">
    <td align="right">主题图片：</td>
    <td><input name="PrizePic" type="Text" id="PrizePic" size="50" value="<%=PrizePic%>" maxlength="50" >
    <button onClick="javascript:OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,$('PrizePic'));">选择图片</button>
    <span id="UpfileSize_Alert"></span></td>
  </tr>
<tr class="hback">
  <td align="right">物品数量：</td>
  <td><input name="storage" type="text" id="storage"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%=storage%>" size="50" maxlength="4"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
  <span id="storage_alert"></span></td>
</tr>
<tr class="hback">
  <td align="right">每人兑换限量：</td>
  <td><input name="perUserNum" type="text" id="perUserNum"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%=perUserNum%>" size="50" maxlength="4"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
  <span id="perUserNum_alert"></span><span id="isRight_alert"></span></td>
</tr>
<tr class="hback">
  <td align="right">物品提供组织：</td>
  <td><input name="provider" type="text" id="provider" value="<%=provider%>" size="50" maxlength="20"></td>
</tr>
      <tr class="hback"> 
          <td align="right">开始日期：</td> 
          <td><input name="startDate" type="text" id="startDate" value="<%if Request.QueryString("act")="editprize" then response.Write(startDate) else Response.Write(Datevalue(Now()))%>" size="50" readonly="true"><button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.PrizePanel.startDate);document.PrizePanel.startDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="startDate_Alert"></span></td> 
  </tr>
        <tr class="hback">
          <td align="right">截止日期：</td>
          <td><input name="EndDate" type="text" id="EndDate" value="<%=endDate%>" size="50" readonly="true"><button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.PrizePanel.EndDate);document.PrizePanel.EndDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="EndDate_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">&nbsp;</td>
          <td><input type="button" name="Submit" value="保存" onClick="ADDEditPrize()">
&nbsp;          <input type="reset" name="Submit2" value="重置"></td>
        </tr>
</table> 
</form>
</body>
<%
if Request.QueryString("Act")="eidtPrize" then
	ChangePrizeRs.close
	set ChangePrizeRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
	Set prizeRs=nothing
end if
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
function ADDEditPrize()
{
	var flag1=isEmpty("PrizeName","PrizeName_Alert");
	var flag2=isEmpty("needpoint","needpoint_Alert");
	var flag3=isNumber("needpoint","needpoint_Alert","积分应该为正整数",true);
	var flag4=isEmpty("storage","storage_alert");
	var flag5=isNumber("storage","storage_Alert","物品数量应该为正整数",true);
	var flag6=isEmpty("perUserNum","perUserNum_alert");
	var flag7=isNumber("perUserNum","perUserNum_alert","物品数量应该为正整数",true);
	var flag8=isEmpty("startDate","startDate_Alert");
	var flag9=isEmpty("EndDate","EndDate_Alert");
	var flag10=isNumber('needpoint','needpoint_Alert','积分应该为正整数',true);
	if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8&&flag9&&flag10)
	{
		document.PrizePanel.submit();
	}
		
}
</script>
</html>






