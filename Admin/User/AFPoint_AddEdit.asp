<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,AID,AFPointRs,ATopic,needPoint,PrizePoint,APic,ADesc,startDate,endDate,AnswerStr,RightAnswerID,AnswerArray,AnswerRs,ArrayIndex
Dim i,str_CurrPath
MF_Default_Conn
MF_Session_TF 
MF_User_Conn
if not MF_Check_Pop_TF("ME_award") then Err_Show 

'************************************Update
if Request.QueryString("Act")="editAFPoint" then
	AID=NoSqlHack(Request.QueryString("AID"))
	Set AFPointRs=User_Conn.execute("Select ATopic,needPoint,PrizePoint,APic,ADesc,AStartDate,AEndDate,AnswerIDS,RightAnswerID from FS_ME_AnswerForPoint where Aid="&AID)
	if not AFPointRs.eof then
		ATopic=AFPointRs("ATopic")
		needPoint=AFPointRs("needPoint")
		PrizePoint=AFPointRs("PrizePoint")
		APic=AFPointRs("APic")
		ADesc=AFPointRs("ADesc")
		startDate=AFPointRs("AstartDate")
		endDate=AFPointRs("AendDate")
		AnswerStr=AFPointRs("AnswerIDS")
		if not isnull(AnswerStr) then
			if AnswerStr<>"" then
				AnswerArray=split(DelHeadAndEndDot(AnswerStr),",")
			end if
		end if
		for ArrayIndex=0 to ubound(AnswerArray)
			if(AFPointRs("RightAnswerID")=Cint(AnswerArray(ArrayIndex))) then
				RightAnswerID=ArrayIndex+1
			end if
		next
	end if
elseif Request.QueryString("Act")="add" then
	if not MF_Check_Pop_TF("ME027") then Err_Show 
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
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes > 
<%
	if NoSqlHack(Request.QueryString("act"))="editAFPoint" then
		Response.Write("<form name='AFPointPanel' id='AFPointPanel' method='post' action='awardAction.asp?act=editAFPointaction&AID="&NoSqlHack(Request.QueryString("AID"))&"'>")
	else
		Response.Write("<form name='AFPointPanel' id='AFPointPanel' method='post' action='awardAction.asp?act=addAFPointaction'>")
	end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <tr class="hback">
    <td align="right" class="xingmu" colspan="2"><div align="left">竞答项目设置&nbsp;&nbsp;| &nbsp;<a href="#" onClick="history.back()">后退</a></div></td></tr> 
        <tr class="hback"> 
          <td align="right">竞答主题：</td> 
          <td> <input name="ATopic" type="text" id="ATopic" value="<%=ATopic%>" size="50" maxlength="40"/>
          <font color="#FF0000">*</font><span id="ATopic_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">需要积分： </td>
    <td><input name="needpoint" type="text" id="needpoint" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%=needPoint%>" size="50" maxlength="4"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
    <font color="#FF0000">*</font><span id="needpoint_Alert"></span>
    </td>
  </tr>
<tr class="hback">
  <td align="right">奖励积分：</td>
  <td><input name="PrizePoint" type="text" id="PrizePoint" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%=PrizePoint%>" size="50" maxlength="4"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
  <font color="#FF0000">*</font><span id="PrizePoint_Alert"></span></td>
</tr>
<tr class="hback">
    <td align="right">主题图片：</td>
    <td><input name="APic" type="text" id="APic" size="50" value="<%=APic%>" maxlength="120">
    <span id="APic_Alert"><button onClick="javascript:OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,$('APic'));">选择图片</button></span></td>
  </tr>
<tr class="hback">
  <td align="right">竞答说明：</td>
  <td><textarea name="ADesc" cols="60" rows="10" id="ADesc"><%=ADesc%></textarea></td>
</tr>
<tr class="hback">
  <td align="right">&nbsp;</td>
  <td>&nbsp;</td>
</tr>
      <tr class="hback"> 
          <td align="right">开始日期：</td> 
          <td><input name="startDate" type="text" id="startDate" value="<%=startDate%>" size="49" readonly="true"> 
            <button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.AFPointPanel.startDate);document.AFPointPanel.startDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="startDate_Alert"></span></td> 
  </tr>
        <tr class="hback">
          <td align="right">截止日期：</td>
          <td><input name="EndDate" type="text" id="EndDate" value="<%=endDate%>" size="49" readonly="true"> 
            <button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.AFPointPanel.EndDate);document.AFPointPanel.EndDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="EndDate_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right"> 竞答答案数 ：</td>
          <td><input name="AnswerNum" type="text" id="AnswerNum" size="50" value="<%if Request.QueryString("act")="editAFPoint" then Response.Write(ubound(AnswerArray)+1) else Response.Write("1")%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"><button onClick="setPrizeGradeNum(AnswerNum.value)">设置</button>竞答答案均为多选一，即只有一个是正确的</td>
        </tr>
</table> 
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class='hback'>
	<td height="21" colspan="2" id="PrizeContainer">
	<%
		if Request.QueryString("act")="editAFPoint" then
			Response.Write("<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' >")
			for ArrayIndex=0 to Ubound(AnswerArray)
				Set AnswerRs=User_Conn.Execute("Select AnswerDesc,AnswerID  from FS_ME_Answer where AnswerID="&NoSqlHack(AnswerArray(ArrayIndex)))
				Response.Write("<tr class='hback'>")
				Response.Write("<td align='right' width='14%'>答案"&(ArrayIndex+1)&"：</td><td><input name='Answer_"&(ArrayIndex+1)&"' type='text' id='Answer_"&(ArrayIndex+1)&"' size='50' maxlength='6' value='"&AnswerRs("AnswerDesc")&"'></td></tr>")
			next
			Response.Write("</table>")
			AnswerRs.close
		end if
	%>	</td>
</tr>
<tr class='hback'>
<td width="14%" align="right">正确答案： </td>
<td width="86%"><input name="rightAnswer" type="text" id="rightAnswer" size="50" value="<%=RightAnswerID%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')" maxlength="5">
<font color="#FF0000">*</font>请填写答案题编号<span id="rightAnswer_Alert"></span></td>
</tr>
<tr class='hback'>
<td>&nbsp;</td>
<td><div align="left"><input type="button" name="ADDEditAFPointButton" onClick="ADDEditAFPoint()" value="保存"> | <input type="reset" name="" value="重置"></div></td>
</tr>
</table>
</form>
</body>
<%
if Request.QueryString("Act")="editAFPoint" then
	AFPointRs.close
	set AFPointRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
	Set AnswerRs=nothing
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
if("<%=Request.QueryString("act")%>"!="editAFPoint")
{
	setPrizeGradeNum(1);
}
function setPrizeGradeNum(Num){
	Num = parseInt(Num);
	if(isNaN(Num))Num=1;
	//alert(Num);
	var i,PrizeGradeNum='';
	for (i=1;i<=Num;i++)
	{
		PrizeGradeNum = PrizeGradeNum +"<tr class='hback'><td align='right' width='14%'>答案"+i+"：</td><td><input name='Answer_"+i+"' type='text' id='Answer_"+i+"' size='50' value=''></td></tr>";
	}
	document.getElementById("PrizeContainer").innerHTML="<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' >"+PrizeGradeNum+"</table>";
}
function ADDEditAFPoint()
{
	var flag1=isEmpty("ATopic","ATopic_Alert");
	var flag2=isEmpty("startDate","startDate_Alert");
	var flag3=isEmpty("EndDate","EndDate_Alert");
	var flag4=isEmpty("needpoint","needpoint_Alert")
	var flag5=isEmpty("Prizepoint","Prizepoint_Alert")
	var flag6=isNumber('needpoint','needpoint_Alert','积分应该为正整数',true);
	var flag7=isNumber('Prizepoint','Prizepoint_Alert','积分应该为正整数',true);
	var flag8=isEmpty('rightAnswer','rightAnswer_Alert')
	if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8)
	{
		document.AFPointPanel.submit();
	}
}
</script>
</html>






