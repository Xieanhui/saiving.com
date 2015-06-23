<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->

<%
'on error resume next
Dim Conn,User_Conn,awardID,awardRs,awardName,awardPic,startDate,endDate,prizeStr,prizeArray,prizeRs,ArrayIndex,needPoint
Dim i,str_CurrPath
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
'************************************Update
if Request("Act")="edit" then
	awardID=NoSqlHack(Request("awardID"))
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	
	if not MF_Check_Pop_TF("ME028") then Err_Show 

	Set awardRs=User_Conn.execute("Select awardName,awardPic,startDate,EndDate,prizeIDS from FS_ME_award where awardid="&CintStr(awardID))
	if not awardRs.eof then
		awardName=awardRs("awardName")
		awardPic=awardRs("awardPic")
		startDate=awardRs("startDate")
		endDate=awardRs("endDate")
		prizeStr=awardRs("PrizeIDS")
		if not isnull(prizeStr) then
			if prizeStr<>"" then
				prizeArray=split(DelHeadAndEndDot(prizeStr),",")
			end if
		end if
	end if
elseif Request("Act")="add" then
if not MF_Check_Pop_TF("ME027") then Err_Show 
 startDate=datevalue(Now())
 needPoint=0
end if
Dim sRootDir
If G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if	

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
	if NoSqlHack(Request("act"))="edit" then
		Response.Write("<form name='AwardPanel' id='AwardPanel' method='post' action='awardAction.asp?act=editaction&awardid="&NoSqlHack(Request("awardid"))&"'>")
	else
		Response.Write("<form name='AwardPanel' id='AwardPanel' method='post' action='awardAction.asp?act=addaction'>")
	end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <tr class="hback">
    <td align="right" class="xingmu" colspan="2"><div align="left">奖品项目设置&nbsp;&nbsp;| &nbsp;<a href="#" onClick="history.back()">后退</a></div></td></tr> 
        <tr class="hback"> 
          <td align="right">主题名称：</td> 
          <td> <input name="awardName" type="text" id="awardName" value="<%=awardName%>" size="50" maxlength="20"/ >
          <font color="#FF0000">*</font><span id="awardName_Alert"></span></td> 
        </tr> 
      <tr class="hback">
    <td align="right">主题图片：</td>
    <td>
	<input type="text" name="awardPic" id="awardPic" value="<%=awardPic%>" size="50" maxlength="120">
	<button onClick="javascript:OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,$('awardPic'));">选择图片</button>	
    <span id="UpfileSize_Alert"></span>
		</td>
  </tr>
		
      <tr class="hback"> 
          <td align="right">开始日期：</td> 
          <td><input name="startDate" type="text" id="startDate" value="<%=startDate%>" size="50" readonly="true"><button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.AwardPanel.startDate);document.AwardPanel.startDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="startDate_Alert"></span></td> 
  </tr>
        <tr class="hback">
          <td align="right">截止日期：</td>
          <td><input name="EndDate" type="text" id="EndDate" value="<%=endDate%>" size="50" readonly="true"><button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.AwardPanel.EndDate);document.AwardPanel.EndDate.focus();">选择时间</button><font color="#FF0000">*</font><span id="EndDate_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">奖品等级数：</td>
          <td><input name="PrizeGradeNum" type="text" id="PrizeGradeNum"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')" value="<%if request("act")="edit" then Response.Write(ubound(prizeArray)+1) else Response.Write("1")%>" size="50" maxlength="3"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
          <button onClick="setPrizeGradeNum(PrizeGradeNum.value)">设置</button></td>
        </tr>
</table> 
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="table"> 
<tr class='hback'>
	<td id="PrizeContainer" colspan="2">
	<%
		if Request("act")="edit" then
			Response.Write("<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' >")
			for ArrayIndex=0 to Ubound(prizeArray)
				Set prizeRs=User_Conn.Execute("Select PrizeName,PrizePic,PrizeNum,NeedPoint from FS_ME_Prize where PrizeID="&NoSqlHack(prizeArray(ArrayIndex)))
				if not prizeRs.eof then
					Response.Write("<tr class='hback'>")
					Response.Write("<td align='right' width='8%'>"&(ArrayIndex+1)&"等奖：</td>")
					Response.Write("<td><input name='Prize_"&(ArrayIndex+1)&"_name' type='text' id='Prize_"&(ArrayIndex+1)&"_name' size='25' value='"&prizeRs("PrizeName")&"'>")
					Response.Write("    |需要积分：<input name='needPoint_"&(ArrayIndex+1)&"' type='text' id='needPoint_"&(ArrayIndex+1)&"' size='5' value='"&prizeRs("NeedPoint")&"'  onKeyUp=""if(isNaN(value)||event.keyCode==32)execCommand('undo')""  onafterpaste=""if(isNaN(value)||event.keyCode==32)execCommand('undo')"">")
					Response.Write("    |图片：<input type=""text"" name=""prize_"&(ArrayIndex+1)&"_pic"" id=""prize_"&(ArrayIndex+1)&"_pic""   value="""&prizeRs("PrizePic")&""" style=""width:30%"" maxlength=""120""><button onClick=""javascript:OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"&f_UserNumber="&session("FS_UserNumber")&"',500,320,window,'document.AwardPanel.prize_"&(ArrayIndex+1)&"_pic');"">选择图片</button>")
					Response.Write("    |数量：<input name='prize_"&(ArrayIndex+1)&"_number' type='text' id='prize_"&(ArrayIndex+1)&"_number' size='5' value='"&prizeRs("PrizeNum")&"'  onKeyUp=""if(isNaN(value)||event.keyCode==32)execCommand('undo')""  onafterpaste=""if(isNaN(value)||event.keyCode==32)execCommand('undo')""></td></tr>")
				End if
			next
			Response.Write("</table>")
		end if
	%>
	</td>
</tr>
<tr class='hback'>
<td width="20">&nbsp;</td>
<td><div align="left"><input type="button" name="ADDEditAwardButton" onClick="ADDEditAward()" value="保存"> | <input type="reset" name="" value="重置"></div></td>
</tr>
</table>
</form>
</body>
<%
if Request("Act")="eidt" then
	awardRs.close
	set awardRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
	Set prizeRs=nothing
end if
%>
<script language="JavaScript" type="text/JavaScript">
if("<%=Request("act")%>"!="edit")
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
		PrizeGradeNum = PrizeGradeNum +"<tr class='hback'><td align='right' width='15%'>"+i+"等奖：</td><td><input name='Prize_"+i+"_name' type='text' id='Prize_"+i+"_name+' size='25' value=''>|需要积分：<input name='needPoint_"+i+"' type='text' id='needPoint_"+i+"' size='5' value='' onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>    |图片：<input name='prize_"+i+"_pic' type='text' id='prize_"+i+"_pic' style='width:30%'><button onClick=\"javascript:OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>&f_UserNumber=<%=session("FS_UserNumber")%>',500,320,window,$('prize_"+i+"_pic'));\">选择图片</button>    |数量：<input name='prize_"+i+"_number' type='text' id='prize_"+i+"_number' size='5' value='' onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');></td></tr>";
	}
	document.getElementById("PrizeContainer").innerHTML="<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' >"+PrizeGradeNum+"</table>";
}
function ADDEditAward()
{
	var flag1=isEmpty("awardName","awardName_Alert");
	var flag2=isEmpty("startDate","startDate_Alert");
	var flag3=isEmpty("EndDate","EndDate_Alert");
	if(flag1&&flag2&&flag3)
	document.AwardPanel.submit();
}
</script>
</html>






