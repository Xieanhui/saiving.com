<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
'on error resume next
Dim Conn,User_Conn,AddGroupRs,if_repage_RS
Dim GroupName,AllowUpFile,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,GroupType,CorpTemplet,LimitInfoNum,GroupDebateNum,GroupDebateNum_Array,JuniorDomain,KeywordsNumber,Ishtml,BcardNumber,Templetwatermark

dim sRootDir,str_CurrPath
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
'************************************Update
if Request.QueryString("Act")="addGroup" then
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_GUser") then Err_Show 
	if not MF_Check_Pop_TF("ME034") then Err_Show 

	Set AddGroupRs=server.CreateObject(G_FS_RS)
	AddGroupRs.open "select GroupName,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,GroupType,CorpTemplet,LimitInfoNum,GroupDebateNum,JuniorDomain,KeywordsNumber,ProductDiscount,isHtml,BcardNumber,Templetwatermark From FS_ME_Group",User_Conn,1,3
	AddGroupRs.addNew
	AddGroupRs("GroupName")=NoSqlHack(Request.Form("GroupName"))
	AddGroupRs("UpfileNum")=NoSqlHack(Request.Form("UpfileNum"))
	if Request.Form("UpfileSize")<1024 then
		Response.Redirect("../error.asp?ErrCodes=<li>�û��ϴ���վ�Ŀռ����Ϊ����1M</li>")
		Response.End()
	else
		AddGroupRs("UpfileSize")=NoSqlHack(Request.Form("UpfileSize"))
	end if
	AddGroupRs("GroupDate")=NoSqlHack(Request.Form("GroupDate"))
	AddGroupRs("GroupPoint")=NoSqlHack(Request.Form("GroupPoint"))
	AddGroupRs("GroupMoney")=NoSqlHack(Request.Form("GroupMoney"))
	AddGroupRs("GroupType")=NoSqlHack(Request.Form("GroupType"))
	AddGroupRs("ProductDiscount")=NoSqlHack(Request.Form("ProductDiscount"))
	'AddGroupRs("CorpTemplet")=Request.Form("CorpTemplet")
	AddGroupRs("LimitInfoNum")=NoSqlHack(Request.Form("LimitInfoNum"))
	AddGroupRs("GroupDebateNum")=NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))
	AddGroupRs("JuniorDomain")=NoSqlHack(Request.Form("JuniorDomain"))
	'AddGroupRs("KeywordsNumber")=Request.Form("KeywordsNumber")
	'AddGroupRs("isHtml")=Request.Form("isHtml")
	'AddGroupRs("BcardNumber")=Request.Form("BcardNumber")
	'AddGroupRs("Templetwatermark")=Request.Form("Templetwatermark")
	Set if_repage_RS=User_Conn.execute("Select GroupName from FS_ME_Group where GroupName='"&NoSqlHack(Request.Form("GroupName"))&"'")
	if not if_repage_RS.eof then
		Response.Redirect("../error.asp?ErrCodes=<li>���������ظ�</li>")
		Response.End()
	End if
	AddGroupRs.update
	if err.number=0 then 
		Response.Redirect("../success.asp?ErrCodes=<li>�����ɹ�</li>&ErrorURL=user/Group_manage.asp")
		Response.End()
	else
		Response.Redirect("../error.asp?ErrCodes=<li>"&err.description&"</li>&ErrorURL="&request.ServerVariables("HTTP_REFERER"))
		Response.End()
	end if
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <tr class="hback"> 
    <td align="right" class="xingmu" colspan="2"><div align="left">����û���&nbsp;&nbsp;<a href="javascript:history.back();">����</a></div></td></tr> 
    <form action="?Act=addGroup" method="post" name="AddGroup" id="AddGroup">  
        <tr class="hback"> 
          <td width="180" align="right"><div align="right">�����ƣ�</div></td> 
          <td width="764"> <input name="GroupName" type="text" id="GroupName" size="50" maxlength="20" />
          <font color="#FF0000">*</font><span id="GroupName_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">�ļ��������ƣ� </td>
    <td><input name="UpfileNum" type="text" id="UpfileNum"  value="0" size="50"></td>
  </tr>
<tr class="hback">
    <td align="right">�ļ���С���ƣ�</td>
    <td><input name="UpfileSize" type="text" id="UpfileSize" value="2048" size="50">
    <span id="UpfileSize_Alert">k<span id="UpfileNum_Alert"> ��Ա�ռ�ռ��</span></span></td>
  </tr>
  <tr class="hback">
  <td align="right">��Ա�ۿۣ�</td>
  <td><input name="ProductDiscount" type="text" id="ProductDiscount" value="1" size="50">
  ���磺����ۣ��ô�����д0.8 <span id="ProductDiscount_Alert"></span></td>
  </tr>
<tr class="hback"> 
                <td align="right">����Ч���ޣ�</td> 
                <td><input name="GroupDate" type="text" id="GroupDate"  value="0" size="50"/> 
                �� <span id="GroupDate_Alert"></span></td> 
        </tr> 
      <tr class="hback"> 
          <td align="right">����������֣�</td> 
          <td><input name="GroupPoint" type="text" id="GroupPoint" value="0" size="50"/>
          <span id="GroupPoint_Alert"></span></td> 
        </tr>
        <tr class="hback">
          <td align="right">���������ң�</td>
          <td><input name="GroupMoney" type="text" id="GroupMoney" value="0" size="50"/>
          <span id="GroupMoney_Alert"></span></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">�����ͣ�</td> 
          <td><label>
            <input name="GroupType" type="radio" value="1" checked > 
            ���˻�Ա��</label>
            <label>
            <input type="radio" name="GroupType" value="0">
��ҵ��Ա��(��ҵ��Ա����ʱӦ�����Ժ���չʹ�á���Ŀǰ����������ʹ��)</label></td> 
        </tr> 
      <tr class="hback"> 
          <td align="right">��Ϣ�����������ޣ�</td> 
          <td><input name="LimitInfoNum" type="text" id="LimitInfoNum" value="10" size="50"/>
          <span id="LimitInfoNum_Alert"></span></td> 
        </tr>
        <tr class="hback" style="display:none">
          
      <td align="right">��Աģ���ַ��</td>
          <td><input name="CorpTemplet" type="text" size="50">
		  <input name="Submit5" type="button" id="selCorpTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.AddGroup.CorpTemplet);document.AddGroup.CorpTemplet.focus();"> 
		  <span id="CorpTemplet_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">��Ⱥ������</td>
          <td>��Ⱥ���������
            <input name="GroupDebateNum_1" type="text" id="GroupDebateNum_1" value="0" size="15"> 
          &nbsp;��Ⱥ�������
          <input name="GroupDebateNum_2" type="text" id="GroupDebateNum_2" value="0" size="15">
          <span id="GroupDebateNum_Alert1"></span>&nbsp; <span id="GroupDebateNum_Alert2"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">��ͨ����������</td>
          <td><p>
            <label>
            <input type="radio" name="JuniorDomain" value="1" >
  ��</label>
            <label>
            <input name="JuniorDomain" type="radio" value="0" checked>
  ��</label>
            <br>
          </p></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">��Ϣ�ؼ��ָ�����</td>
          <td><input name="KeywordsNumber" type="text" id="KeywordsNumber"  value="0" size="50"/>
          <span id="KeywordsNumber_Alert"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">���ɾ�̬�ļ���</td>
          <td><label>
            <input type="radio" name="Ishtml" value="1" >
��</label>
            <label>
            <input name="Ishtml" type="radio" value="0" checked >
��</label></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">��Ƭ�ղظ������ƣ�</td>
          <td><input name="BcardNumber" type="text" id="BcardNumber" value="0" size="50"/>
          <span id="BcardNumber_Alert"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">��ͨˮӡ��</td>
          <td><label>
            <input type="radio" name="Templetwatermark" value="1" >
��</label>
            <label>
            <input name="Templetwatermark" type="radio" value="0" checked >
��</label></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">&nbsp;</td> 
          <td><input type="Button" name="AddGroupButton" value=" ���� " onClick="AddGroupSubmit()"/> 
            <input type="reset" name="Submit2" value=" ���� " /></td> 
        </tr> 
    </form> 
  </tr> 
</table> 
</body>
<%
if Request.QueryString("Act")="addGroup" then
	AddGroupRs.close
	set AddGroupRs=nothing
	if_repage_RS.close
	Set if_repage_RS=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
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
function AddGroupSubmit()
{
	var flag1=isEmpty('GroupName','GroupName_Alert')	
	var flag2=isNumber('UpfileNum','UpfileNum_Alert','�ļ�����ӦΪ������',true)
	var flag3=isNumber('UpfileSize','UpfileSize_Alert','�ļ���СӦΪ������',true)
	var flag4=isNumber('GroupDate','GroupDate_Alert','��Ч����ӦΪ������',true)
	var flag5=isNumber('GroupPoint','GroupPoint_Alert','����ӦΪ������',true)
	var flag6=isNumber('GroupMoney','GroupMoney_Alert','���ӦΪ������',true)
	var flag7=isNumber('LimitInfoNum','CorpTemplet_Alert','���ӦΪ������',true)
	var flag8=isNumber('GroupDebateNum_1','GroupDebateNum_Alert1','��Ⱥ�������ӦΪ������',true)
	var flag9=isNumber('GroupDebateNum_2','GroupDebateNum_Alert2','��Ⱥ�������ӦΪ������',true)
	//var flag10=isNumber('KeywordsNumber','KeywordsNumber_Alert','�ؼ��ָ���ӦΪ������',true)
	//var flag11=isNumber('BcardNumber','BcardNumber_Alert','��Ƭ�ղظ���ӦΪ������',true)	
	//var flag12=isEmpty('CorpTemplet','CorpTemplet_Alert')
	//var flag13=isNumber('ProductDiscount','ProductDiscount_Alert',false)
	if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8&&flag9)
		document.AddGroup.submit();
}
</script>
</html>






