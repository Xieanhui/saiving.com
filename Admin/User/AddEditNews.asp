<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,OperateNewsRs,newsid,title,content,grouptype,groupid,groupIndex,GroupInfoRs,newspoint,isLock
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("ME021") then Err_Show 'Ȩ���ж�
Dim str_CurrPath,sRootDir
MF_User_Conn
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
'************************************Update
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

if Request.QueryString("Act")="edit" then
	newsid=Request.QueryString("newsid")
	Set OperateNewsRs=server.CreateObject(G_FS_RS)
	OperateNewsRs.open "select title,content,addtime,groupid,newspoint,isLock from FS_ME_News where newsid="&NoSqlHack(newsid),User_Conn,1,3
	if not OperateNewsRs.eof then
		title=OperateNewsRs("title")
		content=OperateNewsRs("content")
		groupid=OperateNewsRs("groupid")
		newspoint=OperateNewsRs("newspoint")
		isLock=OperateNewsRs("isLock")
		If Not IsNull(groupid) And Not groupid="" then
			Set GroupInfoRs=User_Conn.execute("select groupid,groupname,grouptype from FS_ME_Group where groupid="&NoSqlHack(groupid))
			if not GroupInfoRs.eof then
				grouptype=GroupInfoRs("GroupType")
			end If
			groupInfoRs.close
			Set groupInfoRs=nothing
		End if
		OperateNewsRs.close
	end if
elseif Request.QueryString("Act")="EditNews" then
	if isNumeric(Request.Form("GroupIndex")) then
		GroupIndex=NoSqlHack(Request.Form("GroupIndex"))
	else
		GroupIndex=0
	end if
	if GroupIndex="" then GroupIndex=0
	User_Conn.Execute("Update FS_ME_News set Title='"&NoSqlHack(Request.Form("title"))&"',content='"&NoSqlHack(Request.Form("Content"))&"',groupid='"&NoSqlHack(GroupIndex)&"',newspoint='"&NoSqlHack(Request.Form("newspoint"))&"',isLock='"&NoSqlHack(Request.Form("isLock"))&"' where newsid="&NoSqlHack(Request.QueryString("newsid"))&"")
	if err.number=0 then 
		Response.Redirect("../success.asp?ErrCodes=<li>�޸ĳɹ�</li>&ErrorURL=./user/news_manage.asp")
	else
		Response.Redirect("../error.asp?ErrCodes=<li>"&err.description&"</li>")
	end if
elseif Request.QueryString("Act")="AddNews" then
	if isNumeric(Request.Form("GroupIndex")) then
		GroupIndex=NoSqlHack(Request.Form("GroupIndex"))
	else
		GroupIndex=0
	end if
	if GroupIndex="" then GroupIndex=0
	Set OperateNewsRs=server.CreateObject(G_FS_RS)
	OperateNewsRs.open "select title,content,addtime,groupid,newspoint,isLock from FS_ME_News",User_Conn,1,3
	OperateNewsRs.addNew
	OperateNewsRs("title")=NoSqlHack(Request.Form("title"))
	OperateNewsRs("content")=NoSqlHack(Request.Form("content"))
	OperateNewsRs("addtime")=now()
	OperateNewsRs("groupid")=NoSqlHack(GroupIndex)
	
	OperateNewsRs("newspoint")=NoSqlHack(Request.Form("newspoint"))
	OperateNewsRs("islock")=NoSqlHack(Request.Form("islock"))
	OperateNewsRs.update
	OperateNewsRs.close
	if err.number=0 then 
		Response.Redirect("../success.asp?ErrCodes=<li>�����ɹ�</li>&ErrorURL=./user/news_manage.asp")
	else
		Response.Redirect("../error.asp?ErrCodes=<li>"&err.description&"</li>")
	end if
end if
function getDebateParentNum(ID)
	if ID="" then exit function
	Dim F_DebateID,F_Rs,FS_count
	FS_count=1
	Set F_Rs=User_Conn.Execute("Select DebateID,title,ParentID from FS_ME_GroupDebate where DebateID="&CintStr(ID))
	if not F_Rs.eof then
		FS_count=FS_count+getDebateParentNum(F_Rs("ParentID"))
	end if
	F_Rs.close
	Set F_Rs=nothing
	getDebateParentNum=FS_count
End function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td class="xingmu">�������</td>
  </tr>
  <tr class="hback"> 
    <td class="hback"><a href="News_manage.asp">��ҳ</a> | <a href="AddEditNews.asp?Act=add">��������</a> | <a href="javascript:history.back();">������һ��</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <%
			if Request.QueryString("Act")="edit" then
				Response.Write("<form action='?Act=EditNews&newsid="&CintStr(newsid)&"' method='post' name='operateNews' id='operateNews'>")
			elseif Request.QueryString("Act")="add" then
				Response.Write("<form action='?Act=AddNews' method='post' name='operateNews' id='operateNews'>")
			end if
		%>
  <tr class="hback"> 
    <td width="107" align="right">������⣺</td>
    <td width="837"> <input name="Title" type="text" id="Title" value="<%=title%>" size="50" maxlength="50" /> 
    <font color="#FF0000">*</font><span id="Title_Alert"></span></td>
  </tr>
  <tr class="hback"> 
    <td align="right">�������ݣ� </td>
    <td><!--�༭����ʼ-->
		<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
		<input type="hidden" name="Content" value="<% = HandleEditorContent(Content) %>">
        <!--�༭������-->
      <font color="#FF0000">*</font><span id="Content_Alert"></span></td>
  </tr>
  <tr class="hback"> 
    <td align="right">�鿴Ȩ�ޣ�</td>
    <td> <select name="GType" id="GType" onChange="getFormInfo(this)">
        <option value="all" <%if grouptype<>1 and grouptype<>0 then Response.Write("selected")%>>���л�Ա��</option>
        <option value="1" <%if grouptype=1 then Response.Write("selected")%>>���˻�Ա��</option>
        <option value="0" <%if grouptype=0 then Response.Write("selected")%>>��ҵ��Ա��</option>
      </select> &nbsp; <span id="GroupIndexContent"> 
      <%
		if Request.QueryString("act")="edit" then
			Response.Write(" |  ��Ա�飺<select name='GroupIndex' id='GroupIndex'>"&Chr(10)&Chr(13))
			if grouptype=1 then
				Set GroupInfoRs=User_Conn.execute("Select GroupID,GroupName from FS_ME_Group where GroupType=1")
				Response.Write("<option value='user'>���и��˻�Ա��</option>")
			elseif grouptype=0 then
				Set GroupInfoRs=User_Conn.execute("Select GroupID,GroupName from FS_ME_Group where GroupType=0")
				Response.Write("<option value='corp'>������ҵ��Ա��</option>")
			end if
			while not GroupInfoRs.eof
				if GroupInfoRs("GroupID")=Groupid then
					Response.Write("<option value='"&GroupInfoRs("GroupID")&"' selected>"&GroupInfoRs("GroupName")&"</option>")
				else
					Response.Write("<option value='"&GroupInfoRs("GroupID")&"'>"&GroupInfoRs("GroupName")&"</option>")
				end if
				GroupInfoRs.movenext
			wend
		end if
	%>
      </span> </td>
  </tr>
  <tr class="hback"> 
    <td align="right">�������ƣ�</td>
    <td><input name="NewsPoint" type="text" id="NewsPoint"  value="<%=newspoint%>" size="50"/>
      �����ִ��ڶ��ٿ��Բ鿴��<span id="NewsPoint_Alert"></span></td>
  </tr>
  <tr class="hback"> 
    <td align="right">�Ƿ�������</td>
    <td> <%
		  	if Request.QueryString("act")="edit" then
				if isLock=1 then
					Response.Write("<input type='radio' name='isLock' value=1 checked>��"&Chr(10)&Chr(13))
					Response.Write("<input type='radio' name='isLock' value=0 >��"&Chr(10)&Chr(13))
				elseif isLock=0 then
					Response.Write("<input type='radio' name='isLock' value=1 >��"&Chr(10)&Chr(13))
					Response.Write("<input type='radio' name='isLock' value=0 checked>��"&Chr(10)&Chr(13))
				end if
			else
				Response.Write("<input type='radio' name='isLock' value=1 >��"&Chr(10)&Chr(13))
				Response.Write("<input type='radio' name='isLock' value=0 checked>��"&Chr(10)&Chr(13))
			end if
		  %> </td>
  </tr>
  <tr class="hback"> 
    <td align="right">&nbsp;</td>
    <td><input type="Button" name="AddNewsButton" value=" ���� " onClick="OperateNewsSubmit(this.form)"/> 
      <input type="reset" name="Submit2" value=" ���� " /></td>
  </tr></form></tr>
</table> 
</body>
<%
if Request.QueryString("Act")="edit" then
	GroupInfoRs.close
	set GroupInfoRs=nothing
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
var request=true;
var result;
var ParamArray;
try
{
	request=new XMLHttpRequest();
}catch(trymicrosoft)
{
try
{
	request=new ActiveXObject("Msxml2.XMLHTTP")
}catch(othermicrosoft)
{
try
{
	request=new ActiveXObject("Microsoft.XMLHTTP")
}catch(filed)
{
	request=false;
}
}
}
if(!request) alert("Error initializing XMLHttpRequest!");
function getFormInfo(Obj)
{
	var typeID=Obj.value;
	if(isNaN(typeID))
	{
		document.getElementById("GroupIndexContent").innerHTML="";
		return ;
	}
	var url="getUserGroup.asp?page=News&id="+typeID+"&r="+Math.random();//����url
	request.open("GET",url,true);//��������
	request.onreadystatechange = getFormInfoResult;
	request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���null
}
function getFormInfoResult()//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			document.getElementById("GroupIndexContent").innerHTML="|&nbsp;&nbsp;��Ա�飺"+result;//����������ʵ�ڿͻ���
		}
	}
}
function OperateNewsSubmit(FormObj)
{	
	var flag1=isEmpty('Title','Title_Alert');
	//var flag2=isEmpty('Content','Content_Alert');
	var flag3=isNumber("NewsPoint",'NewsPoint_Alert','������ʹ��������',true)
	//if(flag1&&flag2&&flag3)
	if(flag1&&flag3)
	{
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
		FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
		FormObj.submit();
	}
}
</script>
</html>