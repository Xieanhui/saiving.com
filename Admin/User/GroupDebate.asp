<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,DebateRs,listnum,page,j,i,nn,n,pagename,DebateID
MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set DebateRs=Server.CreateObject(G_FS_RS)
DebateID=Request("DebateID")
if DebateID<>"" then
	DebateRs.open "select DebateID,title,addtime,ParentID,classid,AddTime,isLock from FS_ME_GroupDebate where parentID="&DebateID,User_Conn,1,1
else
	DebateRs.open "select DebateID,title,addtime,ParentID,classid,AddTime,isLock from FS_ME_GroupDebate where parentID=0",User_Conn,1,1
end if
pagename="GroupDebate.asp?"
if DebateRs.eof and DebateRs.bof Then

else
	listnum=20
	DebateRs.pagesize=listnum
	page=Request("page")
	if (page-DebateRs.pagecount) > 0 then
		page=DebateRs.pagecount
	elseif page = "" or page < 1 then
		page = 1
	end if
	DebateRs.absolutepage=page
	'��ŵ�ʵ��
	j=DebateRs.recordcount
	j=j-(page-1)*listnum
	i=0
	nn=request("page")
	if nn="" then
		n=0
	else
		nn=nn-1
		n=listnum*nn
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
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
function changeLock(Obj1,Obj2)
{
	var url="GroupDebateAction.asp?DebateID="+Obj1+"&value="+Obj2+"&r="+Math.random();//����url
	request.open("GET",url,true);//��������
	request.onreadystatechange = getResult;
	request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���null
}
function getResult()//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			alert("�޸ĳɹ�")

		}
	}
}
function AddGroupDebateSubmit()
{	
	location='AddEditDebate.asp?Act=add'
}
function MyBack(TF)
{
	if(TF==1)
	location="GroupDebate.asp"
	else
	history.back()
}
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;"> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <form name="GroupDebate" id="GroupDebate" method="post" action="GroupDebateAction.asp?act=Del">
  <tr class="hback"> 
	<tr class="xingmu"> 
	<td width="30%" align="center"><div align="left">��Ⱥ����</div></td> 
	<td width="20%" align="center">&nbsp;</td>
	<td width="15%" align="center">&nbsp;</td>
	<td width="15%" align="center">&nbsp;</td> 
	<td>&nbsp;</td> 
    <td width="10%" align="center">&nbsp;</td>
    <td width="10%" align="center"><input type="Button" name="AddNewsSubmit" value="�½���Ⱥ" onClick="AddGroupDebateSubmit()"></td>
  </tr>
        <tr class="xingmu"> 
          <td width="30%" align="center">����</td> 
          <td width="20%" align="center">����ʱ��</td>
		  <td width="20%" align="center">��������</td>
          <td width="15%" align="center">����鿴��</td>
          <td >����</td> 
          <td align="center">����</td>
          <td width="10%" align="center"><input type="checkbox" name="checkAll"></td>
        </tr>
		<%
			do while not DebateRs.eof and i<listnum
				n=n+1
				Response.Write("<tr class='hback'>")
				Response.Write("<td align='center'><a href='GroupDebate.asp?DebateID="&DebateRs("DebateID")&"'>"&DebateRs("title")&"</a></td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&DebateRs("addtime")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&DebateRs("ParentID")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&DebateRs("classid")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>")
				if DebateRs("isLock")=1 then
					Response.Write("<input type='Radio' name='"&DebateRs("DebateID")&"' value=1 onclick='changeLock(this.name,this.value)'checked>��&nbsp;&nbsp;&nbsp;&nbsp;"&Chr(10)&Chr(13))
					Response.Write("<input type='Radio' name='"&DebateRs("DebateID")&"' value=0 onclick='changeLock(this.name,this.value)'>��"&Chr(10)&Chr(13))
				elseif DebateRs("isLock")=0 then
					Response.Write("<input type='Radio' name='"&DebateRs("DebateID")&"' value=1 onclick='changeLock(this.name,this.value)')>��&nbsp;&nbsp;&nbsp;&nbsp;"&Chr(10)&Chr(13))
					Response.Write("<input type='Radio' name='"&DebateRs("DebateID")&"' value=0 onclick='changeLock(this.name,this.value)' checked>��"&Chr(10)&Chr(13))
				end if
				Response.Write("</td>")
				Response.Write("<td align='center'><a href='AddEditDebate.asp?act=edit&DebateID="&DebateRs("DebateID")&"'>����<a></td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'><input type='checkbox' name='Debatecheck' value='"&DebateRs("DebateID")&"'></td>"&Chr(10)&Chr(13))
				Response.Write("</tr>")
				DebateRs.movenext 
				i=i+1 
				j=j-1
			loop
		%>
	<tr class="xingmu"> 
		<td width="30%" align="center"></td> 
		<td width="20%" align="center"></td>
		<td width="20%" align="center"></td>
		<td width="15%" align="center"></td>
		<td><input name="backTop" type="button" id="backTop" value="���ص�һ��" onClick="MyBack(1)"></td> 
		<td align="center">
		<%
			if Request("Debate")<>"" then
				
			End if
		%>
		</td>
		<td width="10%" align="center"><input type="submit" name="Submit" value="ɾ��"></td>
	</tr>
	<tr>
	<td align="right" colspan="7">
	<%=DebateRs.recordcount%> ����Ϣ&nbsp;&nbsp;<%=listnum%> ����Ϣ/ҳ&nbsp;&nbsp;�� <%=DebateRs.pagecount%> ҳ 
	<% if page=1 then %>
	<%else%>
	<a href=<%=pagename%>><strong>|<<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><strong><<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><b>[<%=page-1%>]</b></a>&nbsp; 
	<%end if%>
	<% if DebateRs.pagecount=1 then %>
	<%else%>
	<b>[<%=page%>]</b>
	<%end if%>
	<% if DebateRs.pagecount-page <> 0 then %>
	<a href=<%=pagename%>page=<%=page+1%>><b>[<%=page+1%>]</b></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page+1%>><strong>>></strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=DebateRs.pagecount%>><strong>>>|</strong></a>&nbsp; 
	<%end if%>��
	</td>
	</tr>
	</form> 
</table> 
</body>
<%
if Request("Act")="addGroup" then
	AddGroupRs.close
	set AddGroupRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
end if
%>
</html>






