<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
on error resume next
Dim Conn,User_Conn,NewsRs,listnum,page,j,i,nn,n,pagename
MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set NewsRs=Server.CreateObject(G_FS_RS)
NewsRs.open "select NewsID,title,addtime,groupid,newspoint,isLock from FS_ME_News order by addtime desc",User_Conn,1,3
pagename="news.asp?"
if NewsRs.eof and NewsRs.bof Then

else
	listnum=20
	NewsRs.pagesize=listnum
	page=NoSqlHack(Request("page"))
	if (page-NewsRs.pagecount) > 0 then
		page=NewsRs.pagecount
	elseif page = "" or page < 1 then
		page = 1
	end if
	NewsRs.absolutepage=page
	'��ŵ�ʵ��
	j=NewsRs.recordcount
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
	var url="NewsAction.asp?newsid="+Obj1+"&value="+Obj2+"&r="+Math.random();//����url
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
function AddNewsSubmit()
{	
	location='AddEditNews.asp?Act=add'
}
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;"> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
  <tr class="hback"> 
	<tr class="xingmu"> 
	<td width="30%" align="center"><div align="left">�������</div></td> 
	<td width="20%" align="center">&nbsp;</td>
	<td width="15%" align="center">&nbsp;</td>
	<td width="15%" align="center">&nbsp;</td> 
	<td width="20%" align="center"><input type="Button" name="AddNewsSubmit" value="������Ϣ" onClick="AddNewsSubmit()"></td> 
  </tr>
	</tr> 
    <form action="?Act=addNews" method="post" name="addNewsForm" id="addNewsForm">  
        <tr class="xingmu"> 
          <td width="30%" align="center">����</td> 
          <td width="20%" align="center">����ʱ��</td>
          <td width="15%" align="center">����鿴��</td>
          <td width="15%" align="center">��������</td> 
		  <td width="20%" align="center">����</td> 
        </tr>
		<%
			do while not NewsRs.eof and i<listnum
				n=n+1
				Response.Write("<tr class='hback'>")
				Response.Write("<td align='left'><a href='AddEditNews.asp?act=edit&newsID="&NewsRS("newsid")&"'>"&NewsRs("title")&"</a></td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("addtime")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("GroupID")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("NewsPoint")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>")
				if NewsRs("isLock")=1 then
					Response.Write("<input type='Radio' name='"&NewsRs("NewsID")&"' value=1 onclick='changeLock(this.name,this.value)'checked>��&nbsp;&nbsp;&nbsp;&nbsp;"&Chr(10)&Chr(13))
					Response.Write("<input type='Radio' name='"&NewsRs("NewsID")&"' value=0 onclick='changeLock(this.name,this.value)'>��"&Chr(10)&Chr(13))
				elseif NewsRs("isLock")=0 then
					Response.Write("<input type='Radio' name='"&NewsRs("NewsID")&"' value=1 onclick='changeLock(this.name,this.value)')>��&nbsp;&nbsp;&nbsp;&nbsp;"&Chr(10)&Chr(13))
					Response.Write("<input type='Radio' name='"&NewsRs("NewsID")&"' value=0 onclick='changeLock(this.name,this.value)' checked>��"&Chr(10)&Chr(13))
				end if
				Response.Write("</td>")
				Response.Write("</tr>")
				NewsRs.movenext 
				i=i+1 
				j=j-1
			loop
		%> 
    </form>
	<tr>
	<td align="right" colspan="5">
	<%=NewsRs.recordcount%> ����Ϣ&nbsp;&nbsp;<%=listnum%> ����Ϣ/ҳ&nbsp;&nbsp;�� <%=NewsRs.pagecount%> ҳ 
	<% if page=1 then %>
	<%else%>
	<a href=<%=pagename%>><strong>|<<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><strong><<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><b>[<%=page-1%>]</b></a>&nbsp; 
	<%end if%>
	<% if NewsRs.pagecount=1 then %>
	<%else%>
	<b>[<%=page%>]</b>
	<%end if%>
	<% if NewsRs.pagecount-page <> 0 then %>
	<a href=<%=pagename%>page=<%=page+1%>><b>[<%=page+1%>]</b></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page+1%>><strong>>></strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=NewsRs.pagecount%>><strong>>>|</strong></a>&nbsp; 
	<%end if%>��
	</td>
	</tr> 
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






