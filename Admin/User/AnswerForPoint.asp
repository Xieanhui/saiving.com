<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,AnswerForPoint,listnum,page,j,i,nn,n,pagename,EndDate
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_award") then Err_Show 

Set AnswerForPoint=Server.CreateObject(G_FS_RS)
AnswerForPoint.open "select AID,ATopic,NeedPoint,PrizePoint,ADesc,AStartDate,AEndDate from FS_ME_AnswerForPoint order by AendDate asc",User_Conn,1,1
pagename="AnswerForPoint.asp?"
if AnswerForPoint.eof and AnswerForPoint.bof Then

else
	listnum=10
	AnswerForPoint.pagesize=listnum
	page=CintStr(Request.QueryString("page"))
	if (page-AnswerForPoint.pagecount) > 0 then
		page=AnswerForPoint.pagecount
	elseif page = "" or page < 1 then
		page = 1
	end if
	AnswerForPoint.absolutepage=page
	'��ŵ�ʵ��
	j=AnswerForPoint.recordcount
	j=j-(page-1)*listnum
	i=0
	nn=Request.QueryString("page")
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
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
<form action="awardAction.asp?act=deleteAFPointaction" method="post" name="PrizeForm" id="PrizeForm">  
  <tr>
  <td colspan="6" class="xingmu">�齱����</td>
  </tr>
  <tr class="hback"> 
	<td width="30%" height="20" colspan="6" align="center"><div align="left"><a href="award.asp" target="_self">&nbsp;���ֳ齱</a>&nbsp;|&nbsp;<a href="changePrize.asp">���ֶһ�</a> | ���־���<a href="AnswerForPoint.asp"> |</a> <a href="#" onClick="history.back()">����</a></div></td> 
	</tr>
        <tr class="xingmu"> 
          <td width="40%" align="center">������Ŀ</td> 
          <td width="12%" align="center">��Ҫ����</td>
          <td width="12%" align="center">��������</td>
          <td width="15%" align="center">״̬</td>
		  <td width="15%" align="center">��ֹʱ��</td>
		  <td width="10%" align="center"><input type="checkbox" name="Delete_CheckAll" value="all" onClick="CheckAll(this,'DeleteAFPoint')"></td> 
        </tr>
		<%    
			do while not AnswerForPoint.eof and i<listnum
				Response.Write("<tr class='hback'>"&chr(10)&chr(13)) 
				Response.Write("<td align='center'><a href='AFPoint_AddEdit.asp?act=editAFPoint&AID="&AnswerForPoint("AID")&"'>"&AnswerForPoint("ATopic")&"</a></td>"&chr(10)&chr(13))
				Response.Write("<td align='center' width='20%'>"&AnswerForPoint("NeedPoint")&"</td>")
				Response.Write("<td align='center'>"&AnswerForPoint("PrizePoint")&"</td>")
				EndDate=AnswerForPoint("AEndDate")
				if EndDate<Now() then
					Response.Write("<td align='center'>�ѹ���</td>")
				else
					Response.Write("<td align='center'>δ����</td>")
				end if
				Response.Write("<td align='center'>"&EndDate&"</td>")
				Response.Write("<td align='center'><input type='checkbox' name='DeleteAFPoint' value='"&AnswerForPoint("AID")&"'></td>")
				Response.Write("</tr>")
				AnswerForPoint.movenext
				i=i+1 
				j=j-1
			Loop
		%>
    </form>
	<tr class="hback" height="10"> 
		<td align="right" colspan="5"><input name="AddAward" type="button" value="�� ��" onClick="location='AFPoint_AddEdit.asp?act=addQuestion'"></td> 
		<td width="30%" align="center"><input type="Button" name="deleteAwards" onClick="AlertBeforeSubmite()" value="ɾ ��"> 
	    </td> 
	</tr>
	<tr>
	<td align="right" colspan="6">
	<%=AnswerForPoint.recordcount%> ����Ϣ&nbsp;&nbsp;<%=listnum%> ����Ϣ/ҳ&nbsp;&nbsp;�� <%=AnswerForPoint.pagecount%> ҳ 
	<% if page=1 then %>
	<%else%>
	<a href=<%=pagename%>><strong>|<<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><strong><<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><b>[<%=page-1%>]</b></a>&nbsp; 
	<%end if%>
	<% if AnswerForPoint.pagecount=1 then %>
	<%else%>
	<b>[<%=page%>]</b>
	<%end if%>
	<% if AnswerForPoint.pagecount-page <> 0 then %>
	<a href=<%=pagename%>page=<%=page+1%>><b>[<%=page+1%>]</b></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page+1%>><strong>>></strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=AnswerForPoint.pagecount%>><strong>>>|</strong></a>&nbsp; 
	<%end if%>��
	</td>
	</tr> 
</table> 
</body>
<%
if Request.QueryString("Act")="addGroup" then
	AddGroupRs.close
	set AddGroupRs=nothing
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
function getAwardUser(Obj1,Obj2)
{
	var url="getAwardUser.asp?AwardId="+Obj1+"&PrizeID="+Obj2+"&r="+Math.random();//����url
	request.open("GET",url,true);//��������
	request.onreadystatechange = getResult;
	request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���null
}
function getResult(Obj)//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			var contaner=result.substring(0,result.indexOf("*"));
			var selectContent=result.substring(result.indexOf("*")+1,result.length);
			document.getElementById(contaner).innerHTML=selectContent;

		}
	}
}
function CheckAll(Obj,TargetName)
{
	var CheckBoxArray;
	CheckBoxArray=document.getElementsByName(TargetName);
	for(var i=0;i<CheckBoxArray.length;i++)
	{
		if(Obj.checked)
		{
			CheckBoxArray[i].checked=true;
		}
		else
		{
			CheckBoxArray[i].checked=false;
		}
	}
}
function AlertBeforeSubmite()
{
	if(confirm("ȷ��Ҫɾ���ü�¼?"))
	{
		document.PrizeForm.submit();
	}
}
</script>
</html>






