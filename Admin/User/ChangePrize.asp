<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,ChangePrizeRs,listnum,page,j,i,nn,n,pagename,prizeIDs,EndDate,awardUser,AwardID,ArrayIndex,AwardsUserRs,AwardsUserArray,UserInfoRs
MF_Default_Conn
MF_User_Conn
MF_Session_TF

if not MF_Check_Pop_TF("ME_award") then Err_Show 

Set ChangePrizeRs=Server.CreateObject(G_FS_RS)
ChangePrizeRs.open "select prizeID,PrizeName,PrizePic,NeedPoint,storage,StartDate,EndDate from FS_ME_Prize where isChange=1 order by endDate asc",User_Conn,1,1
pagename="ChangePrize.asp?"
if ChangePrizeRs.eof and ChangePrizeRs.bof Then

else
	listnum=10
	ChangePrizeRs.pagesize=listnum
	page=Request.queryString("page")
	if (page-ChangePrizeRs.pagecount) > 0 then
		page=ChangePrizeRs.pagecount
	elseif page = "" or page < 1 then
		page = 1
	end if
	ChangePrizeRs.absolutepage=page
	'��ŵ�ʵ��
	j=ChangePrizeRs.recordcount
	j=j-(page-1)*listnum
	i=0
	nn=Request.queryString("page")
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
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;"> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
<form action="awardAction.asp?act=deletePrizeaction" method="post" name="PrizeForm" id="PrizeForm">  
    <tr>
  <td colspan="7" class="xingmu">�齱����</td>
  </tr>
	<tr class="hback"> 
	<td width="30%" colspan="7" align="center"><div align="left"><a href="award.asp" target="_self">&nbsp;���ֳ齱</a>&nbsp;|&nbsp;���ֶһ� | <a href="AnswerForPoint.asp">���־���</a> | <a href="#" onClick="history.back()">����</a></div></td> 
	</tr>
        <tr class="xingmu"> 
          <td width="25%" align="center">�һ���Ʒ</td> 
          <td width="15%" align="center">��ƷͼƬ</td>
          <td width="10%" align="center">�������</td>
          <td width="10%" align="center">����</td> 
		  <td width="15%" align="center">״̬</td>
		  <td width="20%" align="center">��ֹʱ��</td>
		  <td width="10%" align="center"><input type="checkbox" name="Delete_CheckAll" value="all" onClick="CheckAll(this,'DeleteChangePrize')"></td> 
        </tr>
		<%
			do while not ChangePrizeRs.eof and i<listnum
				Response.Write("<tr class='hback'>"&chr(10)&chr(13)) 
				Response.Write("<td width='30%' align='center'><a href='ChangePrize_AddEdit.asp?act=editprize&prizeid="&ChangePrizeRs("PrizeID")&"'>"&ChangePrizeRs("PrizeName")&"</a></td>"&chr(10)&chr(13))
				Response.Write("<td align='center' width='20%'><img src='"&ChangePrizeRs("PrizePic")&"' width='40' height='40'>")
				Response.Write("</td>")
				Response.Write("<td align='center'>"&ChangePrizeRs("NeedPoint")&"</td>")
				EndDate=ChangePrizeRs("EndDate")
				if EndDate<Now() then
					Response.Write("<td align='center'>�ѹ���</td>")
				else
					Response.Write("<td align='center'>δ����</td>")
				end if
				Response.Write("<td align='center'>"&ChangePrizeRs("storage")&"</td>")
				Response.Write("<td width='20%' align='center'>"&EndDate&"</td>")
				Response.Write("<td align='center'><input type='checkbox' name='DeleteChangePrize' value='"&ChangePrizeRs("PrizeID")&"'></td>")
				Response.Write("</tr>")
				ChangePrizeRs.movenext
				i=i+1 
				j=j-1
			Loop
		%>
  </form>
	<tr class="hback" height="10"> 
		<td align="right" colspan="6"><input name="AddAward" type="button" value="�� ��" onClick="location='ChangePrize_AddEdit.asp?act=addPrize'"></td> 
		<td width="30%" align="center"><input type="Button" name="deleteAwards" onClick="AlertBeforeSubmite()" value="ɾ ��"> 
	    </td> 
	</tr>
	<tr>
	<td align="right" colspan="7">
	<%=ChangePrizeRs.recordcount%> ����Ϣ&nbsp;&nbsp;<%=listnum%> ����Ϣ/ҳ&nbsp;&nbsp;�� <%=ChangePrizeRs.pagecount%> ҳ 
	<% if page=1 then %>
	<%else%>
	<a href=<%=pagename%>><strong>|<<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><strong><<</strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page-1%>><b>[<%=page-1%>]</b></a>&nbsp; 
	<%end if%>
	<% if ChangePrizeRs.pagecount=1 then %>
	<%else%>
	<b>[<%=page%>]</b>
	<%end if%>
	<% if ChangePrizeRs.pagecount-page <> 0 then %>
	<a href=<%=pagename%>page=<%=page+1%>><b>[<%=page+1%>]</b></a>&nbsp; 
	<a href=<%=pagename%>page=<%=page+1%>><strong>>></strong></a>&nbsp; 
	<a href=<%=pagename%>page=<%=ChangePrizeRs.pagecount%>><strong>>>|</strong></a>&nbsp; 
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
function AddNewsSubmit()
{	
	location='AddEditNews.asp?Act=add'
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






