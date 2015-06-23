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
	'编号的实现
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
  <td colspan="7" class="xingmu">抽奖管理</td>
  </tr>
	<tr class="hback"> 
	<td width="30%" colspan="7" align="center"><div align="left"><a href="award.asp" target="_self">&nbsp;积分抽奖</a>&nbsp;|&nbsp;积分兑换 | <a href="AnswerForPoint.asp">积分竞答</a> | <a href="#" onClick="history.back()">后退</a></div></td> 
	</tr>
        <tr class="xingmu"> 
          <td width="25%" align="center">兑换物品</td> 
          <td width="15%" align="center">物品图片</td>
          <td width="10%" align="center">所需积分</td>
          <td width="10%" align="center">数量</td> 
		  <td width="15%" align="center">状态</td>
		  <td width="20%" align="center">截止时间</td>
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
					Response.Write("<td align='center'>已过期</td>")
				else
					Response.Write("<td align='center'>未过期</td>")
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
		<td align="right" colspan="6"><input name="AddAward" type="button" value="添 加" onClick="location='ChangePrize_AddEdit.asp?act=addPrize'"></td> 
		<td width="30%" align="center"><input type="Button" name="deleteAwards" onClick="AlertBeforeSubmite()" value="删 除"> 
	    </td> 
	</tr>
	<tr>
	<td align="right" colspan="7">
	<%=ChangePrizeRs.recordcount%> 条消息&nbsp;&nbsp;<%=listnum%> 条消息/页&nbsp;&nbsp;共 <%=ChangePrizeRs.pagecount%> 页 
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
	<%end if%>　
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
	var url="getAwardUser.asp?AwardId="+Obj1+"&PrizeID="+Obj2+"&r="+Math.random();//构造url
	request.open("GET",url,true);//建立连接
	request.onreadystatechange = getResult;
	request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是null
}
function getResult(Obj)//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
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
	if(confirm("确认要删除该记录?"))
	{
		document.PrizeForm.submit();
	}
}
</script>
</html>






