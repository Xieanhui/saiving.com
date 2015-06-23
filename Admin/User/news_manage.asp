<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp"-->
<%
on error resume next
Dim Conn,User_Conn,NewsRs,i
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_News") then Err_Show 

Set NewsRs=Server.CreateObject(G_FS_RS)
NewsRs.open "select NewsID,title,addtime,groupid,newspoint,isLock from FS_ME_News order by addtime desc",User_Conn,1,3
'---------------------------------分页定义
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
'-----------------------------------------
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
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu">公告管理</td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="News_manage.asp">首页</a> | <a href="AddEditNews.asp?Act=add">发布公告</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form action="NewsAction.asp?act=delete" method="post" name="NewsForm" id="NewsForm">
    <tr class="xingmu"> 
      <td width="33%" align="center" class="xingmu">标题</td>
      <td width="19%" align="center" class="xingmu">发布时间</td>
      <td width="12%" align="center" class="xingmu">允许查看组</td>
      <td width="17%" align="center" class="xingmu">积分限制</td>
      <td width="12%" align="center" class="xingmu">锁定
        <input type="checkbox" name="Lock_CheckAll" value="all" onClick="CheckAll(this,'lock',true)"></td>
      <td width="7%" align="center" class="xingmu"><input type="checkbox" name="Delete_CheckAll" value="all" onClick="CheckAll(this,'deleteNews',false)"></td>
    </tr>
    <%

			If Not NewsRs.eof then
			'分页使用-----------------------------------
				NewsRs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>NewsRs.PageCount Then cPageNo=NewsRs.PageCount 
				NewsRs.AbsolutePage=cPageNo
			End if
			For i=0 To int_RPP
				If NewsRs.eof Then Exit for
				Response.Write("<tr class='hback'>")
				Response.Write("<td align='left'><a href='AddEditNews.asp?act=edit&newsID="&NewsRS("newsid")&"'>"&NewsRs("title")&"</a></td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("addtime")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("GroupID")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>"&NewsRs("NewsPoint")&"</td>"&Chr(10)&Chr(13))
				Response.Write("<td align='center'>")
				if NewsRs("isLock")=1 then
					Response.Write("<input type='checkbox' id='"&NewsRs("NewsID")&"' name='lock' value=1 onclick='changeLock(this.id,this.checked,false)'checked>"&Chr(10)&Chr(13))
				elseif NewsRs("isLock")=0 then
					Response.Write("<input type='checkbox' id='"&NewsRs("NewsID")&"' name='lock' value=1 onclick='changeLock(this.id,this.checked,false)'>"&Chr(10)&Chr(13))
				end if
				Response.Write("</td>")
				Response.Write("<td align='center'><input type='checkbox' name='deleteNews' value='"&NewsRs("NewsID")&"'></td>")
				Response.Write("</tr>")
				NewsRs.movenext 
			next
		%>
  </form>
  <tr class="hback"> 
    <td align="center" colspan="6"><div align="right">
        <input type="Button" name="DeleteNews" value="删除" onClick="AlertBeforeSubmite()">
        &nbsp;&nbsp;&nbsp;&nbsp;</div></td>
  </tr>
  <tr> 
    <td align="right" colspan="6" class="hback"> <%		Response.Write(fPageCount(NewsRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&vbcrlf)
	%> </td>
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
function changeLock(Obj1,Obj2)
{
	var url="NewsAction.asp?newsid="+Obj1+"&value="+Obj2+"&r="+Math.random();//构造url
	request.open("GET",url,true);//建立连接
	request.onreadystatechange = getResult;
	request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是null
}
function getResult()//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
			alert("修改成功")

		}
	}
}
function CheckAll(Obj,TargetName,isSync)
{
	var CheckBoxArray;
	CheckBoxArray=document.getElementsByName(TargetName);
	for(var i=0;i<CheckBoxArray.length;i++)
	{
		if(Obj.checked)
		{
			CheckBoxArray[i].checked=true;
			if(isSync)
			changeLock('all','true')
		}
		else
		{
			CheckBoxArray[i].checked=false;
			if(isSync)
			changeLock('all','false')
		}
	}
}
function AddNewsSubmit()
{	
	location='AddEditNews.asp?Act=add'
}
function AlertBeforeSubmite()
{
	document.NewsForm.submit();
}
</script>
</html>






