<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr,strpage
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("SS_site") then Err_Show
if not MF_Check_Pop_TF("SS001") then Err_Show

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo

int_RPP=20 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"
Dim Action,Sql,RsVisitObj,ID,IDArray,i
Action = Request("Action")
if Action = "del" then
	ID = Replace(Request("id")," ","")
	if ID=empty then
		strShowErr = "<li>请选择至少一项</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	IDArray = Split(ID,",")
	for i = LBound(IDArray) to UBound(IDArray)
		if IDArray(i) <> "" then
			Conn.Execute("Delete from FS_SS_Stat Where ID="+CintStr(IDArray(i)))
		end if
	next
	strShowErr = "<li>删除成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
elseif Action = "all" then
	Conn.Execute("Delete from FS_SS_Stat")
	strShowErr = "<li>删除所有成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Sql = "Select * from FS_SS_Stat Order By VisitTime Desc"
Set RsVisitObj = Server.CreateObject(G_FS_RS)
RsVisitObj.Open Sql,Conn,1,1
%>
<html>
<head>
<title>来访者信息列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="../../FS_Inc/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" >
<form name="form1" method="post" action="">
  <table width="844" border="0" align="center" cellpadding="5" cellspacing="1"  class="table">
    <tr> 
      <td height="26" colspan="5" valign="middle" class="xingmu"> 来访者信息</td>
    </tr>
  </table>
  <table width="844" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="15%" height="26" class="xingmu"> <div align="center">操作系统</div></td>
      <td width="17%" height="26" class="xingmu"> <div align="center">浏览器</div></td>
      <td width="16%" height="26" class="xingmu"> <div align="center">IP地址</div></td>
      <td width="28%" height="26" class="xingmu"> <div align="center">地区</div></td>
      <td width="18%" height="26" class="xingmu"> <div align="center">访问时间</div></td>
    </tr>
    <%
	if RsVisitObj.eof then
	   RsVisitObj.close
	   set RsVisitObj=nothing
	   Response.Write"<tr  class=""hback""><td colspan=""5""  class=""hback"" height=""40"">没有访问者。</td></tr>"
	else
		RsVisitObj.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>RsVisitObj.PageCount Then cPageNo=RsVisitObj.PageCount 
		RsVisitObj.AbsolutePage=cPageNo
		for i=1 to RsVisitObj.pagesize
			if RsVisitObj.eof Then exit For 
%>
    <tr class="hback"> 
      <td><input name="id" type="checkbox" id="id" value="<%=RsVisitObj("ID")%>">
        <%=RsVisitObj("OSType")%></td>
      <td><div align="center"><%=RsVisitObj("ExploreType")%></div></td>
      <td><div align="center"><%=RsVisitObj("IP")%></div></td>
      <td><div align="center"><%=RsVisitObj("Area")%></div></td>
      <td><div align="center"><%=RsVisitObj("VisitTime")%></div></td>
    </tr>
    <%
		RsVisitObj.MoveNext
	Next
	%>
<tr class="hback">
      <td colspan="5"><table width="833" height="34">
      <tr class="hback"> <td width="466" align="left"> <%
			response.Write "<p>"&  fPageCount(RsVisitObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%> </td>
 <td width="355" colspan="5" align="left"><input name="Action" type="hidden" id="Action" value="del" >
          <input type="button" name="Submit" value="删除" onClick="{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
          <input type="button" name="Submit2" value="删除所有" onClick="{if(confirm('确定删除所有吗？')){window.location.href='Visit_visitorList.asp?Action=all';return true;}return false;}"></td>
    </tr></table></td>
    </tr>
	    
<%
end if
%>
  </table>
</form>
</body>
</html>
<%
Set Conn = Nothing
Set RsVisitObj = Nothing
%>





