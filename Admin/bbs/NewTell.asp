<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session判断
MF_Session_TF 
if not MF_Check_Pop_TF("WS001") then Err_Show
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
 }

</script>
<%
'进行调用处理
if Request.QueryString("Act")="N" and Request.QueryString("ID")<>"" then
	Conn.execute("Update FS_WS_NewsTell Set IsUse=0")
	Conn.Execute("Update FS_WS_NewsTell Set IsUse=1 where ID="&CintStr(Request.QueryString("ID"))&"")
	Response.Redirect("NewTell.asp")
elseif Request.QueryString("Act")="Y" and Request.QueryString("ID")<>"" then
	Conn.Execute("Update FS_WS_NewsTell Set IsUse=0 where ID="&CintStr(Request.QueryString("ID"))&"")
	Response.Redirect("NewTell.asp")
end if
%>
<BODY>
<%
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=10 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"	
%>
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr class="hback">
      <td align="left" colspan="2" class="xingmu">公告设置&nbsp;<a href="../../help?Lable=Message" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
	<tr>
		<td colspan="2" class="hback"><a href="AddNewsTell.asp">添加公告</a></td>
	</tr>
  </table>
<%
dim Rs
Set Rs=server.createobject(G_FS_RS)
Rs.open "Select ID,Topic,Content,Person,IsUse,PV,AddUser,AddDate From FS_WS_NewsTell order by ID DESC ",Conn,1,1
if Rs.eof and Rs.bof then
	Response.Write("<table width=""98%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" class=""table""><tr class=""hback""><td>暂无公告!</td></tr><table>")
	Set Rs=nothing
	Response.End()
else
	RS.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>RS.PageCount Then cPageNo=Rs.PageCount 
	Rs.AbsolutePage=cPageNo
end if
%>
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
      <form action="DelNewsTell.asp?Act=del" method="post" name="myform">

    <tr class="hback">
	   <td align="center" class="xingmu" width="5%">选 择</td>
      <td align="center" class="xingmu" width="25%">公告标题</td>
	  <td class="xingmu" align="center" width="10%">点击次数</td>
	  <td align="center" class="xingmu" width="6%">调  用</td>
      <td align="center" class="xingmu" width="10%">添 加 者</td>
      <td class="xingmu" align="center" width="15%">添加时间</td>
      <td class="xingmu" align="center" width="18%">操    作</td>
    </tr>
	<%
	For  int_Start=1 To int_RPP 
	%>
    <tr class="hback">
	  <td align="center"><input type="checkbox" value="<%=Rs("ID")%>" name="TellID"></td>
	  <td align="center"><a href="EditNewsTell.asp?Act=Edits&ID=<%=Rs("ID")%>"><%=left(Rs("Topic"),20)%></a></td>
	  <td align="center"><%=Rs("PV")%></td>

	        <td align="center">
	  <%
	  if Rs("IsUse")="0" then 
	  	response.write("<a href='?Act=N&ID="&Rs("ID")&"'>否</a>")
	  else
	  	response.write("<a href='?Act=Y&ID="&Rs("ID")&"'><font color='red'>是</font></a>")
	  end if
	  %></td>
	  <td align="center"><%=Rs("AddUser")%></td>
	  <td align="center"><%=Rs("AddDate")%></td>
	  <td align="center">&nbsp;&nbsp;&nbsp;<a href="EditNewsTell.asp?Act=Edits&ID=<%=Rs("ID")%>">修改</a>
｜<a href="DelNewsTell.asp?Act=singledel&ID=<%=Rs("ID")%>"  onClick="{if(confirm('确定要删除吗？')){return true;}return false;}">删除</a>
      </td>
	   </tr>
	 <%
	 Rs.MoveNext
	 if Rs.eof or Rs.bof then exit for
	 Next
	 %>
    <tr class="hback">
	 <td colspan="8" align="right"><table width="100%"><tr><td width="40%" align="center">
	  <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
          选择所有 
          <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('确实要进行删除吗?')){this.document.myform.submit();return true;}return false;}" value=" 删除 "></td>
	 <%
	 Response.Write("<td class=""hback"" colspan=""8"" align=""right"" width=""60%"">"&fPageCount(Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>")
	 Set Rs=nothing
	%>
	 </tr>
	</form>
</table>

<%
Set Conn=nothing
%>
</body>
</html>






