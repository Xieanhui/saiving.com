<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,strShowErr,Fs_news,obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_str_list,str_Id,obj_unite_rs,str_ClassId
MF_Default_Conn
'session判断
MF_Session_TF 
'权限判断
set Fs_news = new Cls_news
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
str_Id = server.HTMLEncode(NoSqlHack(Request.QueryString("Id")))%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.RefreshLen{
	height: 20px;
	width: 400px;
	border: 1px solid #104a7b;
	text-align: left;
	MARGIN-top:50px;
	margin-bottom: 5px;
}
</style>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback">
    <td colspan="2" class="xingmu"><strong>栏目管理</strong><a href="../../help?Lable=NS_News_Manage" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>　　　　　　　　　　　　　　    </td>
  </tr>
  <tr>
    <form name="form1" method="post" action="">
      <td width="48%" height="18" class="hback"><div align="left"><a href="News_Manage.asp">首页</a></div></td>
      <td class="hback"><div align="center"></div>        </td>
    </form>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr>
      <td colspan="2" class="xingmu">复制新闻选择栏目</td>
    </tr>
    <tr>
    <td width="28%" class="hback"><div align="right">选择栏目</div></td>
    <td width="72%" class="hback">
      <input name="id" type="hidden" id="id" value="<% = str_Id %>">
    
      <label>
      <select name="TargetClassId" id="TargetClassId" style="width:80%">
        <%
		tmp_str_list  = ""
		Set obj_unite_rs = server.CreateObject(G_FS_RS)
		obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_NS_NewsClass where Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
		do while Not obj_unite_rs.eof 
			tmp_str_list = tmp_str_list &"<option value="""& obj_unite_rs("ClassID") &","& obj_unite_rs("ParentID") &""">+"& obj_unite_rs("ClassName") &"</option>"& Chr(13) & Chr(10)
			tmp_str_list = tmp_str_list &Fs_news.UniteChildNewsList(obj_unite_rs("ClassID"),"")
			 obj_unite_rs.movenext
		Loop
		obj_unite_rs.close
		set obj_unite_rs = nothing
		Response.Write tmp_str_list
		 %>
      </select>
      </label></td>
  </tr>
  <tr class="hback">
    <td>&nbsp;</td>
    <td><label>
      <input type="submit" name="Submit" value="确定开始拷贝您选择的新闻 ">
    </label></td>
  </tr></form>
</table>
</body>
</html>
<%
set Fs_news = nothing
%>






