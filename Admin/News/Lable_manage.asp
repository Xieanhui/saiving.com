<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF 
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001")
'得到会员组列表
dim Fs_news,strShowErr
set Fs_news = new Cls_News
Fs_News.GetSysParam()
if Request.QueryString("Action") = "del_Ls" then
	Conn.execute("Delete From FS_NS_Lable where id="&CintStr(Request.QueryString("id")))
	strShowErr = "<li>标签删除成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_Manage.asp")
	Response.end
End if
if Request.QueryString("Action") = "del_lable" then
	Conn.execute("Delete From FS_NS_Lable")
	strShowErr = "<li>所有的标签已经置空</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
	Response.end
Elseif Request.QueryString("Action") = "del_lable_dir" then
	Dim obj_all_rs
	set obj_all_rs = Conn.execute("select ID from FS_NS_LableClass order by id desc")
	do while not obj_all_rs.eof 
		Conn.execute("Delete From FS_NS_LableClass where ID="&obj_all_rs("Id"))
		Conn.execute("Update  FS_NS_Lable set LableClassID=0 where LableClassID="&obj_all_rs("Id"))
		obj_all_rs.movenext
	Loop
	strShowErr = "<li>所有标签分类已经删除</li><li>标签分类下的标签已经归位到根目录</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
	Response.end
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">标签管理<a href="../../help?Lable=NS_Lable_Manage" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Lable_Manage.asp">管理首页</a> 
        &nbsp;|&nbsp; <a href="Lable_Create.asp?ClassID=<%=Request.QueryString("ClassID")%>">创建标签</a> 
        &nbsp;|&nbsp; <a href="Lable_Dir_add.asp#Add">添加标签栏目</a> &nbsp;|&nbsp; 
        <a href="Lable_style.asp">标签样式管理</a> &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable"  onClick="{if(confirm('确认删除吗!')){return true;}return false;}">删除所有标签</a> 
        &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable_dir"  onClick="{if(confirm('确认删除吗!')){return true;}return false;}">删除所有标签目录</a> 
        | <a href="../../help?Lable=NS_Lable_Manage_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr class="hback"> 
    <td> 
      <%
  	Dim Lableclass_SQL,obj_Lableclass_rs,icNum,news_count,obj_Lableclass_rs_1
	Lableclass_SQL = "Select ID,ClassName from FS_NS_LableClass   Order by id desc"
	Set obj_Lableclass_rs = server.CreateObject(G_FS_RS)
	obj_Lableclass_rs.Open Lableclass_SQL,Conn,1,3
	Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" >")
	Response.Write("<tr>")
	icNum = 0
	if Not obj_Lableclass_rs.eof then
		Do while Not obj_Lableclass_rs.eof 
				Set obj_Lableclass_rs_1 = server.CreateObject(G_FS_RS)
				obj_Lableclass_rs_1.Open "Select ID from FS_NS_Lable where LableClassID="& obj_Lableclass_rs("ID") &"",Conn,1,3
				news_count = "("&obj_Lableclass_rs_1.recordcount&")"
				obj_Lableclass_rs_1.close:set obj_Lableclass_rs_1 = nothing
			Response.Write"<td height=""22"">"
			Response.Write "<a href=Lable_Manage.asp?ClassID="&obj_Lableclass_rs("id")&">"&obj_Lableclass_rs("ClassName")&news_count&"</a>"
			Response.Write "</td>"
			obj_Lableclass_rs.MoveNext
			icNum = icNum + 1
			if icNum mod 6 = 0 then
				Response.Write("</tr><tr>")
			End if
		loop
	Else
	Response.Write("<td>没有目录</td>")
	End if
	Response.Write("</tr>")
	Response.Write("</table>")
	set obj_Lableclass_rs = nothing
%>
    </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td width="42%" class="xingmu">标签名称</td>
    <td width="16%" class="xingmu">栏目</td>
    <td width="28%" class="xingmu">说明</td>
    <td width="14%" class="xingmu">操作</td>
  </tr>
  <%
  	Dim strpage,i,obj_lable_rs,select_count,select_pagecount,tmp_classid
	strpage=request("page")
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
	Set obj_lable_rs = server.CreateObject(G_FS_RS)
	if Request.QueryString("ClassID")<>"" then:tmp_classid = " and LableClassID="& clng(Request.QueryString("ClassID"))&"":Else:tmp_classid = "":End if
	obj_lable_rs.Open "Select ID,LableName,LableDesc,LableContent,AddDate,LablelType,LableClassID,isDel from FS_NS_Lable where isDel=0 "& tmp_classid &" order by id desc",Conn,1,3
	if obj_lable_rs.eof then
	   obj_lable_rs.close
	   set obj_lable_rs=nothing
	   Response.Write"<TR  class=""hback""><TD colspan=""4""  class=""hback"" height=""40"">没有标签。</TD></TR>"
	Else
		obj_lable_rs.pagesize = 20
		obj_lable_rs.absolutepage=cint(strpage)
		select_count=obj_lable_rs.recordcount
		select_pagecount=obj_lable_rs.pagecount
		for i=1 to obj_lable_rs.pagesize
			if obj_lable_rs.eof Then exit For 
  %>
  <tr> 
    <td class="hback">·<a href="Lable_Manage.asp?ClassID=<%=obj_lable_rs("LableClassID")%>&Action=edit&ID=<%=obj_lable_rs("ID")%>"> 
      <% = obj_lable_rs("LableName") %>
      </a></td>
    <td class="hback"> <%
	if obj_lable_rs("LableClassID")=0 then
		Response.Write("根目录")
	Else	
		Dim obj_tmp_rs
		set obj_tmp_rs = Conn.execute("select classname from FS_NS_LableClass where id="&obj_lable_rs("LableClassID"))
		Response.Write "<a href=Lable_Manage.asp?ClassID="& obj_lable_rs("LableClassID") &">"& obj_tmp_rs("classname") &"</a>"
		obj_tmp_rs.close:set obj_tmp_rs = nothing
	End if
	%></td>
    <td class="hback"> 
      <% = obj_lable_rs("LableDesc") %> </td>
    <td class="hback"><a href="Lable_Create.asp?ClassID=<%=obj_lable_rs("LableClassID")%>&Action=edit&ID=<%=obj_lable_rs("ID")%>">修改</a>｜<a href="Lable_Manage.asp?ClassID=<%=obj_lable_rs("LableClassID")%>&Action=del_Ls&ID=<%=obj_lable_rs("ID")%>" onClick="{if(confirm('确定删除标签吗!')){return true;}return false;}">删除</a></td>
  </tr>
  <%
	   obj_lable_rs.MoveNext
  Next
  %>
  <tr> 
    <td height="35" colspan="4" class="hback"> <%
		 Response.Write("每页:&nbsp;<b>"& obj_lable_rs.pagesize &"</b>&nbsp;个标签,")
		Response.write"&nbsp;共<b>"& select_pagecount &"</b>页&nbsp;<b>" & select_count &"</b>&nbsp;个"& Fs_news.allInfotitle &"，本页是第&nbsp;<b>"& strpage &"</b>&nbsp;页。"
		if int(strpage)>1 then
			Response.Write"&nbsp;<a href=""Lable_Manage.asp?page=1&Classid="&Request("Classid")&""">第一页</a>&nbsp;&nbsp;"
			Response.Write"&nbsp;<a href=""Lable_Manage.asp?page="&cstr(cint(strpage)-1)&"&Classid="&Request("Classid")&""">上一页</a>&nbsp;&nbsp;"
		End if
		If int(strpage)<select_pagecount then
			Response.Write"&nbsp;<a href=""Lable_Manage.asp?page="&cstr(cint(strpage)+1)&"&Classid="&Request("Classid")&""">下一页</a>&nbsp;"
			Response.Write"&nbsp;<a href=""Lable_Manage.asp?page="& select_pagecount &"&Classid="&Request("Classid")&""">最后一页</a>&nbsp;&nbsp;"
		End if
	%> </td>
  </tr>
  <%
  End if
  set obj_lable_rs = nothing
  %>
</table>
</body>
</html>
<%
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = ClassForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = ClassForm.chkall.checked;  
    }  
  }
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





