<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Dim Conn,Constr_stat_Rs,sql_cmd,classid
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS034") then Err_Show
classid=NoSqlHack(request.QueryString("classid"))
Set Constr_stat_Rs=server.CreateObject(G_FS_RS)
sql_cmd="Select id,ClassID,ClassName,ParentID from FS_NS_NewsClass where parentID='0' and isConstr=1"
Constr_stat_Rs.open sql_cmd,Conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" type="text/javascript" src="../../FS_Inc/prototype.js"></script>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
<tr>
<td align="left" class="xingmu">投稿统计<a href="../../help?Lable=NS_News_Stat" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
</tr>
<tr>
<td class="hback">
 <span id="span_constrClass_0">
  <!---联动菜单开始--->	
  <%
  	Response.Write("<select name=""sel_constrClass"" name=""sel_constrClass"" onchange=""getChildClass(this,0)"">")
Response.Write("<option value="""">所有分类</option>")
	Dim p_temp_index
  	while not Constr_stat_Rs.eof
		Response.Write("<option value="""&Constr_stat_Rs("id")&""">"&Constr_stat_Rs("ClassName")&"</option>")
		Constr_stat_Rs.movenext
	wend 
	Response.Write("</select>")
  %>
	</span>
	<span id="span_constrClass_1"></span>
	<span id="span_constrClass_2"></span>
	<span id="span_constrClass_3"></span>
	<span id="span_constrClass_4"></span>
	<span id="span_constrClass_5"></span>
	<span id="span_constrClass_6"></span>
	<span id="span_constrClass_7"></span>
</td>
</tr>
<tr class="hback">
<td colspan="10"><iframe id="contrList" src="Constr_stat_view.asp" frameborder="0" width="100%" height="600"></iframe></td>
</tr>
</table>
</body>
</html>
<script language="JavaScript">
//转到相应的分类下面
function getChildClass(Obj,index)
{
	var classid=Obj.value;
	if(classid=="")
	{
		var elements=Obj.parentNode.parentNode.childNodes;
		for(var i=2;i<elements.length-3;i++)
		{
			elements[i].innerHTML=""
			/*
			if(current.firstChild!=null)
				current.removeChild(current.firstChild);*/
		}
	}
	if("<%=classid%>"!="")
	{
		classid="<%=classid%>";
	}
	var container="span_constrClass_"+(index+1);
	$(container).innerHTML="<img src='../../sys_images/small_loading.gif'/>";
	var AjaxObj=new Ajax.Updater(container,'getClass.asp?and='+Math.random(),{method:'get',parameters:"index="+index+"&classid="+classid});
	$('contrList').src="constr_stat_view.asp?classid="+classid
}
</script>
<%
Constr_stat_Rs.close
Conn.close
Set Conn=nothing
Set Constr_stat_Rs=nothing
%>





