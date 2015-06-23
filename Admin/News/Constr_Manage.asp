<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim Conn,Constr_Rs,sql_cmd,AuditTF,classid
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS_Constr") then Err_Show
Set Constr_Rs=server.CreateObject(G_FS_RS)
sql_cmd="Select id,ClassID,ClassName,ParentID from FS_NS_NewsClass where parentID='0' and isConstr=1"
Constr_Rs.open sql_cmd,Conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language='javascript' src="../../FS_Inc/prototype.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td colspan="10" class="xingmu"><a href="#" onmouseover="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by 刘南兵 <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>投稿管理</strong></a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=News_Manage" target="_blank" style="cursor:help;'" class="sd"> 帮助</a></td>
  </tr>
  <tr class="hback">
  <td colspan="10">选择分类：
  <span id="span_constrClass_0">
  <!---联动菜单开始--->	
  <%
  	Response.Write("<select name=""sel_constrClass"" name=""sel_constrClass"" onchange=""getChildClass(this,0)"">")
    Response.Write("<option value="""">所有分类</option>")
	Dim p_temp_index
  	while not Constr_Rs.eof
		Response.Write("<option value="""&Constr_Rs("id")&""">"&Constr_Rs("ClassName")&"</option>")
		Constr_Rs.movenext
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
	<!---联动菜单结束--->	
  	<span id="span_audit">已审核<input type="checkbox" name="audit" value="1" onClick="auditView(this);"></span>
  </td>
  </tr>
  <tr class="hback">
  <td colspan="10"><iframe id="contrList" src="Constr_list.asp" frameborder="0" width="100%" height="600"></iframe></td>
  </tr>
</table>
<input type="hidden" name="hd_classid" />
<input type="hidden" name="hd_audit" />
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
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
	var container="span_constrClass_"+(index+1);
	$(container).innerHTML="<img src='../../sys_images/small_loading.gif'/>";
	var AjaxObj=new Ajax.Updater(container,'getClass.asp?and='+Math.random(),{method:'get',parameters:"index="+index+"&classid="+classid});
	$('contrList').src="constr_list.asp?classid="+classid+"&audittf="+$('hd_audit').value+"&rnd="+Math.random();
}

function auditView(Obj)
{
	if(Obj.checked)
	{
		$('contrList').src="constr_list.asp?classid="+$('hd_classid').value+"&audittf="+1+"&rnd="+Math.random();
	}else
	{
		$('contrList').src="constr_list.asp?classid="+$('hd_classid').value+"&audittf="+0+"&rnd="+Math.random();
	}
}
</script>
<%
Constr_Rs.close
Conn.close
Set Conn=nothing
Set Constr_Rs=nothing
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





