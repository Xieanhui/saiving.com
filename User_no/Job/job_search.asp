<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Charset="GB2312"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=GetUserSystemTitle%></title>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
</head>
<body>
<form  name="searchForm" id="searchform" action="job_search_result.asp" method="post">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
<tr>
<td class="xingmu">职位搜索  <a href="#" onclick="history.back()" class="sd">后退</a></td>
</tr>
<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
<td class="hback"><img src="../images/folderopened.gif"/><input type="text" name="JobName" style="width:50%" />
<button onclick="showPane('div_JobName')">职位名称</button>
<div id="div_JobName" style="display:none; position:relative"></div>
</td>
</tr>
<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
<td class="hback"><img src="../images/folderopened.gif"/><input type="text" name="txt_WorkCity" style="width:50%" />
<button onclick="showPane('div_WorkCity')">工作地点</button>
<div id="div_WorkCity" style="display:none; position:relative"></div>
<div id="div_WorkCity_2" style="display:none; position:relative"></div>
</td>
</tr>
<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
<td class="hback"><img src="../images/folderopened.gif"/><input type="text" name="PublicDate" style="width:50%"/>
<button onclick="showPane('div_PublicDate')">发布日期</button>
<div id="div_PublicDate" style="display:none; position:relative"></div>
</td>
</tr>
<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
<td class="hback">&nbsp;&nbsp;<input type="submit" name="search" value="搜 索" /></td>
</tr>
<input type="hidden" name="hd_PublicDate" style="width:50%"/>
</table>
</form>
</body>
</html>
<%
Conn.close
User_Conn.close
Set User_Conn=nothing
Set Conn=nothing
%>






