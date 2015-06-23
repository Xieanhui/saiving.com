<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--Copyright (c) 2006 Foosun Inc. Code by Einstein.liu-->
<html xmlns="http://www.w3.org/1999/xhtml">
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td><!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr class="back">
    <td   colspan="2" class="xingmu" height="26"><!--#include file="../Top_navi.asp" -->
    </td>
  </tr>
  <tr class="back">
    <td width="18%" valign="top" class="hback"><div align="left">
        <!--#include file="../menu.asp" -->
      </div></td>
    <td width="82%" valign="top" class="hback">
	<table width="100%" height="25" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
      <tr>
        <td class="xingmu" colspan="6">稿件管理</td>
      </tr>
	  <tr height="25">
		<td class="hback" align="center">
		<div id="myclass1">
		<select name="sel_myclass_1" onChange="getMyChildClass(this,1)">
		<option value="">所有专栏</option>
		<%
			Dim myClassRs
			Set myClassRs=User_Conn.execute("select ClassID,ClassCName,UserNumber from FS_ME_InfoClass where UserNumber='"&session("FS_UserNumber")&"' and parentid=0"&vbcrlf)
			while not myClassRs.eof
				Response.Write("<option value='"&myClassRs("classid")&"'>"&myClassRs("ClassCName")&"</option>"&vbcrlf)
				myClassRs.movenext
			wend
			myClassRs.close
			set myClassRs=nothing
		%>
		</select>
		</div>
		</td>
		<td class="hback" align="center"><div id="myclass2"></div></td>
		<td class="hback" align="center"><div id="myclass3"></div></td>
		<td class="hback" align="center"><div id="myclass4"></div></td>
		<td class="hback" align="center"><div id="myclass5"></div></td>
		<td class="hback" align="center"><div id="myclass6"></div></td>
	  </tr>
      <tr>
        <td class="hback" width="18%" align="center">
		<div id="class1">
		<select name="sel_class_1" onChange="getChildClass(this,1)">
		<option value="">所有稿件</option>
		<%
			Dim classRs
			Set classRs=Conn.execute("select id,ClassName from FS_NS_NewsClass where isConstr=1 and ReycleTF=0 and ParentID='0'")
			while not classRs.eof
				Response.Write("<option value='"&classRs("id")&"'>"&classRs("className")&"</option>"&vbcrlf)
				classRs.movenext
			wend
		%>
		</select>
		</div>
		</td>
		<td class="hback" width="18%" align="center"><div id="class2"></div></td>
		<td class="hback" width="18%" align="center"><div id="class3"></div></td>
		<td class="hback" width="18%" align="center"><div id="class4"></div></td>
		<td class="hback" width="18%" align="center"><div id="class5"></div></td>
		<td class="hback" width="12%" align="center">已审核<input type="checkbox" name="audit" value="1" onClick="audited(this)"></td>
      </tr>
      <tr>
        <td class="hback" align="center" colspan="6"><iframe id="contrListFrame" src="contrList.asp?" frameborder="0" width="100%" height="500"></iframe></td>
      </tr>
    </table>
	</td>
  </tr>
  <tr class="back">
    <td height="20"  colspan="2" class="xingmu"><div align="left">
        <!--#include file="../Copyright.asp" -->
      </div></td>
  </tr>
</table>
<input type="hidden" name="hd_classid" value=""/>
<input type="hidden" name="hd_audit" value=""/>
<input type="hidden" name="hd_myclassid" value=""/>
</body>
</html>
<%
Set Fs_User = Nothing
Set User_Conn=nothing
Set Conn=nothing
%>
<script language="javascript">
<!--
function audited(Obj)
{
	if(Obj.checked)
	{
		$('contrListFrame').src="contrList.asp?audit=1&classid="+$('hd_classid').value+"&myclass="+$('hd_myclassid').value
	}
	else
	{
		$('contrListFrame').src="contrList.asp?classid="+$('hd_classid').value+"&myclassid="+$('hd_myclassid').value
	}
}
function getChildClass(Obj,index)
{
	var classid=Obj.value;
	if(classid=="class")return;
	if(classid=="")
	{
		var elements=Obj.parentNode.parentNode.parentNode.childNodes;
		for(var i=1;i<elements.length-1;i++)
		{
			var current=elements[i].firstChild;
			if(current.firstChild!=null)
				current.removeChild(current.firstChild)
		}
	}
	var container="class"+(index+1);
	$(container).innerHTML="<img src='../../sys_images/small_loading.gif'/>";
	var AjaxObj=new Ajax.Updater(container,'getClass.asp?and='+Math.random(),{method:'get',parameters:"index="+index+"&classid="+classid});
	$('contrListFrame').src="contrList.asp?audit="+$('hd_audit').value+"&classid="+classid+"&myclassid="+$('hd_myclassid').value;
}
//获得专栏子栏目
function getMyChildClass(Obj,index)
{
	var myclassid=Obj.value;
	if(myclassid=="myclass")return;
	if(myclassid=="")
	{
		var elements=Obj.parentNode.parentNode.parentNode.childNodes;
		for(var i=1;i<elements.length;i++)
		{
			var current=elements[i].firstChild;
			if(current.firstChild!=null)
				current.removeChild(current.firstChild);
		}
	}
	var container="myclass"+(index+1);
	$(container).innerHTML="<img src='../../sys_images/small_loading.gif'/>";
	var AjaxObj=new Ajax.Updater(container,'getMyClass.asp?and='+Math.random(),{method:'get',parameters:"index="+index+"&classid="+myclassid});
	$('contrListFrame').src="contrList.asp?audit="+$('hd_audit').value+"&classid="+$('hd_classid').value+"&myclassid="+myclassid;
}
-->
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->
