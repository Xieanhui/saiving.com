<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
<title>新闻标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="41%" class="xingmu"><strong>常规标签创建</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">统计系统</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">统计类型</div></td>
      <td width="81%" class="hback">
	  	<select name="sstype" id="sstype">
        <option value="0" selected>无图标</option>
        <option value="1">有图标</option>
        <option value="2">文字统计</option>
      </select>      </td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">路径</div></td>
      <td class="hback"><select name="Path" id="Path">
        <option value="0" selected>相对路径</option>
        <option value="1">绝对路径</option>
      </select>
      </td>
    </tr>
    
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SS=SSTYPE┆';
		retV+='方式$' + obj.sstype.value + '┆';
		retV+='路径$' + obj.Path.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  </form>
</body>
</html>






