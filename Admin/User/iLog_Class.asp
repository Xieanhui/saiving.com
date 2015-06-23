<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
dim Conn,User_Conn,rs,str_c_isp,str_c_user,str_c_pass,str_c_url,str_domain,rs_param,str_c_gurl,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Log") then Err_Show 
if not MF_Check_Pop_TF("ME039") then Err_Show 
if Request.QueryString("Action")="Edit_save" then
	if Request.Form("classname")="" then
			strShowErr = "<li>请填写完整</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select id,ClassName From FS_ME_iLogClass where id="&CintStr(Request.Form("Aid")),User_Conn,1,3
		rs("ClassName")=NoSqlHack(request.Form("classname"))
		rs.update
		rs.close:set rs=nothing
		strShowErr = "<li>修改成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Class.asp")
		Response.end
	end if
end if
if Request("Action")="del" then
	if Request("id")="" then
			strShowErr = "<li>请选择至少一个目录</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		User_Conn.execute("delete From FS_ME_iLogClass where id in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Class.asp")
		Response.end
	end if
end if
if Request.QueryString("Action")="add_save" then
	if Request.Form("classname")="" then
			strShowErr = "<li>请填写完整</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		set rs= Server.CreateObject(G_FS_RS)
		rs.open "select id,ClassName From FS_ME_iLogClass where ClassName='"&NoSqlHack(Request.Form("ClassName"))&"'",User_Conn,1,3
		if not rs.eof then
			strShowErr = "<li>此目录已经存在</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			rs.addnew
			rs("ClassName")=NoSqlHack(request.Form("classname"))
			rs.update
			rs.close:set rs=nothing
			strShowErr = "<li>添加成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Class.asp")
			Response.end
		end if
	end if
end if
%>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>

<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td width="100%" class="xingmu">日志网摘管理</td>
  </tr>
  <tr> 
    <td class="hback"><a href="iLog.asp">日志管理</a>┆<a href="iLog_Templet.asp">模板设置</a>┆<a href="iLog_Class.asp">系统栏目</a>┆<a href="iLog_SetParam.asp">参数设置</a></td>
  </tr>
</table>
<%
if Request.QueryString("Action")="edit" then
set rs= Server.CreateObject(G_FS_RS)
rs.open "select id,ClassName From FS_ME_iLogClass where id="&CintStr(Request.QueryString("id")),User_Conn,1,3
if Rs.eof then
	response.Write("找不到记录")
	response.end
end if
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="Update" method="post" action="iLog_Class.asp">
    <tr> 
      <td width="33%" class="hback"><div align="right">栏目名称 
          <input name="ClassName" type="text" id="ClassName2" value="<%=rs("classname")%>">
        </div></td>
      <td width="67%" class="hback"><input type="button" name="Submit3" value="修改系统目录" onClick="javascript:UpdateCheck();"> <span id="ClassName_Alert" ></span>
        <input name="Aid" type="hidden" value="<%=rs("id")%>"></td>
    </tr>
  </form>
</table>
<%
rs.close:set rs=nothing
end if%>
<%
if Request.QueryString("Action")="add" then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="Add" method="post" action="iLog_Class.asp">
    <tr> 
      <td width="33%" class="hback"><div align="right">栏目名称 
          <input name="ClassName" type="text" id="ClassName">
        </div></td>
      <td width="67%" class="hback"><input type="button" name="Submit32" value="增加系统目录" onClick="javascript:AddCheck();"><span id="ClassName_Alert"></span>
      </td>
    </tr>
  </form>
</table>
<%
end if%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="17%" class="xingmu">系统栏目名称</td>
      <td width="83%" class="xingmu">操作</td>
    </tr>
    <%
  set rs= Server.CreateObject(G_FS_RS)
  rs.open "select id,ClassName From FS_ME_iLogClass order by id asc",User_Conn,1,3
  do while not rs.eof
  %>
    <tr> 
      <td class="hback"><%=rs("classname")%></td>
      <td class="hback"><a href="iLog_Class.asp?id=<%=rs("id")%>&Action=edit">修改</a>┆<a href="iLog_Class.asp?id=<%=rs("id")%>&Action=del">删除</a> 
        <input type="checkbox" name="id" value="<%=rs("id")%>"></td>
    </tr>
    <%
	  rs.movenext
  loop
  rs.close:Set rs=nothing
  %>
    <tr> 
      <td colspan="2" class="hback"><input name="Action" type="hidden" id="Action" value="del">
        <input type="submit" name="Submit" value="删除栏目">
        <input type="button" name="Submit2" value="增加栏目" onClick="window.location.href='iLog_Class.asp?Action=add';">
        <span class="top_navi">
        <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
        选中/取消 </SPAN></td>
    </tr>
  </form>
</table>

</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form1.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
  }
function AddCheck()
{
	var flag1=isEmpty("ClassName","ClassName_Alert");
	if(flag1)
	{
		document.Add.action="?Action=add_save";
		document.Add.submit();
	}
}
function UpdateCheck()
{
	var flag1=isEmpty("ClassName","ClassName_Alert");
	if(flag1)
	{
		document.Update.action="?Action=Edit_save";
		document.Update.submit();
	}
}

</script>

 





