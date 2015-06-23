<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_SubSite") then Err_Show
if Request.Form("Action")  = "Del" then
		if trim(NoSqlHack(Request.Form("subid")))="" then
			strShowErr = "<li>请至少选择一个进行删除</li>"
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			Conn.execute("Delete From FS_MF_Sub_Sys where id in ("& FormatIntArr(Request.Form("subid")) &")")
			strShowErr = "<li>删除成功</li>"
			Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
End if
Dim p_Sub_List_Rs
Set p_Sub_List_Rs	= CreateObject(G_FS_RS)
p_Sub_List_Rs.Open "select id,Sub_Sys_Name,Sub_Sys_ID,Sub_Sys_Path,Sub_Sys_Index,Sub_Sys_Installed,Sub_Sys_Link from FS_MF_Sub_Sys order by id asc",Conn,1,1

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
  
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
<tr>
<td class="xingmu" colspan="6"><a href="#" class="sd"><strong>子系统维护</strong></a></td>
</tr>
  <form name="SubForm" id="form1" method="post" action="">
    <tr  class="hback"> 
      <td width="18%" align="left" class="xingmu" >子系统名称</td>
      <td width="18%" align="center" class="xingmu">子系统安装目录</td>
      <td width="18%" align="center"  class="xingmu">前台连接地址</td>
      <td width="18%" align="center" class="xingmu">是否启用</td>
      <td width="10%" align="center" class="xingmu">设置</td>
      <td width="5%" align="center" class="xingmu"><input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);"></td>
    </tr>
    <%
	Do While Not p_Sub_List_Rs.Eof
		Response.Write "<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>" & vbcrlf
		Response.Write "<td width=""18%"" align=""left"" class=""hback"" >"& p_Sub_List_Rs("Sub_Sys_Name") & "</td>" & vbcrlf
		Response.Write "<td width=""18%"" align=""center"" class=""hback"" >"& p_Sub_List_Rs("Sub_Sys_Path") & "</td>" & vbcrlf
		Response.Write "<td width=""18%"" align=""center"" class=""hback"" >"& p_Sub_List_Rs("Sub_Sys_Link") & "</td>" & vbcrlf
		If p_Sub_List_Rs("Sub_Sys_Installed") = 1 Then
			Response.Write "<td width=""18%"" align=""center""  class=""hback"" ><span class=""tx"">已启用</span></td>"
			Response.Write "<td width=""10%"" align=""center""  class=""hback"" ><a href=""SubSysSet_Edit.asp?Sub_ID="&p_Sub_List_Rs("Sub_Sys_ID")&"""  class=""otherset"">设置</FONT></a></td>" & vbcrlf
		Else
			Response.Write "<td width=""18%"" align=""center""  class=""hback"" >未启用</td>"
			Response.Write "<td width=""10%"" align=""center"" class=""hback""><a href=""SubSysSet_Edit.asp?Sub_ID="&p_Sub_List_Rs("Sub_Sys_ID")&""" class=""otherset"">设置</a></td>" & vbcrlf
		End If
			Response.Write "<td width=""5%"" align=""center""  class=""hback"" ><input type=""checkbox"" name=""subid"" value="""&p_Sub_List_Rs("id")&""" /></td>"
		Response.Write "</tr>" & vbcrlf
		p_Sub_List_Rs.MoveNext
	Loop

%>
    <tr  class="hback"> 
      <td colspan="6" align="left" class="hback" ><div align="right"><input name="Action" type="hidden" value="">
          <input type="button" name="Submit" value="删除选中的子系统" onClick="document.SubForm.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.SubForm.submit();return true;}return false;}">
        </div></td>
    </tr>
  </form>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
p_Sub_List_Rs.Close
Set p_Sub_List_Rs = Nothing
Conn.Close
Set Conn = Nothing
%><script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = SubForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = SubForm.chkall.checked;  
    }  
	}
</script>






