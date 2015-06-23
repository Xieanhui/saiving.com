<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Log") then Err_Show 
	if not MF_Check_Pop_TF("ME039") then Err_Show 

if request.Form("Action")="save" then
	if trim(Request.Form("isDefault"))="" then
			strShowErr = "<li>请选择默认项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		User_Conn.execute("Update FS_ME_InfoiLogTemplet set isDefault=0")
		User_Conn.execute("Update FS_ME_InfoiLogTemplet set isDefault=1 where Id="&clng(NoSqlHack(Request.Form("isDefault"))))
		strShowErr = "<li>设置成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Templet.asp")
		Response.end
	end if
end if
if Request.Form("action")="Edit_save" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select id,TempletName,TempletSavePath,isDefault From FS_ME_InfoiLogTemplet where id="&CintStr(Request.Form("id")),User_Conn,1,3
	rs("TempletName")=NoSqlHack(Request.Form("TempletName"))
	rs("TempletSavePath")=NoSqlHack(Request.Form("TempletSavePath"))
	rs.update
	rs.close:set rs=nothing
	strShowErr = "<li>修改成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Templet.asp")
	Response.end
end if
if Request.Form("action")="add_save" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select id,TempletName,TempletSavePath,isDefault From FS_ME_InfoiLogTemplet where 1=0",User_Conn,1,3
	rs.addnew
	rs("TempletName")=NoSqlHack(Request.Form("TempletName"))
	rs("TempletSavePath")=NoSqlHack(Request.Form("TempletSavePath"))
	rs.update
	rs.close:set rs=nothing
	strShowErr = "<li>增加成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Templet.asp")
	Response.end
end if
if Request.Form("Action")="Del" then
	if trim(Request.Form("id"))="" then
			strShowErr = "<li>请选择至少一个项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		User_Conn.execute("Delete From FS_ME_InfoiLogTemplet  where Id in ("&FormatIntArr(Request.Form("id"))&")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog_Templet.asp")
		Response.end
	end if
end if
%>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="100%" class="xingmu">日志网摘管理</td>
    </tr>
    <tr> 
      
    <td class="hback"><a href="iLog.asp">日志管理</a>┆<a href="iLog_Templet.asp">模板设置</a>┆<a href="iLog_Class.asp">系统栏目</a>┆<a href="iLog_SetParam.asp">参数设置</a></td>
    </tr>
</table>

  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td class="xingmu"><div align="center">模板名称</div></td>
      <td class="xingmu"><div align="center">目录地址</div></td>
      <td class="xingmu"><div align="center">设置默认</div></td>
      <td class="xingmu"><div align="center">选择</div></td>
    </tr>
    <%
	  dim rs
	  set rs= Server.CreateObject(G_FS_RS)
	  rs.open "select id,TempletName,TempletSavePath,isDefault From FS_ME_InfoiLogTemplet Order by Id desc",User_Conn,1,3
	  do while not rs.eof
  %>
    <tr> 
      <td width="20%" class="hback"><div align="center"><a href="iLog_Templet.asp?id=<%=rs("id")%>&action=edit"><%=rs("TempletName")%></a></div></td>
      <td width="36%" class="hback"><div align="center"><%=rs("TempletSavePath")%></div></td>
      <td width="22%" class="hback"><div align="center"> 
          <input type="radio" name="isDefault" value="<%=rs("id")%>" <%if rs("isDefault")=1 then response.Write("checked")%>>
        </div></td>
      <td width="22%" class="hback"><div align="center"> 
          <input name="id" type="checkbox" id="id" value="<%=rs("id")%>">
        </div></td>
    </tr>
    <%
	  rs.movenext
  loop
  rs.close:set rs=nothing
  %>
    <tr> 
      <td colspan="4" class="hback"> <div align="right">
          <input type="button" name="Submit2" value="增加模板" onClick="window.location.href='iLog_Templet.asp?action=addd';">
          <input type="button" name="Submit22" value="删除"  onClick="document.form1.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
          <input name="s" type="button" id="s" value="设置默认值"  onClick="document.form1.Action.value='save';{if(confirm('确定吗？')){this.document.form1.submit();return true;}return false;}">
          <input name="Action" type="hidden" id="Action" value="">
        </div></td>
    </tr>
    <tr> 
      <td colspan="4" class="hback">说明：目录名默认目录为模板目录下的ilog目录，如果你对目录不太清楚，请到主系统的模板管理中查看目录。<span class="tx">ilog目录不能改名</span></td>
    </tr>
  </form>
</table>
<%
if Request.QueryString("action")="edit" then
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select id,TempletName,TempletSavePath,isDefault From FS_ME_InfoiLogTemplet where id="&CintStr(Request.QueryString("id")),User_Conn,1,3
	if rs.eof then
		response.Write"错误的参数"
		response.end
	end if 
%>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form2" method="post" action="">
    <tr> 
      <td colspan="2" class="xingmu">修改模板</td>
    </tr>
    <tr> 
      <td width="18%" class="hback"><div align="right"> 
          <input type="hidden" name="action" value="Edit_save">
          <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
          模板名称 </div></td>
      <td width="82%" class="hback"><input name="TempletName" type="text" id="TempletName" value="<%=rs("TempletName")%>" size="40"><span id="TempletName_Alert"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">模板的目录</div></td>
      <td class="hback"><input name="TempletSavePath" type="text" id="TempletSavePath" value="<%=rs("TempletSavePath")%>" size="40"> 
        <INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%=replace("/"&G_VIRTUAL_ROOT_DIR&"/" & G_TEMPLETS_DIR&"/iLog","//","/")%>',300,250,window,document.form2.TempletSavePath);document.form2.TempletSavePath.focus();"><span id="TempletSavePath_Alert"></td>
    </tr>
    <tr> 
      <td class="hback">&nbsp;</td>
      <td class="hback"><input type="button" name="Submit" value="更新" onClick="javascript:AddCheck();"></td>
    </tr>
  </form>
</table>
<%
end if
%>
<%
if Request.QueryString("action")="addd" then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form2" method="post" action="">
    <tr> 
      <td colspan="2" class="xingmu">增加模板</td>
    </tr>
    <tr> 
      <td width="18%" class="hback"><div align="right"> 
          <input name="action" type="hidden" id="action" value="add_save">
          模板名称 </div></td>
      <td width="82%" class="hback"><input name="TempletName" type="text" id="TempletName" size="40"onKeyUp="if(event.keyCode==32)execCommand('undo')"  onafterpaste="if(event.keyCode==32)execCommand('undo')"><span id="TempletName_Alert"></span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">模板的目录</div></td>
      <td class="hback"><input name="TempletSavePath" type="text" id="TempletSavePath" size="40"> 
        <INPUT type="button"  name="Submit42" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%=replace("/"&G_VIRTUAL_ROOT_DIR&"/" & G_TEMPLETS_DIR&"/iLog","//","/")%>',300,250,window,document.form2.TempletSavePath);document.form2.TempletSavePath.focus();"><span id="TempletSavePath_Alert"></span></td>
    </tr>
    <tr> 
      <td class="hback">&nbsp;</td>
      <td class="hback"><input type="button" name="Submit3" value="增加模板" onClick="javascript:AddCheck();"></td>
    </tr>
  </form>
</table>
<%
end if
%>
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
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
function AddCheck()
{
	var flag1=isEmpty("TempletName","TempletName_Alert");
	var flag2=isEmpty("TempletSavePath","TempletSavePath_Alert");
	if(flag1&&flag2)
	{
		document.form2.submit();
	}
}
</script>
 





