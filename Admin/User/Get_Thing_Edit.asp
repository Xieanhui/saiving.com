<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
dim strShowErr,conn,user_conn
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Mproducts") then Err_Show 
if not MF_Check_Pop_TF("ME032") then Err_Show 

Function GetFriendUserNumber(f_username)
	Dim RsGetFriendName
	Set RsGetFriendName = User_Conn.Execute("Select UserNumber From FS_ME_Users Where Username = '"& NoSqlHack(f_username) &"'")
	If  Not RsGetFriendName.eof  Then 
		GetFriendUserNumber = RsGetFriendName("UserNumber")
	Else
		GetFriendUserNumber = "用户昵称有问题"
	End If 
	set RsGetFriendName = nothing
End Function 

Function GetFriendName(f_strNumber)
	Dim RsGetFriendName
	Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_strNumber) &"'")
	If  Not RsGetFriendName.eof  Then 
		GetFriendName = RsGetFriendName("UserName")
	Else
		GetFriendName = 0
	End If 
	set RsGetFriendName = nothing
End Function 
dim str_productid,str_version,str_PType,str_MaxNum,str_UserNumber,str_Content,str_URL_1,str_addTime,str_action,str_id
if Request.QueryString("Action")="Edit" then
	dim edit_rs
	if isnumeric(Request.QueryString("ID"))=false then
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
		Response.end
	end if
	set edit_rs= Server.CreateObject(G_FS_RS)
	edit_rs.open "select * From FS_ME_GetThing where id="&CintStr(Request.QueryString("ID")),User_Conn,1,3
	if edit_rs.eof then
		strShowErr = "<li>找不到记录</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
		Response.end
	else
		str_productid=edit_rs("ProductID")
		str_version=edit_rs("Version")
		str_PType=edit_rs("PType")
		str_MaxNum=edit_rs("MaxNum")
		str_UserNumber=edit_rs("UserNumber")
		str_Content=edit_rs("Content")
		str_URL_1=edit_rs("URL_1")
		str_addTime = edit_rs("addTime")
		str_action="Edit"
		str_id=edit_rs("id")
	end if
else
		str_MaxNum=5
		str_addTime = now
		str_action="Add"
end if
if Request.Form("Action")<>"" then
	dim rs_save_thing
	dim user_thing_obj
	set user_thing_obj = User_Conn.execute("select UserNumber From FS_ME_Users where UserName='"&NoSqlHack(Request.Form("UserNumber"))&"'")
	if user_thing_obj.eof then
		strShowErr = "<li>找不到此用户</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set rs_save_thing= Server.CreateObject(G_FS_RS)
	if Request.Form("Action")="Edit" then
		 rs_save_thing.open "select * From FS_ME_GetThing where Id="&CintStr(Request.Form("Id")),User_Conn,1,3
	else
		 rs_save_thing.open "select * From FS_ME_GetThing where 1=0",User_Conn,1,3
		 rs_save_thing.AddNEW
	end if
	rs_save_thing("ProductID")=NoSqlHack(Request.Form("ProductID"))
	rs_save_thing("Version")=NoSqlHack(Request.Form("Version"))
	rs_save_thing("PType")=NoSqlHack(Request.Form("PType"))
	rs_save_thing("MaxNum")=NoSqlHack(Request.Form("MaxNum"))
	rs_save_thing("UserNumber")=GetFriendUserNumber(NoSqlHack(Request.Form("UserNumber")))
	rs_save_thing("URL_1")=NoSqlHack(Request.Form("URL_1"))
	rs_save_thing("Content")=NoSqlHack(Request.Form("Content"))
	rs_save_thing("addTime")=NoSqlHack(Request.Form("addTime"))
	rs_save_thing.update
	rs_save_thing.close:set rs_save_thing=nothing
	strShowErr = "<li>保存成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Get_Thing.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes >
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td class="xingmu">会员商品</td>
  </tr>
  <tr>
    <td class="hback"><a href="Get_Thing.asp">所有</a>&nbsp;|&nbsp;<a href="Get_Thing.asp?Use=1">已使用</a> | <a href="Get_Thing.asp?Use=0">未使用</a> 
      | <a href="Get_Thing.asp?UserDel=1" onClick="history.back()">会员已删除</a> | <a href="Get_Thing.asp?isLock=1" onClick="history.back()">已锁定</a> | <a href="Get_Thing.asp?isLock=0">未锁定</a></td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <form name="form1" method="post" action=""><tr class="xingmu">
      <td width="24%" align="center" class="hback"><div align="right">产品</div></td>
      <td width="76%" align="center" class="hback"><div align="left">
        <input name="ProductID" type="text" id="ProductID" value="<% = str_productid %>" size="36">
      </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">版本号</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="Version" type="text" id="Version" value="<% = str_Version %>" size="36">
        </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">型号</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="PType" type="text" id="PType" value="<% = str_PType %>" size="36">
        </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">最大下载</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="MaxNum" type="text" id="MaxNum" value="<% = str_MaxNum %>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
        次　　　请填写正整数</div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">用户名</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="UserNumber" type="text" id="UserNumber" value="<% = GetFriendName(str_UserNumber) %>">
        请填写用户的用户名</div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">产品描述</div></td>
        <td align="center" class="hback"><div align="left">
          <textarea name="Content" cols="60" rows="8" id="Content"><% = str_Content %></textarea>
        </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">下载地址</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="URL_1" type="text" id="URL_1" value="<% = str_URL_1 %>" size="40">
        </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right">添加日期</div></td>
        <td align="center" class="hback"><div align="left">
          <input name="addTime" type="text" id="addTime" value="<% = str_addTime %>" readonly>
          <button onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',280,120,window,document.all.addTime);document.all.addTime.focus();">选择时间</button><font color="#FF0000">*</font><span id="startDate_Alert"><input name="Action" type="hidden" id="Action" value="<% = str_action %>">
          <input name="Id" type="hidden" id="Id" value="<% = str_Id %>">
        </div></td>
      </tr>
      <tr class="hback">
        <td align="center" class="hback"><div align="right"></div></td>
        <td align="center" class="hback"><div align="left">
          <input type="submit" name="Submit" value="保存用户商品">
          <input type="reset" name="Submit2" value="重置">
        </div></td>
      </tr>
   </form>
  </table>

</body>
</html>






