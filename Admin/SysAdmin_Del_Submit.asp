<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Pop") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo

int_RPP=20 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<%
Dim obj_admin_get_rs
set obj_admin_get_rs= Conn.execute("select ID,Admin_Name,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked From FS_MF_Admin where ID="&CintStr(Request.QueryString("Id")))
Dim obj_admin_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL
strpage=NoSqlHack(request("page"))
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set obj_admin_Rs = server.CreateObject(G_FS_RS)
SQL = "Select ID,Admin_Name,Admin_Parent_Admin,Admin_Is_Super,Admin_Real_Name,Admin_Is_Locked  from FS_MF_Admin where Admin_Name<>'"& obj_admin_get_rs("Admin_Name")&"' and  Admin_Parent_Admin<>'"& obj_admin_get_rs("Admin_Name")&"' Order by id desc"
obj_admin_Rs.Open SQL,Conn,1,3
%>
  <form name="form1" method="post" action="">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu">管理员管理</td>
  </tr>
  <tr class="hback">
    <td><a href="SysAdmin_List.asp">管理员首页</a> | <a href="SysAdmin_List.asp?Is_Super=1">超级管理员</a> 
      | <a href="SysAdmin_List.asp?islock=1">锁定的管理员</a> | <a href="SysAdmin_List.asp?islock=0">开放的管理员</a> 
      | <a href="SysAdmin_List.asp?my=1">我的管理员</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">

    <tr class="hback"> 
      <td width="32%" height="25" class="xingmu"> <div align="left">指定删除此管理员<span class="tx">(<% = obj_admin_get_rs("Admin_Name")%>
          )</span>后的下级管理员所属的管理员,必须选择一项 </div></td>
    </tr>
    <tr class="hback"> 
      <td height="25"><input type="radio" name="Parent_Admin_Name" value="0">
        设置此管理员将不属于任何管理员</td>
    </tr>
    <%
	Response.Write"<tr class=""hback""><td class=""hback""><table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
if obj_admin_Rs.eof then
   obj_admin_Rs.close
   set obj_admin_Rs=nothing
   Response.Write"<table width=""98%"" class=""table"" align=""center""><tr  class=""hback""><td  class=""hback"" height=""40"">没有符合条件的管理员。</td></tr></table>"
else
	obj_admin_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>obj_admin_Rs.PageCount Then cPageNo=obj_admin_Rs.PageCount 
	obj_admin_Rs.AbsolutePage=cPageNo
	Response.Write"<tr class=""hback"">"
	Dim i_tmp_n
	i_tmp_n = 0
	for i=1 to obj_admin_Rs.pagesize
		if obj_admin_Rs.eof Then exit For 
			Response.Write"<td class=""hback""><input type=""radio"" name=""Parent_Admin_Name"" value="""& obj_admin_Rs("Admin_Name") &""">"& obj_admin_Rs("Admin_Name") &"-"& obj_admin_Rs("Admin_Real_Name") &"</td>"
		obj_admin_Rs.movenext
		i_tmp_n = i_tmp_n +1 
		if i_tmp_n mod 4 = 0 then
			Response.Write("</tr>")
		End if
	Next
		Response.Write"</tr>"
		Response.Write"</table></td></tr>"
	%>
    <tr class="hback"> 
      <td height="25"> <%
			response.Write "<p>"&  fPageCount(obj_admin_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
      </td>
    </tr>
    <%
	obj_admin_Rs.close
	set obj_admin_Rs = nothing
End if
%>
</table>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td class="hback"><div align="right">
          <input name="AdminID" type="hidden" id="AdminID" value="<% = Request.QueryString("ID")%>">
          <input name="Action" type="hidden" id="Action" value="del_p">
          <input type="submit" name="Submit" value="确定删除">
      </div></td>
  </tr>
</table></form>
</body>
</html>
<%
if Request.Form("Action")="del_p" then
	if NoSqlHack(Request.Form("Parent_Admin_Name"))="" then
		strShowErr = "<li>请选择一个父级管理员!!!</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		Dim obj_admin_rs_2,tmp_str_d
		Set obj_admin_rs_2 = Conn.execute("Select Admin_Parent_Admin,Admin_Name,Admin_Is_Super,Admin_Add_Admin From FS_MF_Admin where ID="&CintStr(Request.Form("AdminID")))
		tmp_str_d = obj_admin_rs_2("Admin_Name")
		if session("Admin_Is_Super")<>1 then
			if obj_admin_rs_2("Admin_Name")<>session("Admin_Name") then
				if obj_admin_rs_2("Admin_Is_Super")=1 then
					strShowErr = "<li>您不能删除系统管理员</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
				if obj_admin_rs_2("Admin_Add_Admin")<>session("Admin_Name") then
					strShowErr = "<li>您不能删除别人的管理员。</li>"
					Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
					Response.end
				end if
			end if
		end if
		'判断是否有隶属管理员
		Conn.execute("Update FS_MF_Admin set Admin_Parent_Admin ='"&NoSqlHack(Request.Form("Parent_Admin_Name"))&"' where Admin_Parent_Admin='"& NoSqlHack(tmp_str_d) &"'")
		Conn.execute("Delete From FS_MF_Admin where id="&CintStr(Request.Form("AdminID")))
		'插入日志
		'删除静态目录
		Dim p_FSO,tmp_path
		Set p_FSO = Server.CreateObject(G_FS_FSO)
		tmp_path = "..\"& G_UP_FILES_DIR &"\adminFiles\"& tmp_str_d
		tmp_path = Server.MapPath(Replace(tmp_path,"\\","\"))
		if p_FSO.FolderExists(tmp_path) = true then p_FSO.DeleteFolder tmp_path
		set p_FSO = nothing
		Call MF_Insert_oper_Log("删除管理员","删除了管理员ID("& tmp_str_d &")："&Request.Form("AdminID")&",同时锁定了此管理员下所有的隶属管理员",now,session("admin_name"),"MF")
		obj_admin_rs_2.close:set obj_admin_rs_2 = nothing
		strShowErr = "<li>删除管理员成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
		Response.end
	End if
End if
Set Conn = Nothing
%>





