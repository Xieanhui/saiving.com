<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("FL001") then Err_Show
if Request("act") ="del" then
	if not MF_Check_Pop_TF("FL002") then Err_Show
	if request("id")="" then
		strShowErr = "<li>请选择一项</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Conn.execute("Delete from FS_FL_FrendList where id in ("& FormatIntArr(Request("id")) & ")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
		Response.end
	end if
end if
if Request("act") ="lock" then
	if not MF_Check_Pop_TF("FL002") then Err_Show
	if request("id")="" then
		strShowErr = "<li>请选择一项</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
		Response.end
	else
		Conn.execute("Update FS_FL_FrendList set  F_Lock =1 where id in ("& FormatIntArr(Request("id")) &")")
		strShowErr = "<li>锁定成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
		Response.end
	end if
end if
if Request("act") ="unlock" then
	if not MF_Check_Pop_TF("FL002") then Err_Show
	if request("id")="" then
		strShowErr = "<li>请选择一项</li>"
		Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
		Response.end
	else
		Conn.execute("Update FS_FL_FrendList set  F_Lock =0 where id in ("& FormatIntArr(Request("id")) & ")")
		strShowErr = "<li>解锁成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
		Response.end
	end if
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=20 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td align="left" colspan="2" class="xingmu">友情连接管理</td>
  </tr>
  <tr> 
    <td align="left" colspan="2" class="hback"><a href="Flink_Manage.asp">管理首页</a>┆<a href="Flink_Edit.asp?Action=Add">添加连接</a>┆<a href="Flink_Manage.asp?Type=0">图片连接</a>┆<a href="Flink_Manage.asp?Type=1">文字连接</a>┆<a href="Flink_Manage.asp?Lock=1">已锁定</a>┆<a href="Flink_Manage.asp?Lock=0">未锁定</a>┆用户申请的连接:<a href="Flink_Manage.asp?Lock=0&isUser=1">已审核</a>，<a href="Flink_Manage.asp?Lock=1&isUser=1">未审核</a>┆管理员添加的连接:<a href="Flink_Manage.asp?Lock=0&isAdmin=1">已审核</a>，<a href="Flink_Manage.asp?Lock=1&isAdmin=1">未审核</a></td>
  </tr>
</table>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="myForm" method="post" action="">
    <tr> 
      <td width="26%" class="xingmu"><div align="center">站点</div></td>
      <td width="15%" class="xingmu"><div align="center">类别</div></td>
      <td width="13%" class="xingmu"><div align="center">类型</div></td>
      <td width="20%" class="xingmu"><div align="center">作者</div></td>
      <td width="26%" class="xingmu"><div align="center">操作</div></td>
    </tr>
    <%
	Dim strpage,obj_fl_Rs,SQL,i,tmp_lock,tmp_classid,tmp_user,tmp_admin
	strpage=CintStr(request("page"))
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
	if Request.QueryString("Type")="0" then
		tmp_type=" and F_Type=0"
	elseif Request.QueryString("Type")="1" then
		tmp_type=" and F_Type=1"
	Else
		tmp_type=""
	end if
	if Request.QueryString("Lock")="0" then
		tmp_lock=" and F_Lock=0"
	elseif Request.QueryString("Lock")="1" then
		tmp_lock=" and F_Lock=1"
	Else
		tmp_lock=""
	end if
	if trim(Request.QueryString("Classid"))<>""  and trim(Request.QueryString("Classid"))<>"0" then
		tmp_classid = " and classid="&clng(Request.QueryString("Classid") )&""
	else
		tmp_classid = ""
	end if
	if trim(Request.QueryString("isUser"))="1"  then
		tmp_user = " and F_isAdmin=0"
	elseif trim(Request.QueryString("isUser"))="0"  then
		tmp_user = " and F_isAdmin=1"
	else
		tmp_user = ""
	end if
	if trim(Request.QueryString("isAdmin"))="1"  then
		tmp_admin = " and F_isAdmin=1"
	elseif trim(Request.QueryString("isUser"))="0"  then
		tmp_admin = " and F_isAdmin=0"
	else
		tmp_admin = ""
	end if
	Set obj_fl_Rs = server.CreateObject(G_FS_RS)
	SQL = "Select  id,F_Name,F_Type,F_Url,ClassID,F_Author,addtime,F_isAdmin,F_Lock,F_IsUser  from FS_FL_FrendList where id>0 "& tmp_type & tmp_admin & tmp_user & tmp_lock & tmp_classid &" Order by F_OrderID desc,id desc"
	obj_fl_Rs.Open SQL,Conn,1,1
	if obj_fl_Rs.eof then
	   obj_fl_Rs.close
	   set obj_fl_Rs=nothing
	   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">没有友情连接。</td></tr>"
	else
		obj_fl_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo>obj_fl_Rs.PageCount Then cPageNo=obj_fl_Rs.PageCount 
		If cPageNo<=0 Then cPageNo=1
		obj_fl_Rs.AbsolutePage=cPageNo
		for i=1 to obj_fl_Rs.pagesize
			if obj_fl_Rs.eof Then exit For 
	%>
    <tr> 
      <td class="hback"><a href="<% = obj_fl_Rs("F_Url") %>" target="_blank"><% = obj_fl_Rs("F_Name") %></a></td>
      <td class="hback"  align="center">
	  <%
	  dim class_obj_rs
	  set  class_obj_rs = Conn.execute("select F_ClassCName from FS_FL_Class where id="&obj_fl_Rs("Classid"))
	  if Not class_obj_rs.eof then
		  Response.Write "<a href=Flink_Manage.asp?Classid="& obj_fl_Rs("Classid")&">"&class_obj_rs("F_ClassCName")&"</a>"
	  else
		  Response.Write "-----"
	  End if
	  %></td>
      <td class="hback" align="center"><%if obj_fl_Rs("F_Type")=0 then response.Write("图片"):else:response.Write("文字"):end if %></td>
      <td class="hback"> <div align="center">
          <%
	  if  obj_fl_Rs("F_isAdmin") =1 then
	  	Response.Write "管理员:" &obj_fl_Rs("F_Author") 
	  else
	  	if obj_fl_Rs("F_IsUser") =1 then
			dim User_Obj_Rs
			set User_Obj_Rs = User_Conn.execute("select UserName from Fs_me_Users where UserNumber = '"& NoSqlHack(obj_fl_Rs("F_Author"))&"'")
			if  Not User_Obj_Rs.eof then
				Response.Write "<a href=""../../"& G_USER_DIR &"/ShowUser.asp?UserNumber="& obj_fl_Rs("F_Author") &""">" &User_Obj_Rs("UserName") &"</a>"
	  		else
				Response.Write obj_fl_Rs("F_Author")
			end if
			User_Obj_Rs.close:set User_Obj_Rs = nothing
	  	Else
			Response.Write  obj_fl_Rs("F_Author")
		End if
	  end if
	  %></div></td>
      <td class="hback"><div align="center"><a href="Flink_Edit.asp?id=<% = obj_fl_Rs("id") %>&Action=edit">修改</a>┆<a href="Flink_Manage.asp?id=<% = obj_fl_Rs("id") %>&Act=del" onClick="{if(confirm('确定清除您所选择的记录吗？')){return true;}return false;}">删除</a>
<%if obj_fl_Rs("F_Lock")=0 then%>
          ┆<a href="Flink_Manage.asp?act=lock&id=<%=obj_fl_Rs("id")%>" onClick="{if(confirm('确定锁定吗？')){return true;}return false;}">未锁定</a>
<%else%>
          ┆<a href="Flink_Manage.asp?act=unlock&id=<%=obj_fl_Rs("id")%>" onClick="{if(confirm('确定解锁吗？')){return true;}return false;}"><font color="#FF0000">已锁定</font></a> 
          <%end if%>
          <input name="Id" type="checkbox" id="Id" value="<% = obj_fl_Rs("id") %>">
        </div></td>
    </tr>
	 <%
			obj_fl_Rs.movenext
		Next
	 %>
    <tr> 
	<td colspan="5" class="hback"><table width="100%">
	  <tr><td width="42%" align="left">
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);">
          选中/取消所有
          <input name="act" type="hidden" id="act3">
          <input type="button" name="Submit4" value="批量锁定" onClick="document.myForm.act.value='lock';{if(confirm('确定所选择的记录审核吗？')){this.document.myForm.submit();return true;}return false;}">
          <input type="button" name="Submit42" value="通过审核" onClick="document.myForm.act.value='unlock';{if(confirm('确定所选择的记录取消审核吗？')){this.document.myForm.submit();return true;}return false;}">
          <input type="button" name="Submit" value="删除"  onClick="document.myForm.act.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.myForm.submit();return true;}return false;}">
        </td><td width="58%" align="right"><%
			response.Write "<p>"&  fPageCount(obj_fl_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	  %></td></tr></table></td>
    
    <%end if%>
  </form>
</table>

<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
</script>
</body>
</html>






