<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr,sRootDir,str_CurrPath
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("FL002") then Err_Show
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
if trim(Request.Form("edit_save"))<>"" then
	dim obj_save_Rs
	if trim(Request.Form("F_Name"))="" or trim(Request.Form("F_Type"))="" or trim(Request.Form("F_Url"))="" or trim(Request.Form("F_Url"))="http://" or trim(Request.Form("F_Content"))="" then
			strShowErr = "<li>带*的必须填写</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	set obj_save_Rs = Server.CreateObject(G_FS_RS)
	if Request.Form("edit_save")="edit" then
		obj_save_Rs.open "select * from FS_FL_FrendList where id="&CintStr(Request.Form("id")),Conn,1,3
	else
		obj_save_Rs.open "select * from FS_FL_FrendList where 1=2",Conn,1,3
		obj_save_Rs.addnew
	end if
	obj_save_Rs("F_Name") = NoSqlHack(Request.Form("F_Name"))
	if Request.Form("F_Type")="0" then:obj_save_Rs("F_Type") =0:else:obj_save_Rs("F_Type") =1:end if
	obj_save_Rs("F_Url") = NoSqlHack(Request.Form("F_Url"))
	obj_save_Rs("F_Content") = NoSqlHack(Request.Form("F_Content"))
	if Request.Form("F_Type")="0" then:obj_save_Rs("F_PicUrl") = NoSqlHack(Request.Form("F_PicUrl")):end if
	if trim(Request.Form("F_Lock"))<>"" then:obj_save_Rs("F_Lock") =1:else:obj_save_Rs("F_Lock") = 0:end if
	obj_save_Rs("F_OrderID") = clng(Request.Form("F_OrderID"))
	if trim(Request.Form("F_IsUser"))<>"" then:obj_save_Rs("F_IsUser") = 1:else:obj_save_Rs("F_IsUser") = 0:end if
	if Request.Form("edit_save")="edit" then
		obj_save_Rs("F_Author") = NoSqlHack(Request.Form("F_Author"))
	else
		obj_save_Rs("F_Author") = session("Admin_Name")
	end if  	
	obj_save_Rs("F_Mail") = NoSqlHack(Request.Form("F_Mail"))
	obj_save_Rs("F_Content_des") = NoSqlHack(Request.Form("F_Content_des"))
	obj_save_Rs("F_LinkContent") = NoSqlHack(Request.Form("F_LinkContent"))
	IF trim(Request.Form("F_IsUser"))<>"" Then	
		obj_save_Rs("F_isAdmin") = 0
	Else
		obj_save_Rs("F_isAdmin") = 1
	End If	
	if trim(Request.Form("ClassID"))="" then:obj_save_Rs("ClassID") = 0:else:obj_save_Rs("ClassID") = Request.Form("ClassID"):end if
	obj_save_Rs("Addtime") = NoSqlHack(Request.Form("Addtime"))
	obj_save_Rs.update
	obj_save_Rs.close:set obj_save_Rs = nothing
	strShowErr = "<li>保存成功</li>"
	Response.Redirect("../success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Flink/Flink_Manage.asp")
	Response.end
end if
if Request.QueryString("Action")="edit" then
		dim tmp_id,tmp_F_Name,tmp_F_Type,tmp_F_Url,tmp_F_Content,tmp_F_PicUrl,tmp_F_Lock,tmp_F_OrderID,tmp_save,tmp_str,obj_edit_Rs
		dim tmp_F_IsUser,tmp_F_Author,tmp_F_Mail,tmp_F_Content_des,tmp_F_LinkContent,tmp_F_isAdmin,tmp_ClassID,tmp_Addtime
		if trim(Request.QueryString("id"))=empty or not isnumeric(trim(Request.QueryString("id"))) then
			strShowErr = "<li>错误参数</li>"
			Response.Redirect("../success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		set obj_edit_Rs = Server.CreateObject(G_FS_RS)
		obj_edit_Rs.open "select * from FS_FL_FrendList where id="&CintStr(Request.QueryString("id")),Conn,1,3
		if obj_edit_Rs.eof then
			strShowErr = "<li>此记录不存在</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			tmp_id=obj_edit_Rs("id")
			tmp_F_Name=obj_edit_Rs("F_Name")
			tmp_F_Type=obj_edit_Rs("F_Type")
			tmp_F_Url=obj_edit_Rs("F_Url")
			tmp_F_Content=obj_edit_Rs("F_Content")
			tmp_F_PicUrl=obj_edit_Rs("F_PicUrl")
			tmp_F_Lock=obj_edit_Rs("F_Lock")
			tmp_F_OrderID=obj_edit_Rs("F_OrderID")
			tmp_F_IsUser=obj_edit_Rs("F_IsUser")
			tmp_F_Author=obj_edit_Rs("F_Author")
			tmp_F_Mail=obj_edit_Rs("F_Mail")
			tmp_F_Content_des=obj_edit_Rs("F_Content_des")
			tmp_F_LinkContent=obj_edit_Rs("F_LinkContent")
			tmp_F_isAdmin=obj_edit_Rs("F_isAdmin")
			tmp_ClassID=obj_edit_Rs("ClassID")
			tmp_Addtime=obj_edit_Rs("Addtime")
			tmp_save="edit"
		end if
elseif Request.QueryString("Action")="Add" then
			tmp_id=""
			tmp_F_Name=""
			tmp_F_Type=0
			tmp_F_Url="http://"
			tmp_F_Content=""
			tmp_F_PicUrl=""
			tmp_F_Lock=0
			tmp_F_OrderID=0
			tmp_F_IsUser=0
			tmp_F_Author=""
			tmp_F_Mail=""
			tmp_F_Content_des=""
			tmp_F_LinkContent=""
			tmp_F_isAdmin=1
			tmp_ClassID=""
			tmp_Addtime=now
			tmp_save="add"
			tmp_str = "readonly"
else
	response.Write("错误的参数"):response.end
end if
%>
<html>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td align="left" colspan="2" class="xingmu">友情连接管理</td>
  </tr>
  <tr> 
    <td align="left" colspan="2" class="hback"><a href="Flink_Manage.asp">管理首页</a>┆<a href="Flink_Edit.asp?Action=Add">添加连接</a>┆<a href="Flink_Manage.asp?Type=0">图片连接</a>┆<a href="Flink_Manage.asp?Type=1">文字连接</a>┆<a href="Flink_Manage.asp?Lock=1">已锁定</a>┆<a href="Flink_Manage.asp?Lock=0">未锁定</a>┆用户申请的连接:<a href="Flink_Manage.asp?Lock=0&isUser=1">已审核</a>，<a href="Flink_Manage.asp?Lock=1&isUser=1">未审核</a>┆管理员添加的连接:<a href="Flink_Manage.asp?Lock=0&isAdmin=1">已审核</a>，<a href="Flink_Manage.asp?Lock=1&isAdmin=1">未审核</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="">
    <tr> 
      <td class="hback"><div align="right">类别</div></td>
      <td class="hback"> <select name="ClassID" id="ClassID">
          <option value="">不选择栏目</option>
          <%dim obj_list_rs
	  set obj_list_rs=conn.execute("select id,F_ClassCName from FS_FL_Class order by id desc")
	  do while not obj_list_rs.eof
	  	if tmp_classid = obj_list_rs("id") then
			Response.Write "<option value="""& obj_list_rs("id") &""" selected>"& obj_list_rs("F_ClassCName") &"</option>"
		else
			Response.Write "<option value="""& obj_list_rs("id") &""">"& obj_list_rs("F_ClassCName") &"</option>"
		end if
		  obj_list_rs.movenext
	  loop
	  obj_list_rs.close:set obj_list_rs = nothing
	  %>
        </select></td>
    </tr>
    <tr> 
      <td width="27%" class="hback"><div align="right">站点名称</div></td>
      <td width="73%" class="hback"><input name="F_Name" type="text" id="F_Name" value="<% = tmp_F_Name%>" size="40">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">连接类型</div></td>
      <td class="hback"> <input type="radio" name="F_Type" value="1" <%if tmp_F_Type = 1 then response.Write("checked")%>>
        文字连接 <input type="radio" name="F_Type" value="0" <%if tmp_F_Type = 0 then response.Write("checked")%>>
        图片连接*</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">连接地址</div></td>
      <td class="hback"><input name="F_Url" type="text" id="F_Url" value="<% = tmp_F_Url%>" size="40">
        *</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">站点说明</div></td>
      <td class="hback"><textarea name="F_Content" rows="5" id="F_Content" style="width:60%"><% = tmp_F_Content%></textarea>
        最多300个字符*</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">图片地址</div></td>
      <td class="hback"><input name="F_PicUrl" type="text" id="F_PicUrl" size="42" value="<% = tmp_F_PicUrl%>"> 
        <input type="button" name="PPPChoose"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.form1.F_PicUrl);"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">是否锁定</div></td>
      <td class="hback"><input name="F_Lock" type="checkbox" id="F_Lock" value="1" <%if tmp_F_Lock = 1 then response.Write("checked")%>>
        是</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(id_fl);"  language=javascript><a href="###">高级选项</a></td>
    </tr>
    <tr id="id_fl" style="display:none;">
      <td colspan="2" class="hback"> <table width="100%" border="0" cellpadding="5" cellspacing="1" class="table" >
          <tr> 
            <td width="26%" class="hback"><div align="right">权重（排列使用）</div></td>
            <td width="74%" class="hback"><input name="F_OrderID" type="text" id="F_OrderID" value="<% = tmp_F_OrderID%>"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">前台会员申请</div></td>
            <td class="hback"><input name="F_IsUser" type="checkbox" id="F_IsUser" value="1" <%if tmp_F_IsUser= 1 then response.Write("checked")%>>
              是</td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">申请人作者</div></td>
            <td class="hback"><input name="F_Author" type="text" id="F_Author" value="<% = tmp_F_Author%>"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">申请人电子邮件</div></td>
            <td class="hback"><input name="F_Mail" type="text" id="F_Mail" value="<% = tmp_F_Mail%>"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">申请理由</div></td>
            <td class="hback"><textarea name="F_Content_des" rows="5" id="F_Content_des" style="width:60%"><% = tmp_F_Content_des%></textarea>
              最多300个字符</td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">其他联系方式</div></td>
            <td class="hback"><textarea name="F_LinkContent" rows="5" id="textarea2" style="width:60%"><% = tmp_F_LinkContent%></textarea>
              最多300个字符</td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">申请日期</div></td>
            <td class="hback"><input name="Addtime" type="text" id="Addtime" value="<% = tmp_Addtime%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input type="submit" name="Submit" value="保存信息"> <input type="reset" name="Submit2" value="重置"> 
        <input name="id" type="hidden" id="id" value="<% = tmp_id %>"> <input name="edit_save" type="hidden" id="edit_save" value="<% = tmp_save %>"></td>
    </tr>
  </form>
</table>

</body>
</html>






