<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF 
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001")
'得到会员组列表
dim Fs_news
set Fs_news = new Cls_News
Fs_News.GetSysParam()
Dim str_StyleName,txt_Content,dmt_time,strShowErr,Lableclass_SQL,obj_Lableclass_rs
str_StyleName = NoSqlHack(Trim(Request.Form("StyleName")))
txt_Content = NoSqlHack(Trim(Request.Form("TxtFileds")))
if Request.Form("Action") = "add_save" then
		if str_StyleName ="" or txt_Content ="" then
			strShowErr = "<li>所有都是必须填写的</li><li>请重新填写</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		Lableclass_SQL = "Select StyleName,Content,AddDate from FS_NS_Labestyle where StyleName ='"& str_StyleName &"'"
		Set obj_Lableclass_rs = server.CreateObject(G_FS_RS)
		obj_Lableclass_rs.Open Lableclass_SQL,Conn,1,3
		if obj_Lableclass_rs.eof then
			obj_Lableclass_rs.addnew
			obj_Lableclass_rs("StyleName") = str_StyleName
			obj_Lableclass_rs("content") = txt_Content
			obj_Lableclass_rs("AddDate") =now
			obj_Lableclass_rs.update
		else
			strShowErr = "<li>此样式名称重复,请重新输入</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End if
		obj_Lableclass_rs.close:set obj_Lableclass_rs =nothing
		strShowErr = "<li>样式添加成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
		Response.end
Elseif Request.Form("Action") = "edit_save" then
		Lableclass_SQL = "Select StyleName,Content,AddDate from FS_NS_Labestyle where id ="& CintStr(Request.Form("ID")) 
		Set obj_Lableclass_rs = server.CreateObject(G_FS_RS)
		obj_Lableclass_rs.Open Lableclass_SQL,Conn,1,3
		if not obj_Lableclass_rs.eof then
			obj_Lableclass_rs("StyleName") = str_StyleName
			obj_Lableclass_rs("content") = txt_Content
			'obj_Lableclass_rs("AddDate") =now
			obj_Lableclass_rs.update
		End if
		obj_Lableclass_rs.close:set obj_Lableclass_rs =nothing
		strShowErr = "<li>样式修改成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
		Response.end
End if
if Request.QueryString("Action") = "del" then
	if Request.QueryString("id") = "" or isnumeric(Request.QueryString("id"))=false then
		strShowErr = "<li>错误的参数</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
		Response.end
	Else
		Conn.execute("Delete from FS_NS_Labestyle where id ="&CintStr(Request.QueryString("id")))
	End if
		strShowErr = "<li>删除样式成功</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Lable_style.asp")
		Response.end
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">标签管理<a href="../../help?Lable=NS_Lable_Manage" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Lable_Manage.asp">管理首页</a> 
        &nbsp;|&nbsp; <a href="Lable_Create.asp?ClassID=<%=Request.QueryString("ClassID")%>">创建标签</a> 
        &nbsp;|&nbsp; <a href="Lable_Dir_add.asp#Add">添加标签栏目</a> &nbsp;|&nbsp; 
        <a href="Lable_style.asp">标签样式管理</a> &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable"  onClick="{if(confirm('确认删除吗!')){return true;}return false;}">删除所有标签</a> 
        &nbsp;|&nbsp; <a href="Lable_Manage.asp?Action=del_lable_dir"  onClick="{if(confirm('确认删除吗!')){return true;}return false;}">删除所有标签目录</a> 
        | <a href="../../help?Lable=NS_Lable_Manage_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="34%" class="xingmu"> <div align="center">样式名称 </div></td>
    <td width="17%" class="xingmu"><div align="center">查看</div></td>
    <td width="29%" class="xingmu"><div align="center">添加日期</div></td>
    <td width="20%" class="xingmu"><div align="center">操作</div></td>
  </tr>
  <%
		Dim list_SQL,obj_List_rs
		list_SQL = "Select top 20 id,StyleName,Content,AddDate from FS_NS_Labestyle Order by Id desc"
		Set obj_List_rs = server.CreateObject(G_FS_RS)
		obj_List_rs.Open list_SQL,Conn,1,3
		do while not obj_List_rs.eof 
	%>
  <tr class="hback"> 
    <td> · <a href="Lable_style.asp?id=<% = obj_List_rs("id")%>&Action=edit#add"><% = obj_List_rs("StyleName")%></a></td>
    <td  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(sid<% = obj_List_rs("ID")%>);"  language=javascript><div align="center">查看</div></td>
    <td><% = obj_List_rs("adddate")%></td>
    <td><div align="center"><a href="Lable_style.asp?id=<% = obj_List_rs("id")%>&Action=edit#add">修改</a>｜<a href="Lable_style.asp?id=<% = obj_List_rs("id")%>&Action=del" onClick="{if(confirm('确认删除吗!!')){return true;}return false;}">删除</a></div></td>
  </tr>
  <tr class="hback" id="sid<% = obj_List_rs("ID")%>" style="display:none"> 
    <td height="48" colspan="4"> 
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback">
            <% = obj_List_rs("Content")%>
          </td>
        </tr>
      </table></td>
  </tr>
  <%
	  obj_List_rs.movenext
  Loop
  obj_List_rs.close
  set  obj_List_rs = nothing
  %>
</table>
<%
if Request.QueryString("Action")="edit" then
	Dim tmp_obj,str_StyleName_e,str_Content_e,str_add,str_id
	set tmp_obj = Conn.execute("select id,StyleName,Content,adddate from FS_NS_Labestyle where id="&CintStr(Request.QueryString("id")))
	if Not tmp_obj.eof then
		str_StyleName_e = tmp_obj("StyleName")
		str_Content_e = tmp_obj("Content")
		str_id = tmp_obj("id")
	End if
	str_add = "edit_save"
Else
	str_add = "add_save"
End if
%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="xingmu"> 
    <td colspan="2" class="xingmu"><a name="Add" id="Add"></a>创建样式 最多允许建立20个样式</td>
  </tr>
  <form name="Lable_Form" method="post" action="">
    <tr class="hback"> 
      <td width="13%"> <div align="right"> 样式名称</div></td>
      <td width="87%"><input name="StyleName" type="text" id="StyleName" value="<% = str_StyleName_e %>" size="40">
        <input name="id" type="hidden" id="id" value="<% = str_id %>"></td>
    </tr>
    <tr class="hback"> 
      <td><div align="right">插入字段</div></td>
      <td>
	  <select name="NewsFields" style="width:50%">
          <option value="{ID}">自动编号</option>
          <option value="{NewsID}">NewsID</option>
          <option value="{NewsTitle}"> 
          <% = Fs_news.allInfotitle %>
          标题</option>
          <option value="{CurtTitle}"> 
          <% = Fs_news.allInfotitle %>
          副标题</option>
          <option value="{NewsNaviContent}"> 
          <% = Fs_news.allInfotitle %>
          导读</option>
          <option value="{Content}"> 
          <% = Fs_news.allInfotitle %>
          内容</option>
          <option value="{AddTime}"> 
          <% = Fs_news.allInfotitle %>
          添加日期</option>
          <option value="{Author}"> 
          <% = Fs_news.allInfotitle %>
          作者</option>
          <option value="{Editer}"> 
          <% = Fs_news.allInfotitle %>
          责任编辑</option>
          <option value="" style="background:#DEDEDE">----下面是 
          <% = Fs_news.allInfotitle %>
          自定义字段----</option>
          <!--辅助字段信息-->
          <!--AuxiColumnStrNews-->
          <!--辅助字段信息-->
          <option value="" style="background:#DEDEDE">----下面是预定义字段----</option>
          <option value="{hits}">点击数</option>
          <option value="{KeyWords}">关键字</option>
          <option value="{TxtSource}"> 
          <% = Fs_news.allInfotitle %>
          来源</option>
          <option value="{SmallPicPath}">图片 
          <% = Fs_news.allInfotitle %>
          的图片地址(小图)</option>
          <option value="{PicPath}">图片 
          <% = Fs_news.allInfotitle %>
          的图片地址(大图)</option>
          <option value="{FormReview}">评论表单</option>
          <option value="{ReviewURL}">评论字样地址</option>
          <option value="{ShowComment}">显示评论列表</option>
          <option value="{AddFavorite}">加入收藏</option>
          <option value="{SendFriend}">发送给好友</option>
          <option value="{SpecialList}">所属专题列表</option>
          <option value="{NewsURL}"> 
          <% = Fs_news.allInfotitle %>
          访问路径</option>
          <option value="" style="background:#DEDEDE">----下面是栏目可定义字段----</option>
          <option value="{ClassName}">栏目中文名称</option>
          <option value="{ClassURL}">栏目访问路径</option>
          <option value="{ClassNaviPicURL}">栏目导航图片地址</option>
          <option value="" style="background:#DEDEDE">----下面是专题可定义字段----</option>
          <option value="{SpecialName}">栏目中文名称</option>
          <option value="{SpecialURL}">栏目访问路径</option>
          <option value="{SpecialNaviPicURL}">栏目导航图片地址</option>
          <!--option value="{SpecialNaviPicURL}">专题导航图片地址</option-->
        </select> <input name="button" type="button" onClick="insert(document.Lable_Form.NewsFields.value);" value=" 插 入 "></td>
    </tr>
    <tr class="hback"> 
      <td><div align="right">样式内容</div></td>
      <td><textarea name="TxtFileds" rows="15" id="TxtFileds" style="width:90%"><% = str_Content_e %></textarea></td>
    </tr>
    <tr class="hback"> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="保存样式"> 
        <input name="Action" type="hidden" id="Action" value="<% = str_add %>">
        <input type="reset" name="Submit2" value="重置"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
set Fs_news = nothing
%>
<script  language="JavaScript">  
function insert(insertContent)
{
		obj=document.getElementById("TxtFileds");
		obj.focus();
	if(document.selection==null)
	{
		var iStart = obj.selectionStart
		var iEnd = obj.selectionEnd;
		obj.value = obj.value.substring(0, iEnd) +insertContent+ obj.value.substring(iEnd, obj.value.length);
	}else
	{
		var range = document.selection.createRange();
		range.text+=insertContent;
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
</script>  
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





