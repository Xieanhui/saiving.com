<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,obj_Label_Rs,SQL,strShowErr,str_CurrPath,sRootDir
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
MF_Session_TF 
Dim GetSysConfigObj,GetStyleMaxNum
Set GetSysConfigObj = New Cls_SysConfig
GetSysConfigObj.getSysParam()
GetStyleMaxNum = Clng(GetSysConfigObj.Style_MaxNum)
Set GetSysConfigObj = Nothing
Dim str_StyleName,txt_Content,Labelclass_SQL,obj_Labelclass_rs,obj_Count_rs
Dim Label_Sub

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo,tmp_LableClassID,LableClassID
int_RPP=50 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页

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
Label_Sub = NoSqlHack(Request.QueryString("Label_Sub"))
str_StyleName = NoSqlHack(Trim(Request.Form("StyleName")))
txt_Content = Trim(Request.Form("TxtFileds"))
'txt_Content = replace(txt_Content,"{DS_","{DS:")
if Request.Form("Action") = "Add_save" then
	if str_StyleName ="" or txt_Content ="" then
		strShowErr = "<li>请填写完整</li>"
		Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	set obj_Count_rs = server.CreateObject(G_FS_RS)
	obj_Count_rs.Open "Select StyleName,Content,AddDate,LableClassID from FS_MF_Labestyle Order by id desc",Conn,1,3
	if Not obj_Count_rs.eof then
		if obj_Count_rs.recordcount>GetStyleMaxNum then
			strShowErr = "<li>您建立的样式已经超过" & GetStyleMaxNum & "个。你将不能再增加\n如果需要增加，请删除部分样式</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Action=Add&Label_Sub="&Request.Form("Label_Sub")&"")
			Response.end
		end if
	end if
	Labelclass_SQL = "Select StyleName,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where StyleName ='"& str_StyleName &"'"
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if obj_Labelclass_rs.eof then
		obj_Labelclass_rs.addnew
		obj_Labelclass_rs("StyleName") = str_StyleName
		obj_Labelclass_rs("content") = txt_Content
		obj_Labelclass_rs("AddDate") =now
		obj_Labelclass_rs("StyleType") =Request.Form("Label_Sub")
		obj_Labelclass_rs("LableClassID") =Request.Form("LableClassID")'---写入数据库-------2/1 by chen---------
		'tmp_LableClassID=obj_Labelclass_rs("LableClassID")
		obj_Labelclass_rs.update
	else
			strShowErr = "<li>此样式名称重复,请重新输入</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	strShowErr = "<li>添加成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
elseif Request.Form("Action") = "Add_edit" then
	Labelclass_SQL = "Select StyleName,Content,AddDate,LableClassID from FS_MF_Labestyle where id ="& NosqlHack(Request.Form("ID")) 
	Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
	obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
	if not obj_Labelclass_rs.eof then
		obj_Labelclass_rs("StyleName") = str_StyleName
		obj_Labelclass_rs("content") = txt_Content
		'obj_Labelclass_rs("AddDate") =now
	obj_Labelclass_rs("LableClassID")=Request.Form("LableClassID")'--------写入数据库----2/1 by chen--------------------
	'tmp_LableClassID=obj_Labelclass_rs("LableClassID")
		obj_Labelclass_rs.update
	End if
	obj_Labelclass_rs.close:set obj_Labelclass_rs =nothing
	strShowErr = "<li>修改成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
end if
if Request.QueryString("DelTF")="1" then
	Conn.execute("Delete From FS_MF_Labestyle where StyleType='"& NoSqlHack(Request.QueryString("Label_Sub"))&"' and id="&CintStr(Request.QueryString("id")))
	strShowErr = "<li>删除成功</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Label/All_Label_style.asp?Label_Sub="& Request.Form("Label_Sub")&"")
	Response.end
end if
%>
<html>
<style type="text/css">
<!--
.STYLE1 {color: #0066FF}
-->
</style>
<head>
<title>标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<body>
<table width="98%" height="81" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" >
    <td width="100%" height="20"  align="Left" class="xingmu">引用样式管理</td>
  </tr>
  <tr class="hback" >
    <td class="hback" align="center"><div align="left"><a href="../Templets_List.asp">模板管理</a>
        <%
	if Request.QueryString("TF") = "NS" then
		Response.Write("┆<a href=News_Label.asp target=_self>返回标签管理</a>") 
	elseif Request.QueryString("TF") = "DS" then
		Response.Write("┆<a href=Down_Label.asp target=_self>返回标签管理</a>") 
	elseif Request.QueryString("TF") = "SD" then
		Response.Write("┆<a href=supply_Label.asp target=_self>返回标签管理</a>") 
	elseif Request.QueryString("TF") = "HS" then
		Response.Write("┆<a href=House_Label.asp target=_self>返回标签管理</a>") 
	elseif Request.QueryString("TF") = "AP" then
		Response.Write("┆<a href=job_Label.asp target=_self>返回标签管理</a>") 
	elseif Request.QueryString("TF") = "MS" then
		Response.Write("┆<a href=Mall_Label.asp target=_self>返回</a>")
	else
		Response.Write("") 
	end if
	%><!--增加文本编辑 by sicend -->
        ┆<a href="All_Label_Style.asp" target="_self">所有样式</a>┆<a href="Label_Style_Class.asp" target="_self">创建样式分类</a> <a href="../../help?Label=MF_Label_Creat" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a>┆<a href="javascript:history.back();">后退</a></div></td>
  </tr>
  <tr class="hback" >
    <td valign="top" class="hback"><strong>分类:</strong> <a href="All_Label_style.asp?Label_Sub=NS" title="浏览 新闻系统 标签引用样式" target="_self">新闻系统</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=NS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建新闻系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=NS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>┆ <a href="All_Label_style.asp?Label_Sub=DS" title="浏览 下载系统 标签引用样式" target="_self">下载系统</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=DS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建下载系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=DS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
      ┆
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=SD" title="浏览 供求系统 标签引用样式" target="_self">供求系统</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=SD&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建供求系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=SD&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
      ┆
      <%end if%>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=HS" title="浏览 房产楼盘 标签引用样式" target="_self">房产楼盘</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=HS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建房产系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=HS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>
    <%end if%>┆
	  <a href="All_Label_style.asp?Label_Sub=CForm" title="浏览 自定义表单 标签引用样式" target="_self">自定义表单</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=CForm&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建会员登陆标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=CForm&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a></td>
  </tr>
  <tr class="hback" >
    <td valign="top" class="hback"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=AP" title="浏览 招聘求职 标签引用样式" target="_self">招聘求职</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=AP&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建人才系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=AP&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>┆
      <%End if%>
      <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then%>
      <a href="All_Label_style.asp?Label_Sub=MS" title="浏览 商城B2C 标签引用样式" target="_self">商城B2C</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=MS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建商城系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=MS&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>┆
      <%end if%>
      <a href="All_Label_style.asp?Label_Sub=ME" title="浏览 会员系统 标签引用样式" target="_self">会员系统</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=ME&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建会员系统标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=ME&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a>┆
	  <a href="All_Label_style.asp?Label_Sub=Login" title="浏览 会员登陆 标签引用样式" target="_self">会员登陆</a><a href="All_Label_Style.asp?Action=Add&Label_Sub=Login&ClassID=<%= Request.QueryString("ClassId") %>" title="采用编辑器模式创建会员登陆标签引用样式表" target="_self"><img src="../Images/addstyle.gif" border="0"></a><a href="All_Label_Style_txt.asp?Action=Add&Label_Sub=Login&ClassID=<%= Request.QueryString("ClassId") %>" title="采用文本模式创建新闻系统标签引用样式表" target="_self"><img border="0" src="../Images/addstyletxt.gif"></a></td>
  </tr>
</table>
  <form name="Label_Form" method="get" action="" target="_self" style="margin:0;padding:0;" onSubmit="return false;">
        &nbsp;  搜索样式：<input type="text" id="key" name="keyw" value="<% = trim(Request("key")) %>" /><input type="button" name="se" value="搜索样式" onClick="searcha();" />
  </form>
<script type="text/javascript">
    function searcha()
       {
            if(document.getElementById("key").value=="")
            {
                alert("填写关键字");
            return false;
            } 
            window.location.href="all_label_style.asp?key="+escape(document.getElementById("key").value)+"&Label_Sub=<%=request("Label_Sub") %>";
       } 
</script>
<!------2/1 by chen 控制所建立的样式的分类文件夹 并读取建立的样式------------------------------------>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
 <tr>
    <td width="8%" class="xingmu"><div align="center">样式ID号</div></td>
    <td width="32%" class="xingmu"><div align="center">样式名称</div></td>
    <td width="25%" class="xingmu"><div align="center">引用查看</div></td>
	<td width="14%" class="xingmu"><div align="center">所属系统</div></td>
	<td width="21%" class="xingmu"><div align="center">描述/操作</div></td>
</tr>
  <%
  dim rs_class,str_ParentID
  if trim(NoSqlHack(Request.QueryString("ClassID")))<>"" then
		str_ParentID = " and ParentID="&NoSqlHack(Request.QueryString("ClassID"))&""
  elseif not isnumeric(trim(NoSqlHack(Request.QueryString("ClassID")))) then
		str_ParentID = " and ParentID=0"
  else
		str_ParentID = " and ParentID=0"
  end if
  set rs_class=Conn.execute("select id,ClassName,ClassContent,ParentID From FS_MF_StyleClass where 1=1"&str_ParentID&" order by id desc")
  do while not rs_class.eof 
  %>
  <tr class="hback">
    <td valign="top"><div align="center"><img src="../Images/Folder/folder.gif" alt="文件夹" width="20" height="16"></div></td>
    <td><a href="All_Label_Style.asp?ClassId=<% = rs_class("id")%>&ParentID=<%=rs_class("id")%>">
      <% = rs_class("ClassName")%>
      </a></td>
    <td><% = rs_class("ClassContent")%></td>
	<td></td>
	<td></td>
  </tr>
  <%
  rs_class.movenext
  loop
  rs_class.close:set rs_class = nothing
  %>
  <tr class="hback_1">
    <td colspan="7" height="2"></td>
  </tr>
  <%
	dim rs_stock,ClassId,LableClassID1
	if Request.QueryString("ClassId")<>"" then
		LableClassID1 = NoSqlHack(Request.QueryString("ClassID"))
	Else
		LableClassID1=0
	End if
	ClassId = " and LableClassID="&LableClassID1&""
	set rs_stock= Server.CreateObject(G_FS_RS)
	rs_stock.open "select ID,StyleName,Content,LableClassID From FS_MF_Labestyle Where ID > 0" & ClassId  &" order by ID desc",Conn,1,1
	if rs_stock.eof then
	   rs_stock.close
	   set rs_stock=nothing
	   Response.Write"<TR  class=""hback""><TD colspan=""7""  class=""hback"" height=""40"">没有记录。</TD></TR>"
	 end if
	%>
<!---------2/2 by chen------------------------------------------------------------------------------------------>
<%
	  Select Case Request.QueryString("Action")
		Case "Add"
			Call Add()
		Case "Add_Save"
			Call Add_Save()
		Case else
			Call Main()
	End Select
	Sub Main()
	%>
<%
			dim tmp_Label_Sub,LableClassID
			Set obj_Label_Rs = server.CreateObject(G_FS_RS)
			if trim(Label_Sub) <>"" then:tmp_Label_Sub = "and StyleType='"& Label_Sub &"'":else:tmp_Label_Sub = "":end if
			if Request.QueryString("ClassId")<>"" then
				LableClassID = " and LableClassID = " & NoSqlHack(Request.QueryString("ClassID"))
			Else
				LableClassID= " and (LableClassID < 1 Or LableClassID is Null)"
			End if
			dim keys,wh
	        keys = trim(Request("key"))
	        if keys<>"" then
	            wh = " and (StyleName like '%"+keys+"%' or Content like '%"+keys+"%')"
				LableClassID = ""
	        end if

			SQL = "Select  ID,StyleName,LoopContent,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where id>0 "& tmp_Label_Sub & wh & LableClassID & " Order by id desc"
			'response.Write(SQL)
			obj_Label_Rs.Open SQL,Conn,1,3
			If not obj_Label_Rs.Eof Then
				obj_Label_Rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>obj_Label_Rs.PageCount Then cPageNo=obj_Label_Rs.PageCount 
				If cPageNo<=0 Then cPageNo=1		
				obj_Label_Rs.AbsolutePage=cPageNo
			
				For int_Start=1 TO int_RPP  
			%>
  <tr class="hback">
<td width="8%" class="hback" align="center"><% = obj_Label_Rs("ID") %></td>
    <td width="32%" class="hback">· 
      <a href="All_Label_style.asp?Action=Add&type=edit&id=<%= obj_Label_Rs("id")%>&Label_Sub=<%= obj_Label_Rs("StyleType")%>&ClassID=<% If Request.QueryString("ClassId")<>"" And IsNumeric(Request.QueryString("ClassId")) Then : Response.Write Request.QueryString("ClassId") : Else : Response.Write "0" : End IF %>"  target="_self"><% = obj_Label_Rs("StyleName") %></a>
  [<a href="All_Label_style_txt.asp?Action=Add&type=edit&id=<%= obj_Label_Rs("id")%>&Label_Sub=<%= obj_Label_Rs("StyleType")%>&ClassID=<% If Request.QueryString("ClassId")<>"" And IsNumeric(Request.QueryString("ClassId")) Then : Response.Write Request.QueryString("ClassId") : Else : Response.Write "0" : End IF %>"  target="_self">文本编辑</a>]</td>
    <td width="25%" class="hback"   id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Label<%=obj_Label_Rs("ID")%>);" >引用样式查看</td>
    <td width="14%"><% = obj_Label_Rs("StyleType")%>
    </td>
    <td width="21%"><a href="All_Label_style.asp?id=<%=obj_Label_Rs("ID")%>&Label_Sub=<%=obj_Label_Rs("StyleType")%>&DelTF=1" onClick="{if(confirm('确定删除此引用样式吗？')){return true;}return false;}">删除</a> </td>
  </tr>
  <tr id="Label<%=obj_Label_Rs("ID")%>" style="display:none"  class="hback">
    <td height="42" colspan="7"  class="hback">
    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <tr>
          <td height="48" class="hback"><%
			Dim regEx,result
			Set regEx = New RegExp '
			regEx.Pattern = "<img(.+?){(.+?)}(.+?)>" '  
			regEx.IgnoreCase = true ' 
			regEx.Global = True '  
			result = regEx.replace(obj_label_Rs("Content"),"<img src='../images/default.png'/>") 
			Response.Write(result)
		  %>
          </td>
        </tr>
      </table></td>
  </tr>
  <%
				obj_Label_Rs.MoveNext
				If obj_Label_Rs.Eof or obj_Label_Rs.Bof Then Exit For
			Next
		End If
			%>
  <tr  class="hback">
    <td height="21" colspan="7"  class="hback STYLE1">注：点击样式标题将使用 编辑器模式修改样式，点击[文本编辑]将使用文本编辑模式修改标签！ </td>
  </tr>
</table>
<div align="center"></div>
<p>
  <%
	End Sub
	%>
  <%
	  Sub Add()
	  	  dim str_id,tmp_id,tmp_StyleName,tmp_Content,tmp_Action
		  str_id = NoSqlHack(Request.QueryString("id"))
		  if NoSqlHack(Request.QueryString("type"))="edit" then
		  	if NoSqlHack(Request("IsPostBack")) <> "1" then
				Set obj_Label_Rs = server.CreateObject(G_FS_RS)
				obj_Label_Rs.Open "Select  ID,StyleName,LoopContent,Content,AddDate,StyleType,LableClassID from FS_MF_Labestyle where id="& str_id &"",Conn,1,3
				tmp_id = obj_Label_Rs("id")
				tmp_StyleName = obj_Label_Rs("StyleName")
				tmp_Content = obj_Label_Rs("Content")
			end if
			tmp_Action = "Add_edit"
		  Else
			tmp_id = ""
			tmp_StyleName = ""
			tmp_Content = ""
			tmp_Action = "Add_save"
		  End if
		  if Request.Form("IsPostBack") = "1" then
			tmp_id = NoSqlHack(Request("id"))
			tmp_StyleName = NoSqlHack(Request("StyleName"))
			tmp_Content = NoSqlHack(Request("TxtFileds"))
		  end if
	%>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="xingmu">
    <td colspan="2" class="xingmu">创建标签引用样式(最多允许建立<span class="tx"><% = GetStyleMaxNum %></span>个样式)</td>
  </tr>
  <form name="Label_Form" method="post" action="" target="_self">
    <tr class="hback">
      <td width="13%"><div align="right"> 样式名称</div></td>
      <td width="87%"><input name="StyleName" type="text" id="StyleName" size="40" value="<% = tmp_StyleName %>">
        <input name="id" type="hidden" id="id" value="<% = tmp_id %>">
        <select name="LableClassID" id="LableClassID">
          <option value="0">选择所属栏目</option>
          <%
				  dim class_rs_obj
				  set class_rs_obj=Conn.execute("select id,ParentID,ClassName From FS_MF_StyleClass where ParentID=0 order by id desc")
				  do while not class_rs_obj.eof
						If CStr(Request.QueryString("ClassID"))=CStr(class_rs_obj("id")) Then 
							response.Write "<option value="""&class_rs_obj("id")&""" selected >"&class_rs_obj("ClassName")&"</option>"
						Else
							response.Write "<option value="""&class_rs_obj("id")&""">"&class_rs_obj("ClassName")&"</option>"
						End If 
						response.Write get_childList(class_rs_obj("id"),"")
					class_rs_obj.movenext
				  loop
				  class_rs_obj.close:set class_rs_obj=nothing
				  %>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td><div align="right">插入字段</div></td>
      <td><%
			Dim Label_Sub
			Label_Sub = NoSqlHack(Request.QueryString("Label_Sub"))
			select case Label_Sub
					case "NS"
						Call NS_select()
					case "DS"
						Call DS_select()
					case "SD"
						Call SD_select()
					case "HS"
						Call HS_select()
					case "AP"
						Call AP_select()
					case "MS"
						Call MS_select()
					case "ME"
						Call ME_select()
					Case "Login"
						Call ME_Login()	
					Case "CForm"
						Call MF_CustomForm()	
					case else
						Call NS_select()
			end select
			%>
    </tr>
    <tr class="hback" <%if request.QueryString("Label_Sub")<>"DS" then response.Write("style=""display:'none';""") else response.Write("style=""display:'';"" ") end if%>> </tr>
    <tr class="hback">
      <td><div align="right">样式内容</div></td>
      <td><!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=TxtFileds' frameborder=0 scrolling=no width='100%' height='440'></iframe>
				<input type="hidden" name="TxtFileds" value="<% = HandleEditorContent(tmp_Content)%>">
                <!--编辑器结束-->
        </td>
    </tr>
    <tr class="hback">
      <td>&nbsp;</td>
      <td><span class="tx">特别提醒。对于高级技术人员（对html相当熟悉的人员），你可以在标签引用样式里随心所欲的写html样式，达到灵活处理前台页面的效果；对于对html不太熟悉的人员，请查看帮助文档或者插入简单的标签样式就可以！</span></td>
    </tr>
    <tr class="hback">
      <td>&nbsp;</td>
      <td><%
			 if tmp_Action = "Add_save" then
				dim obj_Count_rs_1,tmp_str,tmp_display
				Set obj_Count_rs_1 = server.CreateObject(G_FS_RS)
				obj_Count_rs_1.Open "Select StyleName,Content,AddDate from FS_MF_Labestyle where StyleType='NS' Order by id desc",Conn,1,3
				if Not obj_Count_rs_1.eof then
					if obj_Count_rs_1.recordcount>GetStyleMaxNum then
						tmp_str = "--标签样式已经超过" & GetStyleMaxNum & "个,不能创建。"
						tmp_display = "disabled"
					Else
						tmp_str = ""
						tmp_display = ""
					End if
				Else
					tmp_str = ""
					tmp_display = ""
				End if
			 Else
					tmp_str = ""
					tmp_display = ""
			 End if
			  %>
        <input type="submit" name="Submit" value="以HTML保存样式<% = tmp_str %>"<% = tmp_display %> onClick="return Label_Form_sumit(this.form,1);">
        <input name="Action" type="hidden" id="Action" value="<% = tmp_Action %>" >
        <input name="Label_Sub" type="hidden" value="<%=Request.QueryString("Label_Sub")%>">
        <input type="hidden" name="IsPostBack" value="1">
        <input type="submit" name="Submit3" value="以XHTML保存样式<% = tmp_str %>"<% = tmp_display %> onClick="return Label_Form_sumit(this.form,0);">
        <input type="reset" name="Submit2" value="重置">
      </td>
    </tr>
  </form>

</table>
<script language="JavaScript" type="text/JavaScript">
function Label_Form_sumit(FormObj,IsHTML)
{
	if(FormObj.StyleName.value == "")
	{
		alert("请填写标签名称")
		FormObj.StyleName.focus();
		return false;
	}
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	FormObj.TxtFileds.value=frames["NewsContent"].GetNewsContentArray();
	return true;
}

function Insertlabel_Sel(Lable_obj)
{
	if(Lable_obj.options[Lable_obj.selectedIndex].value==''){
	return false;
	}else{
	InsertEditor(Lable_obj.options[Lable_obj.selectedIndex].value);
	}
}
function InsertEditor(InsertValue)
{
	InsertHTML(InsertValue,"NewsContent");
}
</script>
<%
End Sub
%>
<%
sub NS_select()
			'得到自定义字段
			dim ns_D_rs,ns_list
			ns_list = ""
			set ns_D_rs = Server.CreateObject(G_FS_RS)
			ns_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='NS'",Conn,1,3
			if ns_D_rs.eof then
				ns_list =ns_list& "<option value="""">没有自定义字段</option>"
				ns_D_rs.close:set ns_D_rs=nothing
			else
				do while not ns_D_rs.eof 
					ns_list = ns_list & "<option value=""{NS=Define|"&ns_D_rs("D_Coul")&"}"">"& ns_D_rs("D_Name")&"</option>"
					ns_D_rs.movenext
				loop
				ns_D_rs.close:set ns_D_rs=nothing
			end if
			%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄基本字段┄┄</option>
  <option value="{NS:FS_ID}">自动编号</option>
  <option value="{NS:FS_NewsID}">NewsID</option>
  <option value="<a href={NS:FS_NewsURL}>{NS:FS_NewsTitle}</a>">新闻标题(截断)</option>
  <option value="{NS:FS_NewsTitleAll}">新闻完整标题</option>
  <option value="{NS:FS_NewsURL}"> 新闻访问路径</option>
  <option value="{NS:FS_CurtTitle}"> 新闻副标题</option>
  <option value="{NS:FS_NewsNaviContent}"> 新闻导读</option>
  <option value="{NS:FS_Content}"> 新闻内容</option>
  <option value="{NS:FS_AddTime}"> 新闻添加日期</option>
  <option value="{NS:FS_Author}"> 新闻作者</option>
  <option value="{NS:FS_Editer}"> 新闻责任编辑</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄预定义字段┄┄</option>
  <option value="{NS:FS_hits}">点击数</option>
  <option value="{NS:FS_KeyWords}">关键字</option>
  <option value="{NS:FS_TxtSource}"> 新闻来源</option>
  <option value="{NS:FS_SmallPicPath}">图片新闻的图片地址(小图)</option>
  <option value="{NS:FS_PicPath}">图片新闻的图片地址(大图)</option>
  <option value="{NS:FS_FormReview}">评论表单</option>
  <option value="{NS:FS_ReviewURL}">评论字样(带地址)</option>
  <option value="{NS:FS_ShowComment}">显示评论列表</option>
  <option value="{NS:FS_AddFavorite}">加入收藏</option>
  <option value="{NS:FS_SendFriend}">发送给好友</option>
  <option value="{NS:FS_SpecialList}">所属专题列表</option>
  <option value="{NS:FS_PrevPage}"> 上一篇新闻</option>
  <option value="{NS:FS_NextPage}"> 下一篇新闻</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄专题可定义字段┄┄</option>
  <option value="{NS:FS_SpecialName}">专题中文名称</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄新闻自定义字段┄┄</option>
  <%=ns_list%>
</select>
<select name="SingleClassFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄栏目可定义字段┄┄</option>
  <option value="{NS:FS_ClassName}">栏目中文名称</option>
  <option value="{NS:FS_ClassURL}">栏目访问路径</option>
  <option value="{NS:FS_ClassNaviPicURL}">栏目导航图片地址</option>
  <option value="{NS:FS_ClassNaviDescript}">栏目导航说明</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄单页栏目可定义字段┄┄</option>
  <option value="{NS:FS_PageContent}">栏目内容</option>
  <option value="{NS:FS_Keywords}">栏目META关键字</option>
  <option value="{NS:FS_description}">栏目META描述</option>
</select>
</td>
<%end sub%>
<%sub DS_select()
			'得到自定义字段
			dim ds_D_rs,ds_list
			ds_list = ""
			set ds_D_rs = Server.CreateObject(G_FS_RS)
			ds_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='DS'",Conn,1,3
			if ds_D_rs.eof then
				ds_list =ds_list& "<option value="">没有自定义字段</option>"
				ds_D_rs.close:set ds_D_rs=nothing
			else
				do while not ds_D_rs.eof 
					ds_list = ds_list & "<option value=""{DS=Define|"&ds_D_rs("D_Coul")&"}"">"& ds_D_rs("D_Name")&"</option>"
					ds_D_rs.movenext
				loop
				ds_D_rs.close:set ds_D_rs=nothing
			end if%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄基本字段┄┄</option>
  <option value="{DS:FS_ID}">自动编号</option>
  <option value="{DS:FS_DownLoadID}">DownLoadID</option>
  <option value="<a href={DS:FS_DownURL}>{DS:FS_Name}</a>">下载标题(截断)</option>
  <option value="<a href={DS:FS_DownURL}>{DS:FS_NameAll}</a>">下载标题(完整)</option>
  <option value="{DS:FS_Description}">下载简介</option>
  <option value="{DS:FS_AddTime}">添加时间</option>
  <option value="{DS:FS_EditTime}">修改时间</option>
  <option value="{DS:FS_SystemType}">系统平台</option>
  <option value="{DS:FS_Accredit}">下载授权</option>
  <option value="{DS:FS_Version}">版本</option>
  <option value="{DS:FS_Appraise}">星级评价</option>
  <option value="{DS:FS_FileSize}">文件大小</option>
  <option value="{DS:FS_Language}">语言</option>
  <option value="{DS:FS_PassWord}">解压密码</option>
  <!--<option value="{DS:FS_Property}">下载性质</option>-->
  <option value="{DS:FS_Provider}">开发商</option>
  <option value="{DS:FS_ProviderUrl}">提供者Url地址</option>
  <option value="{DS:FS_EMail}">联系人EMAIL</option>
  <option value="{DS:FS_Types}">下载类型</option>
  <option value="{DS:FS_OverDue}">过期天数</option>
  <option value="{DS:FS_ConsumeNum}">消费点数</option>
  <option value="{DS:FS_Address$&amp;lt;br /&amp;gt;}">下载地址</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄预定义字段┄┄</option>
  <option value="{DS:FS_Hits}">点击数</option>
  <option value="{DS:FS_ClickNum}">下载次数</option>
  <option value="{DS:FS_Pic}">显示图片地址</option>
  <option value="{DS:FS_ReviewURL}">评论字样(带地址)</option>
  <option value="{DS:FS_FormReview}">评论表单</option>
  <option value="{DS:FS_ShowComment}">显示评论列表</option>
  <option value="{DS:FS_SpecialList}">所属专区列表</option>
  <option value="{DS:FS_AddFavorite}">加入收藏</option>
  <option value="{DS:FS_SendFriend}">发送给好友</option>
  <option value="{DS:FS_DownURL}">下载访问路径</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄栏目可定义字段┄┄</option>
  <option value="{DS:FS_ClassName}">栏目中文名称</option>
  <option value="{DS:FS_ClassURL}">栏目访问路径</option>
  <option value="{DS:FS_ClassNaviPicURL}">栏目导航图片地址</option>
  <option value="{DS:FS_ClassNaviDescript}">栏目导航说明</option>
  <option value="{DS:FS_ClassKeywords}">栏目关键字</option>
  <option value="{DS:FS_Classdescription}">栏目描述</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄专区可定义字段┄┄</option>
  <option value="{DS:FS_SpecialName}">专区中文名称</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄下载自定义字段┄┄</option>
  <%=ds_list%>
</select>
<%end sub%>
<%sub SD_select()
'得到自定义字段
			dim sd_D_rs,sd_list
			sd_list = ""
			set sd_D_rs = Server.CreateObject(G_FS_RS)
			sd_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='SD'",Conn,1,3
			if sd_D_rs.eof then
				sd_list =sd_list& "<option value="">没有自定义字段</option>"
				sd_D_rs.close:set sd_D_rs=nothing
			else
				do while not sd_D_rs.eof 
					sd_list = sd_list & "<option value=""{SD=Define|"&sd_D_rs("D_Coul")&"}"">"& sd_D_rs("D_Name")&"</option>"
					sd_D_rs.movenext
				loop
				sd_D_rs.close:set sd_D_rs=nothing
			end if%>
<select name="NewsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄基本字段┄┄</option>
  <option value="&lt;a href=&quot;{SD:FS_URL}&quot; target=_blank&gt;{SD:FS_title}">标题</option>
  <option value="{SD:FS_Alltitle}">完整标题(不截断)</option>
  <option value="{SD:FS_URL}">供求联接路径</option>
  <option value="{SD:FS_PubType}">类型</option>
  <option value="{SD:FS_PubTypeLink}">带连接的类型</option>
  <option value="{SD:FS_PubContent}">内容</option>
  <option value="{SD:FS_AreaID}">所属地区</option>
  <option value="{SD:FS_ClassID}">所属分类</option>
  <option value="{SD:FS_Keyword}">关键字</option>
  <option value="{SD:FS_CompType}">经营方式</option>
  <option value="{SD:FS_PubNumber}">产品数量</option>
  <option value="{SD:FS_PubPrice}">产品价格</option>
  <option value="{SD:FS_PubPack}">包装说明</option>
  <option value="{SD:FS_Pubgui}">产品规格</option>
  <option value="{SD:FS_PubPic_1}">图片一地址</option>
  <option value="{SD:FS_PubPic_2}">图片二</option>
  <option value="{SD:FS_PubPic_3}">图片三</option>
  <option value="{SD:FS_Addtime}">发布时间</option>
  <option value="{SD:FS_EditTime}">最后更新时间</option>
  <option value="{SD:FS_ValidTime}">有效时间</option>
  <option value="{SD:FS_PubAddress}">产地</option>
  <option value="{SD:FS_Fax}">联系传真</option>
  <option value="{SD:FS_User}">发布会员用户名</option>
  <option value="{SD:FS_tel}">联系电话</option>
  <option value="{SD:FS_Mobile}">移动电话</option>
  <option value="{SD:FS_otherLink}">其他联系方式</option>
  <option value="{SD:FS_hits}">点击</option>
  <!--option value="{SD:FS_ReviewURL}">评论字样(带地址)</option-->
  <option value="{SD:FS_review}">发表评论</option>
  <option value="{SD:FS_reviewcontent}">评论内容</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄供求自定义字段┄┄</option>
  <%=sd_list%>
</select>
</span>
<%end sub%>
<%sub HS_select()
				dim hs_D_rs,hs_list
				hs_list = ""
				set hs_D_rs = Server.CreateObject(G_FS_RS)
				hs_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='HS'",Conn,1,3
				if hs_D_rs.eof then
					hs_list =hs_list& "<option value="""">没有自定义字段</option>"
					hs_D_rs.close:set hs_D_rs=nothing
				else
					do while not hs_D_rs.eof 
						hs_list = hs_list & "<option value=""{HS=Define|"&hs_D_rs("D_Coul")&"}"">"& hs_D_rs("D_Name")&"</option>"
						hs_D_rs.movenext
					loop
					hs_D_rs.close:set hs_D_rs=nothing
				end if
			  %>
<select name="HouseFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄基本字段┄┄</option>
  <option value="{all:HS_FS_ID}">自动编号</option>
  <option value="{all:HS_FS_Price}">报价</option>
  <option value="{all:HS_FS_PubDate}">发布时间</option>
  <option value="{all:HS_FS_UserNumber}">发布者编号</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄预定义字段┄┄</option>
  <option value="{HS_FS_FormReview}">评论表单</option>
  <option value="{HS_FS_ReviewURL}">评论字样地址</option>
  <option value="{HS_FS_ShowComment}">显示评论列表</option>
  <option value="{HS_FS_AddFavorite}">加入收藏</option>
  <option value="{HS_FS_SendFriend}">发送给好友</option>
  <option value="{HS_FS_HouseURL}"> 信息访问路径</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄房产自定义字段┄┄</option>
  <%=hs_list%>
</select>
<select name="LouFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄楼盘信息┄┄</option>
  <option value="{Lou:HS_FS_HouseName}">楼盘名称</option>
  <option value="{Lou:HS_FS_KaiFaShang}">开发商</option>
  <option value="{Lou:HS_FS_Position}">楼盘位置</option>
  <option value="{Lou:HS_FS_Direction}">楼盘方位</option>
  <option value="{Lou:HS_FS_Class}">项目类别</option>
  <option value="{Lou:HS_FS_OpenDate}">开盘日期</option>
  <option value="{Lou:HS_FS_PreSaleRange}">预售范围</option>
  <option value="{Lou:HS_FS_Status}">房屋状况</option>
  <option value="{Lou:HS_FS_introduction}">房屋介绍</option>
  <option value="{Lou:HS_FS_Contact}"> 联系方式</option>
  <option value="{Lou:HS_FS_hits}">点击数</option>
  <option value="{Lou:HS_FS_Pic}">图片</option>
</select>
<select name="SecondFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄二手信息┄┄</option>
  <option value="{Second:HS_FS_UseFor}">用途</option>
  <option value="{Second:HS_FS_Label}">房屋编号</option>
  <option value="{Second:HS_FS_FloorType}">住宅类别</option>
  <option value="{Second:HS_FS_BelongType}">产权性质</option>
  <option value="{Second:HS_FS_HouseStyle}">户型</option>
  <option value="{Second:HS_FS_Structure}">建筑结构</option>
  <option value="{Second:HS_FS_Area}">建筑面积</option>
  <option value="{Second:HS_FS_BuildDate}">建筑年代</option>
  <option value="{Second:HS_FS_CityArea}">区县</option>
  <option value="{Second:HS_FS_Address}">地址</option>
  <option value="{Second:HS_FS_Floor}">楼层</option>
  <option value="{Second:HS_FS_Decoration}">装修情况</option>
  <option value="{Second:HS_FS_equip}">配套设施</option>
  <option value="{Second:HS_FS_Remark}">备注</option>
  <option value="{Second:HS_FS_LinkMan}">联系人</option>
  <option value="{Second:HS_FS_Contact}">联系方式</option>
  <option value="{Second:HS_FS_hits}">点击数</option>
  <option value="{Second:HS_FS_Pic}">图片</option>
</select>
<select name="TenancyFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄租赁信息┄┄</option>
  <option value="{Tenancy:HS_FS_UseFor}">使用性质</option>
  <option value="{Tenancy:HS_FS_XingZhi}">房屋性质</option>
  <option value="{Tenancy:HS_FS_Class}">类型</option>
  <option value="{Tenancy:HS_FS_ZaWuJian}">杂物间</option>
  <option value="{Tenancy:HS_FS_CityArea}">区县</option>
  <option value="{Tenancy:HS_FS_HouseStyle}">户型</option>
  <option value="{Tenancy:HS_FS_Area}">建筑面积</option>
  <option value="{Tenancy:HS_FS_Period}">有效期</option>
  <option value="{Tenancy:HS_FS_XiaoQuName}">小区名称</option>
  <option value="{Tenancy:HS_FS_Position}">房源地址</option>
  <option value="{Tenancy:HS_FS_JiaoTong}">交通状况</option>
  <option value="{Tenancy:HS_FS_Floor}">楼层</option>
  <option value="{Tenancy:HS_FS_BuildDate}">建筑年代</option>
  <option value="{Tenancy:HS_FS_Decoration}">装修情况</option>
  <option value="{Tenancy:HS_FS_equip}">配套设施</option>
  <option value="{Tenancy:HS_FS_Remark}">备注</option>
  <option value="{Tenancy:HS_FS_LinkMan}">联系人</option>
  <option value="{Tenancy:HS_FS_Contact}">联系方式</option>
  <option value="{Tenancy:HS_FS_hits}">点击数</option>
  <option value="{Tenancy:HS_FS_Pic}">图片</option>
</select>
<%end sub%>
<%sub AP_select()
				dim ap_D_rs,ap_list
				ap_list = ""
				set ap_D_rs = Server.CreateObject(G_FS_RS)
				ap_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='AP'",Conn,1,3
				if ap_D_rs.eof then
					ap_list =ap_list& "<option value="""">没有自定义字段</option>"
					ap_D_rs.close:set ap_D_rs=nothing
				else
					do while not ap_D_rs.eof 
						ap_list = ap_list & "<option value=""{AP=Define|"&ap_D_rs("D_Coul")&"}"">"& ap_D_rs("D_Name")&"</option>"
						ap_D_rs.movenext
					loop
					ap_D_rs.close:set ap_D_rs=nothing
				end if			  
			  %>
<select name="APFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄预定义字段┄┄</option>
  <option value="{AP:FS_AddFavorite}">加入收藏</option>
  <option value="{AP:FS_SendFriend}">发送给好友</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄人才自定义字段┄┄</option>
  <%=ap_list%>
</select>
<SELECT name="APFields_1" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄招聘信息基本字段┄</option>
  <option value="{INV:AP:FS_ID}">自动编号</option>
  <option value="{INV:AP:FS_UserNumber}">公司名称</option>
  <!-----------------2/2 by chen------------------------------------------->
  <option value="{INV:AP:FS_Mobile}">公司电话</option>
  <option value="{INV:AP:FS_Fax}">公司传真</option>
  <option value="{INV:AP:FS_Address}">公司地址</option>
  <option value="{INV:AP:FS_ConnectPer}">公司联系人</option>
  <option value="{INV:AP:FS_WebSit}">公司网站</option>
  <!------------------2/2  by chen----------------------------------------->
  <option value="{INV:AP:FS_JobName}">职位名称</option>
  <option value="{INV:AP:FS_JobDescription}">职位描述</option>
  <option value="{INV:AP:FS_ResumeLang}">简历接受语言</option>
  <option value="{INV:AP:FS_WorkCity}">工作地点</option>
  <option value="{INV:AP:FS_PublicDate}">发布日期</option>
  <option value="{INV:AP:FS_EndDate}">有效日期</option>
  <option value="{INV:AP:FS_NeedNum}">招聘人数</option>
  <option value="{INV:AP:FS_APURL}"> 信息访问路径</option>
</SELECT>
<SELECT name="APFields_2" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄人才信息基本字段┄</option>
  <option value="{TO:AP:FS_ID}">自动编号</option>
  <option value="{TO:AP:FS_UserNumber}">用户编号</option>
  <option value="{TO:AP:FS_UserName}">用户名</option>
  <option value="{TO:AP:FS_PersonURL}">用户求职简历路径</option>
  <option value="{TO:AP:FS_JobReadURL}">查看用户详情路径</option>
  <option value="{TO:AP:FS_Sex}">性别</option>
  <option value="{TO:AP:FS_Pic}">图片</option>
  <option value="{TO:AP:FS_Birthday}">生日</option>
  <option value="{TO:AP:FS_CertificateClass}">证件类型</option>
  <option value="{TO:AP:FS_CertificateNo}">证件号码</option>
  <option value="{TO:AP:FS_CurrentWage}">目前年薪</option>
  <option value="{TO:AP:FS_CurrencyType}">币种</option>
  <option value="{TO:AP:FS_WorkAge}">工作年限</option>
  <option value="{TO:AP:FS_Province}">所在省</option>
  <option value="{TO:AP:FS_City}">所在市</option>
  <option value="{TO:AP:FS_HomeTel}">家庭电话</option>
  <option value="{TO:AP:FS_CompanyTel}">公司电话</option>
  <option value="{TO:AP:FS_Mobile}">手机</option>
  <option value="{TO:AP:FS_Email}">电子邮件</option>
  <option value="{TO:AP:FS_QQ}">QQ</option>
  <option value="{TO:AP:FS_click}">浏览数</option>
  <option value="{TO:AP:FS_lastTime}">最后修改时间</option>
  <option value="{TO:AP:FS_ShenGao}">身高</option>
  <option value="{TO:AP:FS_XueLi}">学历</option>
  <option value="{TO:AP:FS_HowDay}">多久可以到岗</option>
  <option value="">--所属岗位--</option>
  <option value="{TO:AP_1:FS_Job}">所属岗位</option>
</SELECT>
<%end sub%>
<%sub MS_select()
			'得到自定义字段
			dim ms_D_rs,ms_list
			ms_list = ""
			set ms_D_rs = Server.CreateObject(G_FS_RS)
				ms_D_rs.open "select D_Coul,D_Name From FS_MF_DefineTable where D_SubType='MS'",Conn,1,1
			if ms_D_rs.eof then
				ms_list ="<option value="""">没有自定义字段</option>"
				ms_D_rs.close:set ms_D_rs=nothing
			else
				do while not ms_D_rs.eof 
					ms_list = ms_list & "<option value=""{MS=Define|"&ms_D_rs("D_Coul")&"}"">"& ms_D_rs("D_Name")&"</option>"
					ms_D_rs.movenext
				loop
				ms_D_rs.close:set ms_D_rs=nothing
			end if
			%>
<select name="ProductsFields" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄┄基本字段┄┄</option>
  <option value="{MS:FS_ID}">自动编号</option>
  <option value="{MS:FS_ProductTitle}"> 商品名称(截断)</option>
  <option value="{MS:FS_ProductTitleAll}"> 商品名称（完整）</option>
  <option value="{MS:FS_Barcode}"> 商品条形码</option>
  <option value="{MS:FS_Serialnumber}">产品序列号</option>
  <option value="{MS:FS_ProductURL}">商品浏览访问路径</option>
  <option value="{MS:FS_Stockpile}"> 商品库存</option>
  <option value="{MS:FS_OldPrice}"> 市场价格</option>
  <option value="{MS:FS_NewPrice}"> 实际价格</option>
  <!--运输费用-->
  <option value="{MS:FS_Mail_money}">运输费用</option>
  <option value="{MS:FS_NowMoney}">含运输费用后的价格</option>
  <!--运输费用-->
  <option value="{MS:FS_ProductContent}"> 商品描述</option>
  <option value="{MS:FS_RepairContent}"> 保修条款</option>
  <option value="{MS:FS_AddTime}"> 商品添加日期</option>
  <option value="{MS:FS_AddMember}"> 商品添加者</option>
  <option value="{MS:FS_ProductAddress}"> 产品产地</option>
  <option value="{MS:FS_MakeFactory}"> 生产厂商</option>
  <option value="{MS:FS_MakeTime}"> 制造日期</option>
  <option value="{MS:FS_saleNumber}"> 售出数量</option>
  <option value="{MS:FS_SaleStyle}"> 销售形式</option>
  <option value="{MS:FS_DiscountStartDate}"> 打折开始时间</option>
  <option value="{MS:FS_DiscountEndDate}"> 打折结束时间</option>
  <option value="{MS:FS_Discount}"> 折扣率</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄可定义字段┄┄</option>
  <option value="{MS:FS_hits}">点击数</option>
  <option value="{MS:FS_KeyWords}">关键字(带连接的搜索)</option>
  <option value="{MS:FS_TitleKeyWords}">关键字(用于搜索引擎搜索)</option>
  <!--option value="{MS:FS_TxtSource}"> 商品来源</option-->
  <option value="{MS:FS_SmallPicPath}">商品图片(小图)</option>
  <option value="{MS:FS_PicPath}">商品图片(大图)</option>
  <option value="{MS:FS_ShopBagURL}">购物车地址</option>
  <option value="{MS:FS_FormReview}">评论表单</option>
  <option value="{MS:FS_ReviewTF}">标题后显示评论字样</option>
  <option value="{MS:FS_ShowComment}">显示评论列表</option>
  <option value="{MS:FS_AddFavorite}">加入收藏</option>
  <option value="{MS:FS_SendFriend}">发送给好友</option>
  <option value="{MS:FS_SpecialList}">所属专题列表</option>
  <!--<option value="{MS:FS_ProductURL}"> 新闻访问路径</option>-->
  <!--<option value="" style="background:#88AEFF;color:000000">┄┄栏目可定义字段┄┄</option>-->
  <option value="{MS:FS_ClassName}">栏目中文名称</option>
  <option value="{MS:FS_ClassURL}">栏目浏览访问路径</option>
  <option value="{MS:FS_ClassNaviPicURL}">栏目导航图片</option>
  <option value="{MS:FS_ClassNaviContent}">栏目导航说明</option>
  <option value="{MS:FS_ClassKeywords}">栏目关键字(用于商城栏目的终极列表，优化搜索引擎)</option>
  <option value="{MS:FS_Classdescription}">栏目描述(用于商城栏目的终极列表，优化搜索引擎)</option>
  <!--<option value="{MS:FS_Classdescription}">栏目描述</option>-->
  <!--<option value="" style="background:#88AEFF;color:000000">┄┄专题可定义字段┄┄</option>-->
  <option value="{MS:FS_SpecialName}">专题中文名称</option>
  <!-- <option value="{MS:FS_SpecialURL}">专题浏览访问路径</option>-->
  <option value="{MS:FS_SpecialNaviPicURL}">专题导航图片</option>
  <option value="{MS:FS_SpecialNaviDescript}">专题导航说明</option>
  <option value="" style="background:#88AEFF;color:000000">┄┄自定义字段┄┄</option>
  <%=ms_list%>
</select>
<%end sub
sub ME_select()
%>
<SELECT name="APFields_1" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄个人会员基本字段┄</option>
  <option value="{ME:FS_UserNumber}">用户编号</option>
  <option value="{ME:FS_UserName}">用户名</option>
  <option value="{ME:FS_NickName}">用户昵称</option>
  <option value="{ME:FS_RealName}">真实姓名</option>
  <option value="{ME:FS_Sex}">性别</option>
  <option value="{ME:FS_HeadPic}">头像</option>
  <option value="{ME:FS_tel}">电话</option>
  <option value="{ME:FS_Email}">Email</option>
  <option value="{ME:FS_HomePage}">个人主页</option>
  <option value="{ME:FS_QQ}">QQ</option>
  <option value="{ME:FS_MSN}">MSN</option>
  <option value="{ME:FS_Province}">省份</option>
  <option value="{ME:FS_City}">城市</option>
  <option value="{ME:FS_Address}">地址</option>
  <option value="{ME:FS_PostCode}">邮政编码</option>
  <option value="{ME:FS_Vocation}">职业</option>
  <option value="{ME:FS_BothYear}">出生日期</option>
  <option value="{ME:FS_Age}">年龄</option>
  <option value="{ME:FS_Integral}">积分</option>
  <option value="{ME:FS_FS_Money}">金币</option>
  <option value="{ME:FS_IsMarray}">婚否</option>
  <option value="{ME:FS_RegTime}">注册日期</option>
</SELECT>
<SELECT name="APFields_2" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄企业会员基本字段┄</option>
  <option value="{ME:FS_C_Name}">企业名称</option>
  <option value="{ME:FS_C_ShortName}">企业简称</option>
  <option value="{ME:FS_C_logo}">企业Logo</option>
  <option value="{ME:FS_C_Tel}">电话</option>
  <option value="{ME:FS_C_Fax}">传真</option>
  <option value="{ME:FS_C_VocationClassID}">所在行业</option>
  <option value="{ME:FS_C_WebSite}">公司网站</option>
  <option value="{ME:FS_C_Operation}">业务范围</option>
  <option value="{ME:FS_C_Products}">公司产品</option>
  <option value="{ME:FS_C_Content}">公司简介</option>
  <option value="{ME:FS_C_Province}">省份</option>
  <option value="{ME:FS_C_City}">城市</option>
  <option value="{ME:FS_C_Address}">地址</option>
  <option value="{ME:FS_C_PostCode}">邮政编码</option>
  <option value="{ME:FS_C_Vocation}">联系人职务</option>
  <option value="{ME:FS_C_BankName}">开户银行</option>
  <option value="{ME:FS_C_BankUserName}">银行帐号</option>
  <option value="{ME:FS_C_property}">公司性质</option>
</SELECT>
<SELECT name="APFields_3" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000">┄公共字段┄</option>
  <option value="{ME:FS_UserURL}/ShowUser.asp?UserNumber={ME:FS_UserNumber}">详细信息URL</option>
  <option value="{ME:FS_UserURL}/Message_write.asp?ToUserNumber={ME:FS_UserNumber}">短信URL</option>
  <option value="{ME:FS_UserURL}/book_write.asp?ToUserNumber={ME:FS_UserNumber}&M_Type=0">留言URL</option>
  <option value="{ME:FS_UserURL}/Friend_add.asp?type=0&UserName={ME:FS_UserName}">加为好友URL</option>
  <option value="{ME:FS_UserURL}/UserReport.asp?action=report&ToUserNumber={ME:FS_UserNumber}">举报URL</option>
  <option value="{ME:FS_UserURL}/Corp_card_add.asp?UserNumber={ME:FS_UserNumber}">收藏名片URL</option>
  <option value="{ME:FS_UserURL}/?User={ME:FS_UserNumber}">会员个人主页URL</option>
</SELECT>
<%end sub%>
<% Sub ME_Login() %>
<SELECT name="Select_Login" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">选择会员登录标签所需要的字段</option>
  <option value="{Login_Name}" style="color:#FF0000;">用户名输入框(必选)</option>
  <option value="{Login_Password}" style="color:#FF0000;">密码输入框(必选)</option>
  <option value="{Login_Simbut}" style="color:#FF0000;">登录提交按钮(必选)</option>
  <option value="{Login_Type}">登录方式选择框</option>
  <option value="{Login_Reset}">登录取消按钮</option>
  <option value="{Reg_LinkUrl}">注册新用户连接</option>
  <option value="{Get_PassLink}">取回密码连接</option>
</SELECT>
<SELECT name="Login_Display" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">选择登录后显示内容字段</option>
  <option value="{User_Name}">会员姓名</option>
  <option value="{User_JiFen}">会员积分</option>
  <option value="{User_JinBi}">会员金币</option>
  <option value="{User_LoginTimes}">登录次数</option>
  <option value="{User_TouGao}">投稿数</option>
  <option value="{User_ConCenter}">控制面板连接</option>
  <option value="{User_LogOut}">退出连接</option>
</SELECT>
<br />
<span style="margin-left:10px; text-align:left; font-size:12px; color:#FF0000;">可在此处利用html代码设置显示样式，更多参数在标签设置;登陆样式和显示样式以<font color="#0033FF">"$*$"</font>分隔：登陆样式 $*$ 显示样式，否则会引起显示混乱</span>
<% End Sub %>
<% Sub MF_CustomForm()

%>
<SELECT name="CustomFormID" onChange="this.form.TxtFileds.value=frames['NewsContent'].GetNewsContentArray();this.form.Action.value='';this.form.submit();">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">选择自定义表单</option>
  <%
  Dim CustomFormRS,CustomFormItemRS,CustomFormID,CustomFormItemArray,i
  CustomFormID = NoSqlHack(Request("CustomFormID"))
  Set CustomFormRS = Conn.Execute("Select * from FS_MF_CustomForm")
  Do while Not CustomFormRS.Eof
  %>
  <option <% if CustomFormRS("ID") & "" = CustomFormID then Response.Write("selected") %> value="<% = CustomFormRS("ID") %>" style="color:#FF0000;"><% = CustomFormRS("formname") %></option>
  <%
  	CustomFormRS.MoveNext
  Loop
  CustomFormRS.Close
  Set CustomFormRS = Nothing
  %>
</SELECT>
<%
  if CustomFormID <> "" then
  	SQL = "Select ItemName,FieldName from FS_MF_CustomForm_Item Where FormID=" & CustomFormID
	Set CustomFormItemRS = Server.CreateObject(G_FS_RS)
	CustomFormItemRS.Open SQL,Conn,1,1
	CustomFormItemArray = CustomFormItemRS.GetRows
	CustomFormItemRS.Close
	Set CustomFormItemRS = Nothing
%>
<SELECT name="Select_CustomFormField" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">选择表单字段</option>
  <option value="{CustomFormHeader}">表单头</option>
  <option value="{CustomFormTailor}">表单尾</option>
  <option value="{CustomFormValidate}">验证码</option>
  <option value='<input name="" type="submit" value="提交"/>'>提交按钮</option>
  <option value='<input name="" type="reset" value="重填"/>'>重填按钮</option>
  <option value='<input name="" type="button" value="普通"/>'>普通按钮</option>
  <%
  if IsArray(CustomFormItemArray) then
  	For i = LBound(CustomFormItemArray,2) to UBound(CustomFormItemArray,2)
  %>
  <option value="{CustomForm_<% = CustomFormItemArray(1,i) %>}"><% = CustomFormItemArray(0,i) %></option>
  <%
  	Next
  end if
  %>
</SELECT>
<SELECT name="Select_CustomFormField" onChange="Insertlabel_Sel(this)">
  <option value="" style="background:#88AEFF;color:000000" selected="selected">选择数据显示字段</option>
  <option value="{CustomFormData_form_usernum}">用户ID</option>
  <option value="{CustomFormData_form_username}">用户名</option>
  <option value="{CustomFormData_form_ip}">来源IP地址</option>
  <option value="{CustomFormData_form_time}">添加时间</option>
  <option value="{CustomFormData_form_answer}">回复内容</option>
  <%
  if IsArray(CustomFormItemArray) then
  	For i = LBound(CustomFormItemArray,2) to UBound(CustomFormItemArray,2)
  %>
  <option value="{CustomFormData_<% = CustomFormItemArray(1,i) %>}"><% = CustomFormItemArray(0,i) %></option>
  <%
  	Next
  end if
  %>
</SELECT>
<%
  end if
%>
<% End Sub %>
</body>
</html>
<% 
Sub Add_Save()

End Sub
Set Conn=nothing
Function get_childList(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassName from FS_MF_StyleClass where ParentID=" & CintStr(TypeID) & " order by id desc" )
	f_TempStr =f_CompatStr & "┄"
	do while Not f_ChildNewsRs.Eof
			get_childList = get_childList & "<option value="""& f_ChildNewsRs("id")&""""
			If CStr(Request.QueryString("ClassID"))=CStr(f_ChildNewsRs("id")) then
				get_childList = get_childList & " selected" & Chr(13) & Chr(10)	
			End If
			get_childList = get_childList & ">├" &  f_TempStr & f_ChildNewsRs("ClassName") 
			get_childList = get_childList & "</option>" & Chr(13) & Chr(10)
			get_childList = get_childList &get_childList(f_ChildNewsRs("id"),f_TempStr)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function
%>
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