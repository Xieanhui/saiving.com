<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,obj_Label_Rs,SQL,strShowErr
	MF_Default_Conn
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage
	int_RPP=50 '设置每页显示数目
	int_showNumberLink_=8 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
	toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
	toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
	toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
	toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
	toL_="<font face=webdings title=""最后一页"">:</font>"
	Dim str_StyleName,txt_Content,Labelclass_SQL,obj_Labelclass_rs,obj_Count_rs
	if Request("type")="del" then
		if trim(Request("id"))="" then
			strShowErr = "<li>请选择标签</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
			Conn.execute("Delete From FS_MF_Lable where id in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>删除成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=label/All_Label_stock.asp")
		Response.end
	end if
	if Request("type")="deltobak" then
		if trim(Request("id"))="" then
			strShowErr = "<li>请选择标签</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
			Conn.execute("Update FS_MF_Lable set isDel=1 where id in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>加入标签备份库成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=label/All_Label_stock.asp")
		Response.end
	end if
	if Request("type")="bakto" then
		if trim(Request("id"))="" then
			strShowErr = "<li>请选择标签</li>"
			Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
			Conn.execute("Update FS_MF_Lable set isDel=0 where id in (" & FormatIntArr(Request("id")) & ")")
		strShowErr = "<li>还原成功</li>"
		Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=label/All_Label_stock.asp")
		Response.end
	end if
%>
<html>
<head>
<title>标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<body>
<table width="98%" height="66" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" > 
    <td width="100%" height="20"  align="Left" class="xingmu">标签库</td>
  </tr>
  <tr class="hback" > 
   <td height="27" align="center" class="hback"><div align="left"><a href="All_Label_Stock.asp">所有标签</a>┆<a href="../FreeLabel/FreeLabelList.asp"><font color="#FF0000">自由标签</font></a>┆<a href="All_Label_Stock.asp?isDel=1">备份库</a>┆<a href="label_creat.asp">创建标签</a>┆<a href="label_creat_txt.asp">文本创建标签</a>┆<a href="Label_Class.asp" target="_self">标签分类</a>&nbsp;┆<a href="All_label_style.asp">样式管理</a>&nbsp;<a href="../../help?Label=MF_Label_Stock" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="7%" class="xingmu"><div align="center">序号</div></td>
    <td width="39%" class="xingmu"><div align="center">标签名称</div></td>
    <td width="27%" class="xingmu"><div align="center">描述/操作</div></td>
    <td width="27%" class="xingmu">选择</td>
  </tr>
  <%
  dim rs_class,str_ParentID
  if trim(Request.QueryString("ParentID"))<>"" then
		str_ParentID = " and ParentID="&NoSqlHack(Request.QueryString("ParentID"))&""
  elseif not isnumeric(trim(Request.QueryString("ParentID"))) then
		str_ParentID = " and ParentID=0"
  else
		str_ParentID = " and ParentID=0"
  end if
  set rs_class=Conn.execute("select id,ClassName,ClassContent,ParentID From FS_MF_LableClass where 1=1"&str_ParentID&" order by id desc")
  do while not rs_class.eof 
  %>
  <tr class="hback">
  <td valign="top"><div align="center"><img src="../Images/Folder/folder.gif" alt="文件夹" width="20" height="16"></div></td>
  <td><a href="All_Label_Stock.asp?ClassId=<% = rs_class("id")%>&isdel=<%=Request.QueryString("isdel")%>&ParentID=<%=rs_class("id")%>"><% = rs_class("ClassName")%></a></td>
  <td><% = rs_class("ClassContent")%></td>
  <td>&nbsp;</td>
  </tr>
  <%
  rs_class.movenext
  loop
  rs_class.close:set rs_class = nothing
  %>
  <tr class="hback_1">
	<td colspan="4" height="2"></td>
  </tr>
 <%
	dim rs_stock,i,ClassId,ides
	strpage=NoSqlHack(Trim(request("page")))
	If len(strpage)=0 Or strpage<"1" Or Trim(strpage)="" Then:strpage="1":End If
	if Request.QueryString("ClassId")<>"" then
		ClassId = " and LableClassID="&NoSQLHack(Request.QueryString("ClassID"))&""
	else
		ClassId = " and LableClassID=0"
	end if
	if Request.QueryString("isDel")="1" then
		ides = " and isdel=1"
	else
		ides = " and isdel=0"
	end if
	dim keys,wh
	keys = trim(Request("key"))
	if keys<>"" then
		ClassID = ""
		wh = " and (LableName like '%"&keys&"%' or LableContent like '%"&keys&"%')"
	end if
	dim stocksql
	stocksql="select ID,LableName,LableContent,isDel From FS_MF_Lable Where 1=1 "& ides & ClassId & wh &" order by ID desc"
	'response.Write(stocksql)
	'response.End()
	set rs_stock= Server.CreateObject(G_FS_RS)
	rs_stock.open stocksql,Conn,1,1
	
	if rs_stock.eof then
	   rs_stock.close
	   set rs_stock=nothing
	   Response.Write"<TR  class=""hback""><TD colspan=""4""  class=""hback"" height=""40"">没有记录。</TD></TR>"
	else
		rs_stock.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>rs_stock.PageCount Then cPageNo=rs_stock.PageCount 
		rs_stock.AbsolutePage=cPageNo
		for i=1 to rs_stock.pagesize
			if rs_stock.eof Then exit For 
	%>
 <form name="form1" method="post" action=""> <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td class="hback"><div align="center"><%=i%></div></td>
    <td class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(lable_<% = rs_stock("ID")%>);"  language=javascript><%=rs_stock("LableName")%></td>
    <td class="hback"><a href="label_creat.asp?id=<%=rs_stock("id")%>&type=edit">编辑</a>┆<a href="label_creat_txt.asp?id=<%=rs_stock("id")%>&type=edit">文本编辑</a>┆<a href="All_Label_Stock.asp?id=<%=rs_stock("id")%>&type=del"  onClick="{if(confirm('确定清除您所选择的记录吗？\n此操作不可逆！')){return true;}return false;}">删除</a>┆
	<%if Request.QueryString("isDel")="1" then%>
	<a href="All_Label_Stock.asp?id=<%=rs_stock("id")%>&type=bakto" onClick="{if(confirm('确定还原吗')){return true;}return false;}">还原标签</a></td>
	<%else%>
	<a href="All_Label_Stock.asp?id=<%=rs_stock("id")%>&type=deltobak" onClick="{if(confirm('确定把标签移动到标签备份库中？\n移动后此标签在前台不调用')){return true;}return false;}">加入备份库</a></td>
    <%end if%>
	<td class="hback">
      <label>
        <input name="id" type="checkbox" id="id" value="<%=rs_stock("id")%>">
        </label>    </td>
  </tr>
   <tr id="lable_<% = rs_stock("ID")%>" style="display:none;">
    <td height="43" colspan="4" class="hback_1"><div align="left"><font style="font-family:Courier New"><%=rs_stock("LableContent")%></font></div></td>
  </tr>
<%
	rs_stock.movenext
	next
%>
   <tr>
     <td height="26" colspan="4" class="hback"><div align="right">
       <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
      选中所有<input name="type" type="hidden" id="type">
		<input type="button" name="Submit" value="删除"  onClick="document.form1.type.value='del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
		<%if Request.QueryString("isDel")="1" then%>
		<input type="button" name="Submit" value="还原标签"  onClick="document.form1.type.value='bakto';{if(confirm('确定还原标签吗')){this.document.form1.submit();return true;}return false;}">
		<%else%>
		<input type="button" name="Submit" value="删除到标签备份库"  onClick="document.form1.type.value='deltobak';{if(confirm('确定把标签移动到标签备份库中？\n移动后此标签在前台不调用')){this.document.form1.submit();return true;}return false;}">
     	<%end if%>
	 </div></td>
   </tr>
  </form>
  
<tr class="hback">
<td colspan="4">
<%
	response.Write "<p>"&  fPageCount(rs_stock,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	rs_stock.close:set rs_stock=nothing
end if
%></td>
</tr>
<tr>
<td colspan="4">
  <form name="Label_Form" method="get" action="" target="_self" style="margin:0;padding:0;" onSubmit="return false;">
        &nbsp;  搜索标签：<input type="text" id="key" name="keyw" />
        <input type="button" name="se" value="搜索标签" onClick="searcha();" />
        <input type="button" name="se1" value="导出标签" onClick="outlabel();" />
        <input type="button" name="se1" value="导入标签" onClick="inlabel();" />
  </form>
</td>
</tr>
</table>
</body>
<% 
Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function insert(insertContent)
{
		obj=window.frames.item('NewsContent').EditArea.document.body;
		obj.focus();
	if(document.selection==null)
	{
		var iStart = obj.selectionStart
		var iEnd = obj.selectionEnd;
		obj.value = obj.value.substring(0, iEnd) +insertContent+ obj.value.substring(iEnd, obj.value.length);
	}else
	{
		var range = document.selection.createRange();
		range.text=insertContent;
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
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
  }
function searcha()
{
	if(document.getElementById("key").value=="")
	{
		alert("填写关键字");
		return false;
	} 
	window.location.href="all_label_stock.asp?key="+escape(document.getElementById("key").value)+"";
} 
   
function outlabel()
{
	if(confirm('确定要备份标签吗？'))
	{
		window.location.href="out_in.asp?action=out";
	}
	return false;
}

function inlabel()
{
	if(confirm('确定要导入标签吗？如果标签重复。系统将重命名标签名称。'))
	{
		window.location.href="out_in.asp?action=in";
	}
	return false;
}
</script>