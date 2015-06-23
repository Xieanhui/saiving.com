<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,i

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage
int_RPP=20 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"

MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF 
'权限判断
dim Fs_news
set Fs_news = new Cls_News
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>栏目管理</strong></a><a href="../../help?Lable=DS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">管理首页</a>┆<a href="Class_add.asp?ClassID=&Action=add">添加根栏目</a>┆<a href="Class_Action.asp?Action=one">一级栏目排序</a>┆<a href="Class_Action.asp?Action=n">N级栏目排序</a>┆<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('确认复位所有栏目？\n\n如果选择确定，所有的栏目将设置为一级分类!!')){return true;}return false;}">复位所有栏目</a>┆<a href="Class_Action.asp?Action=unite">栏目合并</a>┆<a href="Class_Action.asp?Action=allmove">栏目转移</a> 
        ┆ <a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('确认清空所有栏目里的数据吗？\n\n如果选择确定,所有的栏目的下载将被删除!!')){return true;}return false;}">删除所有栏目</a>┆<a  href="#" onClick="javascirp:history.back()">后退</a> ┆ <a href="../../help?Lable=DS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="ClassForm" method="post" action="Class_Action.asp">
    <tr class="xingmu"> 
      <td width="5%" height="25" class="xingmu"><div align="center">ID</div></td>
      <td width="30%" class="xingmu"><div align="center">栏目中文[栏目英文]</div></td>
      <td width="7%" class="xingmu"><div align="center">权重</div></td>
      <td width="22%" class="xingmu"><div align="center">属性</div></td>
      <td width="36%" class="xingmu"><div align="center">操作</div></td>
    </tr>
	<%
	Dim obj_news_rs,obj_news_rs_1,isUrlStr
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,isConstr,isShow,[Domain],FileExtName,isPop from FS_DS_Class where Parentid  = '0'   and ReycleTF=0  Order by Orderid desc,ID desc",Conn,1,3
	if obj_news_rs.eof then
	   Response.Write"<TR  class=""hback""><TD colspan=""5""  class=""hback"" height=""40"">没有记录。</TD></TR>"
	else
		obj_news_rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo>obj_news_rs.PageCount Then cPageNo=obj_news_rs.PageCount 
		If cPageNo<=0 Then cPageNo=1
		obj_news_rs.AbsolutePage=cPageNo
		for i=1 to int_RPP
			if obj_news_rs.eof Then exit For 
	%>
    <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)> 
      <td height="22" class="hback"> 
        <div align="center"><% = obj_news_rs("id")%>
          </div></td>
      <td height="22" class="hback"> 
        <% 
		if obj_news_rs("IsUrl") = 1 then
			isUrlStr = ""
		Else 
			isUrlStr = "["&obj_news_rs("ClassEName")&"]"
		End if
		Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
		obj_news_rs_1.Open "Select Count(ID) from FS_DS_Class where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
		if obj_news_rs_1(0)>0 then
			Response.Write  "<img src=""images/+.gif""></img><a href=Class_add.asp?ClassID="&obj_news_rs("ClassID")&"&Action=edit>"&obj_news_rs("ClassName")&"</a>&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
		Else
			Response.Write  "<img src=""images/-.gif""></img><a href=Class_add.asp?ClassID="&obj_news_rs("ClassID")&"&Action=edit>"&obj_news_rs("ClassName")&"</a>&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
		End if
		obj_news_rs_1.close:set obj_news_rs_1 =nothing
		%>
	</td>
      <td class="hback" align="center"><% = obj_news_rs("OrderID")%></td>
      <td class="hback"> <div align="center">
          <%
	  if obj_news_rs("IsUrl") = 1 then
		  Response.Write("<font color=red>外部</font>&nbsp;┆&nbsp;")
	  Else
		  Response.Write("系统&nbsp;┆&nbsp;")
	  End if
	  if obj_news_rs("isConstr") = 1 then
		  Response.Write("<font color=red>稿</font>&nbsp;┆&nbsp;")
	  Else
		  Response.Write("<strike>稿</strike>&nbsp;┆&nbsp;")
	  End if
	  if obj_news_rs("isShow") = 1 then
		  Response.Write("<font color=red>显示</font>&nbsp;┆&nbsp;")
	  Else
		  Response.Write("隐藏&nbsp;┆&nbsp;")
	  End if
	  if len(obj_news_rs("Domain")) >5 then
		  Response.Write("<font color=red>域</font>&nbsp;┆&nbsp;")
	  Else
		  Response.Write("本&nbsp;┆&nbsp;")
	  End if
	  %>
        </div></td>
      <td class="hback"><div align="center"><a href="DownClass_review.asp?id=<% = obj_news_rs("ClassID")%>" target="_blank">预览</a>┆
	  <a href="Class_add.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=add">添加子栏目</a>
	  ┆<a href="Class_add.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=edit">修改</a>
	  ┆<a href="Class_Action.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=clear" onClick="{if(confirm('确定清空此栏目下信息吗?')){return true;}return false;}">清空</a> ┆<a href="Class_Action.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=del"  onClick="{if(confirm('确定删除您所选择的栏目吗?\n\n此栏目下的子类也将被删除!!')){return true;}return false;}">删除</a>
	  <input name="Cid" type="checkbox" id="Cid" value="<% = obj_news_rs("ClassID")%>">
        </div></td>
    </tr>
<%
		Response.Write(Fs_news.GetChildNewsList(obj_news_rs("ClassID"),""))
		obj_news_rs.MoveNext
	next
%>
    <tr> 
      <td height="30" colspan="5" class="hback">
 
<div align="right">
          <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
          选择所有
          <input name="Action" type="hidden" id="Action">
        </div></td>
    </tr>
	<tr>
		<td height="30" colspan="5" class="hback">
<%
	response.Write "<p>"&  fPageCount(obj_news_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
%>
		</td>
	</tr>	
		<%end if%>
  </form>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="18" class="hback"><div align="left"> 
        <p><span class="tx"><strong>说明</strong></span>:<br>
          系统 -------系统目录<br>
          外部-------外部栏目 <br>
          稿--------栏目可以投稿<br>
          显示----导航中显示<br>
          隐藏----导航中隐藏 <br>
          域--------捆绑了二级域名的目录</p>
        </div></td>
  </tr>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
obj_news_rs.close
set obj_news_rs =nothing
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = ClassForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = ClassForm.chkall.checked;  
    }  
  }
</script>

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





