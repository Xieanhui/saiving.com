<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.CacheControl = "no-cache"
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_CustomForm") then Err_Show

Dim CharIndexStr,strShowErr,obj_form_rs,formlist_sql,strpage,i
Dim int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
int_showNumberLink_=8
showMorePageGo_Type_ = 1
str_nonLinkColor_="#999999"
toF_="<font face=webdings title=""首页"">9</font>"
toP10_=" <font face=webdings title=""上十页"">7</font>"
toP1_=" <font face=webdings title=""上一页"">3</font>"
toN1_=" <font face=webdings title=""下一页"">4</font>"
toN10_=" <font face=webdings title=""下十页"">8</font>"
toL_="<font face=webdings title=""最后一页"">:</font>"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自定义表单管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>自定义表单管理</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="FormOperate.asp?act=add">新建表单</a></td>
  </tr>
</table>
<%
formlist_sql = "Select ID,formName,tableName,remark from FS_MF_CustomForm"
Set obj_form_rs = Server.CreateObject(G_FS_RS)
obj_form_rs.Open formlist_sql,Conn,1,1
%>

<table width="98%" border="0" cellspacing="1" cellpadding="5" align="center" class="table">
  <tr class="xingmu">
    <td width="14%" align="center" class="xingmu">表单名称</td>
    <td width="19%" align="center" class="xingmu">表名</td>
    <td width="35%" align="center" class="xingmu">说明</td>
    <td align="center" class="xingmu">操作</td>
  </tr>
  <%
  if Not obj_form_rs.Eof then
	strpage=request("page")
	if strpage = "" then strpage = 1
	if Not IsNumeric(strpage) then strpage = 1
	strpage = CInt(strpage)
	obj_form_rs.PageSize = 10
	if strpage > obj_form_rs.PageCount then strpage = obj_form_rs.PageCount
	obj_form_rs.AbsolutePage = strpage
	  for i=1 to obj_form_rs.pagesize
		if obj_form_rs.eof then exit for
  %>
  <tr class="hback">
    <td align="center" class="hback"><%=obj_form_rs("formName")%></td>
    <td align="center" class="hback"><%=obj_form_rs("tableName")%></td>
    <td align="center" class="hback"><%=obj_form_rs("remark")%></td>
    <td align="center" class="hback"><A href="FormItem.asp?FormID=<%=obj_form_rs("ID")%>">表单项管理</A> <A href="FormOperate.asp?act=edit&ID=<%=obj_form_rs("ID")%>">修改</A> <A href="FormOperate.asp?act=del&ID=<%=obj_form_rs("ID")%>" onClick="return ConfirmGoto('你确定要删除该表单吗?数据将不能被恢复!');">删除</A> <A href="CustomForm_HtmlCode.asp?ID=<%=obj_form_rs("ID")%>&op=2">预览</A> <A href="Form_Data.asp?formid=<%=obj_form_rs("ID")%>">数据</A></td>
  </tr>
  <%
		obj_form_rs.movenext
	  next
  %>
  <tr>
    <td colspan="4"class="hback">
	  <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="79%" colspan="2" align="right" class="hback"> <%
			response.Write "<p>"&  fPageCount(obj_form_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,strpage)
	%> </td>
        </tr>
      </table>
    </td>
  </tr>
 <%
 end if
 %>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function ConfirmGoto(fstrTips){
   if(confirm(fstrTips)) return true;
   else return false;
}
function GetHtml(id,f)
{
	var WWidth = (window.screen.width-600)/2;
	var Wheight = (window.screen.height-500)/2-50;
	window.open('CustomForm_HtmlCode.asp?ID='+id +'&op='+ f, '获取HTML', 'height=500px, width=600px,toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no,top='+Wheight+', left='+WWidth+', status=no'); 
}
</script>