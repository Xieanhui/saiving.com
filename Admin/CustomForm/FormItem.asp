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
Dim Conn,strShowErr,obj_form_rs,formlist_sql,showMorePageGo_Type_,cPageNo,i
Dim strpage,int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_

int_RPP=20'设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"

MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("MF095") then Err_Show
dim FormID
FormID=NoSqlHack(request.QueryString("FormID"))
if formID="" then	
	strShowErr = "<li>非法数据传递</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自定义表单管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu"><a href="FormManage.asp" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by Samjun <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>表单项管理</strong></a><a href="../../help?Lable=NS_Form_Custom" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a>　　　　　　　　　　　　　　    </td>
  </tr>
  <tr>
      <td height="18" class="hback"><a href="FormManage.asp">表单管理</a>&nbsp;&nbsp;<a href="FormItemOperate.asp?act=add&formid=<%=formID%>">新建表单项</a></td>
  </tr>
</table>
<%
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
formlist_sql = "Select FormItemID,orderby,ItemName,FieldName,ItemType,IsNull from FS_MF_CustomForm_Item where formID="&formID&" order by orderby"
Set obj_form_rs = Server.CreateObject(G_FS_RS)
obj_form_rs.Open formlist_sql,Conn,1,1
%>

<table width="98%" border="0" cellspacing="1" cellpadding="5" align="center" class="table">
  <tr class="xingmu">
    <td width="6%" align="center" class="xingmu">顺序</td>
    <td width="25%" align="center" class="xingmu">表单项名</td>
    <td width="23%" align="center" class="xingmu">字段名</td>
    <td width="20%" align="center" class="xingmu">字段类型</td>
    <td width="13%" align="center" class="xingmu">是否必填</td>
    <td width="13%" align="center" class="xingmu">操作</td>
  </tr>
  <%
  for i=1 to obj_form_rs.pagesize
  	if obj_form_rs.eof then exit for
  %>
  <tr class="hback">
    <td align="center" class="hback"><%=obj_form_rs("orderby")%></td>
    <td class="hback"><%=obj_form_rs("ItemName")%></td>
    <td class="hback"><%=obj_form_rs("FieldName")%></td>
    <td class="hback"><%
	select case obj_form_rs("ItemType")
		case "SingleLineText"
			response.Write("单行文本")
		case "MultiLineText"
			response.Write("多行文本")
		case "PassWordText"
			response.Write("密码")
		case "DateTime"
			response.Write("日期时间")
		case "RadioBox"
			response.Write("单选项")
		case "CheckBox"
			response.Write("多选项")
		case "Numberic"
			response.Write("数字")
		case "UploadFile"
			response.Write("附件")
		case "DropList"
			response.Write("下拉框")
		case "List"
			response.Write("列表框")
	end select%></td>
    <td align="center" class="hback"><%if obj_form_rs("IsNull")=0 then response.Write("必填项") else response.Write("非必填项") end if%></td>
    <td align="center" class="hback"><A href="FormItemOperate.asp?act=edit&FormItemID=<%=obj_form_rs("FormItemID")%>">修改</A> <A onClick="if(confirm('确定要删除吗？'))location='FormItemOperate.asp?act=del&FormItemID=<%=obj_form_rs("FormItemID")%>';" href="javascript:void(0);">删除</A> </td>
  </tr>
  <%
  	obj_form_rs.movenext
  next
  %>
  <tr>
    <td colspan="6"class="hback">
	  <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="79%" colspan="2" align="right" class="hback"> <%
			response.Write "<p>"&  fPageCount(obj_form_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%> </td>
        </tr>
      </table>
	 </td>
  </tr>
</table>
</body>
</html>