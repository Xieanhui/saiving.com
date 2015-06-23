<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
Dim User_Conn, info_Rs,sql_cmd,sql_info_cmd,ClassID,AuditTF,SuperClass_Str,ClassPath,Conn
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
Dim ContID,title,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,untread '稿件信息
Response.CacheControl = "no-cache"

MF_Default_Conn
MF_User_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("NS_Constr") then Err_Show
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
'MF_Check_Pop_TF（）参数：文件名命名+代码
'if Not MF_Check_Pop_TF("NS_News_01") then Top_Go_To_Error_Page 
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"				'尾页
'------------------------------------------------
ClassID=NoSqlHack(Request.QueryString("classid"))
AuditTF=NoSqlHack(Request.QueryString("audittf"))
if ClassID="" then
ClassID=0
End if
if AuditTF="" then
AuditTF=0
End if
User_Conn.CursorLocation = 3
Set info_Rs=Server.CreateObject(G_FS_RS)
if ClassID<>0 then'如果classid不为根目录
	if  AuditTF="1" then '如果已经审核
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,MainID,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=1 and MainID="&ClassID
	else'如果未审核
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,isTF,MainID,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=0 and MainID="&ClassID
	End if
else'如果classid为空
	if AuditTF="1" then'如果已经审核  
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,MainID,isTF,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=1"
	else'如果未审核 
		sql_info_cmd="Select ContID,ContTitle,addtime,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,MainID,isTF,untread from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=0"
	End if
End IF
info_Rs.open sql_info_cmd,User_Conn,1,1
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language='javascript' src="../../FS_Inc/prototype.js"></script>
<script language='javascript' src="../../FS_Inc/publicjs.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
	<td class="xingmu">稿件标题</td>
	<td align='center'  class="xingmu">投稿时间</td>
  	<td align='center' class="xingmu">类型</td>
	<td align='center' class="xingmu">信息级</td>
	<td align='center' class="xingmu">发布者</td>
	<td align='center' class="xingmu">总站显示</td>
	<td align='center' class="xingmu">已退稿</td>
  	<td align='center' class="xingmu">推荐</td>
	<td align='center' class="xingmu">锁定</td>
	<td align='center'><input type="checkbox" name="chk_constr" onClick="selectAll(document.all('chk_constr'))" value=""/></td>
  </tr>
  <%
	If Not info_Rs.eof then
		info_Rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>info_Rs.PageCount Then cPageNo=info_Rs.PageCount 
		info_Rs.AbsolutePage=cPageNo
	End if

	FOR int_Start=1 TO int_RPP 
		if info_Rs.eof then exit for
		ContID=info_Rs("ContID")
		title=info_Rs("ContTitle")
		addtime=info_Rs("AddTime")
		ContSytle=info_Rs("ContSytle")
		if ContSytle=0 then
			ContSytle="原创"
		elseif ContSytle=1 then
			ContSytle="转载"
		elseif ContSytle=2 then
			ContSytle="代理"
		End IF
		
		InfoType=info_Rs("InfoType")
		if InfoType=0 then
			InfoType="普通"
		elseif InfoType=1 then
			InfoType="优先"
		elseif InfoType=2 then
			InfoType="加急"
		End IF
		UserNumber=info_Rs("UserNumber")
		isPublic=info_Rs("isPublic")
		if isPublic=1 then
			isPublic=info_Rs("MainID")
		else
			isPublic="否"
		End IF
		isTF=info_Rs("isTF")
		if isTF=1 then
			isTF="推荐"
		else
			isTF="否"
		End IF
		AdminLock=info_Rs("AdminLock")
		if AdminLock=1 then
			AdminLock="<font color=""red"">锁定</font>"
		else
			AdminLock="否"
		End IF
		untread=info_Rs("untread")
		if untread=1 then
			untread="<font color='red'>√</font>"
		else
			untread="-"
		End if
		Response.Write("<tr class='hback'>"&chr(10)&chr(13))
		Response.Write("<td><a href='Audit_Edit.asp?contid="&ContID&"'>"&title&"</a></td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&addtime&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&ContSytle&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&InfoType&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'><a href=../../"&G_USER_DIR&"/ShowUser.asp?UserNumber="&UserNumber&" target=_blank>"&GetFriendName(UserNumber)&"</a></td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&isPublic&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&untread&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&isTF&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'>"&AdminLock&"</td>"&Chr(10)&Chr(13))
		Response.Write("<td align='center'><input type=""checkbox"" name=""chk_constr"" id=""chk_constr"" value="""&ContID&"""></td>"&Chr(10)&Chr(13))
		Response.write("</tr>")
		info_Rs.movenext
	Next
  %>
  <tr>
  <td class="hback" colspan="10" align="right">
	<%if AuditTF="0" then Response.Write("<button onClick=""Operator('audit')"" >审  核</button>&nbsp;")%>
	<button onClick="Operator('lock')">锁  定</button>&nbsp;
	<button onClick="Operator('unlock')">解除锁定</button>&nbsp;
	<button onClick="Operator('tf')">推  荐</button>&nbsp;
	<button onClick="Operator('untf')">取消推荐</button>&nbsp;
	<button onClick="Operator('delete')">删  除</button>&nbsp;
	<button onClick="Operator('deleteAll')">删除所有</button>&nbsp;
	<%if AuditTF="0" then Response.Write("<button onClick=""Operator('untread')"">退  稿</button>")%>
  </td>
  </tr>
  <%
  	Response.Write("<tr>"&vbcrlf)
	Response.Write("<td align='right' colspan='10'  class=""hback"">"&fPageCount(info_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
  %>
</table>
</body>
</html>
<script language="javascript">
//执行提交动作，
//param用来作为选择执行动作的标志

function Operator(param)
{
	var values=getCheckBoxValues();
	if(param!="deleteAll"&&values.length<2)
	{	
		alert("请选择操作对象");
		return;
	}
	switch(param)
	{
		case "audit":location="Constr_Action.asp?act=audit&values="+values;break;//审核
		case "recall":location="Constr_Action.asp?act=recall&values="+values;break//撤消审核
		case "lock":location="Constr_Action.asp?act=lock&values="+values;break//锁定
		case "unlock":location="Constr_Action.asp?act=unlock&values="+values;break//解除锁定
		case "tf":location="Constr_Action.asp?act=tf&values="+values;break//推荐
		case "untf":location="Constr_Action.asp?act=untf&values="+values;break//取消推荐
		case "delete": if(confirm("确认删除？"))location="Constr_Action.asp?act=delete&values="+values;break//删除
		case "deleteAll": if(confirm("确认删除所有记录？"))location="Constr_Action.asp?act=deleteAll";break//删除所有
		case "untread": if(confirm("确认退稿？"))location="Constr_Action.asp?act=untread&values="+values;break//退稿
	}
	
	
}
//获得被选种的checkbox的值的集
function getCheckBoxValues()
{
	var checkBoxValues=""
	for(var i=1;i<document.all("chk_constr").length;i++)
	{
		if(document.all("chk_constr")[i].checked)
		{
			checkBoxValues=checkBoxValues+","+document.all("chk_constr")[i].value;
		}
	}
	return checkBoxValues;
}
parent.document.all.hd_classid.value="<%=Classid%>"
parent.document.all.hd_audit.value="<%=AuditTF%>"
</script>
<%
info_Rs.close
User_Conn.close
Set info_Rs=nothing
Set User_Conn=nothing
%>





