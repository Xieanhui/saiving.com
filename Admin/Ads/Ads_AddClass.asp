<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS003") then Err_Show
Dim lng_ClassID,o_Class_Rs,str_Class_Sql,OpCType,str_ClassName,strShowErr,str_ClassSubmit,str_temp_name,str_Optype_name

str_ClassSubmit=Request.Form("ClassSubmit")
If str_ClassSubmit="Submit" Then
	str_ClassName=Request.Form("ClassName")
	If Trim(str_ClassName)="" or IsNull(str_ClassName) Then
		strShowErr = "<li>栏目名字不能为空!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		str_Class_Sql="Select * from FS_AD_Class where AdClassName='"&NoSqlHack(str_ClassName)&"'"
		Set o_Class_Rs=Conn.execute(str_Class_Sql)
		If o_Class_Rs.Eof Then
			Conn.execute("insert into FS_AD_Class(AdClassName,AddDate,Lock) values('"&NoSqlHack(str_ClassName)&"','"&Now()&"',0)")
			Set o_Class_Rs=Nothing
			strShowErr = "<li>添加成功!</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp")
			Response.end
		Else
			Set o_Class_Rs=Nothing
			strShowErr = "<li>栏目名字已存在!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		End If
	End If
End If

str_temp_name= Request.Querystring("OpCType")
If str_temp_name<>"" Then
	str_temp_name=" 添 加 "
	str_ClassSubmit="Submit"
Else
	str_temp_name=" 修 改 "
End If

lng_ClassID=Request.QueryString("ID")
If lng_ClassID<>"" or Not IsNull(lng_ClassID) Then
	Set o_Class_Rs=Conn.execute("Select AdClassName From FS_AD_Class Where AdClassID="&Clng(lng_ClassID)&"")
	If Not o_Class_Rs.Eof Then
		str_ClassName=o_Class_Rs("AdClassName")
	End If
	Set o_Class_Rs=Nothing
End If 
str_Optype_name=Request.Form("ClassName")
If Trim(str_Optype_name)<>"" Then
	lng_ClassID=Request.Form("ClassID")
	If lng_ClassID="" or IsNull(lng_ClassID) Then
		strShowErr = "<li>参数错误!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Else
		If isnumeric(lng_ClassID)=False Then
			strShowErr = "<li>参数错误!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			str_Class_Sql="Select * from FS_AD_Class where AdClassName='"&NoSqlHack(str_Optype_name)&"'"
			Set o_Class_Rs=Conn.execute(str_Class_Sql)
			If o_Class_Rs.Eof Then
				Conn.execute("Update FS_AD_Class Set AdClassName='"&NoSqlHack(str_Optype_name)&"' where AdClassID="&Clng(lng_ClassID)&"")
				Set o_Class_Rs=Nothing
				strShowErr = "<li>修改成功!</li>"
				Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_ClassManage.asp?Page="&Request.QueryString("Page")&"")
				Response.end
			Else
				Set o_Class_Rs=Nothing
				strShowErr = "<li>栏目名字已存在!</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			End If
			Set o_Class_Rs=Nothing
		End If
	End If
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告统计___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><%=str_temp_name%>广告分类</td>
  </tr>
   <tr class="hback"> 
    <td align="left" class="hback"><a href="javascript:history.back();">返回上一级</a></td>
  </tr>
   <tr class="hback">
     <td align="left" class="hback" height="10"></td>
   </tr>
   <tr class="hback">
     <td align="center" class="hback"><form name="AdClass" method="post" action="">
       <p>　</p>
       <table width="426" border="0" cellspacing="0" cellpadding="0">
       <tr>
         <td width="109" height="20" align="right">栏目名称</td>
         <td width="10" align="right">　</td>
         <td width="307" align="left"><input name="ClassName" type="text" size="30" maxlength="30" title="栏目名称:必填" value="<%=str_ClassName%>"><input type="hidden" value="<%=str_ClassSubmit%>" name="ClassSubmit"><input type="hidden" value="<%=str_ClassName%>" name="Tempname"><input type="hidden" value="<%=lng_ClassID%>" name="ClassID"></td>
       </tr>
       <tr>
         <td height="10" colspan="3">&nbsp;</td>
       </tr>
       <tr>
         <td height="10" colspan="3" align="center"><input name="Tj" type="submit" id="Tj" value="<%=str_temp_name%>">
             <input name="Cx" type="reset" id="Cx" value=" 重 写 "></td>
       </tr>
     </table>
	 </form></td>
   </tr>
</table>
</body>
</html>
<%
Sub CheckAllIDFLag(Showstr)
	If CheckAllID="" or IsNull(CheckAllID) Then
		strShowErr = "<li>请选择要"&Showstr&"的文件!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
End Sub
Set Conn=Nothing
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





