<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/VS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,Vs_Rs,Vs_Sql
Dim AutoDelete,Months
MF_Default_Conn 
MF_Session_TF
if not MF_Check_Pop_TF("VS_site") then Err_Show
if not MF_Check_Pop_TF("VS001") then Err_Show

Sub Save()
	Dim sysid,Str_Tmp,Arr_Tmp
	sysid = NoSqlHack(request.Form("sysid"))
	Str_Tmp = "IPInterView,IsSigned"
	Arr_Tmp = split(Str_Tmp,",")	
	Vs_Sql = "select top 1 "&Str_Tmp&"  from FS_VS_SysPara"
	'response.Write(Vs_Sql)
	Set Vs_Rs = CreateObject(G_FS_RS)
	Vs_Rs.Open Vs_Sql,Conn,3,3
	if Vs_Rs.eof then Vs_Rs.AddNew
	for each Str_Tmp in Arr_Tmp
		'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
		Vs_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
	next
	'response.End()
	Vs_Rs.update
	Vs_Rs.close
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/FS_VS_SysPara.asp" )&"&ErrCodes=<li>恭喜，修改成功。</li>")
End Sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript">
<!--
function chkinput()
{
	return isEmpty('IPInterView','IPInterView_Alt') && isNumber('IPInterView','IPInterView_Alt','必须数字',false) && isEmpty('IsSigned','IsSigned_Alt') && isNumber('IsSigned','IsSigned_Alt','必须数字',false) ;
}
-->
</script>
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            <td colspan="10" align="left" class="xingmu" >系统参数设置</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="FS_VS_SysPara.asp">管理首页</a>	
	</td>
  </tr>
</table>
<%
'******************************************************************
if request.QueryString("Act")="Save" then 
	Call Save
else
	Call Add_Edit_Search		
end if
'******************************************************************
Sub Add_Edit_Search()
Dim Bol_IsEdit,IsSigned
Bol_IsEdit = false
Vs_Sql = "select top 1 sysid,IPInterView,IsSigned from FS_VS_SysPara"
Set Vs_Rs	= CreateObject(G_FS_RS)
Vs_Rs.Open Vs_Sql,Conn,1,1
if not Vs_Rs.eof then 
	Bol_IsEdit = True
	IsSigned = Vs_Rs("IsSigned")
else
	IsSigned = 0	
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" onSubmit="return chkinput();" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >系统参数设置信息</td>
	</tr>
    <tr  class="hback"> 
      <td width="140" align="right">IP时间间隔</td>
      <td width="818">
		<input type="text" size="5" maxlength="5" name="IPInterView" id="IPInterView" value="<%if Bol_IsEdit then response.Write(Vs_Rs("IPInterView")) else response.Write("1") end if%>">
        <span id="IPInterView_Alt"></span>&nbsp;<span class="tx">同一IP间隔多少时间才能再投票（单位分钟）</span></td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否注册才能投票</td>
      <td>
		<select name="IsSigned" id="IsSigned">
		<%=PrintOption(IsSigned,"0:否,1:是")%>
		</select>
        <span id="IsSigned_Alt"></span> <span class="tx">默认为游客可以投票无特殊情况不要选是</span></td>
    </tr>
   <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" 确定提交 "/> 
              &nbsp; <input type="reset" value=" 重置 " />
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%
End Sub
set Vs_Rs = Nothing
Conn.close
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





