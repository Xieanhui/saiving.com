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
'session判断
MF_Session_TF 
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001")
'得到会员组列表
dim Fs_news
set Fs_news = new Cls_News
Fs_News.GetSysParam()
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
    <td class="xingmu">模板查看<a href="../../help?Lable=NS_Class_Templet" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_ToTemplet.asp">首页</a> 
        &nbsp;|&nbsp; <a href="Class_ToTempletRead.asp">栏目模板查看</a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="ClassForm" method="post" action="Class_Action.asp">
    <tr class="xingmu"> 
      <td width="26%" height="25" class="xingmu"><div align="center">栏目中文[栏目英文]</div></td>
      <td width="37%" class="xingmu"> <div align="center">栏目模板</div></td>
      <td width="37%" class="xingmu">新闻模板</td>
    </tr>
    <%
	Dim obj_news_rs,obj_news_rs_1,isUrlStr
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,Templet,NewsTemplet from FS_NS_NewsClass where Parentid  = '0'   and ReycleTF=0  Order by Orderid desc,ID desc",Conn,1,3
	Do while Not obj_news_rs.eof 
	%>
    <tr class="hback"> 
      <td height="22" class="hback"> <% 
		if obj_news_rs("IsUrl") = 1 then
			isUrlStr = ""
		Else
			isUrlStr = "["&obj_news_rs("ClassEName")&"]"
		End if
		Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
		obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
		if obj_news_rs_1(0)>0 then
			Response.Write  "<img src=""images/+.gif""></img>"&obj_news_rs("ClassName")&"&nbsp;<font style=""font-size:9px;"">"& isUrlStr &"</font>"
		Else
			Response.Write  "<img src=""images/-.gif""></img>"&obj_news_rs("ClassName")&"</a>&nbsp;<font style=""font-size:9px;"">"& isUrlStr &"</font>"
		End if
		obj_news_rs_1.close:set obj_news_rs_1 =nothing
		%> </td>
      <td align="center" class="hback"> <div align="left">
          <% = obj_news_rs("Templet")%>
        </div>
        <div align="center"></div></td>
      <td align="center" class="hback"> <div align="left">
          <% = obj_news_rs("NewsTemplet")%>
        </div></td>
    </tr>
    <%
		Response.Write(GetChildNewsList(obj_news_rs("ClassID"),""))
		obj_news_rs.MoveNext
	loop
%>
  </form>
</table>
</body>
</html>
<%
set Fs_news = nothing
%>
<%
	Function GetChildNewsList(TypeID,CompatStr)  
		Dim ChildNewsRs,ChildTypeListStr,TempStr,TmpStr,f_isUrlStr
		Set ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassEName,ClassID,IsUrl,NewsTemplet,Templet from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "'  and ReycleTF=0  order by Orderid desc,id desc" )
		TempStr =CompatStr & "<img src=""images/L.gif""></img>"
		do while Not ChildNewsRs.Eof
			  TmpStr = ""
			  if ChildNewsRs("IsUrl") = 1 then
				  TmpStr = TmpStr & "<font color=red>外部</font>&nbsp;|&nbsp;" 
			  Else
				 TmpStr = TmpStr & "系统&nbsp;|&nbsp;" 
			  End if
	  		GetChildNewsList = GetChildNewsList & "<tr>"&Chr(13) & Chr(10)
			if ChildNewsRs("IsUrl") = 1 then
				f_isUrlStr = ""
			Else
				f_isUrlStr = "["&ChildNewsRs("ClassEName")&"]"
			End if
			GetChildNewsList = GetChildNewsList & "<td class=""hback"">&nbsp;"& TempStr &"<Img src=""images/-.gif""></img><a href=""Class_add.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=edit"">" & ChildNewsRs("ClassName") & "</a>&nbsp;<font style=""font-size:9px;"">"& f_isUrlStr &"</font></td>" & Chr(13) & Chr(10) 
			GetChildNewsList = GetChildNewsList & "<td class=""hback"" align=""left"">"&ChildNewsRs("Templet")&"</td>" & Chr(13) & Chr(10)
			GetChildNewsList = GetChildNewsList & "<td class=""hback"" align=""left"">"&ChildNewsRs("NewsTemplet")&"</td>" & Chr(13) & Chr(10)
			GetChildNewsList = GetChildNewsList & "</tr>" & Chr(13) & Chr(10)
			GetChildNewsList = GetChildNewsList &GetChildNewsList(ChildNewsRs("ClassID"),TempStr)
			ChildNewsRs.MoveNext
		loop
		ChildNewsRs.Close
		Set ChildNewsRs = Nothing
	End Function
	%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





