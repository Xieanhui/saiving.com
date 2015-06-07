<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../MF_Function.asp" -->
<%
	Dim Conn,str_sysId,str_type
	MF_Default_Conn
	MF_Session_TF
	str_sysId = NoSqlHack(Request.QueryString("SysID"))
	str_type = NoSqlHack(Request.QueryString("type"))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择项目</title>
<style type="text/css">
<!--
a{text-decoration: none;} /* 链接无下划线,有为underline */ 
a:link {color: #232323;} /* 未访问的链接 */
a:visited {color: #232323;} /* 已访问的链接 */
a:hover{color: #FFCC00;} /* 鼠标在链接上 */ 
a:active {color: #FFCC00;} /* 点击激活链接 */
td ,body{
	color:#232323;
	font-size:12px;
	line-height: 18px;
}
.selected {
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
</head>
<body  onselectstart="return false;" oncontextmenu="return false;">
<%
select case str_sysId
	case "NS"
		Call NS()
	'Case "MS"
	'	Call MS()
	'Case "DS"
	'	Call DS()
	Case else
		Call NS()
End Select
Sub NS()
	if str_type = "class" then
		dim TypeSql,RsTypeObj
		TypeSql = "Select ClassID,ClassName,ParentID from FS_NS_NewsClass where ParentID='0' and ReycleTF=0 and isUrl=0 order by OrderID desc,id desc"
		Set RsTypeObj = Conn.Execute(TypeSql)
		if Not RsTypeObj.Eof then
			do while Not RsTypeObj.Eof
				Dim str_action,obj_news_rs_1
				Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
				obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& RsTypeObj("ClassID") &"'",Conn,0,1
				if obj_news_rs_1(0)>0 then
					str_action=  "<img src=""../../sys_images/+.gif""></img>"
				Else
					str_action=  "<img src=""../../sys_images/-.gif""></img>"
				End if
				obj_news_rs_1.close:set obj_news_rs_1 =nothing
				Response.Write str_action
				response.Write"<span cont="""&RsTypeObj("ClassId")&""" style=""cursor:default;"" onClick=""clicked(this);"">"&RsTypeObj("ClassName")&"</span><br />"
				Response.Write(GetChildTypeList(RsTypeObj("ClassID"),""," style=""display:none;"" "))
				RsTypeObj.movenext
			loop
		end if
	end if
	if str_type = "special" then
		dim rs
		set rs = Conn.execute("select SpecialCName,SpecialEName From FS_NS_Special where isLock=0 order by SpecialID desc")
		do while not rs.eof
				response.Write"<img src=""../../sys_images/+.gif"" /><span cont="""&rs("SpecialEName")&""" style=""cursor:default;"" onClick=""clicked(this);"">"&rs("SpecialCName")&"</span><br />"
			rs.movenext
		loop
		rs.close:set rs = nothing
	end if 
	if str_type = "style" then
		dim rs_style
		set rs_style = Conn.execute("select ID,StyleName From FS_MF_Labestyle where StyleType='NS' order by ID desc")
		do while not rs_style.eof
				response.Write"<img src=""../../sys_images/+.gif"" /><span cont="""&rs_style("ID")&""" style=""cursor:default;"" onClick=""clicked(this);"">"&rs_style("StyleName")&"</span><br />"
			rs_style.movenext
		loop
		rs_style.close:set rs_style = nothing
	end if 
	if str_type = "ungelnews" then
		dim rs_un,rs1
		set rs_un = Conn.execute("Select DisTinct UnRegulatedMain From [FS_NS_News_Unrgl] order by UnRegulatedMain DESC")
		do while not rs_un.eof
				set rs1 = Conn.execute("Select UnregNewsName From FS_NS_News_Unrgl where UnregulatedMain='"&rs_un("UnRegulatedMain")&"' order by Rows")
				response.Write"<img src=""../../sys_images/+.gif"" /><span cont="""&rs_un("UnRegulatedMain")&""" style=""cursor:default;"" onClick=""clicked(this);"">"&rs1("UnregNewsName")&"</span><br />" 
				rs1.close:set rs1=nothing
			rs_un.movenext
		loop
		rs_un.close:set rs_un=nothing
	end if 
End Sub

Function GetChildTypeList(TypeID,CompatStr,ShowStr)
	Dim ChildTypeListRs,ChildTypeListStr,TempStr
	Set ChildTypeListRs = Conn.Execute("Select ClassID,ClassName,ParentID from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 and IsUrl=0 order by OrderID desc,id desc")
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildTypeListRs.Eof
		Dim str_action_1,obj_news_rs_1s
		Set str_action_1 = server.CreateObject(G_FS_RS)
		str_action_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& ChildTypeListRs("ClassID") &"'",Conn,0,1
		if str_action_1(0)>0 then
			str_action_1=  "<img src=""../../sys_images/+.gif""></img>"
		Else
			str_action_1=  "<img src=""../../sys_images/-.gif""></img>"
		End if
		set obj_news_rs_1s =nothing
		GetChildTypeList = GetChildTypeList & "" & TempStr & ""& str_action_1 &""
		GetChildTypeList = GetChildTypeList & "<span cont="""&ChildTypeListRs("ClassId")&""" style=""cursor:default;"" onClick=""clicked(this);"">"&ChildTypeListRs("ClassName")&"</span><br />"
		GetChildTypeList = GetChildTypeList & GetChildTypeList(ChildTypeListRs("ClassID"),TempStr,ShowStr)
		ChildTypeListRs.MoveNext
	loop
	ChildTypeListRs.Close
	Set ChildTypeListRs = Nothing
End Function
%>
</body>
</html>
<script language="JavaScript">
	function clicked(Obj)
	{
		for(var i=0;i<document.body.all.length;i++)
		{
			var OldObj = document.body.all(i);
			OldObj.className='';
		}
		Obj.className='selected';
	}
</script>





