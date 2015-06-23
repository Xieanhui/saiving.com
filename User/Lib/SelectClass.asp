<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Md5.asp" -->
<!--#include file="strlib.asp" -->
<html>
<head>
	<title>会员注册step 3 of 4 step</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
</head>
<style type="text/css">
	.LableSelectItem
	{
		background-color: highlight;
		cursor: hand;
		color: white;
		text-decoration: underline;
	}
	.LableItem
	{
		cursor: hand;
	}
</style>
<body ondragstart="return false;" onselectstart="return false;" topmargin="12" leftmargin="0">
	<table width="98%" height="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td valign="top">
				<table width="100%" border="0" cellpadding="2" cellspacing="1">
					<%
Dim TypeSql,RsTypeObj
TypeSql = "Select vClassName,ParentID,VCID from FS_ME_VocationClass  where Parentid = 0 Order by VCID desc"
Set RsTypeObj = User_Conn.Execute(TypeSql)
if Not RsTypeObj.Eof then
	do while Not RsTypeObj.Eof
					%>
					<tr parentid="<% = RsTypeObj("ParentID") %>">
						<td>
							<table border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td>
										+
									</td>
									<td>
										<span ondblclick="SubmitLable(this);" extend="False" class="LableItem" typeid="<% = RsTypeObj("VCID")&"***"&RsTypeObj("vClassName") %>" onclick="SelectFolder(this)">
											<% = RsTypeObj("vClassName") %>
										</span>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<%
		Response.Write(GetChildTypeList(RsTypeObj("VCID"),""," style=""display:none;"" "))
		RsTypeObj.MoveNext
	loop
end if
					%>
				</table>
			</td>
		</tr>
	</table>
</body>
</html>
<%
Set User_Conn = Nothing
Function GetChildTypeList(TypeID,CompatStr,ShowStr)
	Dim ChildTypeListRs,ChildTypeListStr,TempStr
	Set ChildTypeListRs = User_Conn.Execute("Select vClassName,ParentID,VCID from FS_ME_VocationClass where ParentID=" & CintStr(TypeID) & "" )
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildTypeListRs.Eof
	  	GetChildTypeList = GetChildTypeList & "<tr TypeFlag=""Class"" ParentID=""" & ChildTypeListRs("ParentID") & """ " & ShowStr & ">" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13) & Chr(10) & "<tr>"  & Chr(13) & Chr(10) & "<td>" & TempStr & "-</td>"
		GetChildTypeList = GetChildTypeList & "<td><span onDblClick=""SubmitLable(this);"" class=""LableItem"" TypeID=""" & ChildTypeListRs("VCID")&"***"&ChildTypeListRs("vClassName") & """ Extend=""False"" onClick=""SelectFolder(this)"">" & ChildTypeListRs("vClassName") & "</span></td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>"
		GetChildTypeList = GetChildTypeList & "</td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "</tr>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & GetChildTypeList(ChildTypeListRs("VCID"),TempStr,ShowStr)
		ChildTypeListRs.MoveNext
	loop
	ChildTypeListRs.Close
	Set ChildTypeListRs = Nothing
End Function
%>

<script type="text/javascript">
	function SelectLable(Obj) {
		var lists = document.getElementsByTagName('span');
		for (var i = 0; i < lists.length; i++) {
			if (lists[i].className == 'LableSelectItem') lists[i].className = 'LableItem';
		}
		Obj.className = 'LableSelectItem';
	}
	function SelectFolder(Obj) {
		var CurrObj;
		var lists = document.getElementsByTagName('span');
		for (var i = 0; i < lists.length; i++) {
			if (lists[i].className == 'LableSelectItem') lists[i].className = 'LableItem';
		}
		Obj.className = 'LableSelectItem';
		if (Obj.Extend == 'True') {
			ShowOrDisplay(Obj, 'none', true);
			Obj.Extend = 'False';
		}
		else {
			ShowOrDisplay(Obj, '', false);
			Obj.Extend = 'True';
		}
	}
	function ShowOrDisplay(Obj, Flag, Tag) {
		var lists = document.getElementsByTagName('span');
		for (var i = 0; i < lists.length; i++) {
			CurrObj = lists[i];
			if (Obj.attributes['typeid'] && CurrObj.ParentID == Obj.attributes['typeid'].value.split('***')[0]) {
				CurrObj.style.display = Flag;
				if (Tag)
					if (CurrObj.TypeFlag == 'Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0), Flag, Tag);
			}
		}
	}
	function SubmitLable(Obj) {
		var LableName = Obj.attributes['typeid'].value;
		window.top.returnValue = LableName;
		window.top.close();
	}
	window.onunload = SetReturnValue;
	function SetReturnValue() {
		if (typeof window.top.returnValue != 'string') window.top.returnValue = '';
	}
</script>

