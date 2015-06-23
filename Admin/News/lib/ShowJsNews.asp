<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="cls_main.asp" -->
<!--#include file="Cls_Js.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
Dim Conn,newsRs,jsName,OpType,JsID,JsTypeSql,JsTypeRs
MF_Default_Conn
MF_Session_TF
jsName=NoSqlHack(request.querystring("JSName"))
'-----------------------------------------------------------------------------
'删除JS列表的一条内容
OpType=NoSqlHack(Request.QueryString("Type"))
JsID=Request.QueryString("JsID")
If OpType="DelJs" And JsID<>"" Then
	Conn.execute("Delete From FS_NS_FreeJsFile Where ID="&Clng(JsID))
	Response.write JsList(jsName)
	'更新JS文件--------------------------------------------------------------
	JsTypeSql="Select Manner From FS_NS_Freejs Where Ename='"&NoSqlHack(jsName)&"'"
	Set JsTypeRs=Conn.execute(JsTypeSql)
	If Not JsTypeRs.Eof Then
		Call UpdateJS(jsName,JsTypeRs("Manner"))
	End If
	JsTypeRs.Close
	Set JsTypeRs=Nothing
	Response.End
'-----------------------------------------------------------------------------
Else
%>
<HTML>
<HEAD>
<TITLE>CMS5.0</TITLE>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<script language="javascript" src=../../../FS_Inc/Prototype.js></script>
<script language="javascript" src=../../../FS_Inc/public.js></script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->
<div id="JsList">
<%
Response.write JsList(jsName)
%>
</div>
</BODY>
</HTML>
<script language="javascript">
//-------------------------------------------------------------------------------
//删除JS新闻的AJAX函数
function DelJS(JsID)
{
	if (!confirm("你确定要删除此新闻吗？")){
		return false;
	}else{
		var  options={  
			   method:'get',  
			   parameters:"Type=DelJs&JsID="+JsID+"&JsName=<%=jsName %>",  
			   onComplete:function(transport)
				{ 
					var returnvalue=transport.responseText;
					$('JsList').innerHTML=returnvalue;
					if (returnvalue.indexOf("??")>-1)
						$('JsList').innerHTML='Error';
					else
						$('JsList').innerHTML=returnvalue;
				}  
		   };  
		new  Ajax.Request('ShowJsNews.asp',options); 
	}
}
//-------------------------------------------------------------------------------
</script>
<%
End If

'-----------------------------------------------------------------------------
'获取JS新闻列表
Function JsList(jsName)
	Dim TempStr
	Dim index
	If jsName<>"" then
		TempStr="<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		TempStr=TempStr&"<tr>" 
		TempStr=TempStr&"<td class=""xingmu"" colspan=""3"">调用新闻:</td>"
		TempStr=TempStr&"</tr>"
		Set newsRs=Conn.execute("Select title,ID,JsName from FS_NS_FreeJsFile where JSName='"&jsName&"'")
		index=0
		While Not newsRs.eof
			index=index+1
			TempStr=TempStr&"<tr>"&vbcrlf
			TempStr=TempStr&"<td width=""8"" class=""hback""><img src=""../../images/all_article_icon.gif""></td><td class=""hback"">"&index&":"&newsRs("title")&"</td><td width=""30"" class=""hback""><a href=""删除此新闻"" onclick=""DelJS('"&newsRs("ID")&"');return false;"" title=""删除此新闻"">删除</a></td>"&vbcrlf
			TempStr=TempStr&"</tr>"
			newsRs.movenext
		wend
		TempStr=TempStr&"</table>"
		newsRs.close()
		Set newsRs=nothing
		JsList = TempStr
	Else
		JsList = ""
	End if
End Function
'-----------------------------------------------------------------------------
'更新JS文件过程
Sub UpdateJS(JSEName,JsType)
  	Dim JSClassObj,ReturnValue,TempRs
	Set TempRs=Conn.Execute("Select NewsDir from FS_NS_SysParam")
	If TempRs.eof Then
		Response.Redirect("/error.asp?ErrCodes=<li>出现异常</li>")
		Response.End()
	End if
	Set JSClassObj = New Cls_Js
	JSClassObj.SysRootDir=TempRs("NewsDir")
	Select Case Cint(JsType)
		case 1   ReturnValue = JSClassObj.WCssA(JSEName,True)
		case 2   ReturnValue = JSClassObj.WCssB(JSEName,True)
		case 3   ReturnValue = JSClassObj.WCssC(JSEName,True)
		case 4   ReturnValue = JSClassObj.WCssD(JSEName,True)
		case 5   ReturnValue = JSClassObj.WCssE(JSEName,True)
		case 6   ReturnValue = JSClassObj.PCssA(JSEName,True)
		case 7   ReturnValue = JSClassObj.PCssB(JSEName,True)
		case 8   ReturnValue = JSClassObj.PCssC(JSEName,True)
		case 9   ReturnValue = JSClassObj.PCssD(JSEName,True)
		case 10  ReturnValue = JSClassObj.PCssE(JSEName,True)
		case 11  ReturnValue = JSClassObj.PCssF(JSEName,True)
		case 12  ReturnValue = JSClassObj.PCssG(JSEName,True)
		case 13  ReturnValue = JSClassObj.PCssH(JSEName,True)
		case 14  ReturnValue = JSClassObj.PCssI(JSEName,True)
		case 15  ReturnValue = JSClassObj.PCssJ(JSEName,True)
		case 16  ReturnValue = JSClassObj.PCssK(JSEName,True)
		case 17  ReturnValue = JSClassObj.PCssL(JSEName,True)
	End Select
	ReturnValue=""
	TempRs.Close
	Set TempRs=Nothing
End Sub
'-----------------------------------------------------------------------------
%>