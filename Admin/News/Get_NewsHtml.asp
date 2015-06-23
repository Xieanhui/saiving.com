<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_InterFace/Refresh_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../FS_InterFace/MS_Public.asp" -->
<!--#include file="../../FS_InterFace/Other_Public.asp" -->
<!--#include file="../../FS_InterFace/ME_Public.asp" -->
<!--#include file="../../FS_InterFace/MF_Public.asp" -->
<%
Dim Conn,User_Conn,strShowErr,Fs_news,obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_c_path
MF_Default_Conn
MF_User_Conn
MF_Session_TF 
dim NewsId,ActType,NowNum
NewsID = NoSqlHack(server.HTMLEncode(Request.QueryString("ID")))
NewsID = Replace(NewsId," ","")
ActType = NoSqlHack(Request.QueryString("type"))
NowNum = NoSqlHack(server.HTMLEncode(Request.QueryString("NowNum")))
If Not NewsId<>"" Or Not ActType<>"" Then
	Response.Write "错误的参数"
	Response.End
End If
Select Case ActType
	Case "RefreshOne"
		RefreshOne NewsID,NowNum
	Case Else
		main(NewsID)
End Select

Sub RefreshOne(NewsID,NowNum)
	Dim varvalue
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Response.Charset="GB2312"
	varvalue = Refresh("NS_news",NewsID)
	If Err Then
		Response.Write "Fail$"&Err.Description
	Else
		If varvalue =true then
			Response.Write "Next$"&NowNum
		Else
			Response.Write "Err$"&NowNum
		End If
	End If
End Sub

Sub main(NewsID)
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.RefreshLen{
	height: 20px;
	width: 400px;
	border: 1px solid #104a7b;
	text-align: left;
	MARGIN-top:50px;
	margin-bottom: 5px;
}
</style>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">生成管理<a href="../../help?Lable=NS_MakHtml" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="News_Manage.asp">返回新闻管理</a>┆</div></td>
  </tr>
</table>
<div id="RefreshSchedule" style="display:none;" align="center"></div>
<script language="JavaScript" type="text/JavaScript">
var Errar=new Array();
$('RefreshSchedule').style.display="";
$('RefreshSchedule').innerHTML="<div class=\"RefreshLen\"><div class=\"xingmu\" id=\"RefreshLen\"></div></div>\
<span id=\"result_str\"></span><br><br>";
$("RefreshLen").style.width ="0%";
$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">0%</span>";
$('Result_str').innerHTML="正在准备...&nbsp;&nbsp;";

function Refresh_GO(act,Nownum)
{
	var G_REFRESH_NUM_TIME=<%= G_REFRESH_NUM_TIME %>;
	var NewsID = <%= "[["&Replace(NewsId,",","],[")&"]]" %>;
	var countnum=NewsID.length;
	var Action='';
	var StrTemp='';
	var percent;
	var goback="<a href=\"News_Manage.asp\">返回</a>";
	
	if (act=="Err"){
		Errar[Errar.length]=NewsID[Nownum-1];
	}
	if ((Nownum+1)>countnum){
		percent=100;
	}else{
		percent=((Nownum+1)/countnum)*100;
	}
	percent=Math.round(percent);
	$("RefreshLen").style.width =percent+"%";
	$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">"+percent+"%</span>";
	
	if ((Nownum+1)>countnum){
		for (i=0;i<Errar.length;i++){
			if (StrTemp==""){
				StrTemp=Errar[i].toString(10);
			}else{
				StrTemp+="；"+Errar[i].toString(10);
			}
		}
		$('result_str').innerHTML="总共要发布" + countnum + "条内容,发布成功" + (countnum-Errar.length) + "条。";
		if (Errar.length>0){
			$('result_str').innerHTML+="<br />发布失败的NewsID："+StrTemp;
		}
		$('result_str').innerHTML=$('result_str').innerHTML+"<br />发布结束&nbsp;&nbsp;"+goback;
		Nownum=0;
	}else{
		
		$('result_str').innerHTML="总共要发布" + countnum + "条内容,正在发布" + (Nownum+1) + "条内容...";
		Action="Type=RefreshOne&ID="+NewsID[Nownum]+"&NowNum="+(Nownum+1);
		if (((Nownum+1) % G_REFRESH_NUM_TIME)==0){
			window.setTimeout("Start_Refresh('Get_NewsHtml.asp','"+Action+"');",1000);
		}else{
			Start_Refresh('Get_NewsHtml.asp',Action);
		}
	}
}

function Start_Refresh(url,Action){
	var myAjax = new Ajax.Request(
		url,
		{method:'get',
		parameters:Action,
		onComplete:Refresh_Receive
		}
		);
}
function Refresh_Receive(OriginalRequest){
	var check,goback;
	var percent=0;
	var goback="<a href=\"News_Manage.asp\">返回</a>";
	if (OriginalRequest.responseText.indexOf("$")>-1){
		check=OriginalRequest.responseText.split("$");
		switch (check[0]) {
			case "Next" :
				Refresh_GO("Next",parseInt(check[1]));
				break;
			case "Err" :
				Refresh_GO("Err",parseInt(check[1]));
				break;
			default :
				$('result_str').innerHTML="发布失败，发布程序异常。&nbsp;&nbsp;"+goback+"<br>错误描述如下："+check[1];
		}
	}
	else{
		$('result_str').innerHTML="发布失败，发布程序异常。&nbsp;&nbsp;"+goback+"<br>错误描述如下："+OriginalRequest.responseText;
	}
}
Refresh_GO('Next',0);
</script>
</body>
</html>
<%
End Sub
%>






