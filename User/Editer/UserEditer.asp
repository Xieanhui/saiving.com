<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim ContentID
ContentID=Request.QueryString("id")
if ContentID="" then ContentID="Content"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ӱ༭��</title>
</head>
<link rel="stylesheet" href="../../Editer/Editer.css">
<script language="JavaScript" src="../../Editer/Editer.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body onLoad="return LoadEditFile('../../Images/Editer/','../../Editer/','<% = ContentID %>',1);">
<div id="Toolbar"></div>
<div class="Toolbar_row" style="border-bottom:none">
<div style="float:left; height:30px; padding-top:10px;"><div id="ShowObject"></div></div>
<div class="Toolbar_row_Button" style="float:right"><a href="javascript:void(0);"><img src="../../Images/Editer/tablemodify.gif" width="23" height="22"  title="����" onClick="ExeEditAttribute();"></a></div>
<div class="Toolbar_row_Button" style="float:right"><a href="javascript:void(0);"><img src="../../Images/Editer/delLable.gif" width="23" height="22"  title="ɾ����ǩ" onClick="DeleteHTMLTag();"></a></div>
</div>
</div>
<div id="Toolbar_Area"><iframe name="EditArea" class="Composition" ID="EditArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="100%;" width="100%" scrolling="yes"></iframe></div>
<div class="Toolbar_row" id="SetModeArea" style="border-bottom:none">
<div class="Toolbar_row_Button2 ModeBarBtnOff" id=Editer_CODE onClick="setMode('CODE');"><img src="../../Images/Editer/CodeMode.GIF" width="50" height="15"></div>
<div class="Toolbar_row_Button2 ModeBarBtnOff" id=Editer_VIEW onClick="setMode('VIEW');"><img src="../../Images/Editer/PreviewMode.gif" width="50" height="15"></div>
<div class="Toolbar_row_Button2 ModeBarBtnOn" id=Editer_EDIT onClick="setMode('EDIT');"><img src="../../Images/Editer/EditMode.GIF" width="50" height="15"></div>
<div class="Toolbar_row_Button2 ModeBarBtnOff" id=Editer_TEXT onClick="setMode('TEXT');"><img src="../../Images/Editer/TextMode.GIF" width="50" height="15"></div>
</div>
</body>
</html>