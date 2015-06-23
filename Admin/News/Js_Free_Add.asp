<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="lib/cls_js.asp"-->
<%
Dim Conn,sRootDir,str_CurrPath,FS_JsObj,jsid
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
'判断用户是否为超级管理员，限定访问路径
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("NS037") then Err_Show
jsid=NoSqlHack(Request.QueryString("jsid"))
Set FS_JsObj=New Cls_Js
if jsid<>"" then
	if isNumeric(jsid) then
		FS_JsObj.getFreeJsParam(jsid)
	End if
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>CMS5.0</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
</head>
<body>
<script src="js/Public.js" language="JavaScript"></script>
<%if NoSqlHack(Request.QueryString("act"))="edit" then%>
<form action="Js_Free_Action.asp?act=edit" method="post" name="JSForm">
<%else%>
<form action="Js_Free_Action.asp?act=add" method="post" name="JSForm">
<%End if%>
  <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<tr> 
	  <td class="xingmu" colspan="6">自由JS添加&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		<a href="../../help?Lable=News_Manage" target="_blank" style="cursor:help;'" class="sd">帮助</a> 
	  </td>
	</tr>
	  <tr>
	  <td colspan="6" class="hback"><a href="JS_Free_Manage.asp">自由Js管理</a></td>
	  </tr>
	<tr class="hback"> 
	  <td width="10%"> <div align="center">名&nbsp;&nbsp;&nbsp;&nbsp;称</div></td>
	  <td colspan="3"> 
	  <input onBlur="<% if NoSqlHack(Request.QueryString("act"))<>"edit" then %>SetClassEName(this.value,document.JSForm.txt_ename);<% end if %>" name="txt_cname" type="text" id="txt_cname" style="width:97%" title="JS的中文名称，便于后台查阅和管理，请不要超过25个字符！" maxlength="25" value="<%=FS_JsObj.cname%>"> 
	  <input type="hidden" name="hid_jsid" id="hid_jsid" value="<%=FS_JsObj.id%>"/><font color="#FF0000">*</font>
		<div align="center"></div></td>
	  <td width="32%" rowspan="11" align="center" valign="middle" id="PreviewArea"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">英文名称</div></td>
	  <td colspan="3"> <input name="txt_ename" type="text" id="txt_ename" style="width:97%" title="JS的英文名称，用于前台调用，请不要超过50个字符且不能与已经存在的JS重名！" value="<%=FS_JsObj.ename%>">
	    <font color="#FF0000">*</font> 
		<div align="center"></div></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">类&nbsp;&nbsp;&nbsp;&nbsp;型</div></td>
	  <td width="20%"> 
	  <input id="rad_Type_word" name="rad_type" type="radio" value="0" onClick="TypeChoose();ChoosePic(document.all.sel_manner.value);" title="JS类型（文字）选择！" <%if FS_JsObj.js_type=0 then Response.Write("checked")%>>
		文字 
	  <input id="rad_Type_pic" type="radio" name="rad_Type" value="1" onClick="TypeChoose();ChoosePic(document.all.sel_manner_pic.value);" title="JS类型（图片）选择！" 
<%if FS_JsObj.js_type=1 then Response.Write("checked")%>>
		图片</td>
	  <td width="10%" valign="middle"> <div align="center">新闻条数</div></td>
	  <td width="28%" valign="middle"><input name="txt_newsNum" type="text" id="txt_newsNum" title="此项设置JS要调用的新闻条数，请务必不要置为‘0’" style="width:100%;" value="<%if FS_JsObj.newsNum="" Then Response.Write("10") else Response.Write(FS_JsObj.newsNum)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">文字样式</div></td>
	  <td> <select name="sel_manner" id="sel_manner" style="width:100% " title="文字JS样式选择，上面有此样式的预览！" onChange="ChoosePic(this.value);">
		  <option value="1" <%if FS_JsObj.manner="1" then Response.Write("selected")%>>样式A</option>
		  <option value="2" <%if FS_JsObj.manner="2" then Response.Write("selected")%>>样式B</option>
		  <option value="3" <%if FS_JsObj.manner="3" then Response.Write("selected")%>>样式C</option>
		  <option value="4" <%if FS_JsObj.manner="4" then Response.Write("selected")%>>样式D</option>
		  <option value="5" <%if FS_JsObj.manner="5" then Response.Write("selected")%>>样式E</option>
		</select> </td>
	  <td valign="middle"> <div align="center">并排条数</div></td>
	  <td valign="middle"> <input name="txt_rowNum" type="text" id="txt_rowNum" style="width:100%;" title="此项设置JS在每行内显示的新闻条数，请务必不要置为‘0’" value="<%if FS_JsObj.rowNum="" then Response.Write("1") else Response.Write(FS_JsObj.rowNum)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">图片样式</div></td>
	  <td> <select name="sel_manner_pic" id="sel_manner_pic" style="width:100% " disabled title="图片JS样式选择，上面有此样式的预览！" onChange="ChoosePic(this.value);">
		  <option value="6"  <%if FS_JsObj.manner="6"  then Response.Write("selected")%>>样式A</option>
		  <option value="7"  <%if FS_JsObj.manner="7"  then Response.Write("selected")%>>样式B</option>
		  <option value="8"  <%if FS_JsObj.manner="8"  then Response.Write("selected")%>>样式C</option>
		  <option value="9"  <%if FS_JsObj.manner="9"  then Response.Write("selected")%>>样式D</option>
		  <option value="10" <%if FS_JsObj.manner="10" then Response.Write("selected")%>>样式E</option>
		  <option value="11" <%if FS_JsObj.manner="11" then Response.Write("selected")%>>样式F</option>
		  <option value="12" <%if FS_JsObj.manner="12" then Response.Write("selected")%>>样式G</option>
		  <option value="13" <%if FS_JsObj.manner="13" then Response.Write("selected")%>>样式H</option>
		  <option value="14" <%if FS_JsObj.manner="14" then Response.Write("selected")%>>样式I</option>
		  <option value="15" <%if FS_JsObj.manner="15" then Response.Write("selected")%>>样式J</option>
		  <option value="16" <%if FS_JsObj.manner="16" then Response.Write("selected")%>>样式K</option>
		</select></td>
	  <td valign="middle"> <div align="center">新闻行距</div></td>
	  <td valign="middle"> <input name="txt_rowSpace" type="text" id="txt_rowSpace" style="width:100%;" title="此项设置上下两条新闻之间的行距，请注意输入数值！" value="<%if FS_JsObj.rowSpace="" Then Response.Write("2") else response.Write(FS_JsObj.rowSpace)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">标题CSS</div></td>
	  <td> <input name="txt_titleCSS" type="text" id="txt_titleCSS" title="新闻标题的CSS样式表。请直接输入样式名称。如果不选用此项设置，请置空！" style="width:100%;" value="<%=Fs_JsObj.titleCss%>"></td>
	  <td valign="middle"> <div align="center">新开窗口</div></td>
	  <td valign="middle"> 
	  <select name="sel_OpenMode" id="sel_OpenMode" style="width:100%;">
		  <option value="1" <%if Fs_JsObj.openMode=1 then Response.Write("selected")%>>是</option>
		  <option value="0" <%if Fs_JsObj.openMode=0 then Response.Write("selected")%>>否</option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">标题字数</div></td>
	  <td> <input name="txt_newsTitleNum" type="text" id="txt_newsTitleNum" title="每条新闻的标题显示字数！;" style="width:100%;" value="<%if Fs_JsObj.newsTitleNum="" Then Response.Write("10") else Response.Write(Fs_JsObj.newsTitleNum)%>"></td>
	  <td valign="middle"> <div align="center">新闻日期</div></td>
	  <td valign="middle"> 
	  <select name="sel_showTimeTF" id="sel_showTimeTF" style="width:100%;" onChange="ChooseDate(this.value);" title="此项设置在新闻标题后面是否显示本条新闻的更新时间！">
		  <option value="1" <%if Fs_JsObj.showTimeTF=1 then Response.Write("selected")%>>调用</option>
		  <option value="0" <%if Fs_JsObj.showTimeTF=0 then Response.Write("selected")%>>不调用</option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">内容CSS</div></td>
	  <td> <input name="txt_contentCSS" type="text" id="txt_contentCSS" title="新闻内容的CSS样式表。请直接输入样式名称。如果不选用此项设置，请置空！" style="width:100%" value="<%=Fs_JsObj.contentCss%>"></td>
	  <td valign="middle"> <div align="center">日期CSS</div></td>
	  <td valign="middle">
	  <select name="sel_dateCSS" id="sel_dateCSS" style="width:100%;" onChange="ChooseDate(this.value);" title="此项设置在新闻标题后面是否显示本条新闻的更新时间！">
          <option value="1" <%If Request("ShowTimeTF")=1 then Response.Write("selected")%>>调用</option>
          <option value="0" <%If Request("ShowTimeTF")=0 then Response.Write("selected")%>>不调用</option>
        </select> </td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">内容字数</div></td>
	  <td> <input name="txt_contentNum" type="text" id="txt_contentNum" style="width:100% " title="为需要显示新闻内容的样式设置每条新闻的内容显示字数！" value="<%if FS_JsObj.contentNum="" then response.write("30") else response.write(FS_JsObj.contentNum)%>"></td>
	  <td valign="middle"> <div align="center">背景CSS</div></td>
	  <td valign="middle"> <input name="txt_backCSS" type="text" id="txt_backCSS" style="width:100%;" title="整体JS的背景样式（表格样式），请直接输入样式名称即可。如果不选用此项设置，请置空！" value="<%=FS_JsObj.backCSS%>" size="14"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">更多链接</div></td>
	  <td> <select name="txt_moreContent" id="txt_moreContent" style="width:100%;" title="此项为有新闻内容的样式在其右下角加一链接到该新闻页的链接，如果不显示此链接，请选择“不显示”！">
		  <option value="1" <%if FS_JsObj.moreContent=1 then Response.Write("selected")%>>显示</option>
		  <option value="0" <%if FS_JsObj.moreContent=0 then Response.Write("selected")%>>不显示</option>
		</select></td>
	  <td valign="middle"> <div align="center">日期格式</div></td>
	  <td valign="middle"> <select name="sel_dateType" id="sel_dateType" style="width:100%;" title="日期调用样式,默认为X月X日！">
		  <option value="1" <%if FS_JsObj.dateType = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
		  <option value="2" <%if FS_JsObj.dateType = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
		  <option value="3" <%if FS_JsObj.dateType = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
		  <option value="4" <%if FS_JsObj.dateType = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
		  <option value="5" <%if FS_JsObj.dateType = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
		  <option value="6" <%if FS_JsObj.dateType = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
		  <option value="7" <%if FS_JsObj.dateType = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
		  <option value="8" <%if FS_JsObj.dateType = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
		  <option value="9" <%if FS_JsObj.dateType = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
		  <option value="10" <%if FS_JsObj.dateType = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
		  <option value="11" <%if FS_JsObj.dateType = "11" then Response.Write("selected") end if%>><%=Month(Now)&"月"&Day(Now)&"日"%></option>
		  <option value="12" <%if FS_JsObj.dateType = "12" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"时"%></option>
		  <option value="13" <%if FS_JsObj.dateType = "13" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"点"%></option>
		  <option value="14" <%if FS_JsObj.dateType = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"时"&Minute(Now)&"分"%></option>
		  <option value="15" <%if FS_JsObj.dateType = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
		  <option value="16" <%if FS_JsObj.dateType = "16" then Response.Write("selected") end if%>><%=Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日"%></option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">链接字样</div></td>
	  <td> <input name="txt_linkWord" type="text" value="<%=FS_JsObj.linkWord%>" id="txt_linkWord" title="为需要显示新闻链接的样式设置链接字样，可以是图片地址，如果是图片地址，请用<br>‘＜img src=../img/1.gif border=0＞’样式，其中‘src=’后为图片路径，‘border=0’为图片无边框！" style="width:100%;"></td>
	  <td valign="middle"> <div align="center">链接CSS</div></td>
	  <td valign="middle"> <input name="txt_linkCSS" type="text" id="txt_linkCSS" style="width:100%;" title="为链接字样选择CSS样式，直接输入CSS样式名称即可！" value="<%=FS_JsObj.linkCss%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">图片宽度</div></td>
	  <td> <input name="txt_picWidth" type="text" disabled id="txt_picWidth"  size="14" style="width:100%;" value="<%if FS_JsObj.picWidth="" Then Response.Write("60") else Response.Write(FS_JsObj.picWidth)%>"></td>
	  <td> <div align="center">图片高度</div></td>
	  <td> <input name="txt_picHeight" type="text" disabled id="txt_picHeight"  size="14" style="width:100%;" value="<%if FS_JsObj.picHeight="" Then Response.Write("60") else Response.Write(FS_JsObj.picHeight)%>"></td>
	  <td>&nbsp;</td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">导航图片</div></td>
	  <td colspan="4"> 
	  <input name="txt_naviPic" type="text" id="txt_naviPic" readonly title="新闻标题前面的导航图标，请选择图片！" style="width:80%;" value="<%=FS_JsObj.naviPic%>"> 
		<input type="button" name="bnt_ChoosePic_naviPic"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_naviPic);">
		<font color="#FF0000">*</font></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">行间图片</div></td>
	  <td colspan="4"> <input name="txt_rowBettween" readonly type="text" id="txt_rowBettween" size="26" title="此项设置上下两条新闻之间的间隔图片，请点击“选择图片”按钮进行设置，亦可为空！" style="width:80%;" value="<%=FS_JsObj.rowBettween%>"> 
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_rowBettween);"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">图片地址</div></td>
	  <td colspan="4"> <input name="txt_picPath" type="text" id="txt_picPath" style="width:80%;" disabled title="为仅需一张图片的样式设置图片，请点击‘选择图片’按钮选择图片！" value="<%=FS_JsObj.picPath%>"> 
		<input type="button" name="bnt_ChoosePic_picPath"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_picPath);"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">备&nbsp;&nbsp;&nbsp;&nbsp;注</div></td>
	  <td colspan="4"> 
		<textarea name="txt_info" rows="6" id="txt_info" style="width:100%" Title="备注，用于代码调用时方便查看属性！"><%=FS_JsObj.info%></textarea></td>
	</tr>
	<tr>
	<td class="hback"></td>
	<td class="hback" colspan="4">
	<input type="button"  name="bnt_addJs" onClick="CheckVaild()" value="保存">&nbsp;
	<input type="button" name="bnt_reset" onClick="AlertBeforReset()" value="重置">
	</td>
	</tr>
  </table>
</form>
</body>
<script language="JavaScript">
//js类型选择，屏蔽不需要的选项
TypeChoose();
function TypeChoose()
{
	if (document.JSForm.rad_Type_word.checked==true)
	{ 
		document.JSForm.sel_manner.disabled=false;
		document.JSForm.sel_manner_pic.disabled=true;
		document.JSForm.txt_picPath.disabled=true;
		document.JSForm.bnt_ChoosePic_picPath.disabled=true;
		document.JSForm.txt_picWidth.disabled=true;
		document.JSForm.txt_picHeight.disabled=true;
	}
	else
	{
		document.JSForm.sel_manner.disabled=true;
		document.JSForm.sel_manner_pic.disabled=false;
		document.JSForm.txt_picPath.disabled=false;
		document.JSForm.bnt_ChoosePic_picPath.disabled=false;
		document.JSForm.txt_picWidth.disabled=false;
		document.JSForm.txt_picHeight.disabled=false;
	}
}
ChoosePic("<%=FS_JsObj.manner%>")
function ChoosePic(style_id)
{
	if(style_id=="")
		style_id=1;
	document.all.PreviewArea.innerHTML="<img src='images/JsStyle/Css"+style_id+".gif' />"
}
ChooseDate("<%=Fs_JsObj.showTimeTF%>")
//若不显示时间，则屏蔽时间调整参数
function ChooseDate(DateStr)
{ 
	if (DateStr==1)
	{
		document.JSForm.sel_dateType.disabled=false;
		document.JSForm.sel_dateCSS.disabled=false;
	}
	else
	{
		document.JSForm.sel_dateType.disabled=true;
		document.JSForm.sel_dateCSS.disabled=true;
	}
}
//验证输入的有效性
function CheckVaild()
{
	var message="";
	var index=1;
	//js名是为空
	if(document.getElementById("txt_cname").value=="")
	{
		message=(index++)+".js名称不能为空\n"
	}
	//js英文名是否为空
	if(document.getElementById("txt_ename").value=="")
	{
		message=message+(index++)+".js英文名称不能为空\n"
	}
	//新闻条数的合法性
	if(document.getElementById("txt_newsNum").value=="")
	{
		message=message+(index++)+".新闻条数不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_newsNum").value))
	{
		message=message+(index++)+".新闻条数只能为数字\n"
	}else if(parseInt(document.getElementById("txt_newsNum").value)<=0||parseInt(document.getElementById("txt_newsNum").value)>30000)
	{
		message=message+(index++)+".新闻条数须大于0且小于30000\n"
	}
	//新闻并列条数的合法性
	if(document.getElementById("txt_rowNum").value=="")
	{
		message=message+(index++)+".新闻并列条数不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_rowNum").value))
	{
		message=message+(index++)+".新闻并列条数只能为数字\n"
	}else if(parseInt(document.getElementById("txt_rowNum").value)<=0||parseInt(document.getElementById("txt_rowNum").value)>30000)
	{
		message=message+(index++)+".新闻并列条数须大于0且小于30000\n"
	}
	//新闻行距的合法性
	if(document.getElementById("txt_rowSpace").value=="")
	{
		message=message+(index++)+".新闻行距不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_rowSpace").value))
	{
		message=message+(index++)+".新闻行距只能为数字\n"
	}else if(parseInt(document.getElementById("txt_rowSpace").value)<0||parseInt(document.getElementById("txt_rowSpace").value)>30000)
	{
		message=message+(index++)+".新闻行距须大于等于0且小于30000\n"
	}
	//新闻标题数字的合法性
	if(document.getElementById("txt_newsTitleNum").value=="")
	{
		message=message+(index++)+".新闻标题字数不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_newsTitleNum").value))
	{
		message=message+(index++)+".新闻标题字数须为数字\n"
	}else if(parseInt(document.getElementById("txt_newsTitleNum").value)<0||parseInt(document.getElementById("txt_newsTitleNum").value)>30000)
	{
		message=message+(index++)+".新闻标题字数须大于等于0且小于30000\n"
	}
	//内容字数的合法性
	if(document.getElementById("txt_contentNum").value=="")
	{
		message=message+(index++)+".内容字数不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_contentNum").value))
	{
		message=message+(index++)+"内容字数须为数字\n"
	}else if(parseInt(document.getElementById("txt_contentNum").value)<0||parseInt(document.getElementById("txt_contentNum").value)>30000)
	{
		message=message+(index++)+".内容须大于等于0且小于30000\n"
	}
	//判断图片宽度值的有效性
	if(document.getElementById("txt_picWidth").value=="")
	{
		message=message+(index++)+".图片宽度不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_picWidth").value))
	{
		message=message+(index++)+".图片宽度须为数字\n"
	}else if(parseInt(document.getElementById("txt_picWidth").value)<0||parseInt(document.getElementById("txt_picWidth").value)>30000)
	{
		message=message+(index++)+".图片宽度须大于0且小于30000\n"
	}
	//判断图片高度值的有效性
	if(document.getElementById("txt_picHeight").value=="")
	{
		message=message+(index++)+".图片高度不能为空\n"
	}
	else if(isNaN(document.getElementById("txt_picHeight").value))
	{
		message=message+(index++)+".图片高度须为数字\n"
	}else if(parseInt(document.getElementById("txt_picHeight").value)<0||parseInt(document.getElementById("txt_picHeight").value)>30000)
	{
		message=message+(index++)+".图片高度须大于0且小于30000\n"
	}
		//导航图片
	if(document.getElementById("txt_naviPic").value=="")
	{
		message=message+(index++)+".导航图片地址不能为空\n"
	}
	if(message!="")
	{
		alert(message+"<%=G_COPYRIGHT%>");
	}else
	{
		document.JSForm.submit();
	}
}
//重置前提示
function AlertBeforReset()
{
	if(confirm("是否要重置整个表单项目？"))
	{
		document.JSForm.reset();
	}
}
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->