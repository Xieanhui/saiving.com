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
'�ж��û��Ƿ�Ϊ��������Ա���޶�����·��
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
	  <td class="xingmu" colspan="6">����JS���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		<a href="../../help?Lable=News_Manage" target="_blank" style="cursor:help;'" class="sd">����</a> 
	  </td>
	</tr>
	  <tr>
	  <td colspan="6" class="hback"><a href="JS_Free_Manage.asp">����Js����</a></td>
	  </tr>
	<tr class="hback"> 
	  <td width="10%"> <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
	  <td colspan="3"> 
	  <input onBlur="<% if NoSqlHack(Request.QueryString("act"))<>"edit" then %>SetClassEName(this.value,document.JSForm.txt_ename);<% end if %>" name="txt_cname" type="text" id="txt_cname" style="width:97%" title="JS���������ƣ����ں�̨���ĺ͹����벻Ҫ����25���ַ���" maxlength="25" value="<%=FS_JsObj.cname%>"> 
	  <input type="hidden" name="hid_jsid" id="hid_jsid" value="<%=FS_JsObj.id%>"/><font color="#FF0000">*</font>
		<div align="center"></div></td>
	  <td width="32%" rowspan="11" align="center" valign="middle" id="PreviewArea"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">Ӣ������</div></td>
	  <td colspan="3"> <input name="txt_ename" type="text" id="txt_ename" style="width:97%" title="JS��Ӣ�����ƣ�����ǰ̨���ã��벻Ҫ����50���ַ��Ҳ������Ѿ����ڵ�JS������" value="<%=FS_JsObj.ename%>">
	    <font color="#FF0000">*</font> 
		<div align="center"></div></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
	  <td width="20%"> 
	  <input id="rad_Type_word" name="rad_type" type="radio" value="0" onClick="TypeChoose();ChoosePic(document.all.sel_manner.value);" title="JS���ͣ����֣�ѡ��" <%if FS_JsObj.js_type=0 then Response.Write("checked")%>>
		���� 
	  <input id="rad_Type_pic" type="radio" name="rad_Type" value="1" onClick="TypeChoose();ChoosePic(document.all.sel_manner_pic.value);" title="JS���ͣ�ͼƬ��ѡ��" 
<%if FS_JsObj.js_type=1 then Response.Write("checked")%>>
		ͼƬ</td>
	  <td width="10%" valign="middle"> <div align="center">��������</div></td>
	  <td width="28%" valign="middle"><input name="txt_newsNum" type="text" id="txt_newsNum" title="��������JSҪ���õ���������������ز�Ҫ��Ϊ��0��" style="width:100%;" value="<%if FS_JsObj.newsNum="" Then Response.Write("10") else Response.Write(FS_JsObj.newsNum)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">������ʽ</div></td>
	  <td> <select name="sel_manner" id="sel_manner" style="width:100% " title="����JS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic(this.value);">
		  <option value="1" <%if FS_JsObj.manner="1" then Response.Write("selected")%>>��ʽA</option>
		  <option value="2" <%if FS_JsObj.manner="2" then Response.Write("selected")%>>��ʽB</option>
		  <option value="3" <%if FS_JsObj.manner="3" then Response.Write("selected")%>>��ʽC</option>
		  <option value="4" <%if FS_JsObj.manner="4" then Response.Write("selected")%>>��ʽD</option>
		  <option value="5" <%if FS_JsObj.manner="5" then Response.Write("selected")%>>��ʽE</option>
		</select> </td>
	  <td valign="middle"> <div align="center">��������</div></td>
	  <td valign="middle"> <input name="txt_rowNum" type="text" id="txt_rowNum" style="width:100%;" title="��������JS��ÿ������ʾ����������������ز�Ҫ��Ϊ��0��" value="<%if FS_JsObj.rowNum="" then Response.Write("1") else Response.Write(FS_JsObj.rowNum)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">ͼƬ��ʽ</div></td>
	  <td> <select name="sel_manner_pic" id="sel_manner_pic" style="width:100% " disabled title="ͼƬJS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic(this.value);">
		  <option value="6"  <%if FS_JsObj.manner="6"  then Response.Write("selected")%>>��ʽA</option>
		  <option value="7"  <%if FS_JsObj.manner="7"  then Response.Write("selected")%>>��ʽB</option>
		  <option value="8"  <%if FS_JsObj.manner="8"  then Response.Write("selected")%>>��ʽC</option>
		  <option value="9"  <%if FS_JsObj.manner="9"  then Response.Write("selected")%>>��ʽD</option>
		  <option value="10" <%if FS_JsObj.manner="10" then Response.Write("selected")%>>��ʽE</option>
		  <option value="11" <%if FS_JsObj.manner="11" then Response.Write("selected")%>>��ʽF</option>
		  <option value="12" <%if FS_JsObj.manner="12" then Response.Write("selected")%>>��ʽG</option>
		  <option value="13" <%if FS_JsObj.manner="13" then Response.Write("selected")%>>��ʽH</option>
		  <option value="14" <%if FS_JsObj.manner="14" then Response.Write("selected")%>>��ʽI</option>
		  <option value="15" <%if FS_JsObj.manner="15" then Response.Write("selected")%>>��ʽJ</option>
		  <option value="16" <%if FS_JsObj.manner="16" then Response.Write("selected")%>>��ʽK</option>
		</select></td>
	  <td valign="middle"> <div align="center">�����о�</div></td>
	  <td valign="middle"> <input name="txt_rowSpace" type="text" id="txt_rowSpace" style="width:100%;" title="��������������������֮����о࣬��ע��������ֵ��" value="<%if FS_JsObj.rowSpace="" Then Response.Write("2") else response.Write(FS_JsObj.rowSpace)%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">����CSS</div></td>
	  <td> <input name="txt_titleCSS" type="text" id="txt_titleCSS" title="���ű����CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�" style="width:100%;" value="<%=Fs_JsObj.titleCss%>"></td>
	  <td valign="middle"> <div align="center">�¿�����</div></td>
	  <td valign="middle"> 
	  <select name="sel_OpenMode" id="sel_OpenMode" style="width:100%;">
		  <option value="1" <%if Fs_JsObj.openMode=1 then Response.Write("selected")%>>��</option>
		  <option value="0" <%if Fs_JsObj.openMode=0 then Response.Write("selected")%>>��</option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��������</div></td>
	  <td> <input name="txt_newsTitleNum" type="text" id="txt_newsTitleNum" title="ÿ�����ŵı�����ʾ������;" style="width:100%;" value="<%if Fs_JsObj.newsTitleNum="" Then Response.Write("10") else Response.Write(Fs_JsObj.newsTitleNum)%>"></td>
	  <td valign="middle"> <div align="center">��������</div></td>
	  <td valign="middle"> 
	  <select name="sel_showTimeTF" id="sel_showTimeTF" style="width:100%;" onChange="ChooseDate(this.value);" title="�������������ű�������Ƿ���ʾ�������ŵĸ���ʱ�䣡">
		  <option value="1" <%if Fs_JsObj.showTimeTF=1 then Response.Write("selected")%>>����</option>
		  <option value="0" <%if Fs_JsObj.showTimeTF=0 then Response.Write("selected")%>>������</option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">����CSS</div></td>
	  <td> <input name="txt_contentCSS" type="text" id="txt_contentCSS" title="�������ݵ�CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�" style="width:100%" value="<%=Fs_JsObj.contentCss%>"></td>
	  <td valign="middle"> <div align="center">����CSS</div></td>
	  <td valign="middle">
	  <select name="sel_dateCSS" id="sel_dateCSS" style="width:100%;" onChange="ChooseDate(this.value);" title="�������������ű�������Ƿ���ʾ�������ŵĸ���ʱ�䣡">
          <option value="1" <%If Request("ShowTimeTF")=1 then Response.Write("selected")%>>����</option>
          <option value="0" <%If Request("ShowTimeTF")=0 then Response.Write("selected")%>>������</option>
        </select> </td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��������</div></td>
	  <td> <input name="txt_contentNum" type="text" id="txt_contentNum" style="width:100% " title="Ϊ��Ҫ��ʾ�������ݵ���ʽ����ÿ�����ŵ�������ʾ������" value="<%if FS_JsObj.contentNum="" then response.write("30") else response.write(FS_JsObj.contentNum)%>"></td>
	  <td valign="middle"> <div align="center">����CSS</div></td>
	  <td valign="middle"> <input name="txt_backCSS" type="text" id="txt_backCSS" style="width:100%;" title="����JS�ı�����ʽ�������ʽ������ֱ��������ʽ���Ƽ��ɡ������ѡ�ô������ã����ÿգ�" value="<%=FS_JsObj.backCSS%>" size="14"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��������</div></td>
	  <td> <select name="txt_moreContent" id="txt_moreContent" style="width:100%;" title="����Ϊ���������ݵ���ʽ�������½Ǽ�һ���ӵ�������ҳ�����ӣ��������ʾ�����ӣ���ѡ�񡰲���ʾ����">
		  <option value="1" <%if FS_JsObj.moreContent=1 then Response.Write("selected")%>>��ʾ</option>
		  <option value="0" <%if FS_JsObj.moreContent=0 then Response.Write("selected")%>>����ʾ</option>
		</select></td>
	  <td valign="middle"> <div align="center">���ڸ�ʽ</div></td>
	  <td valign="middle"> <select name="sel_dateType" id="sel_dateType" style="width:100%;" title="���ڵ�����ʽ,Ĭ��ΪX��X�գ�">
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
		  <option value="11" <%if FS_JsObj.dateType = "11" then Response.Write("selected") end if%>><%=Month(Now)&"��"&Day(Now)&"��"%></option>
		  <option value="12" <%if FS_JsObj.dateType = "12" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"ʱ"%></option>
		  <option value="13" <%if FS_JsObj.dateType = "13" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"��"%></option>
		  <option value="14" <%if FS_JsObj.dateType = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"ʱ"&Minute(Now)&"��"%></option>
		  <option value="15" <%if FS_JsObj.dateType = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
		  <option value="16" <%if FS_JsObj.dateType = "16" then Response.Write("selected") end if%>><%=Year(Now)&"��"&Month(Now)&"��"&Day(Now)&"��"%></option>
		</select></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��������</div></td>
	  <td> <input name="txt_linkWord" type="text" value="<%=FS_JsObj.linkWord%>" id="txt_linkWord" title="Ϊ��Ҫ��ʾ�������ӵ���ʽ��������������������ͼƬ��ַ�������ͼƬ��ַ������<br>����img src=../img/1.gif border=0������ʽ�����С�src=����ΪͼƬ·������border=0��ΪͼƬ�ޱ߿�" style="width:100%;"></td>
	  <td valign="middle"> <div align="center">����CSS</div></td>
	  <td valign="middle"> <input name="txt_linkCSS" type="text" id="txt_linkCSS" style="width:100%;" title="Ϊ��������ѡ��CSS��ʽ��ֱ������CSS��ʽ���Ƽ��ɣ�" value="<%=FS_JsObj.linkCss%>"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">ͼƬ���</div></td>
	  <td> <input name="txt_picWidth" type="text" disabled id="txt_picWidth"  size="14" style="width:100%;" value="<%if FS_JsObj.picWidth="" Then Response.Write("60") else Response.Write(FS_JsObj.picWidth)%>"></td>
	  <td> <div align="center">ͼƬ�߶�</div></td>
	  <td> <input name="txt_picHeight" type="text" disabled id="txt_picHeight"  size="14" style="width:100%;" value="<%if FS_JsObj.picHeight="" Then Response.Write("60") else Response.Write(FS_JsObj.picHeight)%>"></td>
	  <td>&nbsp;</td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">����ͼƬ</div></td>
	  <td colspan="4"> 
	  <input name="txt_naviPic" type="text" id="txt_naviPic" readonly title="���ű���ǰ��ĵ���ͼ�꣬��ѡ��ͼƬ��" style="width:80%;" value="<%=FS_JsObj.naviPic%>"> 
		<input type="button" name="bnt_ChoosePic_naviPic"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_naviPic);">
		<font color="#FF0000">*</font></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">�м�ͼƬ</div></td>
	  <td colspan="4"> <input name="txt_rowBettween" readonly type="text" id="txt_rowBettween" size="26" title="��������������������֮��ļ��ͼƬ��������ѡ��ͼƬ����ť�������ã����Ϊ�գ�" style="width:80%;" value="<%=FS_JsObj.rowBettween%>"> 
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_rowBettween);"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">ͼƬ��ַ</div></td>
	  <td colspan="4"> <input name="txt_picPath" type="text" id="txt_picPath" style="width:80%;" disabled title="Ϊ����һ��ͼƬ����ʽ����ͼƬ��������ѡ��ͼƬ����ťѡ��ͼƬ��" value="<%=FS_JsObj.picPath%>"> 
		<input type="button" name="bnt_ChoosePic_picPath"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.JSForm.txt_picPath);"></td>
	</tr>
	<tr class="hback"> 
	  <td> <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;ע</div></td>
	  <td colspan="4"> 
		<textarea name="txt_info" rows="6" id="txt_info" style="width:100%" Title="��ע�����ڴ������ʱ����鿴���ԣ�"><%=FS_JsObj.info%></textarea></td>
	</tr>
	<tr>
	<td class="hback"></td>
	<td class="hback" colspan="4">
	<input type="button"  name="bnt_addJs" onClick="CheckVaild()" value="����">&nbsp;
	<input type="button" name="bnt_reset" onClick="AlertBeforReset()" value="����">
	</td>
	</tr>
  </table>
</form>
</body>
<script language="JavaScript">
//js����ѡ�����β���Ҫ��ѡ��
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
//������ʾʱ�䣬������ʱ���������
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
//��֤�������Ч��
function CheckVaild()
{
	var message="";
	var index=1;
	//js����Ϊ��
	if(document.getElementById("txt_cname").value=="")
	{
		message=(index++)+".js���Ʋ���Ϊ��\n"
	}
	//jsӢ�����Ƿ�Ϊ��
	if(document.getElementById("txt_ename").value=="")
	{
		message=message+(index++)+".jsӢ�����Ʋ���Ϊ��\n"
	}
	//���������ĺϷ���
	if(document.getElementById("txt_newsNum").value=="")
	{
		message=message+(index++)+".������������Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_newsNum").value))
	{
		message=message+(index++)+".��������ֻ��Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_newsNum").value)<=0||parseInt(document.getElementById("txt_newsNum").value)>30000)
	{
		message=message+(index++)+".�������������0��С��30000\n"
	}
	//���Ų��������ĺϷ���
	if(document.getElementById("txt_rowNum").value=="")
	{
		message=message+(index++)+".���Ų�����������Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_rowNum").value))
	{
		message=message+(index++)+".���Ų�������ֻ��Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_rowNum").value)<=0||parseInt(document.getElementById("txt_rowNum").value)>30000)
	{
		message=message+(index++)+".���Ų������������0��С��30000\n"
	}
	//�����о�ĺϷ���
	if(document.getElementById("txt_rowSpace").value=="")
	{
		message=message+(index++)+".�����о಻��Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_rowSpace").value))
	{
		message=message+(index++)+".�����о�ֻ��Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_rowSpace").value)<0||parseInt(document.getElementById("txt_rowSpace").value)>30000)
	{
		message=message+(index++)+".�����о�����ڵ���0��С��30000\n"
	}
	//���ű������ֵĺϷ���
	if(document.getElementById("txt_newsTitleNum").value=="")
	{
		message=message+(index++)+".���ű�����������Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_newsTitleNum").value))
	{
		message=message+(index++)+".���ű���������Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_newsTitleNum").value)<0||parseInt(document.getElementById("txt_newsTitleNum").value)>30000)
	{
		message=message+(index++)+".���ű�����������ڵ���0��С��30000\n"
	}
	//���������ĺϷ���
	if(document.getElementById("txt_contentNum").value=="")
	{
		message=message+(index++)+".������������Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_contentNum").value))
	{
		message=message+(index++)+"����������Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_contentNum").value)<0||parseInt(document.getElementById("txt_contentNum").value)>30000)
	{
		message=message+(index++)+".��������ڵ���0��С��30000\n"
	}
	//�ж�ͼƬ���ֵ����Ч��
	if(document.getElementById("txt_picWidth").value=="")
	{
		message=message+(index++)+".ͼƬ��Ȳ���Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_picWidth").value))
	{
		message=message+(index++)+".ͼƬ�����Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_picWidth").value)<0||parseInt(document.getElementById("txt_picWidth").value)>30000)
	{
		message=message+(index++)+".ͼƬ��������0��С��30000\n"
	}
	//�ж�ͼƬ�߶�ֵ����Ч��
	if(document.getElementById("txt_picHeight").value=="")
	{
		message=message+(index++)+".ͼƬ�߶Ȳ���Ϊ��\n"
	}
	else if(isNaN(document.getElementById("txt_picHeight").value))
	{
		message=message+(index++)+".ͼƬ�߶���Ϊ����\n"
	}else if(parseInt(document.getElementById("txt_picHeight").value)<0||parseInt(document.getElementById("txt_picHeight").value)>30000)
	{
		message=message+(index++)+".ͼƬ�߶������0��С��30000\n"
	}
		//����ͼƬ
	if(document.getElementById("txt_naviPic").value=="")
	{
		message=message+(index++)+".����ͼƬ��ַ����Ϊ��\n"
	}
	if(message!="")
	{
		alert(message+"<%=G_COPYRIGHT%>");
	}else
	{
		document.JSForm.submit();
	}
}
//����ǰ��ʾ
function AlertBeforReset()
{
	if(confirm("�Ƿ�Ҫ������������Ŀ��"))
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
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->