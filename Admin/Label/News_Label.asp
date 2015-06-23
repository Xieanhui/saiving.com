<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	MF_Default_Conn
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='NS' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	Dim obj_special_Rs,label_special_List
	label_special_List=""
	Set  obj_special_Rs = server.CreateObject(G_FS_RS)
	obj_special_Rs.Open "Select SpecialEName,SpecialCName from FS_NS_Special  Order by  SpecialID desc",Conn,1,3
	do while Not obj_special_Rs.eof 
		label_special_List = label_special_List&"<option value="""& obj_special_Rs(0)&""">"& obj_special_Rs(1)&"</option>"
		obj_special_Rs.movenext
	loop
	obj_special_Rs.close:set obj_special_Rs = nothing
	dim str_CurrPath,sRootDir
	if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
	If Temp_Admin_Is_Super = 1 then
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	Else
		If Temp_Admin_FilesTF = 0 Then
			str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
		Else
			str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
		End If	
	End if
%>
<html>
<head>
	<title>���ű�ǩ����</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css" />
	<base target="self" />
</head>
<body class="hback">

	<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>

	<table width="98%" height="100" border="0" align="center" cellpadding="3" cellspacing="1"
		class="table" valign="absmiddle">
		<form name="form1" method="post">
		<tr class="hback">
			<td colspan="2" align="Left" class="xingmu">
				<a href="News_label.asp" class="sd" target="_self"><strong><font color="#FF0000">������ǩ</font></strong></a>��<a
					href="All_label_style.asp?label_Sub=NS&TF=NS" target="_self" class="sd"><strong>��ʽ����</strong></a>
			</td>
			<td width="6%" align="Left" class="xingmu">
				<div align="right">
					<input name="button4" type="button" onclick="window.returnValue='';window.close();"
						value="�ر�">
				</div>
			</td>
		</tr>
		<tr class="hback" style="font-family: ����">
			<td align="center" class="hback">
				<div align="right">
					��ʾ��ʽ</div>
			</td>
			<td colspan="2" class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					��ǩ����</div>
			</td>
			<td colspan="2" class="hback">
				<select name="labelStyle" onchange="ChooseNewsType(this.options[this.selectedIndex].value);">
					<option value="" style="background: #DEDEDE">---�б���----------</option>
					<option value="ClassNews" selected>������Ŀ�����б�</option>
					<option value="SpecialNews">����ר�������б�</option>
					<!-- <option value="ReadNews">�������(����ҳ��)</option>-->
					<option value="LastNews">������������</option>
					<option value="HotNews">�����ȵ�����</option>
					<option value="RecNews">�����Ƽ�����</option>
					<!--<option value="FiltNews">�����õ�����</option>-->
					<option value="MarNews">������������</option>
					<!-- <option value="CorrNews">�����������</option>-->
					<!--<option value="DayNews">����ͷ������</option>-->
					<option value="BriNews">������������</option>
					<option value="AnnNews">������������</option>
					<option value="ConstrNews">����Ͷ��</option>
					<option value="" style="background: #DEDEDE">---�ռ���----------</option>
					<option value="ClassList">�����ռ������б�</option>
					<option value="subClassList">�������������б�</option>
					<option value="SpecialList">�����ռ�ר���б�</option>
				</select>
				<span class="tx">������ռ����࣬����ѡ����Ŀ</span>
			</td>
		</tr>
		<tr class="hback" style="display: ">
			<td width="14%" align="center" class="hback">
				<div align="right">
					ͼƬ����</div>
			</td>
			<td colspan="2" class="hback">
				<label>
					<select name="NewsPicTF" id="NewsPicTF">
						<option value="0" selected>��</option>
						<option value="1">��</option>
					</select>
				</label>
			</td>
		</tr>
		<tr class="hback" style="display: " id="liststyle">
			<td width="14%" align="center" class="hback">
				<div align="right">
					�б����</div>
			</td>
			<td colspan="2" class="hback">
				<select name="ListType" id="ListType">
					<option value="0" selected>��</option>
					<option value="1">ABC</option>
					<option value="2">abc</option>
					<option value="3">123</option>
				</select>
				<span class="tx">������������26����ĸ�б��������Ч</span>
			</td>
		</tr>
		<tr class="hback" id="specialEName_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					ר���б�</div>
			</td>
			<td colspan="2" class="hback">
				<select id="specialEName" name="specialEName" disabled>
					<option value="">��ѡ��ר��</option>
					<% = label_special_List %>
				</select>
				<span class="tx">�����ѡ����Ϊ����ר�⵼����ר���б���⡣</span>
			</td>
		</tr>
		<tr class="hback" id="ClassName_col">
			<td class="hback" align="center">
				<div align="right">
					��Ŀ�б�</div>
			</td>
			<td colspan="2" class="hback">
				<input name="ClassName" type="text" id="ClassName" size="60" readonly>
				<input name="ClassID" type="hidden" id="ClassID"><input name="button2" type="button"
					onclick="SelectClass();" value="ѡ��������Ŀ">
				<br />
				<span class="tx">�����ѡ������ʾ���е�����,��Ŀ�����б����</span>
			</td>
		</tr>
		<tr class="hback" id="div_id" style="font-family: ����; display: none;">
			<td rowspan="2" align="center" class="hback">
				<div align="right">
				</div>
				<div align="right">
					DIV����</div>
			</td>
			<td colspan="2" class="hback">
				&lt;div id=&quot;
				<input name="DivID" disabled type="text" id="DivID" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass" type="text" id="Divclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; &lt;ul id=&quot;
				<input name="ulid" disabled type="text" id="ulid" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass" type="text" id="ulclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt;
			</td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family: ����; display: none;">
			<td colspan="2" class="hback">
				&lt;li id=&quot;
				<input name="liid" disabled type="text" id="liid" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass" type="text" id="liclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt;
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					��ʾ��Χ</div>
			</td>
			<td colspan="2" class="hback">
				ֻ��ʾ
				<input name="DateNumber" type="text" id="DateNumber" value="0" size="5">
				���ڵ�����. <span class="tx">���Ϊ0������ʾ����ʱ���ڵ�����</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsNum" align="center">
				<div align="right">
					��������</div>
			</td>
			<td colspan="2" class="hback">
				<input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5">
				<span class="tx">������2���ַ�</span> ͼ�ı�־
				<select name="PicTF" id="PicTF">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsNum" align="center">
				<div align="right">
					��������</div>
			</td>
			<td colspan="2" class="hback">
				<input id="NewsNumber" name="NewsNumber" type="text" size="12" value="10">
				<span class="tx">���õ�ǰ̨��ʾ���� �������ѡ���ռ����Ż����ռ�ר�⣬����Ч</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					������������</div>
			</td>
			<td colspan="2" class="hback">
				<select name="ColsNumber" id="select">
					<option value="1" selected>1</option>
					<option value="2">2</option>
					<option value="3">3</option>
					<option value="4">4</option>
					<option value="5">5</option>
					<option value="6">6</option>
					<option value="7">7</option>
					<option value="8">8</option>
					<option value="9">9</option>
					<option value="10">10</option>
					<option value="11">11</option>
					<option value="12">12</option>
					<option value="13">13</option>
					<option value="14">14</option>
					<option value="15">15</option>
					<option value="16">16</option>
					<option value="17">17</option>
					<option value="18">18</option>
					<option value="19">19</option>
					<option value="20">20</option>
				</select>
				<span class="tx">ÿ����ʾ����,�����DIV+CSS,����CSS���������</span>(���������б�һ��ֻ����ʾһ��)
			</td>
		</tr>
		<tr class="hback" id="c_ColsNum" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					��Ŀ��ʾ����</div>
			</td>
			<td colspan="2" class="hback">
				<select name="sub_colsnum" id="sub_colsnum" disabled>
					<option value="1" selected>1</option>
					<option value="2">2</option>
					<option value="3">3</option>
					<option value="4">4</option>
					<option value="5">5</option>
					<option value="6">6</option>
					<option value="7">7</option>
					<option value="8">8</option>
				</select>
				����ͼƬ
				<input name="bg_pic" type="text" id="bg_pic" disabled>
				<span class="tx">�������������б�</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="c_ColsNum" align="center">
				<div align="right">
					��������</div>
			</td>
			<td colspan="2" class="hback">
				<select name="SubTF" id="SubTF">
					<option value="1">��</option>
					<option value="0" selected>��</option>
				</select>
				<span class="tx">������ռ�������Ч.���ѡ��,�����ٶȻ����Ƚ���,ͬʱ������ռ�÷�������Դ���ڴ�</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="Num" align="center">
				<div align="right">
					�����ֶ�</div>
			</td>
			<td colspan="2" class="hback">
				<select id="OrderType" name="OrderType">
					<option value="ID">�Զ����</option>
					<option value="AddTime">�������</option>
					<option value="PopId" selected>����Ȩ��</option>
					<option value="Hits">�������</option>
				</select>
				����ʽ
				<select id="orderby" name="orderby">
					<option value="ASC">����</option>
					<option value="DESC" selected>����</option>
				</select>
				&nbsp;<span id="last_desc"></span></span><span class="tx"> ������ȵ����ţ���ϵͳĬ���Ե������������</span>
			</td>
		</tr>
		<tr class="hback" id="More_col">
			<td class="hback" id="PageStyle_1" align="center">
				<div align="right">
					��������</div>
			</td>
			<td colspan="2" class="hback">
				<input name="More_char" type="text" id="More_char" value="��" size="16">
				<img src="../Images/upfile.gif" width="44" height="22" onclick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.form1.More_char);"
					style="cursor: hand;">���ΪͼƬ����ֱ������ͼƬ��ַ�� �������Ӵ򿪷�ʽ
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>ԭ����</option>
					<option value="1">�´���</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="PageStyle_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					�Ƿ��ҳ</div>
			</td>
			<td colspan="2" class="hback">
				<input name="PageTF" type="checkbox" disabled id="PageTF" value="1" checked>
				��ҳ��ʽ
				<input id="PageStyle" name="PageStyle" type="text" size="9" value='3,CC0066' disabled>
				<input name="button" type="button" id="SetPage" onclick="OpenPageStyle(this.form.PageStyle)"
					value="����" disabled>
				ÿҳ����
				<input id="PageNumber" name="PageNumber" type="text" size="4" value="30" disabled>
				<span id="page_css" style="display: ">��ҳCSS
					<input id="PageCSS" name="PageCSS" type="text" size="5">
				</span><span class="tx">�������ռ�����</span>
			</td>
		</tr>
		<tr class="hback" id="NewsCut_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					�Ƿ�ָ�</div>
			</td>
			<td colspan="2" class="hback">
				�ָ�����
				<input id="CutNum" name="CutNum" type="text" size="2" value='5' disabled>
				�ָ���ʽ
				<input id="CutType" name="CutType" type="text" size="40" value="" disabled>
				<span class="tx"><>��()�滻</span>
			</td>
		</tr>
		<tr class="hback" id="Mar_cols" style="display: none">
			<td class="hback" id="ScrollSpeed" align="center">
				<div align="right">
					�����ٶ�</div>
			</td>
			<td colspan="2" class="hback">
				<input id="MarqueeSpeed" name="MarqueeSpeed" type="text" size="12" value='8' disabled>
				��������
				<select id="MarqueeDirection" name="MarqueeDirection" disabled>
					<option value="up">����</option>
					<option value="down">����</option>
					<option value="left" selected>����</option>
					<option value="right">����</option>
				</select>
				<span style="display: none">ѡ����ʽ</span>
				<select name="Marqueestyle" id="Marqueestyle" disabled style="display: none">
					<option value="0" selected>����</option>
					<option value="1">����</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsStyle" align="center">
				<div align="right">
					���ڸ�ʽ</div>
			</td>
			<td colspan="2" class="hback">
				<input name="DateType" type="text" id="DateType" value="YY02��MM��DD��" size="20">
				<span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					������ʽ</div>
			</td>
			<td class="hback" colspan="2" align="center">
				<div align="left">
					<select id="NewsStyle" name="NewsStyle" style="width: 40%">
						<% = label_style_List %>
					</select>
					<input name="button3" type="button" id="button" onclick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');"
						value="�鿴">
					<span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ��������ʾ��ʽ</span></div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center" height="30">
				<div align="right">
					����</div>
			</td>
			<td height="30" colspan="2" align="center" class="hback">
				<div align="left">
					������ʾ����ַ�
					<input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
					���ŵ�����ʾ����ַ�
					<input name="navinumber" type="text" id="navinumber" value="200" size="12">
					(����2���ַ�)</div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" colspan="3" align="center" height="30">
				<input name="button" type="button" onclick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button" type="button" onclick="window.top.returnValue='';window.top.close();"
					value=" ȡ �� ">
			</td>
		</tr>
		</form>
	</table>
</body>
<% 
Set Conn=nothing
%>
</html>

<script language="JavaScript" type="text/JavaScript">
	function ChooseNewsType(NewsType) {
		switch (NewsType) {
			case "ClassNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "LastNews":
				document.getElementById('OrderType').value = 'AddTime';
				document.getElementById('last_desc').innerHTML = "������������ţ���ѡ��'<b>�������</b>'��'<b>����</b>'����"
				returnChooseType();
				break;
			case "HotNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "RecNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "DayNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "BriNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "AnnNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "ConstrNews":
				document.getElementById('last_desc').innerHTML = ""
				returnChooseType();
				break;
			case "SpecialNews":
				document.getElementById('ClassName_col').disabled = true;
				document.getElementById('PageStyle_col').disabled = true;
				document.getElementById('SetPage').disabled = true;
				document.getElementById('PageTF').disabled = true;
				document.getElementById('PageStyle').disabled = true;
				document.getElementById('PageNumber').disabled = true;
				document.getElementById('NewsCut_col').disabled = true;
				document.getElementById('CutNum').disabled = true;
				document.getElementById('CutType').disabled = true;
				document.getElementById('NewsCut_col').style.display = 'none';
				document.getElementById('Mar_cols').disabled = true;
				document.getElementById('ClassName_col').style.display = 'none';
				document.getElementById('Mar_cols').style.display = 'none';
				document.getElementById('PageStyle_col').style.display = 'none';
				document.getElementById('More_col').style.display = '';
				document.getElementById('More_char').disabled = false;
				document.getElementById('MarqueeSpeed').disabled = true;
				document.getElementById('MarqueeDirection').disabled = true;
				document.getElementById('ColsNumber').value = '1';
				document.getElementById('specialEName_col').style.display = '';
				document.getElementById('specialEName').disabled = false;
				document.getElementById('ClassID').disabled = true;
				document.getElementById('ColsNumber').disabled = false;
				document.getElementById('sub_colsnum').disabled = true;
				document.getElementById('bg_pic').disabled = true;
				document.getElementById('c_ColsNum').style.display = 'none';
				document.getElementById('liststyle').style.display = '';
				break;
			case "MarNews":
				document.getElementById('ClassName_col').disabled = false;
				document.getElementById('PageStyle_col').disabled = true;
				document.getElementById('SetPage').disabled = true;
				document.getElementById('PageTF').disabled = true;
				document.getElementById('PageStyle').disabled = true;
				document.getElementById('PageNumber').disabled = true;
				document.getElementById('NewsCut_col').disabled = true;
				document.getElementById('CutNum').disabled = true;
				document.getElementById('CutType').disabled = true;
				document.getElementById('NewsCut_col').style.display = 'none';
				document.getElementById('Mar_cols').disabled = false;
				document.getElementById('ClassName_col').style.display = '';
				document.getElementById('Mar_cols').style.display = '';
				document.getElementById('PageStyle_col').style.display = 'none';
				document.getElementById('More_col').style.display = '';
				document.getElementById('More_char').disabled = false;
				document.getElementById('MarqueeSpeed').disabled = false;
				document.getElementById('MarqueeDirection').disabled = false;
				//document.getElementById('Marqueestyle').disabled=false;
				//document.getElementById('ColsNumber').value='10';
				document.getElementById('ColsNumber').disabled = true;
				document.getElementById('specialEName_col').style.display = 'none';
				document.getElementById('specialEName').disabled = true;
				document.getElementById('ClassID').disabled = false;
				document.getElementById('sub_colsnum').disabled = true;
				document.getElementById('bg_pic').disabled = true;
				document.getElementById('c_ColsNum').style.display = 'none';
				document.getElementById('liststyle').style.display = '';
				break;
			case "ClassList":
				document.getElementById('ClassName_col').disabled = true;
				document.getElementById('Mar_cols').disabled = true;
				document.getElementById('PageStyle_col').disabled = false;
				document.getElementById('SetPage').disabled = false;
				document.getElementById('PageTF').disabled = false;
				document.getElementById('PageStyle').disabled = false;
				document.getElementById('PageNumber').disabled = false;
				document.getElementById('NewsCut_col').disabled = false;
				document.getElementById('CutNum').disabled = false;
				document.getElementById('CutType').disabled = false;
				document.getElementById('NewsCut_col').style.display = '';
				document.getElementById('ClassName_col').style.display = 'none';
				document.getElementById('Mar_cols').disabled = '';
				document.getElementById('PageStyle_col').style.display = '';
				document.getElementById('More_col').style.display = 'none';
				document.getElementById('MarqueeSpeed').disabled = true;
				document.getElementById('MarqueeDirection').disabled = true;
				document.getElementById('More_char').disabled = true;
				document.getElementById('ColsNumber').value = '1';
				document.getElementById('ClassID').disabled = true;
				document.getElementById('specialEName_col').style.display = 'none';
				document.getElementById('specialEName').disabled = true;
				document.getElementById('ColsNumber').disabled = false;
				document.getElementById('sub_colsnum').disabled = true;
				document.getElementById('bg_pic').disabled = true;
				document.getElementById('c_ColsNum').style.display = 'none';
				document.getElementById('liststyle').style.display = 'none';
				document.getElementById('ListType').disabled = true;
				break;
			case "subClassList":
				document.getElementById('ClassName_col').disabled = true;
				document.getElementById('Mar_cols').disabled = true;
				document.getElementById('PageStyle_col').disabled = false;
				document.getElementById('SetPage').disabled = false;
				document.getElementById('PageTF').disabled = false;
				document.getElementById('PageStyle').disabled = false;
				document.getElementById('PageNumber').disabled = false;
				document.getElementById('NewsCut_col').disabled = true;
				document.getElementById('CutNum').disabled = true;
				document.getElementById('CutType').disabled = true;
				document.getElementById('NewsCut_col').style.display = 'none';
				document.getElementById('ClassName_col').style.display = 'none';
				document.getElementById('Mar_cols').disabled = '';
				document.getElementById('PageStyle_col').style.display = '';
				document.getElementById('More_col').style.display = 'none';
				document.getElementById('MarqueeSpeed').disabled = true;
				document.getElementById('MarqueeDirection').disabled = true;
				document.getElementById('More_char').disabled = true;
				document.getElementById('ColsNumber').value = '1';
				document.getElementById('ClassID').disabled = true;
				document.getElementById('specialEName_col').style.display = 'none';
				document.getElementById('specialEName').disabled = true;
				document.getElementById('ColsNumber').disabled = false;
				document.getElementById('sub_colsnum').disabled = false;
				document.getElementById('bg_pic').disabled = false;
				document.getElementById('c_ColsNum').style.display = '';
				document.getElementById('liststyle').style.display = 'none';
				document.getElementById('ListType').disabled = true;
				break;
			case "SpecialList":
				document.getElementById('ClassName_col').disabled = true;
				document.getElementById('Mar_cols').disabled = true;
				document.getElementById('PageStyle_col').disabled = false;
				document.getElementById('SetPage').disabled = false;
				document.getElementById('PageTF').disabled = false;
				document.getElementById('PageStyle').disabled = false;
				document.getElementById('PageNumber').disabled = false;
				document.getElementById('NewsCut_col').disabled = false;
				document.getElementById('CutNum').disabled = false;
				document.getElementById('CutType').disabled = false;
				document.getElementById('NewsCut_col').style.display = '';
				document.getElementById('ClassName_col').style.display = 'none';
				document.getElementById('Mar_cols').disabled = '';
				document.getElementById('PageStyle_col').style.display = '';
				document.getElementById('More_col').style.display = 'none';
				document.getElementById('MarqueeSpeed').disabled = true;
				document.getElementById('MarqueeDirection').disabled = true;
				document.getElementById('More_char').disabled = true;
				document.getElementById('ColsNumber').value = '1';
				document.getElementById('ClassID').disabled = true;
				document.getElementById('specialEName_col').style.display = 'none';
				document.getElementById('specialEName').disabled = true;
				document.getElementById('ColsNumber').disabled = false;
				document.getElementById('sub_colsnum').disabled = true;
				document.getElementById('bg_pic').disabled = true;
				document.getElementById('c_ColsNum').style.display = 'none';
				document.getElementById('liststyle').style.display = 'none';
				document.getElementById('ListType').disabled = true;
				break;
		}
	}
	function returnChooseType() {
		document.getElementById('ClassName_col').disabled = false;
		document.getElementById('PageStyle_col').disabled = true;
		document.getElementById('SetPage').disabled = true;
		document.getElementById('PageTF').disabled = true;
		document.getElementById('PageStyle').disabled = true;
		document.getElementById('PageNumber').disabled = true;
		document.getElementById('NewsCut_col').disabled = true;
		document.getElementById('CutNum').disabled = true;
		document.getElementById('CutType').disabled = true;
		document.getElementById('NewsCut_col').style.display = 'none';
		document.getElementById('Mar_cols').disabled = true;
		document.getElementById('ClassName_col').style.display = '';
		document.getElementById('Mar_cols').style.display = 'none';
		document.getElementById('PageStyle_col').style.display = 'none';
		document.getElementById('More_col').style.display = '';
		document.getElementById('More_char').disabled = false;
		document.getElementById('ClassID').disabled = false;
		document.getElementById('ColsNumber').value = '1';
		document.getElementById('MarqueeSpeed').disabled = true;
		document.getElementById('MarqueeDirection').disabled = true;
		document.getElementById('specialEName_col').style.display = 'none';
		document.getElementById('specialEName').disabled = true;
		document.getElementById('ColsNumber').disabled = false;
		document.getElementById('sub_colsnum').disabled = true;
		document.getElementById('bg_pic').disabled = true;
		document.getElementById('c_ColsNum').style.display = 'none';
		document.getElementById('liststyle').style.display = '';
		document.getElementById('ListType').disabled = false;
	}
	function selectHtml_express(Html_express) {
		switch (Html_express) {
			case "out_Table":
				document.getElementById('div_id').style.display = 'none';
				document.getElementById('ul_id').style.display = 'none';
				document.getElementById('DivID').disabled = true;
				document.getElementById('Divclass').disabled = true;
				document.getElementById('ulid').disabled = true;
				document.getElementById('ulclass').disabled = true;
				document.getElementById('liid').disabled = true;
				document.getElementById('liclass').disabled = true;
				document.getElementById('page_css').style.display = '';
				break;
			case "out_DIV":
				document.getElementById('div_id').style.display = '';
				document.getElementById('ul_id').style.display = '';
				document.getElementById('page_css').style.display = 'none';
				document.getElementById('DivID').disabled = false;
				document.getElementById('Divclass').disabled = false;
				document.getElementById('ulid').disabled = false;
				document.getElementById('ulclass').disabled = false;
				document.getElementById('liid').disabled = false;
				document.getElementById('liclass').disabled = false;
				break;
		}
	}
	function ShowTempLayer(oDiv) {
		var obj = document.getElementById(oDiv);
		obj.style.display = '';
		obj.style.width = 280;
		obj.style.height = 120;
		obj.style.top = parseInt(document.body.scrollTop + event.clientY) - 100;
		obj.style.left = parseInt(event.clientX) - 280;
	}
	function SelectClass() {
		var ReturnValue = '', TempArray = new Array();
		ReturnValue = OpenWindow('../News/lib/SelectClassFrame.asp?mulit=true', 400, 300, window);
		try {
			document.getElementById('ClassID').value = ReturnValue[0][0];
			document.getElementById('ClassName').value = ReturnValue[1][0];
			if (ReturnValue[0].length > 1) {
				document.getElementById('SubTF').selectedIndex = 1;
				document.getElementById('SubTF').disabled = true;
			} else {
				document.getElementById('SubTF').disabled = false;
			}
		}
		catch (ex) { }
	}
	function OpenPageStyle(obj) {
		var ret = OpenWindow("setPage.asp", 200, 200, "page");
		if (ret != '') obj.value = ret;
	}
	function GetObj(s) {
		return document.getElementById(s);
	}
	function IsDisabled(oName) {
		if (GetObj(oName).disabled == true) {
			return true;
		} else {
			return false;
		}
	}
	var G_CanSave = true;
	var G_Save_Msg = '';
	function ok(obj) {
		//if(obj.labelName.value == ''){alert('�������ǩ����');obj.labelName.focus();return false;}
		if (obj.NewsStyle.value == '') { alert('����д������ʽ!!\n�����û������������ʽ�����н���'); obj.NewsStyle.focus(); return false; }
		if (obj.labelStyle.value == 'ClassNews') {
			if (obj.ClassName.value == '') { alert('��ѡ����Ŀ'); obj.ClassName.focus(); return false; }
		}
		//	if(obj.labelStyle.value=='SpecialNews')
		//		{
		//			if(obj.specialEName.value==''){alert('��ѡ��ר��');obj.specialEName.focus();return false;}
		//		}
		if (obj.labelStyle.value == 'MarNews') {
			if (isNaN(obj.MarqueeSpeed.value) == true) { alert('�����ٶ�����д����'); obj.MarqueeSpeed.focus(); return false; }
		}
		if (obj.contentnumber.value == '') { alert('����д������ʾ�ַ���\n����ռ2���ַ�'); obj.contentnumber.focus(); return false; }
		if (obj.navinumber.value == '') { alert('����д���ŵ�����ʾ�ַ���\n����ռ2���ַ�'); obj.navinumber.focus(); return false; }
		if (isNaN(obj.NewsNumber.value) == true) { alert('������������Ϊ����'); obj.NewsNumber.focus(); return false; }
		if (isNaN(obj.contentnumber.value) == true) { alert('������ʾ������������'); obj.contentnumber.focus(); return false; }
		if (isNaN(obj.navinumber.value) == true) { alert('������ʾ������������'); obj.navinumber.focus(); return false; }
		if (obj.PageStyle.value == '') obj.PageStyle.value = ',CC0066';
		var retV = '{FS:NS=';
		retV += obj.labelStyle.value + '��';
		if (!IsDisabled('NewsPicTF')) retV += 'ͼƬ����$' + obj.NewsPicTF.value + '��';
		if (!IsDisabled('specialEName')) retV += 'ר��$' + obj.specialEName.value + '��';
		if (!IsDisabled('ClassID')) retV += '��Ŀ$' + obj.ClassID.value + '��';
		if (obj.labelStyle.value != 'ClassList' && obj.labelStyle.value != 'SpecialList') {
			if (!IsDisabled('NewsNumber')) retV += 'Loop$' + obj.NewsNumber.value + '��';
		}
		if (!IsDisabled('out_char')) retV += '�����ʽ$' + obj.out_char.value + '��';
		if (!IsDisabled('DivID')) retV += 'DivID$' + obj.DivID.value + '��';
		if (!IsDisabled('Divclass')) retV += 'DivClass$' + obj.Divclass.value + '��';
		if (!IsDisabled('ulid')) retV += 'ulid$' + obj.ulid.value + '��';
		if (!IsDisabled('ulclass')) retV += 'ulclass$' + obj.ulclass.value + '��';
		if (!IsDisabled('liid')) retV += 'liid$' + obj.liid.value + '��';
		if (!IsDisabled('liclass')) retV += 'liclass$' + obj.liclass.value + '��';
		if (!IsDisabled('DateNumber')) retV += '������$' + obj.DateNumber.value + '��';
		if (!IsDisabled('TitleNumber')) retV += '������$' + obj.TitleNumber.value + '��';
		if (!IsDisabled('PicTF')) retV += 'ͼ�ı�־$' + obj.PicTF.value + '��';
		if (!IsDisabled('Openstyle')) retV += '�򿪴���$' + obj.Openstyle.value + '��';
		if (!IsDisabled('SubTF')) retV += '��������$' + obj.SubTF.value + '��';
		if (!IsDisabled('OrderType')) retV += '�����ֶ�$' + obj.OrderType.value + '��';
		if (!IsDisabled('orderby')) retV += '���з�ʽ$' + obj.orderby.value + '��';
		if (!IsDisabled('More_char')) retV += '��������$' + obj.More_char.value + '��';
		if (!IsDisabled('PageTF')) {
			if (obj.PageTF.checked) {
				retV += '��ҳ$' + obj.PageTF.value + '��';
			} else {
				retV += '��ҳ$0��';
			}
		}
		if (!IsDisabled('PageStyle')) retV += '��ҳ��ʽ$' + obj.PageStyle.value + '��';
		if (!IsDisabled('PageNumber')) retV += 'ÿҳ����$' + obj.PageNumber.value + '��';
		if (!IsDisabled('PageCSS')) retV += 'PageCSS$' + obj.PageCSS.value + '��';
		if (!IsDisabled('CutNum')) retV += '�ָ�����$' + obj.CutNum.value + '��';
		if (!IsDisabled('CutType')) retV += '�ָ���ʽ$' + obj.CutType.value + '��';
		if (!IsDisabled('DateType')) retV += '���ڸ�ʽ$' + obj.DateType.value + '��';
		if (!IsDisabled('NewsStyle')) retV += '������ʽ$' + obj.NewsStyle.value + '��';
		if (!IsDisabled('bg_pic')) retV += '��������$' + obj.bg_pic.value + '��';
		if (!IsDisabled('sub_colsnum')) retV += '��Ŀ������$' + obj.sub_colsnum.value + '��';
		retV += '����������$' + obj.ColsNumber.value + '��';
		retV += '��������$' + obj.contentnumber.value + '��';
		if (obj.labelStyle.value == 'MarNews')
		{ retV += '��������$' + obj.navinumber.value; if (!IsDisabled('ListType')) retV += '�� �б����$' + obj.ListType.value + '��'; }
		else
		{ retV += '��������$' + obj.navinumber.value; if (!IsDisabled('ListType')) retV += '�� �б����$' + obj.ListType.value + ''; }
		if (!IsDisabled('MarqueeSpeed')) retV += '�����ٶ�$' + obj.MarqueeSpeed.value + '��';
		if (!IsDisabled('MarqueeDirection')) retV += '��������$' + obj.MarqueeDirection.value + '';
		//if(!IsDisabled('ListType')) retV+='�б���ʽ$' + obj.ListType.value +'';
		retV += '}';
		window.parent.returnValue = retV;
		window.close();
	}
</script>

