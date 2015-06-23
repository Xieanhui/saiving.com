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
	<title>新闻标签管理</title>
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
				<a href="News_label.asp" class="sd" target="_self"><strong><font color="#FF0000">创建标签</font></strong></a>｜<a
					href="All_label_style.asp?label_Sub=NS&TF=NS" target="_self" class="sd"><strong>样式管理</strong></a>
			</td>
			<td width="6%" align="Left" class="xingmu">
				<div align="right">
					<input name="button4" type="button" onclick="window.returnValue='';window.close();"
						value="关闭">
				</div>
			</td>
		</tr>
		<tr class="hback" style="font-family: 宋体">
			<td align="center" class="hback">
				<div align="right">
					显示格式</div>
			</td>
			<td colspan="2" class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					标签类型</div>
			</td>
			<td colspan="2" class="hback">
				<select name="labelStyle" onchange="ChooseNewsType(this.options[this.selectedIndex].value);">
					<option value="" style="background: #DEDEDE">---列表类----------</option>
					<option value="ClassNews" selected>├┄栏目新闻列表</option>
					<option value="SpecialNews">├┄专题新闻列表</option>
					<!-- <option value="ReadNews">新闻浏览(新闻页面)</option>-->
					<option value="LastNews">├┄最新新闻</option>
					<option value="HotNews">├┄热点新闻</option>
					<option value="RecNews">├┄推荐新闻</option>
					<!--<option value="FiltNews">├┄幻灯新闻</option>-->
					<option value="MarNews">├┄滚动新闻</option>
					<!-- <option value="CorrNews">├┄相关新闻</option>-->
					<!--<option value="DayNews">├┄头条新闻</option>-->
					<option value="BriNews">├┄精彩新闻</option>
					<option value="AnnNews">├┄公告新闻</option>
					<option value="ConstrNews">├┄投稿</option>
					<option value="" style="background: #DEDEDE">---终极类----------</option>
					<option value="ClassList">├┄终极新闻列表</option>
					<option value="subClassList">├┄子类新闻列表</option>
					<option value="SpecialList">├┄终极专题列表</option>
				</select>
				<span class="tx">如果是终极分类，不能选择栏目</span>
			</td>
		</tr>
		<tr class="hback" style="display: ">
			<td width="14%" align="center" class="hback">
				<div align="right">
					图片新闻</div>
			</td>
			<td colspan="2" class="hback">
				<label>
					<select name="NewsPicTF" id="NewsPicTF">
						<option value="0" selected>否</option>
						<option value="1">是</option>
					</select>
				</label>
			</td>
		</tr>
		<tr class="hback" style="display: " id="liststyle">
			<td width="14%" align="center" class="hback">
				<div align="right">
					列表序号</div>
			</td>
			<td colspan="2" class="hback">
				<select name="ListType" id="ListType">
					<option value="0" selected>无</option>
					<option value="1">ABC</option>
					<option value="2">abc</option>
					<option value="3">123</option>
				</select>
				<span class="tx">调用数量大于26，字母列表序号则无效</span>
			</td>
		</tr>
		<tr class="hback" id="specialEName_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					专题列表</div>
			</td>
			<td colspan="2" class="hback">
				<select id="specialEName" name="specialEName" disabled>
					<option value="">请选择专题</option>
					<% = label_special_List %>
				</select>
				<span class="tx">如果不选择，则为所有专题导航，专题列表除外。</span>
			</td>
		</tr>
		<tr class="hback" id="ClassName_col">
			<td class="hback" align="center">
				<div align="right">
					栏目列表</div>
			</td>
			<td colspan="2" class="hback">
				<input name="ClassName" type="text" id="ClassName" size="60" readonly>
				<input name="ClassID" type="hidden" id="ClassID"><input name="button2" type="button"
					onclick="SelectClass();" value="选择新闻栏目">
				<br />
				<span class="tx">如果不选择，则显示所有的新闻,栏目新闻列表除外</span>
			</td>
		</tr>
		<tr class="hback" id="div_id" style="font-family: 宋体; display: none;">
			<td rowspan="2" align="center" class="hback">
				<div align="right">
				</div>
				<div align="right">
					DIV控制</div>
			</td>
			<td colspan="2" class="hback">
				&lt;div id=&quot;
				<input name="DivID" disabled type="text" id="DivID" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass" type="text" id="Divclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; &lt;ul id=&quot;
				<input name="ulid" disabled type="text" id="ulid" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass" type="text" id="ulclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt;
			</td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family: 宋体; display: none;">
			<td colspan="2" class="hback">
				&lt;li id=&quot;
				<input name="liid" disabled type="text" id="liid" size="6" style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass" type="text" id="liclass" size="6" disabled style="border-top-width: 0px;
					border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; border-bottom-color: #000000"
					title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt;
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					显示范围</div>
			</td>
			<td colspan="2" class="hback">
				只显示
				<input name="DateNumber" type="text" id="DateNumber" value="0" size="5">
				天内的新闻. <span class="tx">如果为0，则显示所有时间内的新闻</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsNum" align="center">
				<div align="right">
					标题字数</div>
			</td>
			<td colspan="2" class="hback">
				<input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5">
				<span class="tx">中文算2个字符</span> 图文标志
				<select name="PicTF" id="PicTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsNum" align="center">
				<div align="right">
					调用数量</div>
			</td>
			<td colspan="2" class="hback">
				<input id="NewsNumber" name="NewsNumber" type="text" size="12" value="10">
				<span class="tx">调用的前台显示数量 此项如果选择终极新闻或者终极专题，则无效</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					新闻排列列数</div>
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
				<span class="tx">每行显示数量,如果是DIV+CSS,请在CSS里控制条数</span>(子类新闻列表一行只能显示一个)
			</td>
		</tr>
		<tr class="hback" id="c_ColsNum" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					栏目显示列数</div>
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
				背景图片
				<input name="bg_pic" type="text" id="bg_pic" disabled>
				<span class="tx">仅适用于子类列表</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="c_ColsNum" align="center">
				<div align="right">
					包含子类</div>
			</td>
			<td colspan="2" class="hback">
				<select name="SubTF" id="SubTF">
					<option value="1">是</option>
					<option value="0" selected>否</option>
				</select>
				<span class="tx">此项对终极分类无效.如果选择,生成速度会大幅度降低,同时将大量占用服务器资源和内存</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="Num" align="center">
				<div align="right">
					排序字段</div>
			</td>
			<td colspan="2" class="hback">
				<select id="OrderType" name="OrderType">
					<option value="ID">自动编号</option>
					<option value="AddTime">添加日期</option>
					<option value="PopId" selected>新闻权重</option>
					<option value="Hits">点击次数</option>
				</select>
				排序方式
				<select id="orderby" name="orderby">
					<option value="ASC">升序</option>
					<option value="DESC" selected>降序</option>
				</select>
				&nbsp;<span id="last_desc"></span></span><span class="tx"> 如果是热点新闻，则系统默认以点击数降序排列</span>
			</td>
		</tr>
		<tr class="hback" id="More_col">
			<td class="hback" id="PageStyle_1" align="center">
				<div align="right">
					更多连接</div>
			</td>
			<td colspan="2" class="hback">
				<input name="More_char" type="text" id="More_char" value="·" size="16">
				<img src="../Images/upfile.gif" width="44" height="22" onclick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.form1.More_char);"
					style="cursor: hand;">如果为图片，请直接输入图片地址。 更多连接打开方式
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>原窗口</option>
					<option value="1">新窗口</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="PageStyle_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					是否分页</div>
			</td>
			<td colspan="2" class="hback">
				<input name="PageTF" type="checkbox" disabled id="PageTF" value="1" checked>
				分页样式
				<input id="PageStyle" name="PageStyle" type="text" size="9" value='3,CC0066' disabled>
				<input name="button" type="button" id="SetPage" onclick="OpenPageStyle(this.form.PageStyle)"
					value="设置" disabled>
				每页数量
				<input id="PageNumber" name="PageNumber" type="text" size="4" value="30" disabled>
				<span id="page_css" style="display: ">分页CSS
					<input id="PageCSS" name="PageCSS" type="text" size="5">
				</span><span class="tx">仅适用终极分类</span>
			</td>
		</tr>
		<tr class="hback" id="NewsCut_col" style="display: none">
			<td class="hback" align="center">
				<div align="right">
					是否分隔</div>
			</td>
			<td colspan="2" class="hback">
				分隔数量
				<input id="CutNum" name="CutNum" type="text" size="2" value='5' disabled>
				分隔样式
				<input id="CutType" name="CutType" type="text" size="40" value="" disabled>
				<span class="tx"><>用()替换</span>
			</td>
		</tr>
		<tr class="hback" id="Mar_cols" style="display: none">
			<td class="hback" id="ScrollSpeed" align="center">
				<div align="right">
					滚动速度</div>
			</td>
			<td colspan="2" class="hback">
				<input id="MarqueeSpeed" name="MarqueeSpeed" type="text" size="12" value='8' disabled>
				滚动方向
				<select id="MarqueeDirection" name="MarqueeDirection" disabled>
					<option value="up">向上</option>
					<option value="down">向下</option>
					<option value="left" selected>向左</option>
					<option value="right">向右</option>
				</select>
				<span style="display: none">选择样式</span>
				<select name="Marqueestyle" id="Marqueestyle" disabled style="display: none">
					<option value="0" selected>滚动</option>
					<option value="1">闪动</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" id="NewsStyle" align="center">
				<div align="right">
					日期格式</div>
			</td>
			<td colspan="2" class="hback">
				<input name="DateType" type="text" id="DateType" value="YY02年MM月DD日" size="20">
				<span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center">
				<div align="right">
					引用样式</div>
			</td>
			<td class="hback" colspan="2" align="center">
				<div align="left">
					<select id="NewsStyle" name="NewsStyle" style="width: 40%">
						<% = label_style_List %>
					</select>
					<input name="button3" type="button" id="button" onclick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');"
						value="查看">
					<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span></div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" align="center" height="30">
				<div align="right">
					其他</div>
			</td>
			<td height="30" colspan="2" align="center" class="hback">
				<div align="left">
					内容显示最多字符
					<input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
					新闻导读显示最多字符
					<input name="navinumber" type="text" id="navinumber" value="200" size="12">
					(中文2个字符)</div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback" colspan="3" align="center" height="30">
				<input name="button" type="button" onclick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onclick="window.top.returnValue='';window.top.close();"
					value=" 取 消 ">
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
				document.getElementById('last_desc').innerHTML = "如果是最新新闻，请选择'<b>添加日期</b>'和'<b>降序</b>'排列"
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
		//if(obj.labelName.value == ''){alert('请输入标签名称');obj.labelName.focus();return false;}
		if (obj.NewsStyle.value == '') { alert('请填写引用样式!!\n如果还没建立，请在样式管理中建立'); obj.NewsStyle.focus(); return false; }
		if (obj.labelStyle.value == 'ClassNews') {
			if (obj.ClassName.value == '') { alert('请选择栏目'); obj.ClassName.focus(); return false; }
		}
		//	if(obj.labelStyle.value=='SpecialNews')
		//		{
		//			if(obj.specialEName.value==''){alert('请选择专题');obj.specialEName.focus();return false;}
		//		}
		if (obj.labelStyle.value == 'MarNews') {
			if (isNaN(obj.MarqueeSpeed.value) == true) { alert('滚动速度请填写数字'); obj.MarqueeSpeed.focus(); return false; }
		}
		if (obj.contentnumber.value == '') { alert('请填写内容显示字符数\n中文占2个字符'); obj.contentnumber.focus(); return false; }
		if (obj.navinumber.value == '') { alert('请填写新闻导航显示字符数\n中文占2个字符'); obj.navinumber.focus(); return false; }
		if (isNaN(obj.NewsNumber.value) == true) { alert('调用数量必须为数字'); obj.NewsNumber.focus(); return false; }
		if (isNaN(obj.contentnumber.value) == true) { alert('内容显示字数请用数字'); obj.contentnumber.focus(); return false; }
		if (isNaN(obj.navinumber.value) == true) { alert('导航显示字数请用数字'); obj.navinumber.focus(); return false; }
		if (obj.PageStyle.value == '') obj.PageStyle.value = ',CC0066';
		var retV = '{FS:NS=';
		retV += obj.labelStyle.value + '┆';
		if (!IsDisabled('NewsPicTF')) retV += '图片新闻$' + obj.NewsPicTF.value + '┆';
		if (!IsDisabled('specialEName')) retV += '专题$' + obj.specialEName.value + '┆';
		if (!IsDisabled('ClassID')) retV += '栏目$' + obj.ClassID.value + '┆';
		if (obj.labelStyle.value != 'ClassList' && obj.labelStyle.value != 'SpecialList') {
			if (!IsDisabled('NewsNumber')) retV += 'Loop$' + obj.NewsNumber.value + '┆';
		}
		if (!IsDisabled('out_char')) retV += '输出格式$' + obj.out_char.value + '┆';
		if (!IsDisabled('DivID')) retV += 'DivID$' + obj.DivID.value + '┆';
		if (!IsDisabled('Divclass')) retV += 'DivClass$' + obj.Divclass.value + '┆';
		if (!IsDisabled('ulid')) retV += 'ulid$' + obj.ulid.value + '┆';
		if (!IsDisabled('ulclass')) retV += 'ulclass$' + obj.ulclass.value + '┆';
		if (!IsDisabled('liid')) retV += 'liid$' + obj.liid.value + '┆';
		if (!IsDisabled('liclass')) retV += 'liclass$' + obj.liclass.value + '┆';
		if (!IsDisabled('DateNumber')) retV += '多少天$' + obj.DateNumber.value + '┆';
		if (!IsDisabled('TitleNumber')) retV += '标题数$' + obj.TitleNumber.value + '┆';
		if (!IsDisabled('PicTF')) retV += '图文标志$' + obj.PicTF.value + '┆';
		if (!IsDisabled('Openstyle')) retV += '打开窗口$' + obj.Openstyle.value + '┆';
		if (!IsDisabled('SubTF')) retV += '包含子类$' + obj.SubTF.value + '┆';
		if (!IsDisabled('OrderType')) retV += '排列字段$' + obj.OrderType.value + '┆';
		if (!IsDisabled('orderby')) retV += '排列方式$' + obj.orderby.value + '┆';
		if (!IsDisabled('More_char')) retV += '更多连接$' + obj.More_char.value + '┆';
		if (!IsDisabled('PageTF')) {
			if (obj.PageTF.checked) {
				retV += '分页$' + obj.PageTF.value + '┆';
			} else {
				retV += '分页$0┆';
			}
		}
		if (!IsDisabled('PageStyle')) retV += '分页样式$' + obj.PageStyle.value + '┆';
		if (!IsDisabled('PageNumber')) retV += '每页数量$' + obj.PageNumber.value + '┆';
		if (!IsDisabled('PageCSS')) retV += 'PageCSS$' + obj.PageCSS.value + '┆';
		if (!IsDisabled('CutNum')) retV += '分隔数量$' + obj.CutNum.value + '┆';
		if (!IsDisabled('CutType')) retV += '分隔样式$' + obj.CutType.value + '┆';
		if (!IsDisabled('DateType')) retV += '日期格式$' + obj.DateType.value + '┆';
		if (!IsDisabled('NewsStyle')) retV += '引用样式$' + obj.NewsStyle.value + '┆';
		if (!IsDisabled('bg_pic')) retV += '背景底纹$' + obj.bg_pic.value + '┆';
		if (!IsDisabled('sub_colsnum')) retV += '栏目排列数$' + obj.sub_colsnum.value + '┆';
		retV += '新闻排列数$' + obj.ColsNumber.value + '┆';
		retV += '内容字数$' + obj.contentnumber.value + '┆';
		if (obj.labelStyle.value == 'MarNews')
		{ retV += '导航字数$' + obj.navinumber.value; if (!IsDisabled('ListType')) retV += '┆ 列表序号$' + obj.ListType.value + '┆'; }
		else
		{ retV += '导航字数$' + obj.navinumber.value; if (!IsDisabled('ListType')) retV += '┆ 列表序号$' + obj.ListType.value + ''; }
		if (!IsDisabled('MarqueeSpeed')) retV += '滚动速度$' + obj.MarqueeSpeed.value + '┆';
		if (!IsDisabled('MarqueeDirection')) retV += '滚动方向$' + obj.MarqueeDirection.value + '';
		//if(!IsDisabled('ListType')) retV+='列表样式$' + obj.ListType.value +'';
		retV += '}';
		window.parent.returnValue = retV;
		window.close();
	}
</script>

