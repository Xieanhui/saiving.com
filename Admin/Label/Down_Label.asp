<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/DS_Function.asp" -->
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
	Dim obj_Lable_style_Rs,Lable_style_List
	Lable_style_List=""
	Set  obj_Lable_style_Rs = server.CreateObject(G_FS_RS)
	obj_Lable_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='DS' Order by  id desc",Conn,1,3
	do while Not obj_Lable_style_Rs.eof 
		Lable_style_List = Lable_style_List&"<option value="""& obj_Lable_style_Rs(0)&""">"& obj_Lable_style_Rs(1)&"</option>"
		obj_Lable_style_Rs.movenext
	loop
	obj_Lable_style_Rs.close:set obj_Lable_style_Rs = nothing
	
	'---------------------------------专区列表
	Dim obj_special_Rs,label_special_List
	label_special_List=""
	Set  obj_special_Rs = server.CreateObject(G_FS_RS)
	obj_special_Rs.Open "Select SpecialID,SpecialCName from FS_DS_Special  Order by  SpecialID desc",Conn,1,3
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
<title>下载标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" >
      <td colspan="3"  align="Left" class="xingmu"><strong><font color="#FF0000">创建标签</font></strong></td>
      <td width="36%"  align="Left" class="xingmu"><div align="right">
          <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
        </div></td>
    </tr>
    <tr class="hback"  style="font-family:宋体" >
      <td  align="center" class="hback" ><div align="right">显示格式</div></td>
      <td colspan="3" class="hback" ><select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  align="center"><div align="right">标签类型</div></td>
      <td colspan="3" class="hback" ><select  name="labelStyle" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
          <option value="" style="background:#DEDEDE">---列表类----------</option>
          <option value="ClassNews" selected>├┄栏目列表</option>
          <option value="SpecialDown">├┄专区列表</option>
          <option value="LastNews">├┄最新</option>
          <option value="HotNews">├┄点击最多</option>
          <option value="DownHotNews">├┄下载最多</option>
          <option value="DownPicNews">├┄图片</option>
          <option value="RecNews">├┄推荐</option>
          <option value="" style="background:#DEDEDE">---终极类----------</option>
          <option value="ClassList">├┄终极下载列表</option>
          <option value="subClassList">├┄子类下载列表</option>
          <option value="SpecialListDown">├┄终极专区列表</option>
        </select>
        <span class="tx">如果是终极分类，不能选择栏目</span> </td>
    </tr>
    <tr class="hback" style="display:none">
      <td width="14%"  align="center" class="hback"><div align="right">标签名称</div></td>
      <td colspan="3" class="hback" ><input name="labelName"  type="text" size="12" maxlength="25">
        <span class="tx">方便以后查找标签,限制25个字符(只能为中文、数字、英文、下划线和中划线)。</span></td>
    </tr>
    <tr class="hback" id="specialEName_col" style="display:none">
      <td class="hback"  align="center"><div align="right">专区列表</div></td>
      <td colspan="2" class="hback" >
	  	 <select id="specialEName"  name="specialEName" disabled>
          <option value="">请选择专区</option>
          <% = label_special_List %>
        </select>
        <span class="tx">如果不选择，则为所有专区导航</span> </td>
      <td>&nbsp;</td>
    </tr>
    <tr class="hback" id="ClassName_col">
      <td class="hback"  align="center"><div align="right">栏目列表</div></td>
      <td colspan="3" class="hback" ><input  name="ClassName" type="text" id="ClassName" size="12" readonly>
        <input name="ClassID" type="hidden" id="ClassID">
        <input name="button2" type="button" onClick="SelectClass();" value="选择栏目">
        <span class="tx">选择下载栏目，如果不选择，则显示所有的下载</span></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV控制</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot;
        <input name="DivID" disabled type="text" id="DivID" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
        &quot; class=&quot;
        <input name="Divclass"  type="text" id="Divclass" size="6"  disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
        &quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
      <td colspan="3" class="hback" >&lt;ul id=&quot;
        <input name="ulid"  disabled type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
        &quot; class=&quot;
        <input name="ulclass"  type="text" id="ulclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
        &quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid" disabled type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
        &quot; class=&quot;
        <input name="liclass"  type="text" id="liclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
        &quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  align="center"><div align="right">显示范围</div></td>
      <td colspan="3" class="hback" >只显示
        <input name="DateNumber"  type="text" id="DateNumber" value="0" size="5">
        天内的下载. <span class="tx">如果为0，则显示所有时间内的下载</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  id="NewsNum" align="center"><div align="right">标题字数</div></td>
      <td colspan="3" class="hback" ><input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5">
        内容字数
        <input name="ContentNumber" type="text" id="ContentNumber" value="100" size="5">
        <span class="tx">中文算2个字符</span> 　图文标志
        <select name="PicTF" id="PicTF">
          <option value="1" selected>显示</option>
          <option value="0">不显示</option>
        </select></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  id="NewsNum" align="center"><div align="right">调用数量</div></td>
      <td width="35%" class="hback" ><input id="NewsNumber"  name="NewsNumber" type="text" size="12" value="10">
        <span class="tx">调用的前台显示数量　　</span> 　　　　 </td>
      <td colspan="2" class="hback" >排列列数
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
      <span class="tx">下载一行显示数量,对DIV输出格式无效</span></td>
    </tr>
    <tr class="hback"   id="c_ColsNum" style="display:none">
      <td class="hback" align="center"><div align="right">栏目显示列数</div></td>
      <td colspan="4" class="hback"><select name="sub_colsnum" id="sub_colsnum" disabled>
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
        <span class="tx">仅适用于子类列表</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  id="ColsNum" align="center"><div align="right">包含子类</div></td>
      <td colspan="3" class="hback" ><select name="SubTF" id="SubTF">
          <option value="1">是</option>
          <option value="0" selected>否</option>
        </select>
        <span class="tx">建议不要选择，否则下载生成速度会大幅度降低,专区此项无效</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  id="Num" align="center"><div align="right">排序字段</div></td>
      <td class="hback" ><input type="text" name="OrderType" id="OrderType" readonly="" title="此为自动根据标签类型自动产生" size="10" value="ID">
        <span class="tx">自动产生</span></td>
      <td width="15%" align="center" class="hback"  id="NewsType">排序方式</td>
      <td class="hback"><input type="radio" name="orderby" value="ASC" id="ASC">
        升序
        <input type="radio"  name="orderby" checked value="DESC" id="DESC">
        降序 </td>
    </tr>
    <tr class="hback" id="More_col">
      <td class="hback"  id="PageStyle_1" align="center"><div align="right">更多连接</div></td>
      <td colspan="3" class="hback" ><input  name="More_char" type="text" id="More_char" value="·" size="16">
        <img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.form1.More_char);" style="cursor:hand;">如果为图片，请直接输入图片地址。<span class="tx">更多连接打开方式
        <select name="Openstyle" id="Openstyle">
          <option value="0" selected>原窗口</option>
          <option value="1">新窗口</option>
        </select>
        </span></td>
    </tr>
    <tr class="hback" id="PageStyle_col" style="display:none">
      <td class="hback"   align="center"><div align="right">是否分页</div></td>
      <td colspan="3" class="hback" ><input name="PageTF" type="checkbox" id="PageTF" value="1" disabled>
        分页样式
        <input id="PageStyle"  name="PageStyle" type="text" size="9" value='3,CC0066'  disabled>
        <input name="button" type="button" id="SetPage" onClick="OpenPageStyle(this.form.PageStyle)" value="设置" disabled>
        每页数量
        <input id="PageNumber"  name="PageNumber" type="text" size="4" value="30" disabled>
        <span id="page_css" style="display:">分页CSS
        <input id="PageCSS" name="PageCSS" type="text" size="5">
        </span> <span class="tx">仅适用终极分类</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  id="NewsStyle" align="center"><div align="right">日期格式</div></td>
      <td colspan="3" class="hback" ><input name="DateType" type="text" id="DateType" value="YY02-MM-DD" size="15">
        <span class="tx">格式:YY(02)/YY(04)代表年，MM-月，DD-日，HH-小时，MI-分，SS-秒</span></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="3" align="center"><div align="left">
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = Lable_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面下载显示样式</span></div></td>
    </tr>
    <tr class="hback" >
      <td class="hback"  colspan="4" align="center" height="30"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">      </td>
    </tr>
  </form>
</table>
</body>
<% 
Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function ChooseNewsType(NewsType)
{
	switch (NewsType)
	{
		case "ClassNews":
			document.getElementById('OrderType').value='ID';
			returnChooseType();
			break;
		case "LastNews":
			document.getElementById('OrderType').value='AddTime';
			document.getElementById('ASC').checked=false;
			document.getElementById('DESC').checked=true;
			document.getElementById('ASC').readonly=true;
			document.getElementById('DESC').readonly=true;
			returnChooseType();
			break;
		case "HotNews":
			document.getElementById('OrderType').value='Hits';
			document.getElementById('ASC').checked=false;
			document.getElementById('DESC').checked=true;
			document.getElementById('ASC').readonly=true;
			document.getElementById('DESC').readonly=true;
			returnChooseType();
			break;
		case "DownHotNews":
			document.getElementById('OrderType').value='ClickNum';
			document.getElementById('ASC').checked=false;
			document.getElementById('DESC').checked=true;
			document.getElementById('ASC').readonly=true;
			document.getElementById('DESC').readonly=true;
			returnChooseType();
			break;
		case "DownPicNews":
			document.getElementById('OrderType').value='Pic';
			document.getElementById('ASC').checked=false;
			document.getElementById('DESC').checked=true;
			document.getElementById('ASC').readonly=true;
			document.getElementById('DESC').readonly=true;
			returnChooseType();
		case "RecNews":
			document.getElementById('OrderType').value='RecTF';
			document.getElementById('ASC').checked=false;
			document.getElementById('DESC').checked=true;
			document.getElementById('ASC').readonly=true;
			document.getElementById('DESC').readonly=true;
			returnChooseType();
			break;
		case "SpecialDown":
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('specialEName_col').style.display='';
			document.getElementById('specialEName').disabled=false;
			document.getElementById('ClassID').disabled=true;
			break;
		case "ClassList":
			document.getElementById('OrderType').value='ID';
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			break;
		case "subClassList":
			document.getElementById('OrderType').value='ID';
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			document.getElementById('ColsNumber').disabled=false;
			document.getElementById('sub_colsnum').disabled=false;
			document.getElementById('bg_pic').disabled=false;
			document.getElementById('c_ColsNum').style.display='';
			break;
		case "SpecialListDown":
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			break;
	}
 }
function returnChooseType()
	{
			document.getElementById('ClassName_col').disabled=false;
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('ClassName_col').style.display='';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('ClassID').disabled=false;
			document.getElementById('ColsNumber').value='1';
	}
function selectHtml_express(Html_express)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
		document.getElementById('li_id').style.display='none';
		document.getElementById('ul_id').style.display='none';
		document.getElementById('DivID').disabled=true;
		document.getElementById('Divclass').disabled=true;
		document.getElementById('ulid').disabled=true;
		document.getElementById('ulclass').disabled=true;
		document.getElementById('liid').disabled=true;
		document.getElementById('liclass').disabled=true;
		document.getElementById('page_css').style.display='';
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('page_css').style.display='none';
		document.getElementById('DivID').disabled=false;
		document.getElementById('Divclass').disabled=false;
		document.getElementById('ulid').disabled=false;
		document.getElementById('ulclass').disabled=false;
		document.getElementById('liid').disabled=false;
		document.getElementById('liclass').disabled=false;
		break;
	}
}
function ShowTempLayer(oDiv){
	var obj=document.getElementById(oDiv);
	obj.style.display = '';
	obj.style.width = 280;
	obj.style.height = 120;
	obj.style.top = parseInt(document.body.scrollTop+event.clientY)-100;
	obj.style.left = parseInt(event.clientX)-280;
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../Down/lib/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.ClassID.value=TempArray[0]
		document.all.ClassName.value=TempArray[1]
	}
}
function OpenPageStyle(obj){
	var ret = OpenWindow("setPage.asp",200,200,"page");
	if(ret!='')obj.value = ret;
}
function GetObj(s){
	return document.getElementById(s);
}
function IsDisabled(oName){
	if(GetObj(oName).disabled == true){
		return true;
	}else{
		return false;
	}
}
var G_CanSave = true;
var G_Save_Msg = '';
function ok(obj){
	if(obj.NewsStyle.value == ''){alert('请填写引用样式!!\n如果还没建立，请在样式管理中建立');obj.NewsStyle.focus();return false;}
	//if(obj.labelName.value.length >25){alert('标签名称字数太多');obj.labelName.focus();return false;}
	if(obj.labelStyle.value=='ClassNews')
	{
		if(obj.ClassName.value==''){alert('请选择栏目');obj.ClassName.focus();return false;}
	}
	if(obj.labelStyle.value=='SpecialDown')
	{
		if(obj.specialEName.value==''){alert('请选择专题');obj.specialEName.focus();return false;}
	}
	if(obj.ContentNumber.value==''){alert('请填写内容显示字符数\n中文占2个字符');obj.ContentNumber.focus();return false;}
	if(isNaN(obj.NewsNumber.value)==true){alert('调用数量必须为数字');obj.NewsNumber.focus();return false;}
	if(isNaN(obj.ContentNumber.value)==true){alert('内容显示字数请用数字');obj.ContentNumber.focus();return false;}
	//if(obj.out_char.value=='out_DIV')
	//{
	//	if(obj.DivID.value==''){alert('请填写DIV的ID');obj.DivID.focus();return false;}
	//}
	if(obj.PageStyle.value == '')obj.PageStyle.value=',CC0066';
	var retV = '{FS:DS=';
	retV+=obj.labelStyle.value + '┆';
	if(!IsDisabled('labelName')) retV+='名称$' + obj.labelName.value + '┆';
	if(!IsDisabled('specialEName')) retV+='专题$' + obj.specialEName.value + '┆';
	if(!IsDisabled('ClassID')) retV+='栏目$' + obj.ClassID.value + '┆';
	/*if (obj.labelStyle.value!='ClassList'&&obj.labelStyle.value!='SpecialList')
	{
	if(!IsDisabled('NewsNumber')) retV+='Loop$'+ obj.NewsNumber.value + '┆';
	}
	if (obj.labelStyle.value!='ClassList')
	{
	if(!IsDisabled('NewsNumber')) retV+='Loop$'+ obj.NewsNumber.value + '┆';
	}  2006---12---28   by ken */
	retV+='Loop$'+ obj.NewsNumber.value + '┆';
	if(!IsDisabled('out_char')) retV+='输出格式$' + obj.out_char.value + '┆';
	if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '┆';
	if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '┆';
	if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '┆';
	if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '┆';
	if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '┆';
	if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '┆';
	if(!IsDisabled('DateNumber')) retV+='多少天$' + obj.DateNumber.value + '┆';
	if(!IsDisabled('TitleNumber')) retV+='标题数$' + obj.TitleNumber.value + '┆';
	if(!IsDisabled('PicTF')) retV+='图文标志$' + obj.PicTF.value + '┆';
	if(!IsDisabled('Openstyle')) retV+='打开窗口$' + obj.Openstyle.value + '┆';
	if(!IsDisabled('SubTF')) retV+='包含子类$' + obj.SubTF.value + '┆';
	if(!IsDisabled('OrderType')) retV+='排列字段$' + obj.OrderType.value + '┆';
	if(!IsDisabled('ASC')) if(obj.ASC.checked==true) retV+='排列方式$' + 'ASC┆';
	if(!IsDisabled('DESC')) if(obj.DESC.checked==true) retV+='排列方式$' + 'DESC┆';
	if(!IsDisabled('More_char')) retV+='更多连接$' + obj.More_char.value + '┆';
	if(!IsDisabled('PageTF')) retV+='分页$' + obj.PageTF.value + '┆';
	if(!IsDisabled('PageStyle')) retV+='分页样式$' + obj.PageStyle.value + '┆';
	if(!IsDisabled('PageNumber')) retV+='每页数量$' + obj.PageNumber.value + '┆';
	if(!IsDisabled('PageCSS')) retV+='PageCSS$' + obj.PageCSS.value + '┆';
	if(!IsDisabled('DateType')) retV+='日期格式$' + obj.DateType.value + '┆';
	if(!IsDisabled('NewsStyle')) retV+='引用样式$' + obj.NewsStyle.value + '┆';
	if(!IsDisabled('bg_pic')) retV+='背景底纹$' + obj.bg_pic.value + '┆';
	if(!IsDisabled('sub_colsnum')) retV+='栏目排列数$' + obj.sub_colsnum.value + '┆';
	retV+='排列数$' + obj.ColsNumber.value + '┆';
	retV+='内容字数$' + obj.ContentNumber.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






