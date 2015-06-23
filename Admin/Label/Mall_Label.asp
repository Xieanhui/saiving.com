<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/MS_Function.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='MS' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	Dim obj_special_Rs,label_special_List
	label_special_List=""
	Set  obj_special_Rs = server.CreateObject(G_FS_RS)
	obj_special_Rs.Open "Select SpecialID,SpecialCName from FS_MS_Special  Order by  SpecialID desc",Conn,1,3
	do while Not obj_special_Rs.eof 
		label_special_List = label_special_List&"<option value="""& obj_special_Rs(0)&""">"& obj_special_Rs(1)&"</option>"
		obj_special_Rs.movenext
	loop
	obj_special_Rs.close:set obj_special_Rs = nothing
%>
<html>
<head>
<title>新闻标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" > 
      <td colspan="2"  align="Left" class="xingmu"><a href="Malllabel.asp" class="sd" target="_self"><strong><font color="#FF0000">创建标签</font></strong></a>｜<a href="All_label_style.asp?label_Sub=MS&TF=MS" target="_self" class="sd"><strong>样式管理</strong></a></td>
      <td width="38%"  align="Left" class="xingmu"><div align="right"> 
          <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
        </div></td>
    </tr>
    <tr class="hback"  style="font-family:宋体" > 
      <td  align="center" class="hback" ><div align="right">显示格式</div></td>
      <td colspan="2" class="hback" > <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">标签类型</div></td>
      <td colspan="2" class="hback" > <select  name="labelStyle" onChange="ChooseProductType(this.options[this.selectedIndex].value);">
          <option value="lbl" style="background:#DEDEDE" selected>---列表类----------</option>
          <option value="ClassProducts">├┄栏目列表</option>
          <option value="SpecialProducts">├┄专题列表</option>
		  <option value="LastProducts">├┄最新</option>
          <option value="HotProducts">├┄热点</option>
          <option value="RecProducts">├┄推荐</option>		  
          <option value="MarProducts">├┄滚动</option>
          <option value="SpecialOffer">├┄特价</option>
          <option value="Sales">├┄降价</option>
		  <option value="onePrice">├┄一口价</option>
		  <option value="publicsale">├┄竞拍</option>
          <option value="" style="background:#DEDEDE">---终极类----------</option>
          <option value="ClassList">├┄终极商品列表</option>
		 <!-- <option value="subClassList">├┄子类商品列表</option>-->
          <option value="SpecialList">├┄终极专题列表</option>
        </select> <span class="tx">如果是终极分类，不能选择栏目</span> </td>
    </tr>
    <tr class="hback" style="display:none"> 
      <td width="19%"  align="center" class="hback"><div align="right">标签名称</div></td>
      <td colspan="2" class="hback" ><input name="labelName"  type="text" size="12" maxlength="25"> 
        <span class="tx">限制25个字符(只能为中文、数字、英文、下划线和中划线)。</span></td>
    </tr>
    <tr class="hback" id="specialEName_col" style="display:none"> 
      <td class="hback"  align="center"><div align="right">专题列表</div></td>
      <td colspan="2" class="hback" > <select id="specialEName"  name="specialEName" disabled>
          <option value="">请选择专题</option>
          <% = label_special_List %>
        </select> <span class="tx">如果不选择，则为所有专题导航</span> </td>
    </tr>
    <tr class="hback" id="ClassName_col"> 
      <td  align="center" class="hback"><div align="right">栏目列表</div></td>
      <td colspan="2" class="hback" > <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button2" type="button" onClick="SelectClass();" value="选择栏目"> 
        <span class="tx">选择商城栏目，如果不选择，则显示所有的商品</span></td>
    </tr>
    <tr class="hback" id="Noselect" style="display:none">
      <td  align="center" class="hback"><div align="right">序号排列：</div></td>
      <td colspan="2" class="hback" >
        <select name="No_Select" id="No_Select">
			<option value="0">无序列</option>
			<option value="1">A.B.C.</option>
			<option value="2">a.b.c.</option>
			<option value="3">1.2.3.</option>
        </select><span class="tx">序号排序不能应用于“终极类”列表上。</span>
      </td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:宋体;display:none;" > 
      <td rowspan="2"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV控制</div></td>
      <td colspan="2" class="hback" >&lt;div id=&quot; <input name="DivID" disabled type="text" id="DivID" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"  disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt; &lt;ul id=&quot; <input name="ulid"  disabled type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt; </td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:宋体;display:none;"> 
      <td colspan="2" class="hback" >&lt;li id=&quot; <input name="liid" disabled type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!"> 
        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt; </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">显示范围</div></td>
      <td colspan="2" class="hback" >只显示 
        <input name="DateNumber"  type="text" id="DateNumber" value="0" size="5">
        天内的商城. <span class="tx">如果为0，则显示所有时间内的商城</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ProductsNum" align="center"><div align="right">标题字数</div></td>
      <td colspan="2" class="hback" ><input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5"> 
        <span class="tx">中文算2个字符</span> 　　图文标志 
        <select name="PicTF" id="PicTF">
          <option value="1" selected>显示</option>
          <option value="0">不显示</option>
        </select>
        　　打开窗口 
        <select name="Openstyle" id="Openstyle">
          <option value="0" selected>原窗口</option>
          <option value="1">新窗口</option>
        </select> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ProductsNum" align="center"><div align="right">调用数量</div></td>
      <td colspan="2" class="hback" > <input id="ProductsNumber"  name="ProductsNumber" type="text" size="12" value="10"> 
        <span class="tx">调用的前台显示数量　　</span> 　　　　 </td>
    </tr>
    <tr class="hback" > 
      <td class="hback" align="center"><div align="right">排列列数</div></td>
      <td colspan="2" class="hback" > <select name="ColsNumber" id="select">
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
        </select> <span class="tx">文章一行显示数量,如果是DIV+CSS，请在CSS里控制条数</span> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ColsNum" align="center"><div align="right">包含子类</div></td>
      <td colspan="2" class="hback" ><select name="SubTF" id="SubTF">
          <option value="1">是</option>
          <option value="0" selected>否</option>
        </select> <span class="tx">建议不要选择，否则文章生成速度会大幅度降低,同时将大量占用服务器资源和内存</span></td>
    </tr>
    <tr id="sort" class="hback" > 
      <td class="hback" id="Num" align="center"><div align="right">排序字段</div></td>
      <td colspan="2" class="hback" > <select id="OrderType"  name="OrderType">
          <option value="ID">自动编号</option>
          <option value="AddTime">添加日期</option>
          <option value="PopId" selected>商品权重</option>
          <option value="Click">点击次数</option>
        </select> 排序方式 <select id="orderby"  name="orderby">
          <option value="ASC">升序</option>
          <option value="DESC" selected>降序</option>
        </select>&nbsp;<span id="last_desc" class="tx"></span> </td>
    </tr>
    <tr class="hback" id="More_col"> 
      <td class="hback"  id="PageStyle_1" align="center"><div align="right">更多连接</div></td>
      <td colspan="2" class="hback" ><input  name="More_char" type="text" id="More_char" value="·" size="20"> 
        <span class="tx">请输入html语法,如：&lt;img src=&quot;/files/more.gif&quot; border='0'&gt;</span></td>
    </tr>
    <tr class="hback" id="PageStyle_col" style="display:none"> 
      <td class="hback"   align="center"><div align="right">是否分页</div></td>
      <td colspan="2" class="hback" ><input name="PageTF" type="checkbox" id="PageTF" value="1" disabled>
        　 分页样式 
        <input id="PageStyle"  name="PageStyle" type="text" size="9" value='3,CC0066'  disabled> 
        <input name="button" type="button" id="SetPage" onClick="OpenPageStyle(this.form.PageStyle)" value="设置" disabled>
        　每页数量 
        <input id="PageNumber"  name="PageNumber" type="text" size="4" value="30" disabled>
        　<span id="page_css" style="display:">分页CSS 
        <input id="PageCSS" name="PageCSS" type="text" size="5">
        </span> <span class="tx">仅适用终极分类</span></td>
    </tr>
    <tr class="hback" id="Mar_cols" style="display:none"> 
      <td class="hback"  id="ScrollSpeed" align="center"><div align="right">滚动速度</div></td>
      <td colspan="2" class="hback" ><input id="MarqueeSpeed"  name="MarqueeSpeed" type="text" size="12" value='8'  disabled>
        　滚动方向 
        <select id="MarqueeDirection"  name="MarqueeDirection" disabled>
          <option value="up">向上</option>
          <option value="down">向下</option>
          <option value="left" selected>向左</option>
          <option value="right">向右</option>
        </select>
        <span style="display:none">选择样式</span>
        <select name="Marqueestyle" id="Marqueestyle" disabled style="display:none">
          <option value="0" selected>滚动</option>
          <option value="1">闪动</option>
        </select></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ProductsStyle" align="center"><div align="right">日期格式</div></td>
      <td colspan="2" class="hback" ><input name="DateType" type="text" id="DateType" value="YY02年MM月DD日" size="20"> 
        <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="ProductsStyle"  name="ProductsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('Products_label_styleread.asp?ID='+document.form1.ProductsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面文章显示样式</span></div></td>
    </tr>
    <tr class="hback" > 
      <td class="hback" align="center" height="30"><div align="right">其他</div></td>
      <td height="30"  colspan="2" align="center" class="hback"><div align="left"> 
          内容显示最多字符 
          <input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
          文章导读显示最多字符 
          <input name="navinumber" type="text" id="navinumber" value="200" size="12">
        中文2个字符</div></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  colspan="3" align="center" height="30"> <input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签"> 
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
function ChooseProductType(ProductType)
{
	switch (ProductType)
	{
		case "ClassProducts":
			document.getElementById('Noselect').style.display='';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "LastProducts":
			document.getElementById('Noselect').style.display='';
			document.getElementById('sort').style.display='none';
			document.getElementById('OrderType').value='AddTime';
			document.getElementById('orderby').value="DESC"
			returnChooseType();
			break;
		case "HotProducts":
			document.getElementById('Noselect').style.display='';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "RecProducts":
			document.getElementById('Noselect').style.display='';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
			//<--//2008/08/18:Code crazy 目的让用户选择类别时将Noselect显示出来
		case "SpecialOffer"://特价
			document.getElementById('Noselect').style.display='';
			break;
		case "Sales"://降价
			document.getElementById('Noselect').style.display='';
			break;
		case "onePrice"://一口价
			document.getElementById('Noselect').style.display='';
			break;
		case "publicsale"://竞拍
			document.getElementById('Noselect').style.display='';
			break;
			//-->//
		case "DayProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "BriProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "AnnProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "ConstrProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "SpecialProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('Mar_cols').style.display='none';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('specialEName_col').style.display='';
			document.getElementById('specialEName').disabled=false;
			document.getElementById('ClassID').disabled=true;
			break;
		case "MarProducts":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('ClassName_col').disabled=false;
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('Mar_cols').disabled=false;
			document.getElementById('ClassName_col').style.display='';
			document.getElementById('Mar_cols').style.display='';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('MarqueeSpeed').disabled=false;
			document.getElementById('MarqueeDirection').disabled=false;
			//document.getElementById('Marqueestyle').disabled=false;
			document.getElementById('ColsNumber').value='10';
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			document.getElementById('ClassID').disabled=false;
			break;
		case "ClassList":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('Mar_cols').disabled='';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			break;
		case "SpecialList":
			document.getElementById('Noselect').style.display='none';
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('Mar_cols').disabled='';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			break;
/*		case "subClassList":
			document.getElementById('ClassName_col').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('ClassName_col').style.display='none';
			document.getElementById('Mar_cols').disabled='';
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('ClassID').disabled=true;
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
			document.getElementById('ColsNumber').disabled=false;
			break;
*/	}
 }
function returnChooseType()
	{
			document.getElementById('ClassName_col').disabled=false;
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('Mar_cols').disabled=true;
			document.getElementById('ClassName_col').style.display='';
			document.getElementById('Mar_cols').style.display='none';
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('ClassID').disabled=false;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('MarqueeSpeed').disabled=true;
			document.getElementById('MarqueeDirection').disabled=true;
			document.getElementById('specialEName_col').style.display='none';
			document.getElementById('specialEName').disabled=true;
	}
function selectHtml_express(Html_express)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
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
	ReturnValue = OpenWindow('../Mall/lib/SelectClassFrame.asp',400,300,window);
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
	if(obj.labelStyle.value=='lbl')
	{
		alert('请选择“标签类型”');
		obj.labelStyle.focus();
		return false;
	}
	if(obj.ProductsStyle.value == ''){alert('请填写引用样式!!\n如果还没建立，请在样式管理中建立');obj.ProductsStyle.focus();return false;}
	//if(obj.labelName.value.length >25){alert('标签名称字数太多');obj.labelName.focus();return false;}
	if(obj.labelStyle.value=='ClassProducts')
		{
			//if(obj.ClassName.value==''){alert('请选择栏目');obj.ClassName.focus();return false;}
		}
	if(obj.labelStyle.value=='MarProducts')
		{
			if(isNaN(obj.MarqueeSpeed.value)==true){alert('滚动速度请填写数字');obj.MarqueeSpeed.focus();return false;}
		}
	if(obj.contentnumber.value==''){alert('请填写内容显示字符数\n中文占2个字符');obj.contentnumber.focus();return false;}
	if(obj.navinumber.value==''){alert('请填写新闻导航显示字符数\n中文占2个字符');obj.navinumber.focus();return false;}
	if(isNaN(obj.ProductsNumber.value)==true){alert('调用数量必须为数字');obj.ProductsNumber.focus();return false;}
	if(isNaN(obj.contentnumber.value)==true){alert('内容显示字数请用数字');obj.contentnumber.focus();return false;}
	if(isNaN(obj.navinumber.value)==true){alert('导航显示字数请用数字');obj.navinumber.focus();return false;}
	if(obj.PageStyle.value == '')obj.PageStyle.value=',CC0066';
	var No_selects =  '序号排列$' + obj.No_Select.value + '';
	if(obj.labelStyle.value=='SpecialList') No_selects = '';
	if(obj.labelStyle.value=='ClassList') No_selects = '';
	if(obj.labelStyle.value=='MarProducts') No_selects = '';
	if(obj.labelStyle.value=='SpecialProducts') No_selects = '';
	
	var retV = '{FS:MS=';
	retV+=obj.labelStyle.value + '┆';
	if(!IsDisabled('labelName')) retV+='名称$' + obj.labelName.value + '┆';
	if(!IsDisabled('specialEName')) retV+='专题$' + obj.specialEName.value + '┆';
	if(!IsDisabled('ClassID')) retV+='栏目$' + obj.ClassID.value + '┆';
	if (obj.labelStyle.value!='ClassList'&&obj.labelStyle.value!='SpecialList')
	{
	if(!IsDisabled('ProductsNumber')) retV+='Loop$'+ obj.ProductsNumber.value + '┆';
	}
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
	if(!IsDisabled('orderby')) retV+='排列方式$' + obj.orderby.value + '┆';
	if(!IsDisabled('More_char')) retV+='更多连接$' + obj.More_char.value + '┆';
	if(!IsDisabled('PageTF')) retV+='分页$' + obj.PageTF.value + '┆';
	if(!IsDisabled('PageStyle')) retV+='分页样式$' + obj.PageStyle.value + '┆';
	if(!IsDisabled('PageNumber')) retV+='每页数量$' + obj.PageNumber.value + '┆';
	if(!IsDisabled('PageCSS')) retV+='PageCSS$' + obj.PageCSS.value + '┆';
	if(!IsDisabled('DateType')) retV+='日期格式$' + obj.DateType.value + '┆';
	if(!IsDisabled('ProductsStyle')) retV+='引用样式$' + obj.ProductsStyle.value + '┆';
	retV+='排列数$' + obj.ColsNumber.value + '┆';
	retV+='内容字数$' + obj.contentnumber.value + '┆';
	if(obj.labelStyle.value=='MarProducts')
	{retV+='导航字数$' + obj.navinumber.value + '┆';}
	else
	{
		if (No_selects == '')
		{
			retV+='导航字数$' + obj.navinumber.value +'';
		}
		else
		{
			retV+='导航字数$' + obj.navinumber.value +'┆';
		}
	}
	if(!IsDisabled('MarqueeSpeed')) retV+='滚动速度$' + obj.MarqueeSpeed.value + '┆';
	if(!IsDisabled('MarqueeDirection'))
	{ 
		if (No_selects =='')
		{
			retV+='滚动方向$' + obj.MarqueeDirection.value +'';
			
		}
		else
		{
			retV+='滚动方向$' + obj.MarqueeDirection.value +'┆';
		}
	}
	retV+=No_selects
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





