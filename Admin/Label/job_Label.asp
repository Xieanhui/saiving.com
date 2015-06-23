<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
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
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='AP' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	
%>
<html>
<head>
<title>人才标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="javascript" src="../../FS_Inc/prototype.js" type="text/javascript"></script>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <form  name="form1" method="post">
    <tr class="hback" > 
      <td colspan="2"  align="Left" class="xingmu"><a href="Job_label.asp" class="sd" target="_self"><strong><font color="#FF0000">创建标签</font></strong></a>｜<a href="All_label_style.asp?label_Sub=AP&TF=AP" target="_self" class="sd"><strong>样式管理</strong></a></td>
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
      <td colspan="2" class="hback" > <select  name="labelStyle" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
          <option value="" style="background:#DEDEDE">---列表类----------</option>
          <option value="ClassInfo" selected>├┄栏目列表</option>
          <option value="LastInfo">├┄最新</option>
		  <option value="readinfo">├┄信息浏览</option>
          <option value="" style="background:#DEDEDE">---终极类----------</option>
          <option value="ClassList">├┄终极信息列表</option>
        </select> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">标签类型</div></td>
      <td colspan="2" class="hback" >
	  <input type="radio" name="type" value="1" checked="checked" onClick="changePubdate(1)">招聘信息
	  <input type="radio" name="type" value="2" onClick="changePubdate(2)">人才信息	  </td>
    </tr>
    <tr class="hback" >
      <td class="hback"  align="center"><div align="right">工作地区</div></td>
      <td colspan="2" class="hback">
	  <select name="Province" id="Province" onChange="getCity(this.value)">
	  <%
	  	Dim ProvinceRs
		Set ProvinceRs=Conn.execute("Select pid,province from FS_AP_Province")
		Response.Write("<option value="""">选择省份</option>"&vbcrlf)
		while not ProvinceRs.eof
			Response.Write("<option value="""&ProvinceRs("PID")&""">"&ProvinceRs("Province")&"</option>"&vbcrlf)
			ProvinceRs.movenext
		wend
	  %>
	  </select>
	  <span id="span_city"></span>
	  <span id="span_selected"></span>
	  </td>
    </tr>
    <tr class="hback" style="display:none"> 
      <td width="21%"  align="center" class="hback"><div align="right">标签名称</div></td>
      <td colspan="2" class="hback" ><input name="labelName"  type="text" size="12" maxlength="25"> 
        <span class="tx">限制25个字符(只能为中文、数字、英文、下划线和中划线)。</span></td>
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
        天内的新闻. <span class="tx">如果为0，则显示所有时间内的新闻</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="NewsNum" align="center"><div align="right">标题字数</div></td>
      <td colspan="2" class="hback" ><input name="TitleNumber" type="text" id="TitleNumber" value="30" size="5"> 
        <span class="tx">中文算2个字符</span>图文标志 
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
      <td class="hback"  id="NewsNum" align="center"><div align="right">调用数量</div></td>
      <td colspan="2" class="hback" > <input id="NewsNumber"  name="NewsNumber" type="text" size="12" value="10"> 
        <span class="tx">调用的前台显示数量</span></td>
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
        </select> <span class="tx">新闻一行显示数量,如果是DIV+CSS，请在CSS里控制条数</span> </td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="ColsNum" align="center"><div align="right">包含子类</div></td>
      <td colspan="2" class="hback" ><select name="SubTF" id="SubTF">
          <option value="1">是</option>
          <option value="0" selected>否</option>
        </select> <span class="tx">建议不要选择，否则新闻生成速度会大幅度降低,同时将大量占用服务器资源和内存</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  id="Num" align="center"><div align="right">排序字段</div></td>
      <td colspan="2" class="hback" > <select id="OrderType"  name="OrderType">
          <option value="PID">自动编号</option>
          <option value="PublicDate">发布日期</option>
        </select> 排序方式 <select id="orderby"  name="orderby">
          <option value="ASC">升序</option>
          <option value="DESC" selected>降序</option>
        </select>&nbsp;<span id="last_desc" class="tx"></span> </td>
    </tr>
    <tr class="hback" id="More_col" style="display:none;"> 
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
    <tr class="hback" > 
      <td class="hback"  id="NewsStyle" align="center"><div align="right">日期格式</div></td>
      <td colspan="2" class="hback" ><input name="DateType" type="text" id="DateType" value="YY02年MM月DD日" size="20"> 
        <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr class="hback" > 
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面新闻显示样式</span></div></td>
    </tr>
    <tr class="hback" > 
      <td class="hback" align="center" height="30"><div align="right">其他</div></td>
      <td height="30"  colspan="2" align="center" class="hback"><div align="left"> 
          内容显示最多字符 
          <input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
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
function ChooseNewsType(NewsType)
{
	switch (NewsType)
	{
		case "ClassNews":
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "LastNews":
			document.getElementById('OrderType').value='AddTime';
			document.getElementById('last_desc').innerHTML="如果是最新新闻，请选择'<b>添加日期</b>'和'<b>降序</b>'排列"
			returnChooseType();
			break;
		case "DayNews":
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "BriNews":
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "AnnNews":
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "ConstrNews":
			document.getElementById('last_desc').innerHTML=""
			returnChooseType();
			break;
		case "ClassList":
			document.getElementById('PageStyle_col').disabled=false;
			document.getElementById('SetPage').disabled=false;
			document.getElementById('PageTF').disabled=false;
			document.getElementById('PageStyle').disabled=false;
			document.getElementById('PageNumber').disabled=false;
			document.getElementById('PageStyle_col').style.display='';
			document.getElementById('More_col').style.display='none';
			document.getElementById('More_char').disabled=true;
			document.getElementById('ColsNumber').value='1';
			document.getElementById('Province').disabled=true;
			break;
	}
 }
function returnChooseType()
	{
			document.getElementById('PageStyle_col').disabled=true;
			document.getElementById('SetPage').disabled=true;
			document.getElementById('PageTF').disabled=true;
			document.getElementById('PageStyle').disabled=true;
			document.getElementById('PageNumber').disabled=true;
			document.getElementById('PageStyle_col').style.display='none';
			document.getElementById('More_col').style.display='';
			document.getElementById('More_char').disabled=false;
			document.getElementById('ColsNumber').value='1';
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
		//document.getElementById('Province').disabled=true;
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
	if(obj.contentnumber.value==''){alert('请填写内容显示字符数\n中文占2个字符');obj.contentnumber.focus();return false;}
	if(isNaN(obj.NewsNumber.value)==true){alert('调用数量必须为数字');obj.NewsNumber.focus();return false;}
	if(isNaN(obj.contentnumber.value)==true){alert('内容显示字数请用数字');obj.contentnumber.focus();return false;}
	if(obj.PageStyle.value == '')obj.PageStyle.value=',CC0066';
	var retV = '{FS:AP=';
	retV+=obj.labelStyle.value + '┆';
	if(!IsDisabled('labelName')) retV+='名称$' + obj.labelName.value + '┆';
	if (obj.labelStyle.value!='ClassList'&&obj.labelStyle.value!='SpecialList')
	{
	if(!IsDisabled('NewsNumber')) retV+='Loop$'+ obj.NewsNumber.value + '┆';
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
	if(!IsDisabled('NewsStyle')) retV+='引用样式$' + obj.NewsStyle.value + '┆';
	retV+='排列数$' + obj.ColsNumber.value + '┆';
	retV+='内容字数$' + obj.contentnumber.value + '┆';
	if(document.all('type')[0].checked)
		retV+='类型$'+document.all('type')[0].value+''
	else
		retV+='类型$'+document.all('type')[1].value+''
	if (!IsDisabled('Province'))
	{
		retV+='┆';
		if($('span_City').innerHTML!=""&&$('City').options[$('City').selectedIndex].value!="")
			retV+='工作地区$'+$('City').options[$('City').selectedIndex].text+''
		else if($('Province').options[$('Province').selectedIndex].value=="")
			retV+='工作地区$'
		else
			retV+='工作地区$'+$('Province').options[$('Province').selectedIndex].text+''
	}
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
function changePubdate(value)
{	
	var orderType=document.all("OrderType")
	if(value==1)
	{
		orderType.options[0].value="PID"
		orderType.options[1].value="PublicDate"
	}else
	{
		orderType.options[0].value="BID"
		orderType.options[1].value="lastTime"
	}
}
function getCity(pid)
{
	if(pid==""||isNaN(pid))
	{
		$("span_selected").innerHTML="当前选择：<font color=\"red\">无</font>"
		return false;
	}
	$("span_selected").innerHTML="当前选择：<font color=\"red\">"+ $('Province').options[$('Province').selectedIndex].text+"</font>";
	var ajax=new Ajax.Updater("span_city","getCity_Ajax.asp",{method:'get',parameters:"action=getCity&pid="+pid+"&rnd="+Math.random()});
	
}
function selectCity(cid)
{
	if(cid==""||isNaN(cid))
	{
		$("span_selected").innerHTML="当前选择：<font color=\"red\">"+ $('Province').options[$('Province').selectedIndex].text+"</font>";
		return false;
	}
	$("span_selected").innerHTML="当前选择：<font color=\"red\">"+ $('City').options[$('City').selectedIndex].text+"</font>";
}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





