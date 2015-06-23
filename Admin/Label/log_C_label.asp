<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
<title>标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td colspan="2" class="xingmu">分类调用标签</td>
    </tr>
		  <tr> 
            <td width="41%" class="xingmu"><strong>日志标签</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback"> 
      <td height="50%"><div align="left"><a href="FL_C_Label.asp?type=WordFL" target="_self"></a><a href="News_C_Label.asp?type=OldNews" target="_self"></a>日志标签采用固定模式，等待以后完善,如果要修改参数，请直接在生成的标签按照提示修改<br>
        <a href="Log_c_label.asp?type=0" target="_self">常规标签</a> | <a href="Log_c_label.asp?type=1" target="_self">分类标签</a> | <a href="Log_c_label.asp?type=2" target="_self">浏览日志标签</a>  | <a href="Log_c_label.asp?type=3" target="_self">用户列表参数</a> </div></td>
    </tr>
  </table>
  <%If Request.QueryString("type")="1" then%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="21%"><div align="right">选择分类</div></td>
    <td width="79%">
	<select name="ClassID" id="ClassID">
	<%
	dim rs
	set rs = User_Conn.execute("select id,ClassName From FS_ME_iLogClass Order by id asc")
	do while not rs.eof
		response.Write"<option value="""&rs("id")&""">"&rs("ClassName")&"</option>"
	rs.movenext
	loop
	rs.close:set rs=nothing
	%>
    </select>    </td>
  </tr>
  <tr class="hback">
    <td><div align="right">标题CSS</div></td>
    <td><input name="TitleCSS" type="text" id="TitleCSS"></td>
  </tr>
  <tr class="hback">
    <td><div align="right">调用数量</div></td>
    <td><input name="CodeNumber" type="text" id="CodeNumber" value="10"></td>
  </tr>
  <tr class="hback">
    <td><div align="right">标题字数</div></td>
    <td><input name="leftTitle" type="text" id="leftTitle" value="40">
      中文占2个字节</td>
  </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:宋体;display:none;" > 
      <td rowspan="3"  align="center" class="hback">
        <div align="right">DIV控制</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空"> 
        &quot;class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6" disabled  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid" disabled  type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" disabled id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">        &quot; class=&quot; <input name="liclass" disabled type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
  <tr class="hback">
    <td><div align="right">显示日期</div></td>
    <td><select name="DateTF" id="DateTF">
      <option value="1">显示</option>
      <option value="0" selected>不显示</option>
    </select>    </td>
  </tr>
  <tr class="hback">
    <td><div align="right">日期格式</div></td>
    <td><input name="DateType" type="text" id="DateType" value="MM月DD日">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
  </tr>
  <tr class="hback">
    <td><div align="right"></div></td>
    <td><input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
      <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
  </tr>
</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=ClassList┆';
		retV+='调用数量$' + obj.CodeNumber.value + '┆';
		retV+='ClassId$' + obj.ClassID.value + '┆';
		retV+='标题字数$' + obj.leftTitle.value + '┆';
		retV+='标题CSS$' + obj.TitleCSS.value + '┆';
		if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '┆';
		if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '┆';
		if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '┆';
		if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '┆';
		if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '┆';
		if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '┆';
		retV+='显示日期$' + obj.DateTF.value + '┆';
		retV+='日期格式$' + obj.DateType.value + '┆';
		if(!IsDisabled('out_char')) retV+='输出格式$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%elseif Request.QueryString("type")="2" then%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td colspan="2" class="xingmu">日志浏览页标签</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">选择日志标签</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType">
        <option value="Log_title">日志标题</option>
        <option value="Log_Content">日志内容</option>
        <option value="Log_Author">日志作者</option>
        <option value="Log_hits">日志点击率</option>
        <option value="Log_keywords">tags(关键字)</option>
        <option value="Log_AddTime">发布日期</option>
        <option value="Log_LogType">日志分类</option>
        <option value="Log_ReviewList">评论列表</option>
        <option value="Log_ReviewForm">发布评论表单</option>
      </select>      </td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=LogPage:' + obj.LogType.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%elseif Request.QueryString("type")="3" then%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td colspan="2" class="xingmu">用户首页参数</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">选择日志标签</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType">
        <option value="Log_Usertitle">站点名称</option>
        <option value="Log_UserName">用户名</option>
        <option value="Log_NickName">用户昵称</option>
        <option value="Log_UserContent">站点描述</option>
      </select>      </td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=LogUserPage:' + obj.LogType.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%else%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">常规标签</td>
    </tr>
    <tr>
      <td width="26%" class="hback"><div align="right">选择日志标签</div></td>
      <td width="74%" class="hback">
	  	<select name="LogType" id="LogType" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
        <option value="LastLog" selected>最新日志</option>
        <option value="TopLog">日志排行</option>
        <option value="HotLog">热点日志</option>
        <option value="TopSubject">日志专题</option>
        <option value="InfoClass">日志分类</option>
        <option value="InfoList">日志列表(终极)</option>
        <option value="Log_LastReview">最新评论</option>
        <option value="Log_LastForm">评论表单</option>
        <!--<option value="Log_MyInfo">个人资料</option>-->
        <option value="Log_PublicLog">发表日志连接</option>
        <option value="Log_Search">搜索</option>
        <option value="Log_Navi">导航</option>

      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">调用数量,标题字数</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10,40">
        格式&quot;10,40&quot;,请用半角&quot;,&quot;</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">日期格式</div></td>
      <td class="hback"><input name="DateType" type="text" id="DateType" value="MM-DD">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:宋体;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV控制</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空"> 
        &quot;class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6" disabled  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid" disabled  type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" disabled id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">        &quot; class=&quot; <input name="liclass" disabled type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
	<tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=' + obj.LogType.value + '┆';
		if(!IsDisabled('TitleNumber')) retV+='调用数量,标题字数$' + obj.TitleNumber.value + '┆';
		if(!IsDisabled('DivID')) retV+='DivID$' + obj.DivID.value + '┆';
		if(!IsDisabled('Divclass')) retV+='DivClass$' + obj.Divclass.value + '┆';
		if(!IsDisabled('ulid')) retV+='ulid$' + obj.ulid.value + '┆';
		if(!IsDisabled('ulclass')) retV+='ulclass$' + obj.ulclass.value + '┆';
		if(!IsDisabled('liid')) retV+='liid$' + obj.liid.value + '┆';
		if(!IsDisabled('liclass')) retV+='liclass$' + obj.liclass.value + '┆';
		retV+='CSS$┆';
		retV+='显示日期$1┆';
		if(!IsDisabled('DateType')) retV+='日期样式$' + obj.DateType.value + '┆';
		if(!IsDisabled('out_char')) retV+='输出格式$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end if%>
  </form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function IsDisabled(oName){
	if(GetObj(oName).disabled == true){
		return true;
	}else{
		return false;
	}
}
var G_CanSave = true;
var G_Save_Msg = '';
function GetObj(s){
	return document.getElementById(s);
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../News/lib/SelectClassFrame.asp',400,300,window);
	try {
		document.getElementById('ClassID').value = ReturnValue[0][0];
		document.getElementById('ClassName').value = ReturnValue[1][0];
	}
	catch (ex) { }
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
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('DivID').disabled=false;
		document.getElementById('Divclass').disabled=false;
		document.getElementById('ulid').disabled=false;
		document.getElementById('ulclass').disabled=false;
		document.getElementById('liid').disabled=false;
		document.getElementById('liclass').disabled=false;
		break;
	}
}
function ChooseNewsType(NewsType)
{
	switch (NewsType)
	{
		case "LastLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "TopLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "HotLog":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "TopSubject":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "InfoClass":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=false;
			break;
		case "InfoList":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_LastReview":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_LastForm":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_MyInfo":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_PublicLog":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Search":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Navi":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_InfoStat":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_PageTitle":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_title":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Content":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_Author":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_hits":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_keywords":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_AddTime":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_LogType":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
		case "Log_ReviewList":
			document.getElementById('TitleNumber').disabled=false;
			document.getElementById('DateType').disabled=false;
			document.getElementById('out_char').disabled=false;
			break;
		case "Log_ReviewContent":
			document.getElementById('TitleNumber').disabled=true;
			document.getElementById('DateType').disabled=true;
			document.getElementById('out_char').disabled=true;
			break;
	}
 }
</script>






