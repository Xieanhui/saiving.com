<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
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
	MF_Default_Conn
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='DS' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	
	Dim obj_special_Rs,label_special_List
	label_special_List=""
	Set  obj_special_Rs = server.CreateObject(G_FS_RS)
	obj_special_Rs.Open "Select SpecialID,SpecialCName,specialEName from FS_DS_Special  Order by  SpecialID desc",Conn,1,3
	do while Not obj_special_Rs.eof 
		label_special_List = label_special_List&"<option value="""& obj_special_Rs(2)&""">"& obj_special_Rs(1)&"</option>"
		obj_special_Rs.movenext
	loop
	obj_special_Rs.close:set obj_special_Rs = nothing
	'================================
	'获取下载子系统自由标签下拉列表
	'================================
	Function GetNewsFreeList(SysType)
	Dim Rs,Sql
	Sql = "Select LabelID,LabelName From FS_MF_FreeLabel Where ID > 0 And SysType = '" & SysType & "'"
	Set Rs = Conn.ExeCute(Sql)
	GetNewsFreeList = "<select name=""FreeList"" id=""FreeList"">" & vbnewline
	GetNewsFreeList = GetNewsFreeList & "<option value="""">选择自由标签</option>"
	If Rs.Eof Then
		GetNewsFreeList = GetNewsFreeList & ""
	Else
		Do While Not Rs.Eof
			GetNewsFreeList = GetNewsFreeList & "<option value=""" & Rs(0) & """>" & Rs(1) & "</option>" & vbnewline
		Rs.MoveNext
		Loop
	End If
	GetNewsFreeList = GetNewsFreeList & "</select>" & vbnewline
	Rs.Close : Set Rs = NOthing
	End Function
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
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="66%" class="xingmu"><strong>常规标签创建</strong></td>
            <td width="34%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="13%" height="15"><div align="center"><a href="down_C_Label.asp?type=ReadNews" target="_self">下载浏览</a></div></td>
      <td width="12%"><div align="center"><a href="down_C_Label.asp?type=ClassNavi" target="_self">栏目导航</a></div></td>
      <td width="16%"><div align="center"><a href="down_C_Label.asp?type=siteMap" target="_self">站点地图</a></div></td>
      <td width="15%"><div align="center"><a href="down_C_Label.asp?type=Search" target="_self">搜索表单</a></div></td>
      <td width="16%"><div align="center"><a href="down_C_Label.asp?type=infoStat" target="_self">信息统计</a></div></td>
	  <td width="16%"><div align="center"><a href="down_C_Label.asp?type=down_relative" target="_self">相关下载</a></div></td>
    </tr>
	<tr class="hback">
	<td><div align="center"><a href="down_C_Label.asp?type=SpecialNavi" target="_self">专区导航</a></div></td>
	<td><div align="center"><a href="down_C_Label.asp?type=SpecialCode" target="_self">专区调用</a></div></td>
	<td><div align="center"><a href="down_C_Label.asp?type=FreeLabel" target="_self">自由标签</a></div></td>
	<td><div align="center"></div></td>
	<td><div align="center"></div></td>
	<td><div align="center"></div></td>
	</tr>
			
  </table>     
  <%
select case Request.QueryString("type")
		case "ReadNews"
			call readnews()
		case "siteMap"
			call siteMap()
		case "Search"
			call Search()
		case "infoStat"
			call infoStat()
		case "ClassNavi"
			call ClassNavi()
		case "SpecialNavi"
			call SpecialNavi()
		case "SpecialCode"
			call SpecialCode()
		case "down_relative"
		    call down_relative()
		Case "FreeLabel"
			call FreeLabel()	
		case else
			call readnews()
end select
%>
  <%sub readnews()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="21%" class="hback"><div align="right">引用样式</div></td>
      <td width="79%" class="hback"> <select id="NewsStyle"  name="NewsStyle" style="width:40%">
          <% = label_style_List %>
        </select> <input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
        <span class="tx">请在各个子系统中建立前台页面下载显示样式</span> </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示日期格式</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
        <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签"> 
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:DS=ReadNews┆';
	retV+='引用样式$' + obj.NewsStyle.value + '┆';
	retV+='日期格式$' + obj.DateStyle.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
<%end sub%>
<%sub siteMap()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">选择栏目</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button22" type="button" onClick="SelectClass();" value="选择栏目"> 
        <span class="tx"></span></td>
    </tr>
    <tr style="display:none"> 
      <td class="hback"><div align="right">排列方式</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>横向</option>
          <option value="1">纵向</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">标题CSS</div></td>
      <td class="hback"><input  name="Titlecss" type="text" id="Titlecss" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:DS=siteMap┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  <%end sub%>
 <%sub SpecialNavi()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">输出格式</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">专题CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题导航图片/文字</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
				请使用html语法</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:DS=SpecialNavi┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='方向$' + obj.cols.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='CSS$' + obj.Titlecss.value + '┆';
		retV+='导航$' + obj.TitleNavi.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub SpecialCode()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">选择专区</div>
			</td>
			<td width="78%" class="hback">
				<select id="specialEName"  name="specialEName">
					<option value="">请选择专区</option>
					<% = label_special_List %>
				</select>
				<span class="tx">必须选择</span></td>
		</tr>
		<tr>
			<td width="22%" class="hback">
				<div align="right">输出格式</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示图片</div>
			</td>
			<td class="hback">
				<select name="PicTF" id="PicTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				图片高度及宽度
				<input name="PicSize" type="text" id="PicSize" value="120,100" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示专区导航内容</div>
			</td>
			<td class="hback">
				<select name="NaviTF" id="NaviTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				<input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片CSS</div>
			</td>
			<td class="hback">
				<input name="PicCSS" type="text" id="PicCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">名称CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
				内容CSS
				<input name="ContentCSS" type="text" id="ContentCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
				只对table格式有效 　　导航
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		if(obj.specialEName.value=='')
		{
		alert('请选择专区');
		obj.specialEName.focus();
		return false;
		}
		var retV = '{FS:DS=SpecialCode┆';
		retV+='专题$' + obj.specialEName.value + '┆';
		retV+='图片显示$' + obj.PicTF.value + '┆';
		retV+='图片尺寸$' + obj.PicSize.value + '┆';
		retV+='导航内容$' + obj.NaviTF.value + '┆';
		retV+='导航内容字数$' + obj.NaviNumber.value + '┆';
		retV+='专题名称CSS$' + obj.TitleCSS.value + '┆';
		retV+='导航内容CSS$' + obj.ContentCSS.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='排列方式$' + obj.cols.value + '┆';
		retV+='导航$' + obj.TitleNavi.value + '┆';
		retV+='图片css$' + obj.PicCSS.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
<%sub Search()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">日期搜索</div></td>
      <td width="78%" class="hback"><select name="DateShow"  id="DateShow">
          <option value="1" selected>显示</option>
          <option value="0">不显示</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">搜索栏目</div></td>
      <td class="hback"><select name="classShow"  id="classShow">
          <option value="1" selected>显示</option>
          <option value="0">不显示</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:DS=Search┆';
		retV+='显示日期$' + obj.DateShow.value + '┆';
		retV+='显示栏目$' + obj.classShow.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
<%end sub%>
<%sub infoStat()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td class="hback"><div align="right">排列方式</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>横向</option>
          <option value="1">纵向</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:DS=infoStat┆';
		retV+='显示方向$' + obj.cols.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
<%end sub%>

<%sub down_relative()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">相关条件</div>
			</td>
			<td class="hback">
				<select name="ifelse" id="ifelse" >
					<option value="0" selected>标题相关</option>
					<option value="1">开发商相关</option>
				</select>
			</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right">显示数量</div>
			</td>
			<td width="79%" class="hback">
				<label>
				<input name="titleNumber" type="text" id="titleNumber" value="10" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题字数</div>
			</td>
			<td class="hback">
				<label>
				<input name="leftTitle" type="text" id="leftTitle" value="40" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
				中文占2个字符</label>
			</td>
		</tr>
		<tr> 
		  <td width="21%" class="hback"><div align="right">引用样式</div></td>
		  <td width="79%" class="hback"> <select id="NewsStyle"  name="NewsStyle" style="width:40%">
			  <% = label_style_List %>
			</select> <input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
			</td>
		</tr>
		<tr> 
		  <td width="21%" class="hback"><div align="right">简介字数</div></td>
		  <td width="79%" class="hback">
		   <input name="ContentNumber" type="text" id="ContentNumber" value="100" size="5" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
           <span class="tx">中文算2个字符</span>
		   图文标志
			<select name="PicTF" id="PicTF">
			  <option value="1">显示</option>
			  <option value="0" selected>不显示</option>
			</select>
		  </td>
		</tr>
		 <tr class="hback" >
		  <td class="hback"  align="center"><div align="right">显示范围</div></td>
		  <td colspan="3" class="hback" >只显示
			<input name="DateNumber"  type="text" id="DateNumber" value="0" size="5" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
			天内的下载. <span class="tx">如果为0，则显示所有时间内的下载</span></td>
		</tr>
		 <tr class="hback" >
		  <td class="hback"  id="Num" align="center"><div align="right">排序字段</div></td>
		  <td class="hback" >
		     <select name="OrderBy" id="OrderBy">
			  <option value="ID" selected>自动编号</option>
			  <option value="AddTime">添加时间</option>
			  <option value="EditTime">修改时间</option>
			  <option value="Hits">点击次数</option>
			  <option value="ClickNum">下载次数</option>
			</select>
		   排序方式
			<select name="OrderDesc" id="OrderDesc">
			  <option value="Desc" selected>降序</option>
			  <option value="Asc">升序</option>
			</select>
			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">显示日期格式</div></td>
		  <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
		</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:DS=down_relative┆';   
	retV+='相关条件$' + obj.ifelse.value + '┆';
	retV+='显示数量$' + obj.titleNumber.value + '┆';
	retV+='标题字数$' + obj.leftTitle.value + '┆';
	retV+='引用样式$' + obj.NewsStyle.value + '┆';
	retV+='简介字数$' + obj.ContentNumber.value + '┆';
	retV+='图文标记$' + obj.PicTF.value + '┆';
	retV+='日期范围$' + obj.DateNumber.value + '┆';
	retV+='排序字段$' + obj.OrderBy.value + '┆';
	retV+='排序方式$' + obj.OrderDesc.value + '┆';
	retV+='日期格式$' + obj.DateStyle.value;
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
<%end sub%>

<%sub ClassNavi()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx">如果不选择，那么在某个类就调用某个类的导航</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">输出格式</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题导航</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
				请使用html语法</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:DS=ClassNavi┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='方向$' + obj.cols.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '┆';
		retV+='标题导航$' + obj.TitleNavi.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<% Sub FreeLabel() %>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	  <tr>
	    <td class="hback">
		  <div align="right">自由标签</div>
	    </td>
	    <td class="hback">
	    	<% = GetNewsFreeList("DS") %>
		</td>
	  </tr>
	  <tr>
	    <td class="hback">
		  <div align="right"></div>
	    </td>
	    <td class="hback">
		  <input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
		  <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
	    </td>
	  </tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=FreeLabel┆';
		retV+='自由标签$' + obj.FreeList.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<% End Sub %>
</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
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
function selectHtml_express(Html_express)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
		document.getElementById('li_id').style.display='none';
		document.getElementById('ul_id').style.display='none';
		document.getElementById('DivID').disabled=true;
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('DivID').disabled=false;
		break;
	}
}
</script>






