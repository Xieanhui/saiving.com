<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
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
		obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='HS' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
%>
<html>
<head>
<title>房产标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<form  name="form1" method="post">
	<table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
		<tr class="hback" >
			<td height="27"  align="Left" class="xingmu">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="20%" class="xingmu"><strong>常规标签创建</strong></td>
						<td width="80%">
							<div align="right">
								<input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
							</div>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr class="hback">
			<td width="13%" height="15">
				<div align="center"><a href="House_C_Label.asp?type=ReadInfo" target="_self">房产信息浏览</a></div>
			</td>
			<td width="12%">
				<div align="center"><a href="House_C_Label.asp?type=FlashFilt" target="_self">FLASH幻灯片</a></div>
			</td>
			<td width="16%">
				<div align="center"><a href="House_C_Label.asp?type=NorFilt" target="_self">轮换图片幻灯片</a></div>
			</td>
			<td width="15%">
				<div align="center"><a href="House_C_Label.asp?type=infoStat" target="_self">信息统计</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="House_C_Label.asp?type=Search" target="_self">搜索表单</a></div>
			</td>
			<td>
				<div align="center">
			</td>
			<td colspan="2">
				<div align="center"> 栏目：
					<select name="ClassID">
						<option value="Quotation">楼盘信息</option>
						<option value="Second">二手信息</option>
						<option value="Tenancy">租赁信息</option>
						<option value="ToRent">---出租信息</option>
						<option value="Rent">---求租信息</option>
						<option value="ToSell">---出售信息</option>
						<option value="Sell">---求购信息</option>
						<option value="AddRent">---合租信息</option>
						<option value="Transfer">---转让信息</option>
					</select>
				</div>
			</td>
		</tr>
	</table>
	<%
select case Request.QueryString("type")
		case "ReadInfo"
			call ReadInfo()
		case "FlashFilt"
			call FlashFilt()
		case "NorFilt"
			call NorFilt()
		case "siteMap"
			call siteMap()
		case "Search"
			call Search()
		case "infoStat"
			call infoStat()
		case else
			call ReadInfo()
end select
%>
	<%sub ReadInfo()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="21%" class="hback">
				<div align="right">引用样式</div>
			</td>
			<td width="79%" class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
				<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span> </td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示日期格式</div>
			</td>
			<td class="hback">
				<input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
				<span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
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
	var retV = '{FS:HS=ReadInfo┆';
	retV+='引用样式$' + obj.NewsStyle.value + '┆';
	retV+='日期格式$' + obj.DateStyle.value + '┆';
	retV+='栏目$' + obj.ClassID.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
	<%end sub%>
	<%sub OldNews()%>
	<%end sub%>
	<%sub FlashFilt()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">调用数量</div>
			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题字数</div>
			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片尺寸(高度,宽度)</div>
			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				格式120,100。请正确使用格式</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">文本高度</div>
			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">
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
		var retV = '{FS:HS=FlashFilt┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='数量$' + obj.NewsNumber.value + '┆';
		retV+='标题字数$' + obj.TitleNumber.value + '┆';
		retV+='图片尺寸$' + obj.p_size.value +  '┆';
		retV+='文本高度$' + obj.TextSize.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub NorFilt()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">显示标题</div>
			</td>
			<td class="hback">
				<select name="ShowTitle" id="ShowTitle">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">调用数量</div>
			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题字数</div>
			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片尺寸(高度,宽度)</div>
			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				格式120,100。请正确使用格式</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">文本高度</div>
			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>
			</td>
			<td class="hback">
				<input  name="CSS" type="text" id="CSS" size="12">
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
		var retV = '{FS:HS=NorFilt┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='数量$' + obj.NewsNumber.value + '┆';
		retV+='标题字数$' + obj.TitleNumber.value + '┆';
		retV+='图片尺寸$' + obj.p_size.value +  '┆';
		retV+='CSS样式$' + obj.CSS.value +  '┆';
		retV+='文本高度$' + obj.TextSize.value +  '┆';
		retV+='显示标题$' + obj.ShowTitle.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub siteMap()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr style="display:none">
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
		var retV = '{FS:HS=siteMap┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub Search()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">日期搜索</div>
			</td>
			<td width="78%" class="hback">
				<select name="DateShow"  id="DateShow">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示方式</div>
			</td>
			<td class="hback">
				<select name="classShow"  id="classShow">
					<option value="1" selected>横向</option>
					<option value="0">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">文本筐样式</div>
			</td>
			<td class="hback">
				<input type="text" name="TextCss" id="TextCss" maxlength="20">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">下拉菜单样式</div>
			</td>
			<td class="hback">
				<input type="text" name="SelectCss" id="SelectCss" maxlength="20">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">按钮样式</div>
			</td>
			<td class="hback">
				<input type="text" name="ButtonCss" id="ButtonCss" maxlength="20">
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
		var retV = '{FS:HS=Search┆';
		retV+='显示日期$' + obj.DateShow.value + '┆';
		retV+='显示方式$' + obj.classShow.value + '┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='文本筐样式$' + obj.TextCss.value + '┆';
		retV+='下拉菜单样式$' + obj.SelectCss.value + '┆';
		retV+='按钮样式$' + obj.ButtonCss.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub infoStat()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
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
		var retV = '{FS:HS=infoStat┆';
		retV+='显示方向$' + obj.cols.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
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






