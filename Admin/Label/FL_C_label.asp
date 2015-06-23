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
	Dim Conn
	MF_Default_Conn
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
	<title>新闻标签管理</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css"
		rel="stylesheet" type="text/css">
	<base target="self">
</head>
<body class="hback">

	<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>

	<form name="form1" method="post">
	<table width="98%" height="29" border="0" align="center" cellpadding="3" cellspacing="1"
		class="table" valign="absmiddle">
		<tr class="hback">
			<td height="27" align="Left" class="xingmu">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="41%" class="xingmu">
							<strong>常规标签创建</strong>
						</td>
						<td width="59%">
							<div align="right">
								<input name="button4" type="button" onclick="window.returnValue='';window.close();"
									value="关闭">
							</div>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr class="hback">
			<td width="50%">
				<div align="center">
					<a href="FL_C_Label.asp?type=PicFL" target="_self">图片友情连接</a></div>
			</td>
			<td width="50%" height="50%">
				<div align="center">
					<a href="FL_C_Label.asp?type=WordFL" target="_self">文字友情连接</a><a href="News_C_Label.asp?type=OldNews"
						target="_self"></a></div>
			</td>
		</tr>
	</table>
	<%
  dim str_type
  str_type = Request.QueryString("type")
  select case str_type
  		case "PicFL"
			Call PicFL()
		Case "WordFL"
  			Call WordFL()
		Case else
			Call PicFL()
  end select
  Sub PicFL()
	%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="2" class="xingmu">
				图片友情连接
			</td>
		</tr>
		<!-------------------->
		<tr>
			<td width="19%" align="right" class="hback">
				<div align="right">
					选择类别</div>
			</td>
			<td width="81%" class="hback">
				<%
		dim frind_RS
		set frind_RS=Conn.execute("Select ID,F_ClassCName from FS_FL_Class")
		response.Write("<select name=""Select_class_pic"" id=""Select_class_pic"">"&vbcrlf)
		response.Write("<option vaule="""">选择类别</option>"&vbcrlf)
		while not frind_RS.eof
			response.Write("<option value="""&frind_RS("ID")&""">"&frind_RS("F_ClassCName")&"</option>"&vbcrlf)
			frind_RS.movenext
		wend
		response.Write("</select>"&vbcrlf)
		Conn.close()
		Set Conn=nothing
				%>
				<span id="pic" style="color: #FF0000">*请选择友情连接的类别</span>
			</td>
		</tr>
		<!-----------2/7 by chen 友情连接打开方式--------->
		<tr>
			<td width="19%" class="hback">
				<div align="right">
					调用数量</div>
			</td>
			<td width="81%" class="hback">
				<input name="CodeNumber" type="text" id="CodeNumber" value="12" size="10">
				请用正整数&nbsp;&nbsp;&nbsp; 打开方式
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>原窗口</option>
					<option value="1">新窗口</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					站点名称</div>
			</td>
			<td class="hback">
				<select name="ShowTitle" id="ShowTitle">
					<option value="1">显示</option>
					<option value="0" selected>不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					每行数量</div>
			</td>
			<td class="hback">
				<input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
				尺寸
				<input name="NcharCSS" type="text" id="PicSize" value="88,31" size="12">
				(宽,高)
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					输出格式</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="div_style" style="font-family: 宋体; display: none;">
			<td align="right" class="hback">
				DIV CSS
			</td>
			<td class="hback">
				<input name="Divclass" type="text" id="Divclass" size="10" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
			</td>
		</tr>
		<tr>
			<td class="hback">
			</td>
			<td class="hback">
				<input name="button" type="button" onclick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onclick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>

	<script language="JavaScript" type="text/JavaScript">
		function ok(obj) {
			var retV = '{FS:FL=PicFL┆';
			retV += '调用数量$' + obj.CodeNumber.value + '┆';
			retV += '站点名称$' + obj.ShowTitle.value + '┆';
			retV += '每行数量$' + obj.ColsNumber.value + '┆';
			retV += '图片尺寸$' + obj.PicSize.value + '┆';
			retV += '输出格式$' + obj.out_char.value + '┆';
			retV += 'Divclass$' + obj.Divclass.value + '┆';
			retV += '连接类别$' + obj.Select_class_pic.value + '┆';
			retV += '打开窗口$' + obj.Openstyle.value;
			retV += '}';
			window.parent.returnValue = retV;
			window.close();
		}
	</script>

	<%End Sub%>
	<%Sub WordFL()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="2" class="xingmu">
				文字友情连接
			</td>
		</tr>
		<!-------------------->
		<tr>
			<td width="19%" align="right" class="hback">
				<div align="right">
					选择类别</div>
			</td>
			<td width="81%" class="hback">
				<%
		dim frind_RS
		set frind_RS=Conn.execute("Select ID,F_ClassCName from FS_FL_Class")
		response.Write("<select name=""Select_class_word"" id=""Select_class_word"">"&vbcrlf)
		response.Write("<option vaule="""">选择类别</option>"&vbcrlf)
		while not frind_RS.eof
			response.Write("<option value="""&frind_RS("ID")&""">"&frind_RS("F_ClassCName")&"</option>"&vbcrlf)
			frind_RS.movenext
		wend
		response.Write("</select>"&vbcrlf)
		Conn.close()
		Set Conn=nothing
				%><span id="pic" style="color: #FF0000">*请选择友情连接的类别</span>
			</td>
		</tr>
		<!-------------2/7 by chen 友情连接打开方式----------->
		<tr>
			<td width="21%" class="hback">
				<div align="right">
					调用数量</div>
			</td>
			<td width="79%" class="hback">
				<input name="CodeNumber" type="text" id="CodeNumber" value="12" size="10">
				请用正整数&nbsp;&nbsp;&nbsp; 打开方式
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>原窗口</option>
					<option value="1">新窗口</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					每行数量</div>
			</td>
			<td class="hback">
				<input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
				对DIV+CSS框架无效
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					站点显示字数</div>
			</td>
			<td class="hback">
				<input name="leftTitle" type="text" id="leftTitle" value="20" size="10">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					标题CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="10">
				对DIV+CSS框架无效
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					输出格式</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="div_style" style="font-family: 宋体; display: none;">
			<td align="right" class="hback">
				DIV CSS
			</td>
			<td class="hback">
				<input name="Divclass" type="text" id="Text1" size="10" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
			</td>
		</tr>
		<tr>
			<td align="right" class="hback">
			</td>
			<td class="hback">
				<input name="button" type="button" onclick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onclick="window.returnValue='';window.close();"
					value=" 取 消 ">
			</td>
		</tr>
	</table>

	<script language="JavaScript" type="text/JavaScript">
		function ok(obj) {
			var retV = '{FS:FL=WordFL┆';
			retV += '调用数量$' + obj.CodeNumber.value + '┆';
			retV += '每行数量$' + obj.ColsNumber.value + '┆';
			retV += '标题CSS$' + obj.TitleCSS.value + '┆';
			retV += '站点显示字数$' + obj.leftTitle.value + '┆';
			retV += '输出格式$' + obj.out_char.value + '┆';
			retV += 'Divclass$' + obj.Divclass.value + '┆';
			retV += '连接类别$' + obj.Select_class_word.value + '┆';
			retV += '打开窗口$' + obj.Openstyle.value;
			retV += '}';
			window.parent.returnValue = retV;
			window.close();
		}
	</script>

	<%End Sub%>
	</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">
	function SelectClass() {
		var ReturnValue = '', TempArray = new Array();
		ReturnValue = OpenWindow('../News/lib/SelectClassFrame.asp', 400, 300, window);
		try {
			document.getElementById('ClassID').value = ReturnValue[0][0];
			document.getElementById('ClassName').value = ReturnValue[1][0];
		}
		catch (ex) { }
	}
	function selectHtml_express(Html_express) {
		switch (Html_express) {
			case "out_Table":
				document.getElementById('div_style').style.display = 'none';
				document.getElementById('Divclass').disabled = true;
				break;
			case "out_DIV":
				document.getElementById('div_style').style.display = '';
				document.getElementById('Divclass').disabled = false;
				break;
		}
	}
</script>

