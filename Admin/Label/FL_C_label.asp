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
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
	<title>���ű�ǩ����</title>
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
							<strong>�����ǩ����</strong>
						</td>
						<td width="59%">
							<div align="right">
								<input name="button4" type="button" onclick="window.returnValue='';window.close();"
									value="�ر�">
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
					<a href="FL_C_Label.asp?type=PicFL" target="_self">ͼƬ��������</a></div>
			</td>
			<td width="50%" height="50%">
				<div align="center">
					<a href="FL_C_Label.asp?type=WordFL" target="_self">������������</a><a href="News_C_Label.asp?type=OldNews"
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
				ͼƬ��������
			</td>
		</tr>
		<!-------------------->
		<tr>
			<td width="19%" align="right" class="hback">
				<div align="right">
					ѡ�����</div>
			</td>
			<td width="81%" class="hback">
				<%
		dim frind_RS
		set frind_RS=Conn.execute("Select ID,F_ClassCName from FS_FL_Class")
		response.Write("<select name=""Select_class_pic"" id=""Select_class_pic"">"&vbcrlf)
		response.Write("<option vaule="""">ѡ�����</option>"&vbcrlf)
		while not frind_RS.eof
			response.Write("<option value="""&frind_RS("ID")&""">"&frind_RS("F_ClassCName")&"</option>"&vbcrlf)
			frind_RS.movenext
		wend
		response.Write("</select>"&vbcrlf)
		Conn.close()
		Set Conn=nothing
				%>
				<span id="pic" style="color: #FF0000">*��ѡ���������ӵ����</span>
			</td>
		</tr>
		<!-----------2/7 by chen �������Ӵ򿪷�ʽ--------->
		<tr>
			<td width="19%" class="hback">
				<div align="right">
					��������</div>
			</td>
			<td width="81%" class="hback">
				<input name="CodeNumber" type="text" id="CodeNumber" value="12" size="10">
				����������&nbsp;&nbsp;&nbsp; �򿪷�ʽ
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>ԭ����</option>
					<option value="1">�´���</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					վ������</div>
			</td>
			<td class="hback">
				<select name="ShowTitle" id="ShowTitle">
					<option value="1">��ʾ</option>
					<option value="0" selected>����ʾ</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					ÿ������</div>
			</td>
			<td class="hback">
				<input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
				�ߴ�
				<input name="NcharCSS" type="text" id="PicSize" value="88,31" size="12">
				(��,��)
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					�����ʽ</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="div_style" style="font-family: ����; display: none;">
			<td align="right" class="hback">
				DIV CSS
			</td>
			<td class="hback">
				<input name="Divclass" type="text" id="Divclass" size="10" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
			</td>
		</tr>
		<tr>
			<td class="hback">
			</td>
			<td class="hback">
				<input name="button" type="button" onclick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button" type="button" onclick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>

	<script language="JavaScript" type="text/JavaScript">
		function ok(obj) {
			var retV = '{FS:FL=PicFL��';
			retV += '��������$' + obj.CodeNumber.value + '��';
			retV += 'վ������$' + obj.ShowTitle.value + '��';
			retV += 'ÿ������$' + obj.ColsNumber.value + '��';
			retV += 'ͼƬ�ߴ�$' + obj.PicSize.value + '��';
			retV += '�����ʽ$' + obj.out_char.value + '��';
			retV += 'Divclass$' + obj.Divclass.value + '��';
			retV += '�������$' + obj.Select_class_pic.value + '��';
			retV += '�򿪴���$' + obj.Openstyle.value;
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
				������������
			</td>
		</tr>
		<!-------------------->
		<tr>
			<td width="19%" align="right" class="hback">
				<div align="right">
					ѡ�����</div>
			</td>
			<td width="81%" class="hback">
				<%
		dim frind_RS
		set frind_RS=Conn.execute("Select ID,F_ClassCName from FS_FL_Class")
		response.Write("<select name=""Select_class_word"" id=""Select_class_word"">"&vbcrlf)
		response.Write("<option vaule="""">ѡ�����</option>"&vbcrlf)
		while not frind_RS.eof
			response.Write("<option value="""&frind_RS("ID")&""">"&frind_RS("F_ClassCName")&"</option>"&vbcrlf)
			frind_RS.movenext
		wend
		response.Write("</select>"&vbcrlf)
		Conn.close()
		Set Conn=nothing
				%><span id="pic" style="color: #FF0000">*��ѡ���������ӵ����</span>
			</td>
		</tr>
		<!-------------2/7 by chen �������Ӵ򿪷�ʽ----------->
		<tr>
			<td width="21%" class="hback">
				<div align="right">
					��������</div>
			</td>
			<td width="79%" class="hback">
				<input name="CodeNumber" type="text" id="CodeNumber" value="12" size="10">
				����������&nbsp;&nbsp;&nbsp; �򿪷�ʽ
				<select name="Openstyle" id="Openstyle">
					<option value="0" selected>ԭ����</option>
					<option value="1">�´���</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					ÿ������</div>
			</td>
			<td class="hback">
				<input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
				��DIV+CSS�����Ч
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					վ����ʾ����</div>
			</td>
			<td class="hback">
				<input name="leftTitle" type="text" id="leftTitle" value="20" size="10">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					����CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="10">
				��DIV+CSS�����Ч
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">
					�����ʽ</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onchange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
				</select>
			</td>
		</tr>
		<tr class="hback" id="div_style" style="font-family: ����; display: none;">
			<td align="right" class="hback">
				DIV CSS
			</td>
			<td class="hback">
				<input name="Divclass" type="text" id="Text1" size="10" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
			</td>
		</tr>
		<tr>
			<td align="right" class="hback">
			</td>
			<td class="hback">
				<input name="button" type="button" onclick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button" type="button" onclick="window.returnValue='';window.close();"
					value=" ȡ �� ">
			</td>
		</tr>
	</table>

	<script language="JavaScript" type="text/JavaScript">
		function ok(obj) {
			var retV = '{FS:FL=WordFL��';
			retV += '��������$' + obj.CodeNumber.value + '��';
			retV += 'ÿ������$' + obj.ColsNumber.value + '��';
			retV += '����CSS$' + obj.TitleCSS.value + '��';
			retV += 'վ����ʾ����$' + obj.leftTitle.value + '��';
			retV += '�����ʽ$' + obj.out_char.value + '��';
			retV += 'Divclass$' + obj.Divclass.value + '��';
			retV += '�������$' + obj.Select_class_word.value + '��';
			retV += '�򿪴���$' + obj.Openstyle.value;
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

