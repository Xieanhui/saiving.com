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
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<form name="form1" method="post">
	<table width="98%" height="85" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
		<tr class="hback">
			<td height="27" colspan="2" class="xingmu">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="50%" class="xingmu"><strong>���������ű�ǩ����</strong></td>
						<td width="50%" align="right">
							<input name="button4" type="button" onClick="window.top.returnValue='';window.top.close();" value="�ر�">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="hback" >
			<td width="20%" height="27"  align="right" class="hback">
				ѡ�񲻹�������
			</td>
			<td width="80%" align="Left" class="hback">
					<select name="UnId" id="UnId">
						<%
		dim rs,rs1
		set rs = Conn.execute("Select DisTinct UnRegulatedMain From [FS_NS_News_Unrgl] order by UnRegulatedMain DESC")
		do while not rs.eof
				set rs1 = Conn.execute("Select UnregNewsName From FS_NS_News_Unrgl where UnregulatedMain='"&rs("UnRegulatedMain")&"' order by Rows")
				response.Write"<option value="""&rs("UnRegulatedMain")&""">"&rs1("UnregNewsName")&"</option>" 
				rs1.close:set rs1=nothing
			rs.movenext
		loop
		rs.close:set rs=nothing
		%>
					</select>
			</td>
		</tr>
		<tr class="hback">
			<td align="right" class="hback">
				���ű���CSS
			</td>
			<td colspan="3" class="hback" >
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
			</td>
		</tr>
		<tr class="hback">
			<td align="right" class="hback">
				��������(ͼƬ)
			</td>
			<td colspan="3" class="hback" >
				<input name="TitleNavi" type="text" id="TitleNavi">
				��ʹ��html�﷨</td>
		</tr>
		<tr class="hback">
			<td  align="center" class="hback">&nbsp;</td>
			<td colspan="3" class="hback" >
				<label>
					<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
					<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
				</label>
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<script type="text/JavaScript">
	function ok(obj) {
		var retV = '{FS:NS=DefineNews��';
		retV += '������ID$' + obj.UnId.value + '��';
		retV += '����CSS$' + obj.TitleCSS.value + '��';
		retV += '����$' + obj.TitleNavi.value + '';
		retV += '}';
		window.parent.returnValue = retV;
		window.close();
	}

	function SelectClass() {
		var ReturnValue = '', TempArray = new Array();
		ReturnValue = OpenWindow('../News/lib/SelectClassFrame.asp', 400, 300, window);
		try {
			document.getElementById('ClassID').value = ReturnValue[0][0];
			document.getElementById('ClassName').value = ReturnValue[1][0];
		}
		catch (ex) { }
	}

</script>
