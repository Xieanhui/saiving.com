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
	'session�ж�
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
<title>������ǩ����</title>
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
						<td width="20%" class="xingmu"><strong>�����ǩ����</strong></td>
						<td width="80%">
							<div align="right">
								<input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
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
				<div align="center"><a href="House_C_Label.asp?type=ReadInfo" target="_self">������Ϣ���</a></div>
			</td>
			<td width="12%">
				<div align="center"><a href="House_C_Label.asp?type=FlashFilt" target="_self">FLASH�õ�Ƭ</a></div>
			</td>
			<td width="16%">
				<div align="center"><a href="House_C_Label.asp?type=NorFilt" target="_self">�ֻ�ͼƬ�õ�Ƭ</a></div>
			</td>
			<td width="15%">
				<div align="center"><a href="House_C_Label.asp?type=infoStat" target="_self">��Ϣͳ��</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="House_C_Label.asp?type=Search" target="_self">������</a></div>
			</td>
			<td>
				<div align="center">
			</td>
			<td colspan="2">
				<div align="center"> ��Ŀ��
					<select name="ClassID">
						<option value="Quotation">¥����Ϣ</option>
						<option value="Second">������Ϣ</option>
						<option value="Tenancy">������Ϣ</option>
						<option value="ToRent">---������Ϣ</option>
						<option value="Rent">---������Ϣ</option>
						<option value="ToSell">---������Ϣ</option>
						<option value="Sell">---����Ϣ</option>
						<option value="AddRent">---������Ϣ</option>
						<option value="Transfer">---ת����Ϣ</option>
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
				<div align="right">������ʽ</div>
			</td>
			<td width="79%" class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
				<span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ��������ʾ��ʽ</span> </td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾ���ڸ�ʽ</div>
			</td>
			<td class="hback">
				<input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
				<span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:HS=ReadInfo��';
	retV+='������ʽ$' + obj.NewsStyle.value + '��';
	retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
	retV+='��Ŀ$' + obj.ClassID.value + '';
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
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬ�ߴ�(�߶�,���)</div>
			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				��ʽ120,100������ȷʹ�ø�ʽ</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�ı��߶�</div>
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
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:HS=FlashFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value + '';
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
				<div align="right">��ʾ����</div>
			</td>
			<td class="hback">
				<select name="ShowTitle" id="ShowTitle">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬ�ߴ�(�߶�,���)</div>
			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				��ʽ120,100������ȷʹ�ø�ʽ</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�ı��߶�</div>
			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����CSS</div>
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
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:HS=NorFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='CSS��ʽ$' + obj.CSS.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value +  '��';
		retV+='��ʾ����$' + obj.ShowTitle.value + '';
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
				<div align="right">���з�ʽ</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>����</option>
					<option value="1">����</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����CSS</div>
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
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:HS=siteMap��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����CSS$' + obj.Titlecss.value + '';
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
				<div align="right">��������</div>
			</td>
			<td width="78%" class="hback">
				<select name="DateShow"  id="DateShow">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾ��ʽ</div>
			</td>
			<td class="hback">
				<select name="classShow"  id="classShow">
					<option value="1" selected>����</option>
					<option value="0">����</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�ı�����ʽ</div>
			</td>
			<td class="hback">
				<input type="text" name="TextCss" id="TextCss" maxlength="20">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�����˵���ʽ</div>
			</td>
			<td class="hback">
				<input type="text" name="SelectCss" id="SelectCss" maxlength="20">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ť��ʽ</div>
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
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)  
	{
		var retV = '{FS:HS=Search��';
		retV+='��ʾ����$' + obj.DateShow.value + '��';
		retV+='��ʾ��ʽ$' + obj.classShow.value + '��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='�ı�����ʽ$' + obj.TextCss.value + '��';
		retV+='�����˵���ʽ$' + obj.SelectCss.value + '��';
		retV+='��ť��ʽ$' + obj.ButtonCss.value + '';
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
				<div align="right">���з�ʽ</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>����</option>
					<option value="1">����</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:HS=infoStat��';
		retV+='��ʾ����$' + obj.cols.value + '';
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






