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
	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='NS' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	Dim obj_special_Rs,label_special_List
	label_special_List=""
	Set  obj_special_Rs = server.CreateObject(G_FS_RS)
	obj_special_Rs.Open "Select SpecialID,SpecialCName,specialEName from FS_NS_Special  Order by  SpecialID desc",Conn,1,3
	do while Not obj_special_Rs.eof 
		label_special_List = label_special_List&"<option value="""& obj_special_Rs(2)&""">"& obj_special_Rs(1)&"</option>"
		obj_special_Rs.movenext
	loop
	obj_special_Rs.close:set obj_special_Rs = nothing
	
	'================================
	'��ȡ������ϵͳ���ɱ�ǩ�����б�
	'================================
	Function GetNewsFreeList(SysType)
	Dim Rs,Sql
	Sql = "Select LabelID,LabelName From FS_MF_FreeLabel Where ID > 0 And SysType = '" & SysType & "'"
	Set Rs = Conn.ExeCute(Sql)
	GetNewsFreeList = "<select name=""FreeList"" id=""FreeList"">" & vbnewline
	GetNewsFreeList = GetNewsFreeList & "<option value="""">ѡ�����ɱ�ǩ</option>"
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
<title>���ű�ǩ����</title>
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
				<div align="center"><a href="News_C_Label.asp?type=ReadNews" target="_self">�������</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=AllCode" target="_self">ͨ�õ���</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=FlashFilt" target="_self">FLASH�õ�</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=NorFilt" target="_self">�ֻ��õ�</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=siteMap" target="_self">վ���ͼ</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=TodayWord" target="_self">����ͷ��</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=TodayPic" target="_self">ͼƬͷ��</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=Search" target="_self">������</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassNavi" target="_self">��Ŀ����</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SpecialNavi" target="_self">ר�⵼��</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=RssFeed" target="_self">RSS�ۺ�</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SpecialCode" target="_self">ר�����</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassCode" target="_self">��Ŀ����</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=infoStat" target="_self">��Ϣͳ��</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=c_news" target="_self">�������</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=c_ext" target="_self"></a><a href="News_C_Label.asp?type=OldNews" target="_self">�鵵��ǩ</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassInfo" target="_self">��Ŀ��Ϣ</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=FreeLabel" target="_self">���ɱ�ǩ</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SingleClass" target="_self">��ҳ��Ŀ</a></div>
			</td>
			<td>
				<div align="center"></div>
			</td>
			<td>
				<div align="center"></div>
			</td>
			<td>
				<div align="center"></div>
			</td>
			<td>
				<div align="center"></div>
			</td>
			<td>
				<div align="center"></div>
			</td>
		</tr>
	</table>
	<%
select case Request.QueryString("type")
		case "ReadNews"
			call readnews()
		case "OldNews"
			call OldNews()
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
		case "TodayPic"
			call TodayPic()
		case "TodayWord"
			call TodayWord()
		case "ClassNavi"
			call ClassNavi()
		case "SpecialNavi"
			call SpecialNavi()
		case "RssFeed"
			call RssFeed()
		case "SpecialCode"
			call SpecialCode()
		case "ClassCode"
			call ClassCode()
		Case "OldNews"
			call OldNews()
		Case "c_news"
			call c_news()
		Case  "AllCode"
			call AllCode()
		Case "ClassInfo"
			call ClassInfo()
		Case "FreeLabel"
			call FreeLabel()	
		Case "SingleClass"
			Call SingleClass()
		case else
			call readnews()
end select
%>
	<%sub readnews()%>
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
	var retV = '{FS:NS=ReadNews��';
	retV+='������ʽ$' + obj.NewsStyle.value + '��';
	retV+='���ڸ�ʽ$' + obj.DateStyle.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
	<%end sub%>
	<%sub c_news()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<label>
				<select name="ifelse" id="ifelse" onChange="if(this.options[this.selectedIndex].value=='3')ifelsearea.style.display=''; else ifelsearea.style.display='none';">
					<option value="0">���ű������</option>
					<option value="1" selected>���Źؼ������</option>
					<option value="2">�ؼ��ֺͱ������</option>
					<option value="3">�������ĳ���ؼ���</option>
				</select>
				<span id="ifelsearea" style="display:none">
				<select name="ifelsetype" id="ifelsetype">
					<%=PrintOption("=","==:����,&lt;&gt;:������,*s:�Կ�ͷ,*e:�Խ�β,**:��")%>
				</select>
				<input type="text" maxlength="3" name="ifelsevalue" id="ifelsevalue" title="ָ��Դ���ŵĹؼ������.��ؼ���:����,��ׯͼƬ,ָ��1��ʾ����,2��ʾ��ׯͼƬ"
		onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))" onMouseOver="this.focus();">
				</span> </label>
			</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right">��ʾ����</div>
			</td>
			<td width="79%" class="hback">
				<label>
				<input name="titleNumber" type="text" id="titleNumber" value="10">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾ��������</div>
			</td>
			<td class="hback">
				<label>
				<input name="leftTitle" type="text" id="leftTitle" value="40">
				����ռ2���ַ�</label>
			</td>
		</tr>
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
	var value3='';
	if(obj.ifelse.value=='3')
	{
		if (obj.ifelsevalue.value=='')	{alert('��ָ��Դ���ŵĹؼ������.');obj.ifelsevalue.focus();return false;}
		value3 = obj.ifelse.value+''+obj.ifelsetype.value+obj.ifelsevalue.value;
	}
	else
	value3 = obj.ifelse.value;
	
	var retV = '{FS:NS=c_news��';
	retV+='��������$' + value3 + '��';
	retV+='��ʾ����$' + obj.titleNumber.value + '��';
	retV+='��������$' + obj.leftTitle.value + '��';
	retV+='������ʽ$' + obj.NewsStyle.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
	<%end sub%>
	<%sub OldNews()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">&nbsp;</td>
			<td class="hback">�˱�ǩû����</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right"></div>
			</td>
			<td width="79%" class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:NS=OldNews��';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
	<%end sub%>
	<%sub FlashFilt()%>
	    <script language="javascript" type="text/javascript">
        function getColorOptions()
        {
	        var color= new Array("00","33","66","99","CC","FF");
	        for (var i=0;i<color.length ;i++ )
	        {
		        for (var j=0;j<color.length ;j++ )
		        {
			        for (var k=0;k<color.length ;k++ )
			        {
				        if(k==0&&j==0&&i==0)
				        document.write('<option style="background:#'+color[j]+color[k]+color[i]+'" value="#'+color[j]+color[k]+color[i]+'" selected>����</option>');
				        else
				        document.write('<option style="background:#'+color[j]+color[k]+color[i]+'" value="#'+color[j]+color[k]+color[i]+'"></option>');
			        }
		        }
	        }
        }
    </script>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName4" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID4">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx"></span>
				<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">��</option>
					<option value="1">��</option>
				</select>
				��������				</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬ�ߴ�(�߶�,���)</div>			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				��ʽ120,100������ȷʹ�ø�ʽ</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�ı��߶�</div>			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">������ɫ</div></td>
		  <td class="hback"><div align="left"><select name="FlashBG" id="FlashBG" style="width:200px;" class="form"><script language="javascript" type="text/javascript">getColorOptions();</script></select>
		    </div></td>
	  </tr>
		<tr>
			<td class="hback">
				<div align="right"></div>			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=FlashFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value + '��';
		retV+='��������$' + obj.containSubClass.value + '��';
		retV+='������ɫ$' + obj.FlashBG.value +  '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub NorFilt()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName4" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID4">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx"></span>
				<span class="tx"></span>
				<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">��</option>
					<option value="1">��</option>
				</select>
				��������
		  </td>
		</tr>
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
		  <td class="hback"><div align="right">���ڴ򿪷�ʽ</div></td>
		  <td class="hback"><div align="left">
		    <select name="OpenType" id="OpenType">
		      <option value="0">ԭ����</option>
		      <option value="1">�´���</option>
	        </select>
		    </div></td>
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
		var retV = '{FS:NS=NorFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='CSS��ʽ$' + obj.CSS.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value +  '��';
		retV+='��ʾ����$' + obj.ShowTitle.value +  '��';
		retV+='��������$' + obj.containSubClass.value +  '��';
		retV+='���ڴ򿪷�ʽ$' + obj.OpenType.value +  '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub siteMap()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx"></span></td>
		</tr>
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
		var retV = '{FS:NS=siteMap��';
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
				<div align="right">������Ŀ</div>
			</td>
			<td class="hback">
				<select name="classShow"  id="classShow">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾ��ѯ����</div>
			</td>
			<td class="hback">
				<select name="selectShow"  id="selectShow">
					<option value="1" selected>��</option>
					<option value="0">��</option>
				</select>
				  ѡ������ʾ������,����,���ߵȵ������������Ĭ��	
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
		var retV = '{FS:NS=Search��';
		retV+='��ʾ����$' + obj.DateShow.value + '��';
		retV+='��ʾ��Ŀ$' + obj.classShow.value + '��';
		retV+='��ʾ��ѯ����$' + obj.selectShow.value + '';
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
		var retV = '{FS:NS=infoStat��';
		retV+='��ʾ����$' + obj.cols.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub TodayPic()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button222" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx"></span></td>
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
		var retV = '{FS:NS=TodayPic��';
		retV+='��Ŀ$' + obj.ClassID.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub TodayWord()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx"></span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�����ʽ</div>			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
				</select>			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����</div>			</td>
			<td class="hback">
				<select name="cols" id="cols">
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
				</select>
				��ʾ����
				<label>
				<select name="ShowReview" id="ShowReview">
					<option value="1">��ʾ</option>
					<option value="0" selected>����ʾ</option>
				</select>
				</label>			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����CSS</div>			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��������</div>			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="1" size="12" >
				��������
				<input  name="lefttitle" type="text" id="lefttitle" value="30" size="12" >			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">���뷽ʽ</div></td>
		  <td class="hback"><select name="Align" id="Align">
            <option value="left">��</option>
            <option value="center">��</option>
            <option value="right">��</option>
          </select>
		  (������ͨ��ʾ��ʽ��Ч�������DIV������ʽ����)</td>
	  </tr>
		<tr>
			<td class="hback">
				<div align="right"></div>			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� ">			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=TodayWord��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='����$' + obj.cols.value + '��';
		retV+='����CSS$' + obj.Titlecss.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='��������$' + obj.lefttitle.value + '��';
		retV+='��ʾ����$' + obj.ShowReview.value + '��';
		retV+='���뷽ʽ$' + obj.Align.value + '';
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
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx">�����ѡ����ô��ĳ����͵���ĳ����ĵ���</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">�����ʽ</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
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
				<div align="right">����CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">���⵼��</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="��">
		 ���Ҫ����ͼƬ��ֱ����дͼƬ��ַ,ע�ⲻ�ܸ��� IMG��ǩ��</label></td>
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
		var retV = '{FS:NS=ClassNavi��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.cols.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='����CSS$' + obj.Titlecss.value + '��';
		retV+='���⵼��$' + obj.TitleNavi.value + '';
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
				<div align="right">�����ʽ</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
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
				<div align="right">ר��CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">���⵼��ͼƬ/����</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="��">
				��ʹ��html�﷨</label>
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
		var retV = '{FS:NS=SpecialNavi��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='����$' + obj.cols.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='CSS$' + obj.Titlecss.value + '��';
		retV+='����$' + obj.TitleNavi.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub RssFeed()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx">�����ѡ����ô��ĳ����͵���ĳ�����RSS</span></td>
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
		var retV = '{FS:NS=RssFeed��';
		retV+='��Ŀ$' + obj.ClassID.value + '';
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
				<div align="right">ѡ��ר��</div>
			</td>
			<td width="78%" class="hback">
				<select id="specialEName"  name="specialEName">
					<option value="">��ѡ��ר��</option>
					<% = label_special_List %>
				</select>
				<span class="tx">����ѡ��</span></td>
		</tr>
		<tr>
			<td width="22%" class="hback">
				<div align="right">�����ʽ</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾͼƬ</div>
			</td>
			<td class="hback">
				<select name="PicTF" id="PicTF">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
				ͼƬ�߶ȼ����
				<input name="PicSize" type="text" id="PicSize" value="120,100" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾר�⵼������</div>
			</td>
			<td class="hback">
				<select name="NaviTF" id="NaviTF">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
				<input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬCSS</div>
			</td>
			<td class="hback">
				<input name="PicCSS" type="text" id="PicCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
				����CSS
				<input name="ContentCSS" type="text" id="ContentCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">���з�ʽ</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>����</option>
					<option value="1">����</option>
				</select>
				ֻ��table��ʽ��Ч ��������
				<input name="TitleNavi" type="text" id="TitleNavi" value="��">
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
		if(obj.specialEName.value=='')
		{
		alert('��ѡ��ר��');
		obj.specialEName.focus();
		return false;
		}
		var retV = '{FS:NS=SpecialCode��';
		retV+='ר��$' + obj.specialEName.value + '��';
		retV+='ͼƬ��ʾ$' + obj.PicTF.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.PicSize.value + '��';
		retV+='��������$' + obj.NaviTF.value + '��';
		retV+='������������$' + obj.NaviNumber.value + '��';
		retV+='ר������CSS$' + obj.TitleCSS.value + '��';
		retV+='��������CSS$' + obj.ContentCSS.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='���з�ʽ$' + obj.cols.value + '��';
		retV+='����$' + obj.TitleNavi.value + '��';
		retV+='ͼƬcss$' + obj.PicCSS.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub ClassCode()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td width="22%" class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button223" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
				<span class="tx">�����ѡ�����������</span></td>
		</tr>
		<tr>
			<td width="22%" class="hback">
				<div align="right">�����ʽ</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾͼƬ</div>
			</td>
			<td class="hback">
				<select name="PicTF" id="PicTF">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
				ͼƬ�߶ȼ����
				<input name="PicSize" type="text" id="PicSize" value="120,100" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬCSS</div>
			</td>
			<td class="hback">
				<input name="PicCSS" type="text" id="PicCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">��ʾ��Ŀ��������</div>
			</td>
			<td class="hback">
				<select name="NaviTF" id="NaviTF">
					<option value="1" selected>��ʾ</option>
					<option value="0">����ʾ</option>
				</select>
				<input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">����CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
				����CSS
				<input name="ContentCSS2" type="text" id="ContentCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">���з�ʽ</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>����</option>
					<option value="1">����</option>
				</select>
				����
				<input name="TitleNavi" type="text" id="TitleNavi" value="��">
				ֻ��table��ʽ��Ч</td>
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
		var retV = '{FS:NS=ClassCode��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='ͼƬ��ʾ$' + obj.PicTF.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.PicSize.value + '��';
		retV+='��������$' + obj.NaviTF.value + '��';
		retV+='������������$' + obj.NaviNumber.value + '��';
		retV+='��Ŀ����CSS$' + obj.TitleCSS.value + '��';
		retV+='��������CSS$' + obj.ContentCSS.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='���з�ʽ$' + obj.cols.value + '��';
		retV+='����$' + obj.TitleNavi.value + '��';
		retV+='ͼƬCSS$' + obj.PicCSS.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<%sub AllCode()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">&nbsp;</td>
			<td class="hback">˵�����˱�ǩ���Բ����κεط������������Ŀ����ģ�壬�͵��ñ���Ŀ�����ݣ�����������е�����</td>
		</tr>
		<tr>
			<td width="20%" class="hback">
				<div align="right">����</div>
			</td>
			<td width="80%" class="hback">
				<label>
				<select name="NewsType" id="NewsType">
					<option value="HotNews">�ȵ�</option>
					<option value="LastNews">����</option>
					<option value="RecNews">�Ƽ�</option>
					<option value="MarqueeNews">����</option>
					<option value="JcNews">����</option>
				</select>
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������</div>
			</td>
			<td class="hback">
				<label>
				<input name="CodeNumber" type="text" id="CodeNumber" value="10">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������ʾ����</div>
			</td>
			<td class="hback">
				<label>
				<input name="leftNumber" type="text" id="leftNumber" value="40">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������ڶ��ٴ�</div>
			</td>
			<td class="hback">
				<label>
				<input name="ClickNum" type="text" id="ClickNum" value="0">
				����ֻ���ȵ�������Ч�����ѡ��Ϊ0��������</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">ͼƬ����</div>
			</td>
			<td class="hback">
				<label>
				<select name="PicTF" id="PicTF">
					<option value="1">��</option>
					<option value="0" selected>��</option>
				</select>
				ѡ���ǣ���ֻ��ʾͼƬ���š�ѡ�������ʾ����ͼƬ����������</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������������</div>
			</td>
			<td class="hback">
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
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">���ڸ�ʽ</div>
			</td>
			<td class="hback">
				<label>
				<input name="DateStyle" type="text" id="DateStyle" value="YY02��MM��DD��">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������ʽ</div>
			</td>
			<td class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
				<span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ��������ʾ��ʽ</span></td>
		</tr>
		<tr class="hback" >
			<td class="hback" align="center" height="30">
				<div align="right">����</div>
			</td>
			<td height="30"  colspan="2" align="center" class="hback">
				<div align="left"> ������ʾ����ַ�
					<input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
					���ŵ�����ʾ����ַ�
					<input name="navinumber" type="text" id="navinumber" value="200" size="12">
					(����2���ַ�)</div>
			</td>
		</tr>
		<tr>
			<td width="20%" class="hback">
				<div align="right">�����ʽ</div>
			</td>
			<td width="80%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">��ͨ��ʽ</option>
					<option value="out_DIV">DIV+CSS��ʽ</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:����;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV����</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:����;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!">
				&quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
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
		var retV = '{FS:NS=AllCode��';
		retV+='����$' + obj.NewsType.value + '��';
		retV+='������$' + obj.CodeNumber.value + '��';
		retV+='������$' + obj.leftNumber.value + '��';
		retV+='�����$' + obj.ClickNum.value + '��';
		retV+='ͼƬ����$' + obj.PicTF.value + '��';
		retV+='������$' + obj.ColsNumber.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='DivClass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='������ʾ����$' + obj.contentnumber.value + '��';
		retV+='����������ʾ����$' + obj.navinumber.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%end sub%>
	<% Sub ClassInfo() %>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">��������</div>
			</td>
			<td class="hback">
				<select id="InfoType" name="InfoType">
					<option value="ClassName" selected>��Ŀ����</option>
					<option value="Keywords">��Ŀ�ؼ���</option>
					<option value="Description">��Ŀ����</option>
				</select></td>
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
		var retV = '{FS:NS=ClassInfo��';
		retV+='��������$' + obj.InfoType.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<% End Sub %>
	<% Sub FreeLabel() %>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	  <tr>
	    <td class="hback">
		  <div align="right">���ɱ�ǩ</div>
	    </td>
	    <td class="hback">
	    	<% = GetNewsFreeList("NS") %>
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
		var retV = '{FS:MF=FreeLabel��';
		retV+='���ɱ�ǩ$' + obj.FreeList.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<% End Sub %>
	<% Sub SingleClass() %>
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
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=ClassPage��';
		retV+='������ʽ$' + obj.NewsStyle.value ;
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
				document.getElementById('div_id').style.display = 'none';
				document.getElementById('li_id').style.display = 'none';
				document.getElementById('ul_id').style.display = 'none';
				document.getElementById('DivID').disabled = true;
				break;
			case "out_DIV":
				document.getElementById('div_id').style.display = '';
				document.getElementById('li_id').style.display = '';
				document.getElementById('ul_id').style.display = '';
				document.getElementById('DivID').disabled = false;
				break;
		}
	}
</script>