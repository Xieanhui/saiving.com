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
	'session�ж�
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
	obj_special_Rs.Open "Select SpecialID,SpecialCName,specialEName from FS_MS_Special  Order by  SpecialID desc",Conn,1,3
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
	Sql = "Select LabelID,LabelName From FS_MF_FreeLabel Where ID > 0 And SysType = '" & NoSqlHack(SysType) & "'"
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
<title>��ǩ����</title>
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
            <td width="20%" class="xingmu"><strong>�����ǩ����</strong></td>
            <td width="80%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="13%" height="15"><div align="center"><a href="Mall_C_Label.asp?type=ReadProducts" target="_self">��Ʒ���</a></div></td>
      <td width="12%"><div align="center"><a href="Mall_C_Label.asp?type=OldNews" target="_self"></a><a href="Mall_C_Label.asp?type=Search" target="_self">������</a></div></td>
      <td width="16%"><div align="center"><a href="Mall_C_Label.asp?type=FlashFilt" target="_self">FLASH�õ�Ƭ</a></div></td>
      <td width="15%"><div align="center"><a href="Mall_C_Label.asp?type=NorFilt" target="_self">�ֻ�ͼƬ�õ�Ƭ</a></div></td>
      <td width="16%"><div align="center"><a href="Mall_C_Label.asp?type=siteMap" target="_self">վ���ͼ</a></div></td>
      <td width="14%"><div align="center"><a href="Mall_C_Label.asp?type=TodayWord" target="_self"></a><a href="Mall_C_Label.asp?type=infoStat" target="_self">��Ϣͳ��</a></div></td>
    </tr>
    <tr class="hback"> 
      <td><div align="center"><a href="Mall_C_Label.asp?type=ClassNavi" target="_self">��Ŀ����</a></div></td>
      <td><div align="center"><a href="Mall_C_Label.asp?type=SpecialNavi" target="_self">ר�⵼��</a></div></td>
      <td><div align="center"><a href="Mall_C_Label.asp?type=RssFeed" target="_self">RSS�ۺ�</a></div></td>
      <td><div align="center"><a href="Mall_C_Label.asp?type=SpecialCode" target="_self">ר�����</a></div></td>
      <td><div align="center"><a href="Mall_C_Label.asp?type=ClassCode" target="_self">��Ŀ����</a></div></td>
      <td><div align="center"><a href="Mall_C_Label.asp?type=ClassInfo" title="�̳���Ŀ��Ϣ����" target="_self">��Ŀ��Ϣ</a></div></td>
    </tr>
	 <tr class="hback"> 
      <td><div align="center"><a href="Mall_C_Label.asp?type=FreeLabel" target="_self">���ɱ�ǩ</a></div></td>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
    </tr>
  </table>
  <%
select case Request.QueryString("type")
		case "ReadProducts"
			call ReadProducts()
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
		Case "ClassInfo"
			Call ClassInfo()
		Case "FreeLabel"
			call FreeLabel()
		case else
			call ReadProducts()
end select
%>
  <%sub ReadProducts()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="21%" class="hback"><div align="right">������ʽ</div></td>
      <td width="79%" class="hback"> <select id="NewsStyle"  name="NewsStyle" style="width:40%">
          <% = label_style_List %>
        </select> <input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴"> 
        <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ��������ʾ��ʽ</span> </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
        <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:MS=ReadProducts��';
	retV+='������ʽ$' + obj.NewsStyle.value + '��';
	retV+='���ڸ�ʽ$' + obj.DateStyle.value + '';
	retV+='}';
	window.parent.returnValue = retV;
	window.close();
}
</script>
<%end sub%>
<%sub OldNews()%>
OldNews
<%end sub%>
<%sub FlashFilt()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName4" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID4"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx"></span>
		<!--------�̳ǻõư�������  by chen 2/2----------------------------->
		<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">��</option>
					<option value="1">��</option>
				</select>
				��������
		<!------------------------------------------------------------------>
				</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" ></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ͼƬ�ߴ�(�߶�,���)</div></td>
      <td class="hback"><input  name="p_size" type="text" id="p_size" value="120,100" size="12">
        ��ʽ120,100������ȷʹ�ø�ʽ</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�ı��߶�</div></td>
      <td class="hback"><input  name="TextSize" type="text" id="Picsize" value="20" size="12"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=FlashFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value + '��';
		retV+='��������$' + obj.containSubClass.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  <%end sub%>
<%sub NorFilt()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName4" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID4"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx"></span>
		<!------------�̳ǰ������� by chen 2/2------------------->
		<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">��</option>
					<option value="1">��</option>
				</select>
				��������
		<!------------------------------------------------------>
		</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><select name="ShowTitle" id="ShowTitle">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ͼƬ�ߴ�(�߶�,���)</div></td>
      <td class="hback"><input  name="p_size" type="text" id="p_size" value="120,100" size="12">
        ��ʽ120,100������ȷʹ�ø�ʽ</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�ı��߶�</div></td>
      <td class="hback"><input  name="TextSize" type="text" id="Picsize" value="20" size="12"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input  name="CSS" type="text" id="CSS" size="12"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=NorFilt��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='����$' + obj.NewsNumber.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='ͼƬ�ߴ�$' + obj.p_size.value +  '��';
		retV+='CSS��ʽ$' + obj.CSS.value +  '��';
		retV+='�ı��߶�$' + obj.TextSize.value +  '��';
		retV+='��ʾ����$' + obj.ShowTitle.value + '��';
		retV+='��������$' + obj.containSubClass.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  <%end sub%>
<%sub siteMap()%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx"></span></td>
    </tr>
    <tr style="display:none"> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input  name="Titlecss" type="text" id="Titlecss" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=siteMap��';
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
      <td width="22%" class="hback"><div align="right">��������</div></td>
      <td width="78%" class="hback"><select name="DateShow"  id="DateShow">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">������Ŀ</div></td>
      <td class="hback"><select name="classShow"  id="classShow">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=Search��';
		retV+='��ʾ����$' + obj.DateShow.value + '��';
		retV+='��ʾ��Ŀ$' + obj.classShow.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
<%end sub%>
<%sub infoStat()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=infoStat��';
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
      <td class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td class="hback"><input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button222" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx"></span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=TodayPic��';
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
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx"></span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����</div></td>
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
	  </label></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input  name="Titlecss" type="text" id="Titlecss" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input  name="TitleNumber" type="text" id="TitleNumber" value="1" size="12" >��
        ��������
        <input  name="lefttitle" type="text" id="lefttitle" value="30" size="12" ></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
      <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=TodayWord��';
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
		retV+='��ʾ����$' + obj.ShowReview.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
  <%end sub%>
<%sub ClassNavi()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
        <span class="tx">�����ѡ����ô��ĳ����͵���ĳ����ĵ���</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>    <tr> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input  name="Titlecss" type="text" id="Titlecss" size="12" ></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���⵼��</div></td>
      <td class="hback"><label>
        <input name="TitleNavi" type="text" id="TitleNavi" value="��">
      ��ʹ��html�﷨</label></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=ClassNavi��';
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
      <td width="22%" class="hback"><div align="right">�����ʽ</div></td>
      <td width="78%" class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>    <tr> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר��CSS</div></td>
      <td class="hback"><input  name="Titlecss" type="text" id="Titlecss" size="12" ></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���⵼��ͼƬ/����</div></td>
      <td class="hback"><label>
        <input name="TitleNavi" type="text" id="TitleNavi" value="��">
      ��ʹ��html�﷨</label></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=SpecialNavi��';
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
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button22" type="button" onClick="SelectClass();" value="ѡ����Ŀ">
        <span class="tx">�����ѡ����ô��ĳ����͵���ĳ�����RSS</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MS=RssFeed��';
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
      <td width="22%" class="hback"><div align="right">ѡ��ר��</div></td>
      <td width="78%" class="hback"> <select id="specialEName"  name="specialEName">
          <option value="">��ѡ��ר��</option>
          <% = label_special_List %>
        </select></td>
    </tr>
    <tr> 
      <td width="22%" class="hback"><div align="right">�����ʽ</div></td>
      <td width="78%" class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>     <tr> 
      <td class="hback"><div align="right">��ʾͼƬ</div></td>
      <td class="hback"><select name="PicTF" id="PicTF">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
        ͼƬ�߶ȼ���� <input name="PicSize" type="text" id="PicSize" value="120,100" size="12"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾר�⵼������</div></td>
      <td class="hback"><select name="NaviTF" id="NaviTF">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
        <input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬCSS</div></td>
      <td class="hback"><input name="PicCSS" type="text" id="PicCSS" size="12"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="12">
        ����CSS
      <input name="ContentCSS" type="text" id="ContentCSS" size="12"></td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select>
        ֻ��table��ʽ��Ч ��������
        <input name="TitleNavi" type="text" id="TitleNavi" value="��"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
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
		var retV = '{FS:MS=SpecialCode��';
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
      <td width="22%" class="hback"><div align="right">ѡ����Ŀ</div></td>
      <td width="78%" class="hback"> <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button223" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        </td>
    </tr>
    <tr> 
      <td width="22%" class="hback"><div align="right">�����ʽ</div></td>
      <td width="78%" class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt; <span class="tx">�������б���ж�λ����ʽ����,ID����</span></td>
    </tr>     <tr> 
      <td class="hback"><div align="right">��ʾͼƬ</div></td>
      <td class="hback"><select name="PicTF" id="PicTF">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
        ͼƬ�߶ȼ���� <input name="PicSize" type="text" id="PicSize" value="120,100" size="12"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬCSS</div></td>
      <td class="hback"><input name="PicCSS" type="text" id="PicCSS" size="12"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ��Ŀ��������</div></td>
      <td class="hback"><select name="NaviTF" id="NaviTF">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
        <input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12"></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="12">
        ����CSS
      <input name="ContentCSS2" type="text" id="ContentCSS" size="12"></td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">���з�ʽ</div></td>
      <td class="hback"><select name="cols"  id="cols">
          <option value="0" selected>����</option>
          <option value="1">����</option>
        </select>
        ����
        <input name="TitleNavi" type="text" id="TitleNavi" value="��">
        ֻ��table��ʽ��Ч</td>
    </tr>
    <tr> 
      <td class="hback"><div align="right"></div></td>
      <td class="hback"> <input name="button2" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ"> 
        <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		//if(obj.ClassID.value=='')
		//{
		//alert('��ѡ����Ŀ');
		//obj.ClassName.focus();
		//return false;
		//}
		var retV = '{FS:MS=ClassCode��';
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
  <!--�̳���Ŀ��Ϣ���ֿ�ʼ-->
  <% Sub ClassInfo() %>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td class="hback">
				<div align="right">��Ŀ��Ϣ��������</div>
			</td>
			<td class="hback">
				<select id="InfoType" name="InfoType">
					<option value="ClassCName" selected>��Ŀ����</option>
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
		var retV = '{FS:MS=ClassInfo��';
		retV+='��������$' + obj.InfoType.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<% End Sub %>
	<!--��Ŀ��Ϣ���ֽ���-->
	<% Sub FreeLabel() %>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	  <tr>
	    <td class="hback">
		  <div align="right">���ɱ�ǩ</div>
	    </td>
	    <td class="hback">
	    	<% = GetNewsFreeList("MS") %>
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
  
</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
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






