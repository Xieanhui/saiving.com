<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
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
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	MF_Default_Conn
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
	'session�ж�
	MF_Session_TF 
	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='SD' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	
	Dim sRootDir,str_CurrPath
	
	if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
	If Temp_Admin_Is_Super = 1 then
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	Else
		If Temp_Admin_FilesTF = 0 Then
			str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
		Else
			str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
		End If	
	End if
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
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="41%" class="xingmu"><strong>�����ǩ����</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback">
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDList" target="_self">��Ŀ��Ϣ</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=AreaList" target="_self">�������</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDClass" target="_self">��Ŀ����</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDClassList" target="_self">�ռ���Ŀ</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDAreaList" target="_self">�ռ�����</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDPage" target="_self">���ҳ��</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDSearch" target="_self">��������</a></div></td>
    </tr>
	 <tr class="hback">
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDPubTypeList" target="_self">�ռ����</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDChildClass" target="_self">�������</a></div></td>
      <td width="14%"><div align="center"></div></td>
      <td width="14%"><div align="center"></div></td>
      <td width="14%"><div align="center"></div></td>
      <td width="14%"><div align="center"></div></td>
      <td width="14%"><div align="center"></div></td>
    </tr>
  </table>
  <%
	dim str_type
	str_type = Request.QueryString("type")
	select case str_type
		case "SDList"
			Call SDList()
		case "AreaList"
			Call AreaList()
		case "SDClass"
			Call SDClass()
		case "SDClassList"
			Call SDClassList()
		case "SDAreaList"
			Call SDAreaList()
		Case "SDPage"
			Call SDPage()
		Case "SDSearch"
			Call SDSearch()
		Case "SDPubTypeList"
			Call SDPubTypeList()
		Case "SDChildClass"
			Call SDChildClass()		
		case else
			Call SDList()
	End select
  %>
  <%Sub SDList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ŀ�б�</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">����</div></td>
      <td width="81%" class="hback">
        <select name="PubType" id="PubType">
			  <option value="" selected>ȫ��</option>
			  <option value="0">��Ӧ</option>
			  <option value="1">��</option>
			  <option value="2">����</option>
			  <option value="3">����</option>
			  <option value="4">����</option>
        </select>
		<select name="PubPop" id="PubPop">
			<option value="0">����</option>
			<option value="1">�Ƽ�</option>
			<option value="2">����</option>
       </select>
       <input type="button" name="Submit" value="ѡ�������滻ͼƬ����ʽ" onClick="showhide(document.getElementById('typeTR'));">
      </td>
    </tr>
    <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">��ʽ��ֱ����д�Լ������css��ʽ����ͼƬ����ѡ��ͼƬ��Ҳ����ֱ����д��ɫ���룬��:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">��Ӧ����ͼƬ����ʽ</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�󹺴���ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
<!--�����������--->
<!----->	
<!--������Ŀ����--->
<!----->	

    <tr>
      <td class="hback"><div align="right">ÿ����ʽ</div></td>
      <td class="hback">
	   ����<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   ż��<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		ֻ��Ա�����,��ֱ������ɫ#FF0000	  </td>
    </tr>
	
    <tr class="hback" id="ClassName_col"> 
      <td class="hback"  align="center"><div align="right">ѡ����Ŀ</div></td>
      <td colspan="2" class="hback" > <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button2" type="button" onClick="SelectClass();" value="ѡ����Ŀ"> 
        <span class="tx">ѡ����Ŀ�������ѡ������ʾ���е�</span></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback">
	  	<select name="ChildTF" id="ChildTF">
			<option value="1" selected>��</option>
			<option value="0">��</option>
		</select>
	  </td>
    </tr>
	<tr> 
      <td class="hback"><div align="right">ʱ�䷶Χ</div></td>
      <td class="hback"><input name="DayNum" type="text" id="DayNum" value="0">�����ڵ���Ϣ��0Ϊ���ޡ�</td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  <input type="button" name="Submit" value="����������Ϣnew���" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">���ֲ���Ϊ�գ�����ʾ��������Ϊ0����������Ϊ0ȡ�������</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">�������</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			����<input name="NewDayNum" type="text" id="NewDayNum" value="1">
			�����ڵ���Ϣ�������ʾnew���			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">�������</td>
            <td height="24" align="left" valign="middle" class="hback">
			����<input name="NewInfoNum" type="text" id="NewInfoNum" value="10">
			����Ϣ�������ʾnew���</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">���ͼƬ</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <label>
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
        </label>
      ��ʽ���е���������������Ч</div></td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		
		var retV = '{FS:SD=SDList��';
		retV+='��������$' + obj.PubType.value + '��';
		retV+='��������$' + obj.PubPop.value + '��';
		retV+='������ʽ��ͼƬ$' + PubType_Style + '��';
		retV+='��������ʽ$' + obj.PubType_JiTR.value + '��';
		retV+='ż������ʽ$' + obj.PubType_OuTR.value + '��';
		retV+='��Ŀ$' + obj.ClassID.value + '��';
		retV+='��������$' + obj.ChildTF.value + '��';
		retV+='ʱ�䷶Χ$' + obj.DayNum.value + '��';
		retV+='��������$' + obj.TitleNumber.value + '��';
		retV+='�������$' + obj.NewDayNum.value + '��';
		retV+='�������$' + obj.NewInfoNum.value + '��';
		retV+='���ͼƬ$' + obj.NewpicUrl.value + '��';
		retV+='��ʾ����$' + obj.ColsNumber.value + '��';
		retV+='��������$' + obj.leftTitle.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='��������$' + obj.ContentNumber.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub AreaList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�����б�</td>
    </tr>
	<!-- Ken 2006-12-25 -->
     <tr> 
      <td width="19%" class="hback" align="right">��ʾ��ʽ</td>
      <td width="81%" class="hback">
	  <select name="DisType" id="DisType" onChange="DisTypeTF(this.options[this.selectedIndex].value)">
		<option value="0" selected="selected">����Ӧ������ʾ</option>
	  	<option value="1">����������ʾ</option>
	  </select></td>
    </tr>
	<tr style="display:;" id="ShowInfoTR"> 
      <td colspan="2" height="20" align="center" class="hback">
	  <span id="ShowInfo"><font color=red>�˷�ʽָ�Զ�ʶ��ǰ����Ȼ����ʾ����һ����������.ѡ��˷�ʽ,�˱�ǩֻ�ܷ��ڹ����������ģ����,����������ʾ����</font></span>
	  </td>
    </tr>
	<tr id="FG_RE" style="display:;">
      <td width="19%" class="hback" align="right">�ָ����</td>
      <td width="81%" class="hback">
	  <input type="text" name="FG_Info" id="FG_Info" value="" maxlength="20">
	  �벻Ҫ����html����	  </td>
    </tr>
	<tr id="Class_ROW_Num" style="display:;"> 
      <td width="19%" class="hback" align="right">��ʾ����</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassRow" id="ClassRow" value="1" maxlength="20">
	  </td>
    </tr>
	<tr id="ClassLive_TF" style="display:none;"> 
      <td width="19%" class="hback" align="right">������</td>
      <td width="81%" class="hback">
	  <input type="text" name="DisClassLive" id="DisClassLive" value="0" maxlength="20">
	  0Ϊ����,������Ϊ1��,�Դ�����
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">������ʽ</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassStyle" id="ClassStyle" value="" maxlength="20">
	  ��������ʽ��������ɫ����,��ɫ��������#��ͷ	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">�����о�</td>
      <td width="81%" class="hback">
	  <input type="text" name="TRHeight" id="TRHeight" value="20" maxlength="20">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">����ͼƬ������</td>
      <td width="81%" class="hback">
	  <input type="text" name="NaviPic" id="NaviPic" value="" size=30 maxlength="50">
	  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�" onClick="SelectFile();">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">��ʾ����</td>
      <td width="81%" class="hback">
	  <select name="DisNumTF" id="DisNumTF">
	  	<option value="1" selected="selected">��</option>
	  	<option value="0">��</option>
	  </select>
	  �Ƿ���ʾ����������Ϣ����	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">������ʽ</td>
      <td width="81%" class="hback">
	  <input type="text" name="InfoNumCss" id="InfoNumCss" value="" maxlength="20">
	  ��������ʽ��������ɫ����,��ɫ��������#��ͷ	  </td>
    </tr>
    <tr> 
      <td width="19%" class="hback" align="right">��������</td>
      <td width="81%" class="hback">
	  <select name="OpenMode" id="OpenModes">
	  	<option value="1" selected="selected">��</option>
	  	<option value="0">��</option>
	  </select>
	  �Ƿ����´����д�����</td>
    </tr>
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function DisTypeTF(str)
	{
		if(str==0)
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = '';
			document.getElementById('Class_ROW_Num').style.display = '';
			document.getElementById('ClassLive_TF').style.display = 'none';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>�˷�ʽָ�Զ�ʶ��ǰ����Ȼ����ʾ����һ����������.ѡ��˷�ʽ,�˱�ǩֻ�ܷ��ڹ����������ģ����,����������ʾ����</font>';
		}
		else
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = 'none';
			document.getElementById('Class_ROW_Num').style.display = 'none';
			document.getElementById('ClassLive_TF').style.display = '';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>�˷�ʽ��������״��ʽ��ʾ��������</font>';
		}
	}	
	function ok(obj)
	{
		if(isNaN(obj.ClassRow.value))
		{alert('��ʾ�����������֡�');obj.ClassRow.focus();return false;}
		if(isNaN(obj.TRHeight.value))
		{alert('�о�������֡�');obj.TRHeight.focus();return false;}
		if(obj.FG_Info.value.indexOf('��')>-1)
		{alert('�ָ���Ų���ʹ���ض����ũ�');obj.FG_Info.focus();return false;}
		if(obj.NaviPic.value.indexOf('��')>-1)
		{alert('�������ݲ���ʹ��������ũ�');obj.NaviPic.focus();return false;} 
		//-------------------------------------------------------------
		var retV = '{FS:SD=AreaList��';
		retV+='��ʾ��ʽ$' + obj.DisType.value + '��';
		retV+='�ָ�����$' + obj.FG_Info.value + '��';
		retV+='��ʾ����$' + obj.ClassRow.value + '��';
		retV+='������ʽ$' + obj.ClassStyle.value + '��';
		retV+='�о�$' + obj.TRHeight.value + '��';
		retV+='����$' + obj.NaviPic.value + '��';
		retV+='��ʾ����$' + obj.DisNumTF.value + '��';
		retV+='������ʽ$' + obj.InfoNumCss.value + '��';
		retV+='��������$' + obj.OpenMode.value+ '��';
		retV+='������$' + obj.DisClassLive.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDClass()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ŀ�б�</td>
    </tr>
	<!-- Ken 2006-12-25 -->
     <tr> 
      <td width="19%" class="hback" align="right">��ʾ��ʽ</td>
      <td width="81%" class="hback">
	  <select name="DisType" id="DisType" onChange="DisTypeTF(this.options[this.selectedIndex].value)">
		<option value="0" selected="selected">����Ӧ��Ŀ��ʾ</option>
	  	<option value="1">������Ŀ��ʾ</option>
	  </select></td>
    </tr>
	<tr style="display:;" id="ShowInfoTR"> 
      <td colspan="2" height="20" align="center" class="hback">
	  <span id="ShowInfo"><font color=red>�˷�ʽָ�Զ�ʶ��ǰ��Ŀ��Ȼ����ʾ����һ��������Ŀ.ѡ��˷�ʽ,�˱�ǩֻ�ܷ��ڹ�����Ŀģ����,����������ʾ����</font></span>
	  </td>
    </tr>
	<tr id="FG_RE" style="display:;">
      <td width="19%" class="hback" align="right">�ָ����</td>
      <td width="81%" class="hback">
	  <input type="text" name="FG_Info" id="FG_Info" value="" maxlength="20">
	  �벻Ҫ����html����	  </td>
    </tr>
	<tr id="Class_ROW_Num" style="display:;"> 
      <td width="19%" class="hback" align="right">��ʾ����</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassRow" id="ClassRow" value="1" maxlength="20">
	  </td>
    </tr>
	<tr id="ClassLive_TF" style="display:none;"> 
      <td width="19%" class="hback" align="right">��Ŀ����</td>
      <td width="81%" class="hback">
	  <input type="text" name="DisClassLive" id="DisClassLive" value="0" maxlength="20">
	  0Ϊ����,����ĿΪ1��,�Դ�����
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">��Ŀ��ʽ</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassStyle" id="ClassStyle" value="" maxlength="20">
	  ��������ʽ��������ɫ����,��ɫ��������#��ͷ	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">�����о�</td>
      <td width="81%" class="hback">
	  <input type="text" name="TRHeight" id="TRHeight" value="20" maxlength="20">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">����ͼƬ������</td>
      <td width="81%" class="hback">
	  <input type="text" name="NaviPic" id="NaviPic" value="" size=30 maxlength="50">
	  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�" onClick="SelectFile();">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">��ʾ����</td>
      <td width="81%" class="hback">
	  <select name="DisNumTF" id="DisNumTF">
	  	<option value="1" selected="selected">��</option>
	  	<option value="0">��</option>
	  </select>
	  �Ƿ���ʾ����Ŀ����Ϣ����	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">������ʽ</td>
      <td width="81%" class="hback">
	  <input type="text" name="InfoNumCss" id="InfoNumCss" value="" maxlength="20">
	  ��������ʽ��������ɫ����,��ɫ��������#��ͷ	  </td>
    </tr>
    <tr> 
      <td width="19%" class="hback" align="right">��������</td>
      <td width="81%" class="hback">
	  <select name="OpenMode" id="OpenModes">
	  	<option value="1" selected="selected">��</option>
	  	<option value="0">��</option>
	  </select>
	  �Ƿ����´����д�����	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function DisTypeTF(str)
	{
		if(str==0)
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = '';
			document.getElementById('Class_ROW_Num').style.display = '';
			document.getElementById('ClassLive_TF').style.display = 'none';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>�˷�ʽָ�Զ�ʶ��ǰ��Ŀ��Ȼ����ʾ����һ��������Ŀ.ѡ��˷�ʽ,�˱�ǩֻ�ܷ��ڹ�����Ŀģ����,����������ʾ����</font>';
		}
		else
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = 'none';
			document.getElementById('Class_ROW_Num').style.display = 'none';
			document.getElementById('ClassLive_TF').style.display = '';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>�˷�ʽ��������״��ʽ��ʾ������Ŀ</font>';
		}
	}	
	function ok(obj)
	{
		if(isNaN(obj.ClassRow.value))
		{alert('��ʾ�����������֡�');obj.ClassRow.focus();return false;}
		if(isNaN(obj.TRHeight.value))
		{alert('�о�������֡�');obj.TRHeight.focus();return false;}
		if(obj.FG_Info.value.indexOf('��')>-1)
		{alert('�ָ���Ų���ʹ���ض����ũ�');obj.FG_Info.focus();return false;}
		if(obj.NaviPic.value.indexOf('��')>-1)
		{alert('�������ݲ���ʹ��������ũ�');obj.NaviPic.focus();return false;} 
		//-------------------------------------------------------------
		var retV = '{FS:SD=SDClass��';
		retV+='��ʾ��ʽ$' + obj.DisType.value + '��';
		retV+='�ָ�����$' + obj.FG_Info.value + '��';
		retV+='��ʾ����$' + obj.ClassRow.value + '��';
		retV+='��Ŀ��ʽ$' + obj.ClassStyle.value + '��';
		retV+='�о�$' + obj.TRHeight.value + '��';
		retV+='����$' + obj.NaviPic.value + '��';
		retV+='��ʾ����$' + obj.DisNumTF.value + '��';
		retV+='������ʽ$' + obj.InfoNumCss.value + '��';
		retV+='��������$' + obj.OpenMode.value+ '��';
		retV+='��Ŀ����$' + obj.DisClassLive.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
  <%Sub SDAreaList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�����ռ�����</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">��ҳ����</div></td>
      <td width="81%" class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��Ϣ��������</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч 
	  <input type="button" name="Submit" value="ѡ�������滻ͼƬ����ʽ" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">��ʽ��ֱ����д�Լ������css��ʽ����ͼƬ����ѡ��ͼƬ��Ҳ����ֱ����д��ɫ���룬��:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">��Ӧ����ͼƬ����ʽ</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�󹺴���ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <label>
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
        </label>
      ��ʽ���е���������������Ч</div></td>
    </tr>
     <tr>
      <td class="hback"  align="center"><div align="right">��ҳ��ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">��ʽ1</option>
			  <option value="2">��ʽ2</option>
			  <option value="3">��ʽ3</option>
			  <option value="4" selected="selected">��ʽ4</option>
         </select>
		 <span class="tx">��ҳ������ʽ</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">��Ϊ��ɫ����,����Ҫ#����</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDAreaList��';
		retV+='����$' + obj.TitleNumber.value + '��';
		retV+='����$' + obj.ColsNumber.value + '��';
		retV+='����$' + obj.leftTitle.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='��������$' + obj.ContentNumber.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value + '��';
		retV+='��ҳ��ʽ$' + obj.Flg_PageType.value + '��';
		retV+='��ҳ������ʽ$' + obj.Flg_PageCss.value + '��';
		retV+='������ʽ����ͼƬ$' + PubType_Style;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDClassList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ŀ�ռ�����</td>
    </tr>
    <tr>
      <td width="19%" class="hback">
        <div align="right">��ҳ����</div>
      </td>
      <td width="81%" class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��Ϣ��������</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч 
	  <input type="button" name="Submit" value="ѡ�������滻ͼƬ����ʽ" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">��ʽ��ֱ����д�Լ������css��ʽ����ͼƬ����ѡ��ͼƬ��Ҳ����ֱ����д��ɫ���룬��:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">��Ӧ����ͼƬ����ʽ</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�󹺴���ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <label>
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
        </label>
      ��ʽ���е���������������Ч</div></td>
    </tr>
     <tr>
      <td class="hback"  align="center"><div align="right">��ҳ��ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">��ʽ1</option>
			  <option value="2">��ʽ2</option>
			  <option value="3">��ʽ3</option>
			  <option value="4" selected="selected">��ʽ4</option>
         </select>
		 <span class="tx">��ҳ������ʽ</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">��Ϊ��ɫ����,����Ҫ#����</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDClassList��';
		retV+='����$' + obj.TitleNumber.value + '��';
		retV+='����$' + obj.ColsNumber.value + '��';
		retV+='����$' + obj.leftTitle.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='��������$' + obj.ContentNumber.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value + '��';
		retV+='��ҳ��ʽ$' + obj.Flg_PageType.value + '��';
		retV+='��ҳ������ʽ$' + obj.Flg_PageCss.value + '��';
		retV+='������ʽ����ͼƬ$' + PubType_Style;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDPage()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�ռ�ҳ�����</td>
    </tr>    
    <tr>
      <td width="19%"  align="center" class="hback"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td width="81%" class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SD=SDPage��';
		retV+='������ʽ$' + obj.NewsStyle.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDSearch()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">����</td>
    </tr>    
    
    <tr>
      <td class="hback"><div align="right">��ʾ��Ŀ</div></td>
      <td class="hback"><label>
        <select name="ShowClass" id="ShowClass">
          <option value="1" selected>��ʾ</option>
          <option value="0">����ʾ</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><select name="ShowArea" id="ShowArea">
        <option value="1" selected>��ʾ</option>
        <option value="0">����ʾ</option>
      </select></td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right"></div></td>
      <td width="81%" class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SD=SDSearch��';
		retV+='��ʾ��Ŀ$' + obj.ShowClass.value + '��';
		retV+='��ʾ����$' + obj.ShowArea.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
<%
'-----2006-12-19 by ken
Sub SDPubTypeList()
%> 
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">���������ռ��б�</td>
    </tr>
    <tr>
      <td width="19%" class="hback">
        <div align="right">��ҳ����</div>
      </td>
      <td width="81%" class="hback">
	  <input name="PageInfoNum" type="text" id="PageInfoNum" value="10">
	  <input type="button" name="Submit" value="����������Ϣnew���" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">���ֲ���Ϊ�գ�����ʾ��������Ϊ0����������Ϊ0ȡ�������</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">�������</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			����<input name="NewDayNum" type="text" id="NewDayNum" value="1">�����ڵ���Ϣ�������ʾnew���			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">�������</td>
            <td height="24" align="left" valign="middle" class="hback">
			����<input name="NewInfoNum" type="text" id="NewInfoNum" value="10">
			����Ϣ�������ʾnew���</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">���ͼƬ</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��������ʽ</div></td>
      <td class="hback">
	   <input type="text" name="PubType_JiTR" size="20" maxlength="100" value="">
		DIV��Ч,��Ϊ��ɫ#FF0000����ʽ����ͼƬ·��	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">ż������ʽ</div></td>
      <td class="hback">
	   <input type="text" name="PubType_OuTR" size="20" maxlength="100" value="">
		DIV��Ч,��Ϊ��ɫ#FF0000����ʽ����ͼƬ·��	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��Ϣ��������</div></td>
      <td class="hback"><input name="TitleNum" type="text" id="TitleNum" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч 
	  <input type="button" name="Submit" value="ѡ�������滻ͼƬ����ʽ" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">��ʽ��ֱ����д�Լ������css��ʽ����ͼƬ����ѡ��ͼƬ��Ҳ����ֱ����д��ɫ���룬��:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">��Ӧ����ͼƬ����ʽ</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�󹺴���ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <label>
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
        </label>
      ��ʽ���е���������������Ч</div></td>
    </tr>
     <tr>
      <td class="hback"  align="center"><div align="right">��ҳ��ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">��ʽ1</option>
			  <option value="2">��ʽ2</option>
			  <option value="3">��ʽ3</option>
			  <option value="4" selected="selected">��ʽ4</option>
         </select>
		 <span class="tx">��ҳ������ʽ</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">��Ϊ��ɫ����,����Ҫ#����</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDPubTypeList��';
		retV+='��ҳ����$' + obj.PageInfoNum.value + '��'; 
		retV+='���±��$' + obj.NewDayNum.value + ','+ obj.NewInfoNum.value + ','+ obj.NewpicUrl.value + '��';
		retV+='��ż����ʽ$' + obj.PubType_JiTR.value + ','+ obj.PubType_OuTR.value + '��';
		retV+='��������$' + obj.TitleNum.value + '��';
		retV+='��ʾ����$' + obj.ColsNumber.value + '��';
		retV+='�����滻��ʽ$' + PubType_Style + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='Div��ʽ$' + obj.DivID.value + ','+ obj.Divclass.value + ','+ obj.ulid.value + ','+ obj.ulclass.value + ','+ obj.liid.value + ','+ obj.liclass.value + '��';
		retV+='��������$' + obj.ContentNumber.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value + '��';
		retV+='��ҳ��ʽ$' + obj.Flg_PageType.value + '��';
		retV+='��ҳ������ʽ$' + obj.Flg_PageCss.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script> 

<% End Sub 
Sub SDChildClass()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�������</td>
    </tr>
    <tr>
      <td width="20%" class="hback"><div align="right">���෽ʽ</div></td>
      <td width="81%" class="hback">
        <select name="TypeRuler" id="TypeRuler">
			  <option value="0">����Ŀ����</option>
			  <option value="1">���������</option>
       </select>
      </td>
    </tr>
	<tr>
      <td colspan="2" align="center" valign="middle" class="hback"><span style="color:red">ע��:���ѡ����Ŀ����,�뽫�˱�ǩ���ڹ�����Ŀģ����;���ѡ���������,�뽫�˱�ǩ���빩���������ģ����,����ģ�岻Ҫ�ô˱�ǩ�������������ʾ����.</span></td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��ʾ�ӷ�����</div></td>
      <td class="hback">
	  	<select name="ChCNameDisTF" id="ChCNameDisTF">
			<option value="1" selected>��</option>
			<option value="0">��</option>
		</select>
		�ӷ�������ʾ��ʽ
		<input name="ChCNameCss" type="text" id="ChCNameCss" value="">
	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">������Ϣ��������</div></td>
      <td class="hback"><input name="InfoNum" type="text" id="InfoNum" value="10">
	  <input type="button" name="Submit" value="����������Ϣnew���" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">���ֲ���Ϊ�գ�����ʾ��������Ϊ0����������Ϊ0ȡ�������</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">�������</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			����<input name="NewDayNum" type="text" id="NewDayNum" value="1">�����ڵ���Ϣ�������ʾnew���			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">�������</td>
            <td height="24" align="left" valign="middle" class="hback">
			����<input name="NewInfoNum" type="text" id="NewInfoNum" value="3">
			����Ϣ�������ʾnew���</td>
          </tr> 
          <tr>
            <td height="24" align="right" valign="middle" class="hback">���ͼƬ</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr> 
      <td class="hback"><div align="right">ʱ�䷶Χ</div></td>
      <td class="hback"><input name="DayNum" type="text" id="DayNum" value="0">�����ڵ���Ϣ��0Ϊ���ޡ�</td>
    </tr>
	<tr>
      <td width="20%" class="hback"><div align="right">��Ϣ����</div></td>
      <td width="81%" class="hback">
        <select name="PubType" id="PubType">
			  <option value="" selected>ȫ��</option>
			  <option value="0">��Ӧ</option>
			  <option value="1">��</option>
			  <option value="2">����</option>
			  <option value="3">����</option>
			  <option value="4">����</option>
        </select>
		<select name="PubPop" id="PubPop">
			<option value="0">����</option>
			<option value="1">�Ƽ�</option>
			<option value="2">����</option>
       </select>
       <input type="button" name="Submit" value="ѡ�������滻ͼƬ����ʽ" onClick="showhide(document.getElementById('TypeTR'));">
      </td>
    </tr>
    <tr id="TypeTR" style="display:none">
      <td class="hback" colspan="2">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">��ʽ��ֱ����д�Լ������css��ʽ����ͼƬ����ѡ��ͼƬ��Ҳ����ֱ����д��ɫ���룬��:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">��Ӧ����ͼƬ����ʽ</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�󹺴���ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">�������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">��������ͼƬ����ʽ</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input name="TitleNum" type="text" id="TitleNum" value="30"></td>
    </tr>
	<tr>
      <td class="hback"><div align="right">ÿ����ʽ</div></td>
      <td class="hback">
	   ����<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   ż��<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		ֻ��Ա�����,��ֱ������ɫ#FF0000	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">������ʾ����</div></td>
      <td class="hback"><input name="ClassRowNum" type="text" id="ClassRowNum" value="1" size="10">
      ��DIV+CSS�����Ч </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">��Ϣ��ʾ����</div></td>
      <td class="hback"><input name="RowNum" type="text" id="RowNum" value="1" size="10">
      ��DIV+CSS�����Ч </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ڸ�ʽ</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">��ʽ:YY02����2λ�����(��06��ʾ2006��),YY04��ʾ4λ�������(2006)��MM�����£�DD�����գ�HH����Сʱ��MI����֣�SS������</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,1);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr id="Table_ID" style="display:;">
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback"><input name="TableCss" type="text" id="TableCss" value="">
      </td>
    </tr>
	<tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">������������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
      ��ʽ���е���������������Ч</div></td>
    </tr>
	 <tr>
      <td class="hback"  align="center"><div align="right">��������</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <input name="MoreLinkStr" type="text" id="MoreLinkStr" value="����..>>">
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
		����������ʽ
		<input name="MoreLinkCss" type="text" id="MoreLinkCss" value="">
      </div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback">
	  	<select name="OpenMode" id="OpenMode">
			<option value="1" selected>��</option>
			<option value="0">��</option>
		</select>
	  </td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	//----------------------------------
	function ok(obj)
	{
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//-------
		var NewsInfo_BJ_Str = '';              
		NewsInfo_BJ_Str = obj.NewDayNum.value + ',' + obj.NewInfoNum.value + ',' + obj.NewpicUrl.value;
		var InfoType_Str = '';
		InfoType_Str = obj.PubType.value + ',' + obj.PubPop.value;
		var TR_Style_Str = '';
		TR_Style_Str = obj.PubType_JiTR.value + ',' + obj.PubType_OuTR.value;
		var DIV_Style_Str = '';
		DIV_Style_Str = obj.DivID.value + ',' + obj.Divclass.value + ',' + obj.ulid.value + ',' + obj.ulclass.value + ',' + obj.liid.value + ',' + obj.liclass.value;
		//----------
		var retV = '{FS:SD=SDChildClass��';
		retV+='���෽ʽ$' + obj.TypeRuler.value + '��';
		retV+='��ʾ��������$' + obj.ChCNameDisTF.value + '��';  
		retV+='����������ʽ$' + obj.ChCNameCss.value + '��';
		retV+='������Ϣ����$' + obj.InfoNum.value + '��';  
		retV+='������Ϣ���$' + NewsInfo_BJ_Str + '��';
		retV+='ʱ�䷶Χ$' + obj.DayNum.value + '��';  
		retV+='��Ϣ����$' + InfoType_Str + '��'; 
		retV+='������ʽ��ͼƬ$' + PubType_Style + '��';
		retV+='��������$' + obj.TitleNum.value + '��';
		retV+='��ż����ʽ$' + TR_Style_Str + '��';  
		retV+='��ʾ����$' + obj.RowNum.value + '��';
		retV+='���ڸ�ʽ$' + obj.DateStyle.value + '��';
		retV+='�����ʽ$' + obj.out_char.value + '��';
		retV+='Div��ʽ$' + DIV_Style_Str + '��';  
		retV+='������������$' + obj.ContentNumber.value + '��'; 
		retV+='��������$' + obj.MoreLinkStr.value + '��';
		retV+='����������ʽ$' + obj.MoreLinkCss.value + '��';
		retV+='������ʽ$' + obj.NewsStyle.value + '��';
		retV+='��������$' + obj.OpenMode.value + '��';     
		retV+='��������$' + obj.ClassRowNum.value + '��';
		retV+='�����ʽ$' + obj.TableCss.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
<%
End Sub
%>
</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../Supply/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.ClassID.value=TempArray[0]
		document.all.ClassName.value=TempArray[1]
	}
}
function selectHtml_express(Html_express,Dis_TF)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
		document.getElementById('li_id').style.display='none';
		document.getElementById('ul_id').style.display='none';
		if (Dis_TF == "1")
		{
			document.getElementById('DivID').disabled=true;
			document.getElementById('Table_ID').style.display='';
		}	
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		if (Dis_TF == "1")
		{
			document.getElementById('DivID').disabled=false;
			document.getElementById('Table_ID').style.display='none';
		}	
		break;
	}
}
function SelectFile()     
{
 var returnvalue = OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);
 if (returnvalue!='')
 {
 	event.srcElement.parentNode.firstChild.value=returnvalue;
 }
}
function showhide(obj)
{
	if(obj.style.display=='') obj.style.display='none'; else obj.style.display='';
	return false;
}
</script>






