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
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show

	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	
	'--------------ѡ����ҵ��Աʱ����ҵ����--�������ݿ���������------------------------
	
	Dim User_Conn,FS_UserConnection_Str
	if G_IS_SQL_User_DB=0 then
	FS_UserConnection_Str = "DBQ=" + Server.MapPath(Add_Root_Dir(G_User_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	else
	FS_UserConnection_Str = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_User_DATABASE_CONN_STR &";"
	end if
	Set User_Conn = Server.CreateObject(G_FS_CONN)
	User_Conn.Open FS_UserConnection_Str
	'--------------�������ݿ�-----------------------------------------------------------
	Dim obj_VClass_Rs,VC_Class_List
	VC_Class_List="<option value=""0"">��ѡ����ҵ</option>"
	Set  obj_VClass_Rs = server.CreateObject(G_FS_RS)
	obj_VClass_Rs.Open "Select VCID,vClassName from FS_ME_VocationClass Order by  VCID desc",User_Conn,1,3
	do while Not obj_VClass_Rs.eof 
		VC_Class_List = VC_Class_List&"<option value="""& obj_VClass_Rs("VCID")&""">"& obj_VClass_Rs("vClassName")&"</option>"
		obj_VClass_Rs.movenext
	loop
	obj_VClass_Rs.close:set obj_VClass_Rs = nothing
	
	
	'----2007-02-08 ��Ա��½��ǩ��ʽ�б�
	Dim Select_Login,GetLoginStrRs,End_Se_Str
	Set GetLoginStrRs = Conn.ExeCute("Select [ID],StyleName From FS_MF_Labestyle Where StyleType = 'Login' Order By ID Desc")
	Select_Login = "<select name=""Login_StyleID"" id=""Login_StyleID"">" & vbnewline
	Select_Login = Select_Login & "<option value="""" selected>ѡ���½��ǩ��ʽ</option>" & vbnewline
	End_Se_Str = "</select>"
	If GetLoginStrRs.Eof Then
		Select_Login = Select_Login & End_Se_Str
	Else
		Do While Not GetLoginStrRs.Eof
			Select_Login = Select_Login & "<option value=""" & GetLoginStrRs(0) & """>" & GetLoginStrRs(1) & "</option>"
		GetLoginStrRs.MoveNExt
		Loop
		Select_Login = Select_Login & End_Se_Str
	End If
	GetLoginStrRs.Close : Set GetLoginStrRs = NOthing
	'------
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
      <td width="13%" height="15"><div align="center"><a href="All_Label.asp?type=PostionNavi" target="_self">λ�õ���</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=PageTitle" target="_self">ҳ�����</a><a href="News_C_Label.asp?type=OldNews" target="_self"></a></div></td>
      <td width="13%" style="display:none"><div align="center"><a href="All_Label.asp?type=SiteMap" target="_self">վ���ͼ</a></div></td>
      <td width="13%"><div align="center"><a href="News_C_Label.asp?type=NorFilt" target="_self"></a><a href="All_Label.asp?type=Search" target="_self">������</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=InfoStat" target="_self">��Ϣͳ��</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=UserLogin" target="_self">�û���½</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=CopyRight" target="_self">��Ȩ��Ϣ</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=SubList" target="_self">��վ����</a></div></td>
    </tr>
    <tr class="hback" align="center">
      <td height="15"><a href="All_Label.asp?type=UserList" target="_self">��Ա�б�</a></td>
      <td><a href="All_Label.asp?type=CustomForm" target="_self">�Զ����</a></td>
      <td style="display:none">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <%
  dim str_type
  str_type = Request.QueryString("type")
  select case str_type
  		case "PostionNavi"
			Call PostionNavi()
		Case "PageTitle"
  			Call PageTitle()
		Case "SiteMap"
			Call SiteMap()
		Case "Search"
			Call Search()
		Case "InfoStat"
			Call InfoStat()
		Case "UserLogin"
			Call UserLogin()
		Case "CopyRight"
			Call CopyRight()
		Case "SubList"
			Call SubList()
		Case "CustomForm"
			Call CustomForm()
		Case "UserList"
			Call UserList()
			
		Case else
			Call PostionNavi()
  end select
  Sub PostionNavi()
  %>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">λ�õ���</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">�ָ��ַ�(ͼƬ)</div></td>
      <td width="72%" class="hback"><input name="NaviChar" type="text" id="NaviChar" value=" &gt;&gt; ">
      ��ʹ��html�﷨,�벻Ҫʹ�á�$�����������ַ�</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="LinkCSS" type="text" id="LinkCSS">
      �벻Ҫʹ�á�$�����������ַ�</td>
    </tr>
		 <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback" align="left" valign="middle"><select name="OpenMode" id="OpenMode" style="width:130px;">
				<option value="0" selected>��</option>
        <option value="1">��</option>
			</select>
		   </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ǰλ������</div></td>
      <td class="hback"><input name="Nchar" type="text" id="Nchar" value="����">
        CSS
        <input name="NcharCSS" type="text" id="NcharCSS"></td>
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
		var retV = '{FS:MF=PostionNavi��';
		retV+='�ָ��ַ�$' + obj.NaviChar.value + '��';
		retV+='λ������$' + obj.Nchar.value + '��';
		retV+='λ������css$' + obj.NcharCSS.value + '��';
		retV+='����CSS$' + obj.LinkCSS.value + '��';
		retV+='��������$' + obj.OpenMode.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub PageTitle()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">ҳ�����</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">��������</div></td>
      <td width="72%" class="hback"><input name="O_Char" type="text" id="O_Char" value="��Ѷ__Foosun.CN">
      �벻Ҫʹ�á�$�����������ַ�</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������λ��</div></td>
      <td class="hback"><select name="O_Char_dir" id="O_Char_dir">
        <option value="0">ǰ׺</option>
        <option value="1" selected>��׺</option>
      </select></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">�ָ��ַ�</div></td>
      <td class="hback"><input name="split_char" type="text" id="split_char" value="__">
        �벻Ҫʹ�á�$�����������ַ�</td>
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
		var retV = '{FS:MF=PageTitle��';
		retV+='��������$' + obj.O_Char.value + '��';
		retV+='��������λ��$' + obj.O_Char_dir.value + '��';
		retV+='�ָ��ַ�$' + obj.split_char.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SiteMap()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">վ���ͼ</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">�������CSS</div></td>
      <td width="72%" class="hback"><input name="TitleCSS" type="text" id="TitleCSS">
      �벻Ҫʹ�á�$�����������ַ�</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ϵͳ����CSS</div></td>
      <td class="hback"><input name="SubCSS" type="text" id="SubCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͬ������ָ�</div></td>
      <td class="hback"><input name="split_char" type="text" id="split_char">
        �벻Ҫʹ�á�$�����������ַ�</td>
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
		var retV = '{FS:MF=SiteMap��';
		retV+='�������CSS$' + obj.TitleCSS.value + '��';
		retV+='��ϵͳ����CSS$' + obj.SubCSS.value + '��';
		retV+='ͬ������ָ�$' + obj.split_char.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%end Sub%>
 <%Sub Search()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">������</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">��������</div></td>
      <td width="72%" class="hback"><select name="DateShow" id="DateShow">
        <option value="1">��ʾ</option>
        <option value="0" selected>����ʾ</option>
      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ģ������</div></td>
      <td class="hback"><select name="SearchType" id="SearchType">
        <option value="0" selected>��</option>
        <option value="1">��</option>
      </select>      </td>
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
		var retV = '{FS:MF=Search��';
		retV+='��������$' + obj.DateShow.value + '��';
		retV+='ģ������$' + obj.SearchType.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub InfoStat()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ϣͳ��</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">���з�ʽ</div></td>
      <td width="72%" class="hback"><select name="cols" id="cols">
        <option value="1">����</option>
        <option value="0" selected>����</option>
      </select>      </td>
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
		var retV = '{FS:MF=InfoStat��';
		retV+='���з�ʽ$' + obj.cols.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub UserLogin()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ա��½</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">ѡ���ǩ��ʽ</div></td>
      <td width="72%" class="hback"><select name="LableType" id="LableType" onChange="SelectLOginType(this.options[this.selectedIndex].value);">
        <option value="0" selected="selected">�̶���ʽ</option>
        <option value="1">�Զ�����ʽ</option>
      </select></td>
    </tr>
	<tr id="HaveTF" style="display:;">
      <td width="28%" class="hback"><div align="right">��ʾ��ʽ</div></td>
      <td width="72%" class="hback"><select name="LoginDisStyle" id="LoginDisStyle">
        <option value="vertical">����</option>
        <option value="transverse">����</option>
      </select>      </td>
    </tr>
	<tr id="Se_Style" style="display:none;">
      <td colspan="2" class="xingmu">
	  	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="table">
			<tr >
			  <td width="28%" class="hback"><div align="right">��ǩ������ʽ</div></td>
			  <td width="72%" class="hback">
			  <% = Select_Login %>
			  <span id="Txt_loginType"></span>
			  </td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">��½��ǩ������ʽ</div></td>
			  <td width="72%" class="hback"><input name="BGStyle" type="text" id="BGStyle" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			  <span style="color:#ff0000;">����Ϊ�����css��ʽ����Ҳ����ΪͼƬ</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">ѡ�����ʽ</div></td>
			  <td width="72%" class="hback"><input name="SelectStyle" type="text" id="SelectStyle" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">ѡ��������˵���ʽ</div></td>
			  <td width="72%" class="hback"><input name="SelectBGCss" type="text" id="SelectBGCss" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">�ı�����ʽ</div></td>
			  <td width="72%" class="hback"><input name="TextStyle" type="text" id="TextStyle" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">�ύ��ť��ʽ</div></td>
			  <td width="72%" class="hback"><input name="ButtonStyle" type="text" id="ButtonStyle" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			  <span style="color:#ff0000;">����Ϊ�����css��ʽ����Ҳ����ΪͼƬ</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">ȡ����ť��ʽ</div></td>
			  <td width="72%" class="hback"><input name="ResestCss" type="text" id="ResestCss" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			  <span style="color:#ff0000;">����Ϊ�����css��ʽ����Ҳ����ΪͼƬ</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">ע�������ַ���ʽ</div></td>
			  <td width="72%" class="hback"><input name="Reg_LinkCss" type="text" id="Reg_LinkCss" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="SelectFile();">
			  <span style="color:#ff0000;">����Ϊ�����css��ʽ����Ҳ����ΪͼƬ</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">ȡ�����������ַ���ʽ</div></td>
			  <td width="72%" class="hback"><input name="Get_PassCss" type="text" id="Get_PassCss" value="">
				<input type="button" name="bnt_ChoosePic_rowBettween2"  value="ѡ��ͼƬ" onClick="SelectFile();">
				<span style="color:#ff0000;">����Ϊ�����css��ʽ����Ҳ����ΪͼƬ</span>
			  </td>
			</tr>
		</table>
	  </td>
	</tr>		
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
<script language="JavaScript" type="text/JavaScript">
<!--
function SelectLOginType(STRID)
{
	if (STRID == '0')
	{
		document.getElementById('HaveTF').style.display = '';
		document.getElementById('Se_Style').style.display = 'none';
	}
	else
	{
		document.getElementById('HaveTF').style.display = 'none';
		document.getElementById('Se_Style').style.display = '';
	}
}
function ok(obj)
{
	var Str_disType = obj.LableType.value;
	var Str_StyleID = obj.Login_StyleID.value;
	if (Str_disType == '1')
	{
		if (Str_StyleID == '')
		{
			document.getElementById('Txt_loginType').innerHTML = '<font color=red>��ʽ����ѡ��,��û�����Ƚ�����ʽ</font>';
			obj.Login_StyleID.focus();
			return false;
		}
	}
	switch (Str_disType)
	{
	case '0':
		var retV = '{FS:MF=UserLogin��';
		retV+='��ǩ��ʽ$' + Str_disType + '��';
		retV+='��ʾ��ʽ$' + obj.LoginDisStyle.value;	
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
		break;
	case '1':
		var retV = '{FS:MF=UserLogin��';
		retV+='��ǩ��ʽ$' + Str_disType + '��';
		retV+='������ʽ$' + Str_StyleID + '��';
		retV+='��ǩ����$' + obj.BGStyle.value + '��';
		retV+='ѡ�����ʽ$' + obj.SelectStyle.value + '��';
		retV+='ѡ���˵���ʽ$' + obj.SelectBGCss.value + '��';
		retV+='�ı�����ʽ$' + obj.TextStyle.value + '��';
		retV+='�ύ��ť��ʽ$' + obj.ButtonStyle.value + '��';
		retV+='ȡ����ť��ʽ$' + obj.ResestCss.value + '��';
		retV+='ע��������ʽ$' + obj.Reg_LinkCss.value + '��';
		retV+='ȡ������������ʽ$' + obj.Get_PassCss.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
		break;
	}		
}
-->
</script>
 <%End Sub

  Sub UserList()
  %>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ա�б�</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">��Ա����</div></td>
      <td width="72%" class="hback">
	  <select name="UserType"  onChange="Select_VC_Class(this.options[this.selectedIndex].value);">
	   <option value="All">����</option>
	    <option value=0>���˻�Ա</option>
        <option value=1>��ҵ��Ա</option>
	  </select>
	  <!--Select_VC_Class���������ó���ѡ����ҵ��Ա��ʱ����ʾ��ҵ����-->
 <script language="JavaScript" type="text/JavaScript">
function Select_VC_Class(Html_express)
{
	switch (Html_express)
	{
	case "All":
		document.getElementById('VC_Class').style.display='none';
		//document.getElementById('VC_Class').disabled=false;
		break;
	case "0":
		//document.getElementById('VC_Class').disabled=true;
		document.getElementById('VC_Class').style.display='none';
		break;
	case "1":
		//document.getElementById('VC_Class').disabled=true;
		document.getElementById('VC_Class').style.display='';
		break;
	}
}
</script>��
	  </td>
    </tr>
	    <tr name="VC_Class" id="VC_Class" style="font-family:����;display:none;">
      <td class="hback"><div align="right"><span style="color:#FF0000">��ҵ��Ա��ҵ����</span></div></td>
      <td class="hback">
	  <select id="VClass"  name="VClass" style="width:20%">
            <% = VC_Class_List %>
        </select><span style="color:#FF0000">*����ѡ������������</span>
	  </td>	  
    </tr>
    <tr>
      <td class="hback"><div align="right">�б�����</div></td>
      <td class="hback">
	  <select name="OrderBy">
	   <%=PrintOption("","RegTime:����,LoginNum:��¼����,Hits:����,Integral:��Ա����,FS_Money:��Ա���")%>
	  </select>
	  </td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">��Ա�Ա�</div></td>
      <td width="72%" class="hback">
	  <select name="UserSex">
	   <option value="All">����</option>
	  <%=PrintOption("","0:��,1:Ů")%>
	  </select>
	  </td>
    </tr>


<!----------------------------->

    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ÿ������</div></td>
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
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select> </td>
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
            <%
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='ME' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	Response.Write(label_style_List)
			 %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">���ڸ�����ϵͳ�н���ǰ̨ҳ����ʾ��ʽ</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ÿ����ʽ</div></td>
      <td class="hback">
	   ����<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   ż��<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		ֻ��Ա�����,��ֱ������ɫ#FF0000   
	  </td>
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
		if(isNaN(obj.TitleNumber.value))
		{alert('�б������������֡�');obj.TitleNumber.focus();return false;}
		if(obj.NewsStyle.value=='')
		{alert('������ʽ������д��');obj.NewsStyle.focus();return false;}

		var retV = '{FS:ME=UserList��';
		retV+='��Ա����$' + obj.UserType.value + '��';
		retV+='�б�����$' + obj.OrderBy.value + '��';
		retV+='��Ա�Ա�$' + obj.UserSex.value + '��';
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
		retV+='������ʽ$' + obj.NewsStyle.value+ '��';
		retV+='��������ʽ$' + obj.PubType_JiTR.value + '��';
		retV+='ż������ʽ$' + obj.PubType_OuTR.value + '��';
		//��ҵ����
		retV+='��ҵ����$' + obj.VClass.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 
 
 <%Sub CopyRight()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��Ȩ��Ϣ</td>
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
		var retV = '{FS:MF=CopyRight��';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SubList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">��ϵͳ����</td>
    </tr>
	<tr>
      <td width="28%" class="hback"><div align="right">�ָ����</div></td>
      <td width="72%" class="hback"><label>
        <input name="SubName" type="text" id="SubName">
      </label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">CSS</div></td>
      <td class="hback"><input name="SubCSS" type="text" id="SubCSS"></td>
    </tr>
    <tr>
      <td class="hback">&nbsp;</td>
      <td class="hback"><span class="tx">(������ϵͳ�ĵ�����������ϵͳ����--��ϵͳά��������)���ر�ע�⣺������վϵͳ��ǰ̨�������ӱ�������ϵͳ�Ĳ����������������ͬ</span></td>
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
		var retV = '{FS:MF=SubList��';
		retV+='�ָ��-����ʹ��html�﷨$' + obj.SubName.value + '��';
		retV+='CSS$' + obj.SubCSS.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub CustomForm()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">�Զ����</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">���ñ�</div></td>
      <td class="hback">
<SELECT name="CustomFormID" id="CustomFormID" style="width:40%">
  <%
  Dim CustomFormRS
  Set CustomFormRS = Conn.Execute("Select * from FS_MF_CustomForm where state=0")
  Do while Not CustomFormRS.Eof
  %>
  <option value="<% = CustomFormRS("ID") %>" style="color:#FF0000;"><% = CustomFormRS("formname") %></option>
  <%
  	CustomFormRS.MoveNext
  Loop
  CustomFormRS.Close
  Set CustomFormRS = Nothing
  %>
</SELECT>
	  </td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">����ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="CustomFormSytleID"  name="CustomFormSytleID" style="width:40%">
		  	<option value="" selected>ѡ�����ʽ</option>
            <%
	Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
	obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='CForm' Order by  id desc",Conn,1,3
	do while Not obj_label_style_Rs.eof 
		label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
		obj_label_style_Rs.movenext
	loop
	obj_label_style_Rs.close:set obj_label_style_Rs = nothing
	Response.Write(label_style_List)
			 %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.CustomFormSytleID.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">��ʾ�Զ������������ʽ</span></div></td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">��������ʽ</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="CustomDataSytleID"  name="CustomDataSytleID" style="width:40%">
		  	<option value="" selected>ѡ���������ʽ</option>
            <%
			Response.Write(label_style_List)
			 %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.CustomDataSytleID.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="�鿴">
          <span class="tx">��ʾ�Զ�������ݵ�������ʽ</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">�ı���CSS</div></td>
      <td class="hback"><input name="CustomFormTextCSS" type="text" id="CustomFormTextCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">������CSS</div></td>
      <td class="hback"><input name="CustomFormSelectCSS" type="text" id="CustomFormSelectCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��������CSS</div></td>
      <td class="hback"><input name="CustomFormOtherCSS" type="text" id="CustomFormOtherCSS"></td>
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
		if(obj.CustomFormID.value!=''){
			var retV = '{FS:MF=CustomForm��';
			retV+='���ñ�$' + obj.CustomFormID.value + '��';
			retV+='����ʽ$' + obj.CustomFormSytleID.value + '��';
			retV+='������ʽ$' + obj.CustomDataSytleID.value + '';
			if(obj.CustomFormSytleID.value!=''){
				retV+='���ı���CSS$' + obj.CustomFormTextCSS.value + '��';
				retV+='������CSS$' + obj.CustomFormSelectCSS.value + '��';
				retV+='��������CSS$' + obj.CustomFormOtherCSS.value + '';
			}
			retV+='}';
			window.parent.returnValue = retV;
			window.close();
		}else{alert('��ѡ����ñ�');}
	}
	</script>
 <%End Sub%>
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
function SelectFile()     
{
 var returnvalue = OpenWindow('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);
 if (returnvalue!='')
 {
 	event.srcElement.parentNode.firstChild.value=returnvalue;
 }
}
</script>