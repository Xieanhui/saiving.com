<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
'------2006-12-29 �жϵ�ǰ�û������ͽ���Ƿ��㹻������Ϣ   by ken
Dim MustPoint,MustMoney,Get_RulerRs
Set Get_RulerRs = Conn.ExeCute("Select Top 1 PublicPoint,PublicMoney From FS_SD_Config Where ID > 0 Order By ID")
If Get_RulerRs.Eof Then
	MustPoint = 0
	MustMoney = 0
Else
	MustPoint = Clng(Get_RulerRs(0))	
	MustMoney = Clng(Get_RulerRs(1))
End if
Get_RulerRs.Close : Set Get_RulerRs = Nothing
'-------------------------------------------------------
If Clng(Fs_User.NumIntegral) < MustPoint Or Clng(Fs_User.NumFS_Money) < MustMoney Then
	strShowErr = "<li><font color=red>���ĵ������Ҳ���,���ܷ�����Ϣ!!</font></li>"
	Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End If
'�������
dim book_rs,new_count
set book_rs= Server.CreateObject(G_FS_RS)
book_rs.open "Select count(BookID) From FS_ME_Book where M_ReadUserNumber='"&Fs_User.UserNumber&"' and M_ReadTF=0 and M_Type=2",User_Conn,1,3
 if book_rs(0)>0 then
	 new_count = "<span class=""tx""><b>��������"& book_rs(0) &"</b></span>"
 else
	 new_count =  book_rs(0)
 end if
book_rs.close:set book_rs = nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../Editor/FS_scripts/editor.js"></script>
</head>
<body>

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">���λ�ã�<a href="../../">��վ��ҳ</a> &gt;&gt; <a href="../main.asp">��Ա��ҳ</a> 
            &gt;&gt; <a href="PublicManage.asp">����ϵͳ</a> &gt;&gt;��������</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="PublicSupply.asp">������Ϣ</a>��<a href="PublicManage.asp">������Ϣ</a>��<a href="PublicManage.asp#top10">�����Ϣ�������(TOP10)</a>��<a href="PublicManage.asp#new">���¹�����Ϣ</a>��<a href="PublicManage.asp#rec">�����Ƽ�</a>��<a href="../Book.asp?M_type=2">�ҵ�������(<%=new_count%>)</a></td>
        </tr>
      </table>
      <table width="98%"  border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form action="PubLic_Save.asp" method="post" name="s_form">
          <tr class="hback"> 
            <td width="100"> <div align="right">*��Ϣ����</div></td>
            <td > <input name="PubTitle" type="text" id="title2"  value="" size="40" maxlength="50">
              *���� 
              <select name="PubType" id="PubType">
                <option value="0" selected>��Ӧ</option>
                <option value="1">��</option>
                <option value="2">����</option>
                <option value="3">����</option>
                <option value="4">����</option>
              </select>
              ��*���ڵ��� 
              <select name="AreaID" id="AreaID">
                <%
				  dim obj_zl_rs
				  set obj_zl_rs= Server.CreateObject(G_FS_RS)
				  obj_zl_rs.open "select ID,PID,ClassName From FS_SD_Address  where PID=0  order by ClassLevel desc,id asc",Conn,1,3
				  do while not obj_zl_rs.eof 
						Response.Write"<option value="""& obj_zl_rs("ID")&""">��"& obj_zl_rs("ClassName")&"</option>"
					  Response.Write childAreaclassList(obj_zl_rs("ID")," ")
					  obj_zl_rs.movenext
				  Loop
				  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td width="100"> <div align="right">*����</div></td>
            <td > <select name="ClassID" id="ClassID">
                <%
				  set obj_zl_rs= Server.CreateObject(G_FS_RS)
				  obj_zl_rs.open "select ID,PID,GQ_ClassName,PID,classorder From FS_SD_Class  where PID=0  order by classorder desc,id desc",Conn,1,3
				  do while not obj_zl_rs.eof 
						Response.Write"<option value="""& obj_zl_rs("ID")&""">��"& obj_zl_rs("GQ_ClassName")&"</option>"
					  Response.Write childclassList(obj_zl_rs("ID")," ")
					  obj_zl_rs.movenext
				  Loop
				  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select>
              �����ҵ�ר�� 
              <select name="MyClassID" id="MyClassID">
                <option value="">ѡ�����Ĺ���</option>
                <%
			  set obj_zl_rs= Server.CreateObject(G_FS_RS)
			  obj_zl_rs.open "select ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes From FS_ME_InfoClass  where ParentID=0 and ClassTypes=4 and UserNumber='"& Fs_User.UserNumber&"' order by ClassID desc",User_Conn,1,3
			  do while not obj_zl_rs.eof 
			  		Response.Write"<option value="""& obj_zl_rs("ClassID")&""">"& obj_zl_rs("ClassCName")&"</option>"
				  obj_zl_rs.movenext
			  Loop
			  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td><div align="right">�ؼ���</div></td>
            <td ><input name="keyword1" type="text" id="keyword1"  value="" size="17" maxlength="15">
              , 
              <input name="keyword2" type="text" id="keyword2"  value="" size="17" maxlength="15">
              , 
              <input name="keyword3" type="text" id="keyword3"  value="" size="17" maxlength="15"></td>
          </tr>
          <tr class="hback"> 
            <td width="100"> <div align="right">��Ӫ��ʽ</div></td>
            <td > <input name="CompType" type="radio" value="0">
              ���� 
              <input name="CompType" type="radio" value="1" checked>
              ���� 
              <input type="radio" name="CompType" value="2">
              �����Ӫ</td>
          </tr>
		  
		<tr class="hback"> 
      <td align="center"><div align="right">��ϵ�绰</div></td>
      <td  align="center"><div align="left"> 
          <input name="Tel" type="text" id="Tel" value="" size="50">
        </div></td>
    </tr>
    <tr class="hback"> 
      <td align="center"><div align="right">����</div></td>
      <td  align="center"><div align="left"> 
          <input name="Fax" type="text" id="Fax" value="" size="50">
        </div></td>
    </tr>
    <tr class="hback"> 
      <td align="center"><div align="right">�ƶ��绰</div></td>
      <td  align="center"><div align="left"> 
          <input name="Mobile" type="text" id="Mobile" value="" size="50">
        </div></td>
    </tr>
	
          <tr class="hback"> 
            <td width="100"> <div align="right">����</div></td>
            <td ><input name="PubNumber" type="text" id="PubNumber" value="0" size="23">
              ����Ч���� 
              <input name="ValidTime" type="text" id="ValidTime" size="24"  value="15">
              �졣��ЧֵΪ1~360</td>
          </tr>
          <tr class="hback"> 
            <td rowspan="2" align="center"><div align="right">��Ʒ����</div></td>
            <td  align="center"><div align="left">��װ˵�� 
                <input name="PubPack" type="text" id="PubPack" maxlength="100">
                ��Ʒ�۸� 
                <input name="PubPrice" type="text" id="PubPrice" value="0" maxlength="10">
                ,0��ʾ����</div></td>
          </tr>
		  
          <tr class="hback"> 
            <td  align="center"><div align="left">��Ʒ��� 
                <input name="Pubgui" type="text" id="title3"  value="" size="53" maxlength="50">
              </div></td>
          </tr>
          <tr class="hback">
            <td align="center"><div align="right">����</div></td>
            <td  align="center"><div align="left">
                <input name="PubAddress" type="text" id="Pubgui"  value="" size="53" maxlength="50">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td width="100" align="center"> <div align="right">*����</div></td>
            <td  align="center">
				<!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/UserEditer.asp?id=PubContent' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="PubContent">
                <!--�༭������-->
				<span id="span_content"></span>
              </td>
          </tr>
          <tr class="hback"> 
            <td align="center"><div align="right">ͼƬ</div></td>
            <td align="center"><table width="100%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                  <td width="29%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_1"></td>
                        </tr>
                      </table>
                      <input name="pic_1" type="hidden" id="pic_1" size="40" >
                    </div></td>
                  <td width="36%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_2"></td>
                        </tr>
                      </table>
                      <input name="pic_2" type="hidden" id="pic_2" size="40">
                    </div></td>
                  <td width="35%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_3"></td>
                        </tr>
                      </table>
                      <input name="pic_3" type="hidden" id="pic_3" size="40">
                    </div></td>
                </tr>
                <tr> 
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_1);" style="cursor:hand;"> 
                      ��<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_2);" style="cursor:hand;"> 
                      ��<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_3);" style="cursor:hand;"> 
                      ��<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_3();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr class="hback"> 
            <td align="center">&nbsp;</td>
            <td align="center"> <div align="left"> 
                <input type="button" name="Submit1" id="Submit1" value="������Ϣ" onClick="checkinput(this.form);">
                <input type="reset" name="Submit" value="����">
                <input name="Action" type="hidden" id="Action" value="add">
              </div></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<%
Function childclassList(f_classid,f_tmp)
	Dim f_Child_c_Rs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
		Set f_Child_c_Rs = Conn.Execute("Select ID,PID,GQ_ClassName,classorder From FS_SD_Class  where PID=" & CintStr(f_classid) & "  order by classorder desc,id desc" )
		f_TempStr =f_tmp & "��"
		do while Not f_Child_c_Rs.Eof
				childclassList = childclassList & "<option value="""& f_Child_c_Rs("ID") &""">"
				childclassList = childclassList & "��" & f_TempStr &  f_Child_c_Rs("GQ_ClassName") 
				childclassList = childclassList & "</option>" & Chr(13) & Chr(10)
				childclassList = childclassList &childclassList(f_Child_c_Rs("ID"),f_TempStr)
			f_Child_c_Rs.MoveNext
		loop
		f_Child_c_Rs.Close
		Set f_Child_c_Rs = Nothing
end function
Function childAreaclassList(f_classid,f_tmp)
		dim rs_1,f_TempStr
		Set rs_1 = Conn.Execute("Select ID,PID,ClassName From FS_SD_Address  where PID=" & CintStr(f_classid) & "  order by ClassLevel desc,id asc" )
		f_TempStr =f_tmp & "��"
		do while Not rs_1.Eof
				childAreaclassList = childAreaclassList & "<option value="""& rs_1("ID") &""">"
				childAreaclassList = childAreaclassList & "��" & f_TempStr &  rs_1("ClassName") 
				childAreaclassList = childAreaclassList & "</option>" & Chr(13) & Chr(10)
				childAreaclassList = childAreaclassList &childAreaclassList(rs_1("ID"),f_TempStr)
			rs_1.MoveNext
		loop
		rs_1.Close
		Set rs_1 = Nothing
end function
Set Fs_User = Nothing
Set User_Conn = Nothing
Set Conn = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
	new Form.Element.Observer($('pic_1'),1,pics_1);
		function pics_1()
			{
				if ($('pic_1').value=='')
				{
					$('pic_p_1').src='../Images/nopic_supply.gif'
				}
				else
				{
				$('pic_p_1').src=$('pic_1').value
				}
			} 
	new Form.Element.Observer($('pic_2'),1,pics_2);
		function pics_2()
			{
				if($('pic_2').value=='')
				{
				$('pic_p_2').src='../Images/nopic_supply.gif'
				}
				else
				{
				$('pic_p_2').src=$('pic_2').value
				}
			} 
	new Form.Element.Observer($('pic_3'),1,pics_3);
		function pics_3()
			{
				if($('pic_3').value=='')
				{
				$('pic_p_3').src='../Images/nopic_supply.gif'
				}
				else
				{
				$('pic_p_3').src=$('pic_3').value
				}
			}
	function dels_1()
		{
			document.s_form.pic_1.value=''
		}
	function dels_2()
		{
			document.s_form.pic_2.value=''
		}
	function dels_3()
		{
			document.s_form.pic_3.value=''
		}
function checkinput(FormObj)
{
	s_form.PubContent.value=frames["NewsContent"].GetNewsContentArray();
	var submieTF = true;
	if(document.s_form.PubTitle.value=='') {
	alert("����д��Ϣ����");
	s_form.PubTitle.focus();
	submieTF = false;
	return;
	}
	/*if(document.s_form.PubAddress.value=='') {
	alert("��д���ڵ���");
	s_form.PubAddress.focus();
	submieTF = false;
	return;
	}*/
	if(document.s_form.ClassID.value=='') {
	alert("����д����");
	s_form.ClassID.focus();
	submieTF = false;
	return;
	}
	if(document.s_form.PubContent.value=='') {
	alert("��д��Ϣ����");
	s_form.PubContent.focus();
	submieTF = false;
	return;
	}
	if (submieTF == false)
	{
		alert('������������д����');
		return;
	}
	else if (submieTF == true)
	{
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
		document.s_form.PubContent.value=frames["NewsContent"].GetNewsContentArray();
		document.s_form.submit();
	}
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





