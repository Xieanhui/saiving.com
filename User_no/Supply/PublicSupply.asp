<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
'------2006-12-29 判断当前用户点数和金币是否足够发布信息   by ken
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
	strShowErr = "<li><font color=red>您的点数或金币不足,不能发布信息!!</font></li>"
	Response.Redirect(""& s_savepath &"/lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End If
'获得留言
dim book_rs,new_count
set book_rs= Server.CreateObject(G_FS_RS)
book_rs.open "Select count(BookID) From FS_ME_Book where M_ReadUserNumber='"&Fs_User.UserNumber&"' and M_ReadTF=0 and M_Type=2",User_Conn,1,3
 if book_rs(0)>0 then
	 new_count = "<span class=""tx""><b>您有留言"& book_rs(0) &"</b></span>"
 else
	 new_count =  book_rs(0)
 end if
book_rs.close:set book_rs = nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
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
          <td  valign="top">你的位置：<a href="../../">网站首页</a> &gt;&gt; <a href="../main.asp">会员首页</a> 
            &gt;&gt; <a href="PublicManage.asp">供求系统</a> &gt;&gt;发布供求</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="PublicSupply.asp">发布信息</a>┆<a href="PublicManage.asp">管理信息</a>┆<a href="PublicManage.asp#top10">你的信息浏览排行(TOP10)</a>┆<a href="PublicManage.asp#new">最新供求信息</a>┆<a href="PublicManage.asp#rec">供求推荐</a>┆<a href="../Book.asp?M_type=2">我的新留言(<%=new_count%>)</a></td>
        </tr>
      </table>
      <table width="98%"  border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form action="PubLic_Save.asp" method="post" name="s_form">
          <tr class="hback"> 
            <td width="100"> <div align="right">*信息标题</div></td>
            <td > <input name="PubTitle" type="text" id="title2"  value="" size="40" maxlength="50">
              *类型 
              <select name="PubType" id="PubType">
                <option value="0" selected>供应</option>
                <option value="1">求购</option>
                <option value="2">合作</option>
                <option value="3">代理</option>
                <option value="4">其他</option>
              </select>
              　*所在地区 
              <select name="AreaID" id="AreaID">
                <%
				  dim obj_zl_rs
				  set obj_zl_rs= Server.CreateObject(G_FS_RS)
				  obj_zl_rs.open "select ID,PID,ClassName From FS_SD_Address  where PID=0  order by ClassLevel desc,id asc",Conn,1,3
				  do while not obj_zl_rs.eof 
						Response.Write"<option value="""& obj_zl_rs("ID")&""">├"& obj_zl_rs("ClassName")&"</option>"
					  Response.Write childAreaclassList(obj_zl_rs("ID")," ")
					  obj_zl_rs.movenext
				  Loop
				  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td width="100"> <div align="right">*分类</div></td>
            <td > <select name="ClassID" id="ClassID">
                <%
				  set obj_zl_rs= Server.CreateObject(G_FS_RS)
				  obj_zl_rs.open "select ID,PID,GQ_ClassName,PID,classorder From FS_SD_Class  where PID=0  order by classorder desc,id desc",Conn,1,3
				  do while not obj_zl_rs.eof 
						Response.Write"<option value="""& obj_zl_rs("ID")&""">├"& obj_zl_rs("GQ_ClassName")&"</option>"
					  Response.Write childclassList(obj_zl_rs("ID")," ")
					  obj_zl_rs.movenext
				  Loop
				  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select>
              　　我的专栏 
              <select name="MyClassID" id="MyClassID">
                <option value="">选择您的归类</option>
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
            <td><div align="right">关键字</div></td>
            <td ><input name="keyword1" type="text" id="keyword1"  value="" size="17" maxlength="15">
              , 
              <input name="keyword2" type="text" id="keyword2"  value="" size="17" maxlength="15">
              , 
              <input name="keyword3" type="text" id="keyword3"  value="" size="17" maxlength="15"></td>
          </tr>
          <tr class="hback"> 
            <td width="100"> <div align="right">经营方式</div></td>
            <td > <input name="CompType" type="radio" value="0">
              批发 
              <input name="CompType" type="radio" value="1" checked>
              零售 
              <input type="radio" name="CompType" value="2">
              批零兼营</td>
          </tr>
		  
		<tr class="hback"> 
      <td align="center"><div align="right">联系电话</div></td>
      <td  align="center"><div align="left"> 
          <input name="Tel" type="text" id="Tel" value="" size="50">
        </div></td>
    </tr>
    <tr class="hback"> 
      <td align="center"><div align="right">传真</div></td>
      <td  align="center"><div align="left"> 
          <input name="Fax" type="text" id="Fax" value="" size="50">
        </div></td>
    </tr>
    <tr class="hback"> 
      <td align="center"><div align="right">移动电话</div></td>
      <td  align="center"><div align="left"> 
          <input name="Mobile" type="text" id="Mobile" value="" size="50">
        </div></td>
    </tr>
	
          <tr class="hback"> 
            <td width="100"> <div align="right">数量</div></td>
            <td ><input name="PubNumber" type="text" id="PubNumber" value="0" size="23">
              　有效期限 
              <input name="ValidTime" type="text" id="ValidTime" size="24"  value="15">
              天。有效值为1~360</td>
          </tr>
          <tr class="hback"> 
            <td rowspan="2" align="center"><div align="right">产品参数</div></td>
            <td  align="center"><div align="left">包装说明 
                <input name="PubPack" type="text" id="PubPack" maxlength="100">
                产品价格 
                <input name="PubPrice" type="text" id="PubPrice" value="0" maxlength="10">
                ,0表示面议</div></td>
          </tr>
		  
          <tr class="hback"> 
            <td  align="center"><div align="left">产品规格 
                <input name="Pubgui" type="text" id="title3"  value="" size="53" maxlength="50">
              </div></td>
          </tr>
          <tr class="hback">
            <td align="center"><div align="right">产地</div></td>
            <td  align="center"><div align="left">
                <input name="PubAddress" type="text" id="Pubgui"  value="" size="53" maxlength="50">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td width="100" align="center"> <div align="right">*描述</div></td>
            <td  align="center">
				<!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/UserEditer.asp?id=PubContent' frameborder=0 scrolling=no width='100%' height='280'></iframe>
				<input type="hidden" name="PubContent">
                <!--编辑器结束-->
				<span id="span_content"></span>
              </td>
          </tr>
          <tr class="hback"> 
            <td align="center"><div align="right">图片</div></td>
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
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_2);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_3);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_3();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr class="hback"> 
            <td align="center">&nbsp;</td>
            <td align="center"> <div align="left"> 
                <input type="button" name="Submit1" id="Submit1" value="保存信息" onClick="checkinput(this.form);">
                <input type="reset" name="Submit" value="重置">
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
		f_TempStr =f_tmp & "┄"
		do while Not f_Child_c_Rs.Eof
				childclassList = childclassList & "<option value="""& f_Child_c_Rs("ID") &""">"
				childclassList = childclassList & "├" & f_TempStr &  f_Child_c_Rs("GQ_ClassName") 
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
		f_TempStr =f_tmp & "┄"
		do while Not rs_1.Eof
				childAreaclassList = childAreaclassList & "<option value="""& rs_1("ID") &""">"
				childAreaclassList = childAreaclassList & "├" & f_TempStr &  rs_1("ClassName") 
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
	alert("请填写信息标题");
	s_form.PubTitle.focus();
	submieTF = false;
	return;
	}
	/*if(document.s_form.PubAddress.value=='') {
	alert("填写所在地区");
	s_form.PubAddress.focus();
	submieTF = false;
	return;
	}*/
	if(document.s_form.ClassID.value=='') {
	alert("请填写分类");
	s_form.ClassID.focus();
	submieTF = false;
	return;
	}
	if(document.s_form.PubContent.value=='') {
	alert("填写信息描述");
	s_form.PubContent.focus();
	submieTF = false;
	return;
	}
	if (submieTF == false)
	{
		alert('必填资料请填写完整');
		return;
	}
	else if (submieTF == true)
	{
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
		document.s_form.PubContent.value=frames["NewsContent"].GetNewsContentArray();
		document.s_form.submit();
	}
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





