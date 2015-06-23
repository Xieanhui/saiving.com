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
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show

	Dim obj_label_style_Rs,label_style_List
	label_style_List=""
	
	'--------------选择企业会员时按行业分类--载入数据库驱动程序------------------------
	
	Dim User_Conn,FS_UserConnection_Str
	if G_IS_SQL_User_DB=0 then
	FS_UserConnection_Str = "DBQ=" + Server.MapPath(Add_Root_Dir(G_User_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	else
	FS_UserConnection_Str = "Provider=SQLOLEDB.1;Persist Security Info=false;"& G_User_DATABASE_CONN_STR &";"
	end if
	Set User_Conn = Server.CreateObject(G_FS_CONN)
	User_Conn.Open FS_UserConnection_Str
	'--------------操作数据库-----------------------------------------------------------
	Dim obj_VClass_Rs,VC_Class_List
	VC_Class_List="<option value=""0"">请选择行业</option>"
	Set  obj_VClass_Rs = server.CreateObject(G_FS_RS)
	obj_VClass_Rs.Open "Select VCID,vClassName from FS_ME_VocationClass Order by  VCID desc",User_Conn,1,3
	do while Not obj_VClass_Rs.eof 
		VC_Class_List = VC_Class_List&"<option value="""& obj_VClass_Rs("VCID")&""">"& obj_VClass_Rs("vClassName")&"</option>"
		obj_VClass_Rs.movenext
	loop
	obj_VClass_Rs.close:set obj_VClass_Rs = nothing
	
	
	'----2007-02-08 会员登陆标签样式列表
	Dim Select_Login,GetLoginStrRs,End_Se_Str
	Set GetLoginStrRs = Conn.ExeCute("Select [ID],StyleName From FS_MF_Labestyle Where StyleType = 'Login' Order By ID Desc")
	Select_Login = "<select name=""Login_StyleID"" id=""Login_StyleID"">" & vbnewline
	Select_Login = Select_Login & "<option value="""" selected>选择登陆标签样式</option>" & vbnewline
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
<title>新闻标签管理</title>
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
            <td width="20%" class="xingmu"><strong>常规标签创建</strong></td>
            <td width="80%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="13%" height="15"><div align="center"><a href="All_Label.asp?type=PostionNavi" target="_self">位置导航</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=PageTitle" target="_self">页面标题</a><a href="News_C_Label.asp?type=OldNews" target="_self"></a></div></td>
      <td width="13%" style="display:none"><div align="center"><a href="All_Label.asp?type=SiteMap" target="_self">站点地图</a></div></td>
      <td width="13%"><div align="center"><a href="News_C_Label.asp?type=NorFilt" target="_self"></a><a href="All_Label.asp?type=Search" target="_self">搜索表单</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=InfoStat" target="_self">信息统计</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=UserLogin" target="_self">用户登陆</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=CopyRight" target="_self">版权信息</a></div></td>
      <td width="13%"><div align="center"><a href="All_Label.asp?type=SubList" target="_self">子站导航</a></div></td>
    </tr>
    <tr class="hback" align="center">
      <td height="15"><a href="All_Label.asp?type=UserList" target="_self">会员列表</a></td>
      <td><a href="All_Label.asp?type=CustomForm" target="_self">自定义表单</a></td>
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
      <td colspan="2" class="xingmu">位置导航</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">分割字符(图片)</div></td>
      <td width="72%" class="hback"><input name="NaviChar" type="text" id="NaviChar" value=" &gt;&gt; ">
      请使用html语法,请不要使用“$，┆”禁用字符</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">联接CSS</div></td>
      <td class="hback"><input name="LinkCSS" type="text" id="LinkCSS">
      请不要使用“$，┆”禁用字符</td>
    </tr>
		 <tr>
      <td class="hback"><div align="right">弹出窗口</div></td>
      <td class="hback" align="left" valign="middle"><select name="OpenMode" id="OpenMode" style="width:130px;">
				<option value="0" selected>否</option>
        <option value="1">是</option>
			</select>
		   </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">当前位置文字</div></td>
      <td class="hback"><input name="Nchar" type="text" id="Nchar" value="正文">
        CSS
        <input name="NcharCSS" type="text" id="NcharCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=PostionNavi┆';
		retV+='分割字符$' + obj.NaviChar.value + '┆';
		retV+='位置文字$' + obj.Nchar.value + '┆';
		retV+='位置文字css$' + obj.NcharCSS.value + '┆';
		retV+='联接CSS$' + obj.LinkCSS.value + '┆';
		retV+='弹出窗口$' + obj.OpenMode.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub PageTitle()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">页面标题</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">附属文字</div></td>
      <td width="72%" class="hback"><input name="O_Char" type="text" id="O_Char" value="风讯__Foosun.CN">
      请不要使用“$，┆”禁用字符</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">附属文字位置</div></td>
      <td class="hback"><select name="O_Char_dir" id="O_Char_dir">
        <option value="0">前缀</option>
        <option value="1" selected>后缀</option>
      </select></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">分割字符</div></td>
      <td class="hback"><input name="split_char" type="text" id="split_char" value="__">
        请不要使用“$，┆”禁用字符</td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=PageTitle┆';
		retV+='附属文字$' + obj.O_Char.value + '┆';
		retV+='附属文字位置$' + obj.O_Char_dir.value + '┆';
		retV+='分割字符$' + obj.split_char.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SiteMap()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">站点地图</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">分类标题CSS</div></td>
      <td width="72%" class="hback"><input name="TitleCSS" type="text" id="TitleCSS">
      请不要使用“$，┆”禁用字符</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">子系统标题CSS</div></td>
      <td class="hback"><input name="SubCSS" type="text" id="SubCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">同级分类分割</div></td>
      <td class="hback"><input name="split_char" type="text" id="split_char">
        请不要使用“$，┆”禁用字符</td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=SiteMap┆';
		retV+='分类标题CSS$' + obj.TitleCSS.value + '┆';
		retV+='子系统标题CSS$' + obj.SubCSS.value + '┆';
		retV+='同级分类分割$' + obj.split_char.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%end Sub%>
 <%Sub Search()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">搜索表单</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">日期搜索</div></td>
      <td width="72%" class="hback"><select name="DateShow" id="DateShow">
        <option value="1">显示</option>
        <option value="0" selected>不显示</option>
      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">模糊搜索</div></td>
      <td class="hback"><select name="SearchType" id="SearchType">
        <option value="0" selected>否</option>
        <option value="1">是</option>
      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=Search┆';
		retV+='日期搜索$' + obj.DateShow.value + '┆';
		retV+='模糊搜索$' + obj.SearchType.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub InfoStat()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">信息统计</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">排列方式</div></td>
      <td width="72%" class="hback"><select name="cols" id="cols">
        <option value="1">纵向</option>
        <option value="0" selected>横向</option>
      </select>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=InfoStat┆';
		retV+='排列方式$' + obj.cols.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub UserLogin()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">会员登陆</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">选择标签方式</div></td>
      <td width="72%" class="hback"><select name="LableType" id="LableType" onChange="SelectLOginType(this.options[this.selectedIndex].value);">
        <option value="0" selected="selected">固定样式</option>
        <option value="1">自定义样式</option>
      </select></td>
    </tr>
	<tr id="HaveTF" style="display:;">
      <td width="28%" class="hback"><div align="right">显示方式</div></td>
      <td width="72%" class="hback"><select name="LoginDisStyle" id="LoginDisStyle">
        <option value="vertical">纵向</option>
        <option value="transverse">横向</option>
      </select>      </td>
    </tr>
	<tr id="Se_Style" style="display:none;">
      <td colspan="2" class="xingmu">
	  	<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="table">
			<tr >
			  <td width="28%" class="hback"><div align="right">标签引用样式</div></td>
			  <td width="72%" class="hback">
			  <% = Select_Login %>
			  <span id="Txt_loginType"></span>
			  </td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">登陆标签背景样式</div></td>
			  <td width="72%" class="hback"><input name="BGStyle" type="text" id="BGStyle" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			  <span style="color:#ff0000;">可以为定义的css样式名，也可以为图片</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">选择筐样式</div></td>
			  <td width="72%" class="hback"><input name="SelectStyle" type="text" id="SelectStyle" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">选择筐下拉菜单样式</div></td>
			  <td width="72%" class="hback"><input name="SelectBGCss" type="text" id="SelectBGCss" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">文本筐样式</div></td>
			  <td width="72%" class="hback"><input name="TextStyle" type="text" id="TextStyle" value=""></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">提交按钮样式</div></td>
			  <td width="72%" class="hback"><input name="ButtonStyle" type="text" id="ButtonStyle" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			  <span style="color:#ff0000;">可以为定义的css样式名，也可以为图片</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">取消按钮样式</div></td>
			  <td width="72%" class="hback"><input name="ResestCss" type="text" id="ResestCss" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			  <span style="color:#ff0000;">可以为定义的css样式名，也可以为图片</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">注册连接字符样式</div></td>
			  <td width="72%" class="hback"><input name="Reg_LinkCss" type="text" id="Reg_LinkCss" value="">
			  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			  <span style="color:#ff0000;">可以为定义的css样式名，也可以为图片</span></td>
			</tr>
			<tr>
			  <td width="28%" class="hback"><div align="right">取回密码连接字符样式</div></td>
			  <td width="72%" class="hback"><input name="Get_PassCss" type="text" id="Get_PassCss" value="">
				<input type="button" name="bnt_ChoosePic_rowBettween2"  value="选择图片" onClick="SelectFile();">
				<span style="color:#ff0000;">可以为定义的css样式名，也可以为图片</span>
			  </td>
			</tr>
		</table>
	  </td>
	</tr>		
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
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
			document.getElementById('Txt_loginType').innerHTML = '<font color=red>样式必须选择,如没有请先建立样式</font>';
			obj.Login_StyleID.focus();
			return false;
		}
	}
	switch (Str_disType)
	{
	case '0':
		var retV = '{FS:MF=UserLogin┆';
		retV+='标签方式$' + Str_disType + '┆';
		retV+='显示方式$' + obj.LoginDisStyle.value;	
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
		break;
	case '1':
		var retV = '{FS:MF=UserLogin┆';
		retV+='标签方式$' + Str_disType + '┆';
		retV+='引用样式$' + Str_StyleID + '┆';
		retV+='标签背景$' + obj.BGStyle.value + '┆';
		retV+='选择筐样式$' + obj.SelectStyle.value + '┆';
		retV+='选择筐菜单样式$' + obj.SelectBGCss.value + '┆';
		retV+='文本筐样式$' + obj.TextStyle.value + '┆';
		retV+='提交按钮样式$' + obj.ButtonStyle.value + '┆';
		retV+='取消按钮样式$' + obj.ResestCss.value + '┆';
		retV+='注册连接样式$' + obj.Reg_LinkCss.value + '┆';
		retV+='取回密码连接样式$' + obj.Get_PassCss.value;
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
      <td colspan="2" class="xingmu">会员列表</td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">会员类型</div></td>
      <td width="72%" class="hback">
	  <select name="UserType"  onChange="Select_VC_Class(this.options[this.selectedIndex].value);">
	   <option value="All">所有</option>
	    <option value=0>个人会员</option>
        <option value=1>企业会员</option>
	  </select>
	  <!--Select_VC_Class函数，调用出当选择企业会员的时候显示行业分类-->
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
</script>　
	  </td>
    </tr>
	    <tr name="VC_Class" id="VC_Class" style="font-family:宋体;display:none;">
      <td class="hback"><div align="right"><span style="color:#FF0000">企业会员行业分类</span></div></td>
      <td class="hback">
	  <select id="VClass"  name="VClass" style="width:20%">
            <% = VC_Class_List %>
        </select><span style="color:#FF0000">*若不选，则按所有排序</span>
	  </td>	  
    </tr>
    <tr>
      <td class="hback"><div align="right">列表类型</div></td>
      <td class="hback">
	  <select name="OrderBy">
	   <%=PrintOption("","RegTime:最新,LoginNum:登录次数,Hits:人气,Integral:会员积分,FS_Money:会员金币")%>
	  </select>
	  </td>
    </tr>
    <tr>
      <td width="28%" class="hback"><div align="right">会员性别</div></td>
      <td width="72%" class="hback">
	  <select name="UserSex">
	   <option value="All">所有</option>
	  <%=PrintOption("","0:男,1:女")%>
	  </select>
	  </td>
    </tr>


<!----------------------------->

    <tr>
      <td class="hback"><div align="right">调用数量</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示字数</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">每列数量</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      对DIV+CSS框架无效 </td>
    </tr>

    <tr>
      <td class="hback"><div align="right">日期格式</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select> </td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:宋体;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV控制</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:宋体;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">描述内容字数</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <label>
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
        </label>
      样式表中调用了描述此项有效</div></td>
    </tr>
	
    <tr>
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
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
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">每行样式</div></td>
      <td class="hback">
	   奇行<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   偶行<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		只针对表格而言,可直接填颜色#FF0000   
	  </td>
    </tr>

    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		if(isNaN(obj.TitleNumber.value))
		{alert('列表数量必须数字。');obj.TitleNumber.focus();return false;}
		if(obj.NewsStyle.value=='')
		{alert('引用样式必须填写。');obj.NewsStyle.focus();return false;}

		var retV = '{FS:ME=UserList┆';
		retV+='会员类型$' + obj.UserType.value + '┆';
		retV+='列表类型$' + obj.OrderBy.value + '┆';
		retV+='会员性别$' + obj.UserSex.value + '┆';
		retV+='数量$' + obj.TitleNumber.value + '┆';
		retV+='列数$' + obj.ColsNumber.value + '┆';
		retV+='字数$' + obj.leftTitle.value + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='输入格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='内容字数$' + obj.ContentNumber.value + '┆';
		retV+='引用样式$' + obj.NewsStyle.value+ '┆';
		retV+='奇数行样式$' + obj.PubType_JiTR.value + '┆';
		retV+='偶数行样式$' + obj.PubType_OuTR.value + '┆';
		//行业分类
		retV+='行业分类$' + obj.VClass.value + '';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 
 
 <%Sub CopyRight()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">版权信息</td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
  <script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=CopyRight┆';
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SubList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">子系统导航</td>
    </tr>
	<tr>
      <td width="28%" class="hback"><div align="right">分割符号</div></td>
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
      <td class="hback"><span class="tx">(各个子系统的导航连接请在系统参数--子系统维护里设置)，特别注意：各个子站系统的前台导航连接必须与子系统的参数配置里的域名相同</span></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=SubList┆';
		retV+='分割符-可以使用html语法$' + obj.SubName.value + '┆';
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
      <td colspan="2" class="xingmu">自定义表单</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">调用表单</div></td>
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
      <td class="hback"  align="center"><div align="right">表单样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="CustomFormSytleID"  name="CustomFormSytleID" style="width:40%">
		  	<option value="" selected>选择表单样式</option>
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
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.CustomFormSytleID.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">显示自定义表单的引用样式</span></div></td>
    </tr>
    <tr>
      <td class="hback"  align="center"><div align="right">表单数据样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="CustomDataSytleID"  name="CustomDataSytleID" style="width:40%">
		  	<option value="" selected>选择表单数据样式</option>
            <%
			Response.Write(label_style_List)
			 %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.CustomDataSytleID.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">显示自定义表单数据的引用样式</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">文本框CSS</div></td>
      <td class="hback"><input name="CustomFormTextCSS" type="text" id="CustomFormTextCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">下拉框CSS</div></td>
      <td class="hback"><input name="CustomFormSelectCSS" type="text" id="CustomFormSelectCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">其它对象CSS</div></td>
      <td class="hback"><input name="CustomFormOtherCSS" type="text" id="CustomFormOtherCSS"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		if(obj.CustomFormID.value!=''){
			var retV = '{FS:MF=CustomForm┆';
			retV+='调用表单$' + obj.CustomFormID.value + '┆';
			retV+='表单样式$' + obj.CustomFormSytleID.value + '┆';
			retV+='数据样式$' + obj.CustomDataSytleID.value + '';
			if(obj.CustomFormSytleID.value!=''){
				retV+='┆文本框CSS$' + obj.CustomFormTextCSS.value + '┆';
				retV+='下拉框CSS$' + obj.CustomFormSelectCSS.value + '┆';
				retV+='其它对象CSS$' + obj.CustomFormOtherCSS.value + '';
			}
			retV+='}';
			window.parent.returnValue = retV;
			window.close();
		}else{alert('请选择调用表单');}
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