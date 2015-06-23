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
	'session判断
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
            <td width="41%" class="xingmu"><strong>供求标签创建</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="hback">
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDList" target="_self">栏目信息</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=AreaList" target="_self">区域分类</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDClass" target="_self">栏目分类</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDClassList" target="_self">终极栏目</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDAreaList" target="_self">终极区域</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDPage" target="_self">浏览页面</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDSearch" target="_self">供求搜索</a></div></td>
    </tr>
	 <tr class="hback">
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDPubTypeList" target="_self">终极类别</a></div></td>
      <td width="14%"><div align="center"><a href="supply_C_Label.asp?type=SDChildClass" target="_self">子类调用</a></div></td>
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
      <td colspan="2" class="xingmu">栏目列表</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">类型</div></td>
      <td width="81%" class="hback">
        <select name="PubType" id="PubType">
			  <option value="" selected>全部</option>
			  <option value="0">供应</option>
			  <option value="1">求购</option>
			  <option value="2">合作</option>
			  <option value="3">代理</option>
			  <option value="4">其他</option>
        </select>
		<select name="PubPop" id="PubPop">
			<option value="0">正常</option>
			<option value="1">推荐</option>
			<option value="2">排行</option>
       </select>
       <input type="button" name="Submit" value="选择类型替换图片或样式" onClick="showhide(document.getElementById('typeTR'));">
      </td>
    </tr>
    <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">样式请直接填写自己定义的css样式名。图片请点击选择图片。也可以直接填写颜色代码，如:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">供应代替图片或样式</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">求购代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">合作代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">代理代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">其他代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
<!--所属区域分类--->
<!----->	
<!--所属栏目分类--->
<!----->	

    <tr>
      <td class="hback"><div align="right">每行样式</div></td>
      <td class="hback">
	   奇行<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   偶行<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		只针对表格而言,可直接填颜色#FF0000	  </td>
    </tr>
	
    <tr class="hback" id="ClassName_col"> 
      <td class="hback"  align="center"><div align="right">选择栏目</div></td>
      <td colspan="2" class="hback" > <input  name="ClassName" type="text" id="ClassName" size="12" readonly> 
        <input name="ClassID" type="hidden" id="ClassID"> <input name="button2" type="button" onClick="SelectClass();" value="选择栏目"> 
        <span class="tx">选择栏目，如果不选择，则显示所有的</span></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">包含子类</div></td>
      <td class="hback">
	  	<select name="ChildTF" id="ChildTF">
			<option value="1" selected>是</option>
			<option value="0">否</option>
		</select>
	  </td>
    </tr>
	<tr> 
      <td class="hback"><div align="right">时间范围</div></td>
      <td class="hback"><input name="DayNum" type="text" id="DayNum" value="0">天以内的信息，0为不限。</td>
    </tr>
	<tr>
      <td class="hback"><div align="right">调用数量</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  <input type="button" name="Submit" value="设置最新信息new标记" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">数字不能为空，不显示可以设置为0，两个都不为0取标记数量</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">标记天数</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			最新<input name="NewDayNum" type="text" id="NewDayNum" value="1">
			天以内的信息标题后显示new标记			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记数量</td>
            <td height="24" align="left" valign="middle" class="hback">
			最新<input name="NewInfoNum" type="text" id="NewInfoNum" value="10">
			条信息标题后显示new标记</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记图片</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">标题字数</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示列数</div></td>
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
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
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
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		
		var retV = '{FS:SD=SDList┆';
		retV+='供求类型$' + obj.PubType.value + '┆';
		retV+='供求属性$' + obj.PubPop.value + '┆';
		retV+='类型样式或图片$' + PubType_Style + '┆';
		retV+='奇数行样式$' + obj.PubType_JiTR.value + '┆';
		retV+='偶数行样式$' + obj.PubType_OuTR.value + '┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='包含子类$' + obj.ChildTF.value + '┆';
		retV+='时间范围$' + obj.DayNum.value + '┆';
		retV+='调用数量$' + obj.TitleNumber.value + '┆';
		retV+='标记天数$' + obj.NewDayNum.value + '┆';
		retV+='标记数量$' + obj.NewInfoNum.value + '┆';
		retV+='标记图片$' + obj.NewpicUrl.value + '┆';
		retV+='显示列数$' + obj.ColsNumber.value + '┆';
		retV+='标题字数$' + obj.leftTitle.value + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='内容字数$' + obj.ContentNumber.value + '┆';
		retV+='引用样式$' + obj.NewsStyle.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub AreaList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">区域列表</td>
    </tr>
	<!-- Ken 2006-12-25 -->
     <tr> 
      <td width="19%" class="hback" align="right">显示样式</td>
      <td width="81%" class="hback">
	  <select name="DisType" id="DisType" onChange="DisTypeTF(this.options[this.selectedIndex].value)">
		<option value="0" selected="selected">自适应区域显示</option>
	  	<option value="1">所有区域显示</option>
	  </select></td>
    </tr>
	<tr style="display:;" id="ShowInfoTR"> 
      <td colspan="2" height="20" align="center" class="hback">
	  <span id="ShowInfo"><font color=red>此方式指自动识别当前区域，然后显示其下一级所有区域.选择此方式,此标签只能放在供求区域分类模板中,否则会造成显示混乱</font></span>
	  </td>
    </tr>
	<tr id="FG_RE" style="display:;">
      <td width="19%" class="hback" align="right">分割符号</td>
      <td width="81%" class="hback">
	  <input type="text" name="FG_Info" id="FG_Info" value="" maxlength="20">
	  请不要包含html代码	  </td>
    </tr>
	<tr id="Class_ROW_Num" style="display:;"> 
      <td width="19%" class="hback" align="right">显示列数</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassRow" id="ClassRow" value="1" maxlength="20">
	  </td>
    </tr>
	<tr id="ClassLive_TF" style="display:none;"> 
      <td width="19%" class="hback" align="right">区域级数</td>
      <td width="81%" class="hback">
	  <input type="text" name="DisClassLive" id="DisClassLive" value="0" maxlength="20">
	  0为所有,根区域为1级,以此类推
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">区域样式</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassStyle" id="ClassStyle" value="" maxlength="20">
	  可以是样式名或者颜色代码,颜色代码请以#开头	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">上下行距</td>
      <td width="81%" class="hback">
	  <input type="text" name="TRHeight" id="TRHeight" value="20" maxlength="20">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">导航图片或文字</td>
      <td width="81%" class="hback">
	  <input type="text" name="NaviPic" id="NaviPic" value="" size=30 maxlength="50">
	  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择文件" onClick="SelectFile();">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">显示数字</td>
      <td width="81%" class="hback">
	  <select name="DisNumTF" id="DisNumTF">
	  	<option value="1" selected="selected">是</option>
	  	<option value="0">否</option>
	  </select>
	  是否显示该区域中信息总数	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">数字样式</td>
      <td width="81%" class="hback">
	  <input type="text" name="InfoNumCss" id="InfoNumCss" value="" maxlength="20">
	  可以是样式名或者颜色代码,颜色代码请以#开头	  </td>
    </tr>
    <tr> 
      <td width="19%" class="hback" align="right">弹出窗口</td>
      <td width="81%" class="hback">
	  <select name="OpenMode" id="OpenModes">
	  	<option value="1" selected="selected">是</option>
	  	<option value="0">否</option>
	  </select>
	  是否在新窗口中打开连接</td>
    </tr>
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
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
			document.getElementById('ShowInfo').innerHTML = '<font color=red>此方式指自动识别当前区域，然后显示其下一级所有区域.选择此方式,此标签只能放在供求区域分类模板中,否则会造成显示混乱</font>';
		}
		else
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = 'none';
			document.getElementById('Class_ROW_Num').style.display = 'none';
			document.getElementById('ClassLive_TF').style.display = '';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>此方式以竖形树状方式显示所有区域</font>';
		}
	}	
	function ok(obj)
	{
		if(isNaN(obj.ClassRow.value))
		{alert('显示列数必须数字。');obj.ClassRow.focus();return false;}
		if(isNaN(obj.TRHeight.value))
		{alert('行距必须数字。');obj.TRHeight.focus();return false;}
		if(obj.FG_Info.value.indexOf('┆')>-1)
		{alert('分割符号不能使用特定符号┆');obj.FG_Info.focus();return false;}
		if(obj.NaviPic.value.indexOf('┆')>-1)
		{alert('导航内容不能使用特殊符号┆');obj.NaviPic.focus();return false;} 
		//-------------------------------------------------------------
		var retV = '{FS:SD=AreaList┆';
		retV+='显示样式$' + obj.DisType.value + '┆';
		retV+='分隔符号$' + obj.FG_Info.value + '┆';
		retV+='显示列数$' + obj.ClassRow.value + '┆';
		retV+='区域样式$' + obj.ClassStyle.value + '┆';
		retV+='行距$' + obj.TRHeight.value + '┆';
		retV+='导航$' + obj.NaviPic.value + '┆';
		retV+='显示数字$' + obj.DisNumTF.value + '┆';
		retV+='数字样式$' + obj.InfoNumCss.value + '┆';
		retV+='弹出窗口$' + obj.OpenMode.value+ '┆';
		retV+='区域级数$' + obj.DisClassLive.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDClass()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">栏目列表</td>
    </tr>
	<!-- Ken 2006-12-25 -->
     <tr> 
      <td width="19%" class="hback" align="right">显示样式</td>
      <td width="81%" class="hback">
	  <select name="DisType" id="DisType" onChange="DisTypeTF(this.options[this.selectedIndex].value)">
		<option value="0" selected="selected">自适应栏目显示</option>
	  	<option value="1">所有栏目显示</option>
	  </select></td>
    </tr>
	<tr style="display:;" id="ShowInfoTR"> 
      <td colspan="2" height="20" align="center" class="hback">
	  <span id="ShowInfo"><font color=red>此方式指自动识别当前栏目，然后显示其下一级所有栏目.选择此方式,此标签只能放在供求栏目模板中,否则会造成显示混乱</font></span>
	  </td>
    </tr>
	<tr id="FG_RE" style="display:;">
      <td width="19%" class="hback" align="right">分割符号</td>
      <td width="81%" class="hback">
	  <input type="text" name="FG_Info" id="FG_Info" value="" maxlength="20">
	  请不要包含html代码	  </td>
    </tr>
	<tr id="Class_ROW_Num" style="display:;"> 
      <td width="19%" class="hback" align="right">显示列数</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassRow" id="ClassRow" value="1" maxlength="20">
	  </td>
    </tr>
	<tr id="ClassLive_TF" style="display:none;"> 
      <td width="19%" class="hback" align="right">栏目级数</td>
      <td width="81%" class="hback">
	  <input type="text" name="DisClassLive" id="DisClassLive" value="0" maxlength="20">
	  0为所有,根栏目为1级,以此类推
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">栏目样式</td>
      <td width="81%" class="hback">
	  <input type="text" name="ClassStyle" id="ClassStyle" value="" maxlength="20">
	  可以是样式名或者颜色代码,颜色代码请以#开头	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">上下行距</td>
      <td width="81%" class="hback">
	  <input type="text" name="TRHeight" id="TRHeight" value="20" maxlength="20">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">导航图片或文字</td>
      <td width="81%" class="hback">
	  <input type="text" name="NaviPic" id="NaviPic" value="" size=30 maxlength="50">
	  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择文件" onClick="SelectFile();">
	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">显示数字</td>
      <td width="81%" class="hback">
	  <select name="DisNumTF" id="DisNumTF">
	  	<option value="1" selected="selected">是</option>
	  	<option value="0">否</option>
	  </select>
	  是否显示该栏目中信息总数	  </td>
    </tr>
	<tr> 
      <td width="19%" class="hback" align="right">数字样式</td>
      <td width="81%" class="hback">
	  <input type="text" name="InfoNumCss" id="InfoNumCss" value="" maxlength="20">
	  可以是样式名或者颜色代码,颜色代码请以#开头	  </td>
    </tr>
    <tr> 
      <td width="19%" class="hback" align="right">弹出窗口</td>
      <td width="81%" class="hback">
	  <select name="OpenMode" id="OpenModes">
	  	<option value="1" selected="selected">是</option>
	  	<option value="0">否</option>
	  </select>
	  是否在新窗口中打开连接	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
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
			document.getElementById('ShowInfo').innerHTML = '<font color=red>此方式指自动识别当前栏目，然后显示其下一级所有栏目.选择此方式,此标签只能放在供求栏目模板中,否则会造成显示混乱</font>';
		}
		else
		{
			document.getElementById('ShowInfoTR').style.display = '';
			document.getElementById('FG_RE').style.display = 'none';
			document.getElementById('Class_ROW_Num').style.display = 'none';
			document.getElementById('ClassLive_TF').style.display = '';
			document.getElementById('ShowInfo').innerHTML = '<font color=red>此方式以竖形树状方式显示所有栏目</font>';
		}
	}	
	function ok(obj)
	{
		if(isNaN(obj.ClassRow.value))
		{alert('显示列数必须数字。');obj.ClassRow.focus();return false;}
		if(isNaN(obj.TRHeight.value))
		{alert('行距必须数字。');obj.TRHeight.focus();return false;}
		if(obj.FG_Info.value.indexOf('┆')>-1)
		{alert('分割符号不能使用特定符号┆');obj.FG_Info.focus();return false;}
		if(obj.NaviPic.value.indexOf('┆')>-1)
		{alert('导航内容不能使用特殊符号┆');obj.NaviPic.focus();return false;} 
		//-------------------------------------------------------------
		var retV = '{FS:SD=SDClass┆';
		retV+='显示样式$' + obj.DisType.value + '┆';
		retV+='分隔符号$' + obj.FG_Info.value + '┆';
		retV+='显示列数$' + obj.ClassRow.value + '┆';
		retV+='栏目样式$' + obj.ClassStyle.value + '┆';
		retV+='行距$' + obj.TRHeight.value + '┆';
		retV+='导航$' + obj.NaviPic.value + '┆';
		retV+='显示数字$' + obj.DisNumTF.value + '┆';
		retV+='数字样式$' + obj.InfoNumCss.value + '┆';
		retV+='弹出窗口$' + obj.OpenMode.value+ '┆';
		retV+='栏目级数$' + obj.DisClassLive.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
  <%Sub SDAreaList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">区域终极分类</td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right">分页数量</div></td>
      <td width="81%" class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">信息标题字数</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示列数</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      对DIV+CSS框架无效 
	  <input type="button" name="Submit" value="选择类型替换图片或样式" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">样式请直接填写自己定义的css样式名。图片请点击选择图片。也可以直接填写颜色代码，如:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">供应代替图片或样式</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">求购代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">合作代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">代理代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">其他代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">日期格式</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
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
      <td class="hback"  align="center"><div align="right">分页方式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">方式1</option>
			  <option value="2">方式2</option>
			  <option value="3">方式3</option>
			  <option value="4" selected="selected">方式4</option>
         </select>
		 <span class="tx">分页连接样式</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">可为颜色代码,不需要#符号</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDAreaList┆';
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
		retV+='引用样式$' + obj.NewsStyle.value + '┆';
		retV+='分页方式$' + obj.Flg_PageType.value + '┆';
		retV+='分页连接样式$' + obj.Flg_PageCss.value + '┆';
		retV+='类型样式或者图片$' + PubType_Style;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDClassList()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">栏目终极分类</td>
    </tr>
    <tr>
      <td width="19%" class="hback">
        <div align="right">分页数量</div>
      </td>
      <td width="81%" class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10">
	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">信息标题字数</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示列数</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      对DIV+CSS框架无效 
	  <input type="button" name="Submit" value="选择类型替换图片或样式" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">样式请直接填写自己定义的css样式名。图片请点击选择图片。也可以直接填写颜色代码，如:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">供应代替图片或样式</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">求购代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">合作代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">代理代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">其他代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">日期格式</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
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
      <td class="hback"  align="center"><div align="right">分页方式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">方式1</option>
			  <option value="2">方式2</option>
			  <option value="3">方式3</option>
			  <option value="4" selected="selected">方式4</option>
         </select>
		 <span class="tx">分页连接样式</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">可为颜色代码,不需要#符号</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDClassList┆';
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
		retV+='引用样式$' + obj.NewsStyle.value + '┆';
		retV+='分页方式$' + obj.Flg_PageType.value + '┆';
		retV+='分页连接样式$' + obj.Flg_PageCss.value + '┆';
		retV+='类型样式或者图片$' + PubType_Style;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDPage()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">终极页面浏览</td>
    </tr>    
    <tr>
      <td width="19%"  align="center" class="hback"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td width="81%" class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SD=SDPage┆';
		retV+='引用样式$' + obj.NewsStyle.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 <%End Sub%>
 <%Sub SDSearch()%>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="xingmu">搜索</td>
    </tr>    
    
    <tr>
      <td class="hback"><div align="right">显示栏目</div></td>
      <td class="hback"><label>
        <select name="ShowClass" id="ShowClass">
          <option value="1" selected>显示</option>
          <option value="0">不显示</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示地区</div></td>
      <td class="hback"><select name="ShowArea" id="ShowArea">
        <option value="1" selected>显示</option>
        <option value="0">不显示</option>
      </select></td>
    </tr>
    <tr>
      <td width="19%" class="hback"><div align="right"></div></td>
      <td width="81%" class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:SD=SDSearch┆';
		retV+='显示栏目$' + obj.ShowClass.value + '┆';
		retV+='显示地区$' + obj.ShowArea.value;
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
      <td colspan="2" class="xingmu">供求类型终极列表</td>
    </tr>
    <tr>
      <td width="19%" class="hback">
        <div align="right">分页数量</div>
      </td>
      <td width="81%" class="hback">
	  <input name="PageInfoNum" type="text" id="PageInfoNum" value="10">
	  <input type="button" name="Submit" value="设置最新信息new标记" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">数字不能为空，不显示可以设置为0，两个都不为0取标记数量</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">标记天数</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			最新<input name="NewDayNum" type="text" id="NewDayNum" value="1">天以内的信息标题后显示new标记			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记数量</td>
            <td height="24" align="left" valign="middle" class="hback">
			最新<input name="NewInfoNum" type="text" id="NewInfoNum" value="10">
			条信息标题后显示new标记</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记图片</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">奇数行样式</div></td>
      <td class="hback">
	   <input type="text" name="PubType_JiTR" size="20" maxlength="100" value="">
		DIV无效,可为颜色#FF0000、样式名、图片路径	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">偶数行样式</div></td>
      <td class="hback">
	   <input type="text" name="PubType_OuTR" size="20" maxlength="100" value="">
		DIV无效,可为颜色#FF0000、样式名、图片路径	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">信息标题字数</div></td>
      <td class="hback"><input name="TitleNum" type="text" id="TitleNum" value="30"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">显示列数</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      对DIV+CSS框架无效 
	  <input type="button" name="Submit" value="选择类型替换图片或样式" onClick="showhide(document.getElementById('typeTR'));">
	  </td>
    </tr>
	 <tr id="typeTR" style="display:none">
      <td class="hback" colspan="2">
	  
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">样式请直接填写自己定义的css样式名。图片请点击选择图片。也可以直接填写颜色代码，如:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">供应代替图片或样式</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">求购代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">合作代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">代理代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">其他代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">日期格式</div></td>
      <td class="hback"><input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD">
      <span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">输出格式</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,'');">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
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
      <td class="hback"  align="center"><div align="right">分页方式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select name="Flg_PageType" id="Flg_PageType">
			  <option value="1">方式1</option>
			  <option value="2">方式2</option>
			  <option value="3">方式3</option>
			  <option value="4" selected="selected">方式4</option>
         </select>
		 <span class="tx">分页连接样式</span>
		 <input name="Flg_PageCss" type="text" id="Flg_PageCss">
		 <span class="tx">可为颜色代码,不需要#符号</span>
		 </div></td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
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
		var PubType_Style='';
		for (var ii=0;ii<obj.PubType_pic.length;ii++)
		{
			PubType_Style+=','+obj.PubType_pic[ii].value;			
		}
		PubType_Style=PubType_Style.substring(1,PubType_Style.length);
		//******************************************
		var retV = '{FS:SD=SDPubTypeList┆';
		retV+='分页数量$' + obj.PageInfoNum.value + '┆'; 
		retV+='最新标记$' + obj.NewDayNum.value + ','+ obj.NewInfoNum.value + ','+ obj.NewpicUrl.value + '┆';
		retV+='奇偶行样式$' + obj.PubType_JiTR.value + ','+ obj.PubType_OuTR.value + '┆';
		retV+='标题字数$' + obj.TitleNum.value + '┆';
		retV+='显示列数$' + obj.ColsNumber.value + '┆';
		retV+='类型替换样式$' + PubType_Style + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='输入格式$' + obj.out_char.value + '┆';
		retV+='Div样式$' + obj.DivID.value + ','+ obj.Divclass.value + ','+ obj.ulid.value + ','+ obj.ulclass.value + ','+ obj.liid.value + ','+ obj.liclass.value + '┆';
		retV+='内容字数$' + obj.ContentNumber.value + '┆';
		retV+='引用样式$' + obj.NewsStyle.value + '┆';
		retV+='分页方式$' + obj.Flg_PageType.value + '┆';
		retV+='分页连接样式$' + obj.Flg_PageCss.value;
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
      <td colspan="2" class="xingmu">子类调用</td>
    </tr>
    <tr>
      <td width="20%" class="hback"><div align="right">分类方式</div></td>
      <td width="81%" class="hback">
        <select name="TypeRuler" id="TypeRuler">
			  <option value="0">按栏目分类</option>
			  <option value="1">按区域分类</option>
       </select>
      </td>
    </tr>
	<tr>
      <td colspan="2" align="center" valign="middle" class="hback"><span style="color:red">注意:如果选择栏目分类,请将此标签放于供求栏目模板上;如果选择区域分类,请将此标签放与供求区域分类模板上,其他模板不要用此标签，否则会引起显示错误.</span></td>
    </tr>
	<tr>
      <td class="hback"><div align="right">显示子分类名</div></td>
      <td class="hback">
	  	<select name="ChCNameDisTF" id="ChCNameDisTF">
			<option value="1" selected>是</option>
			<option value="0">否</option>
		</select>
		子分类名显示样式
		<input name="ChCNameCss" type="text" id="ChCNameCss" value="">
	  </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">子类信息调用数量</div></td>
      <td class="hback"><input name="InfoNum" type="text" id="InfoNum" value="10">
	  <input type="button" name="Submit" value="设置最新信息new标记" onClick="showhide(document.getElementById('NewsTF'));">
	  </td>
    </tr>
    <tr id="NewsTF" style="display:none;">
      <td colspan="2" align="center" valign="middle" class="hback">
        <table width="98%" border="0" cellspacing="1" cellpadding="5" class="table">
          <tr>
            <td height="24" colspan="2" align="center" valign="middle" class="hback">数字不能为空，不显示可以设置为0，两个都不为0取标记数量</td>
          </tr>
		  <tr>
            <td width="18%" height="24" align="right" valign="middle" class="hback">标记天数</td>
            <td width="82%" height="24" align="left" valign="middle" class="hback">
			最新<input name="NewDayNum" type="text" id="NewDayNum" value="1">天以内的信息标题后显示new标记			</td>
          </tr>
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记数量</td>
            <td height="24" align="left" valign="middle" class="hback">
			最新<input name="NewInfoNum" type="text" id="NewInfoNum" value="3">
			条信息标题后显示new标记</td>
          </tr> 
          <tr>
            <td height="24" align="right" valign="middle" class="hback">标记图片</td>
            <td height="24" align="left" valign="middle" class="hback">
			<input type="text" name="NewpicUrl" onMouseOver="title=value;" size=30 maxlength="50" value="">
		    <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
			</td>
          </tr>
        </table>
      </td>
    </tr>
	<tr> 
      <td class="hback"><div align="right">时间范围</div></td>
      <td class="hback"><input name="DayNum" type="text" id="DayNum" value="0">天以内的信息，0为不限。</td>
    </tr>
	<tr>
      <td width="20%" class="hback"><div align="right">信息类型</div></td>
      <td width="81%" class="hback">
        <select name="PubType" id="PubType">
			  <option value="" selected>全部</option>
			  <option value="0">供应</option>
			  <option value="1">求购</option>
			  <option value="2">合作</option>
			  <option value="3">代理</option>
			  <option value="4">其他</option>
        </select>
		<select name="PubPop" id="PubPop">
			<option value="0">正常</option>
			<option value="1">推荐</option>
			<option value="2">排行</option>
       </select>
       <input type="button" name="Submit" value="选择类型替换图片或样式" onClick="showhide(document.getElementById('TypeTR'));">
      </td>
    </tr>
    <tr id="TypeTR" style="display:none">
      <td class="hback" colspan="2">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
		  <td class="xingmu" colspan="2">样式请直接填写自己定义的css样式名。图片请点击选择图片。也可以直接填写颜色代码，如:<font color="red">#FFFFFF</font></td>
		</tr>
		 <tr>
		  <td class="hback">
		  <div align="right">供应代替图片或样式</div></td>
		  <td class="hback">		  
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();"></td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">求购代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">合作代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">代理代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
		
		 <tr>
		  <td class="hback">
		  <div align="right">其他代替图片或样式</div></td>
		  <td class="hback">
		  <input type="text" name="PubType_pic" onMouseOver="title=value;" size=30 maxlength="50" value="">
		  <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">		  </td>
		</tr>
	   </table>	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">标题字数</div></td>
      <td class="hback"><input name="TitleNum" type="text" id="TitleNum" value="30"></td>
    </tr>
	<tr>
      <td class="hback"><div align="right">每行样式</div></td>
      <td class="hback">
	   奇行<input type="text" name="PubType_JiTR" size="10" maxlength="50" value="">
	   偶行<input type="text" name="PubType_OuTR" size="10" maxlength="50" value="">
		只针对表格而言,可直接填颜色#FF0000	  </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">子类显示列数</div></td>
      <td class="hback"><input name="ClassRowNum" type="text" id="ClassRowNum" value="1" size="10">
      对DIV+CSS框架无效 </td>
    </tr>
	<tr>
      <td class="hback"><div align="right">信息显示列数</div></td>
      <td class="hback"><input name="RowNum" type="text" id="RowNum" value="1" size="10">
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
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value,1);">
          <option value="out_Table">普通格式</option>
          <option value="out_DIV">DIV+CSS格式</option>
          
        </select></td>
    </tr>
    <tr id="Table_ID" style="display:;">
      <td class="hback"><div align="right">表格样式</div></td>
      <td class="hback"><input name="TableCss" type="text" id="TableCss" value="">
      </td>
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
        <input name="ContentNumber" type="text" id="ContentNumber" value="200">
      样式表中调用了描述此项有效</div></td>
    </tr>
	 <tr>
      <td class="hback"  align="center"><div align="right">更多连接</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left">
        <input name="MoreLinkStr" type="text" id="MoreLinkStr" value="更多..>>">
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="SelectFile();">
		更多连接样式
		<input name="MoreLinkCss" type="text" id="MoreLinkCss" value="">
      </div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">弹出窗口</div></td>
      <td class="hback">
	  	<select name="OpenMode" id="OpenMode">
			<option value="1" selected>是</option>
			<option value="0">否</option>
		</select>
	  </td>
    </tr>
	<tr>
      <td class="hback"  align="center"><div align="right">引用样式</div></td>
      <td class="hback"  colspan="2" align="center"><div align="left"> 
          <select id="NewsStyle"  name="NewsStyle" style="width:40%">
            <% = label_style_List %>
          </select>
          <input name="button3" type="button" id="button" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
          <span class="tx">请在各个子系统中建立前台页面显示样式</span></div></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
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
		var retV = '{FS:SD=SDChildClass┆';
		retV+='分类方式$' + obj.TypeRuler.value + '┆';
		retV+='显示分类名称$' + obj.ChCNameDisTF.value + '┆';  
		retV+='分类名称样式$' + obj.ChCNameCss.value + '┆';
		retV+='分类信息数量$' + obj.InfoNum.value + '┆';  
		retV+='最新信息标记$' + NewsInfo_BJ_Str + '┆';
		retV+='时间范围$' + obj.DayNum.value + '┆';  
		retV+='信息类型$' + InfoType_Str + '┆'; 
		retV+='类型样式或图片$' + PubType_Style + '┆';
		retV+='标题字数$' + obj.TitleNum.value + '┆';
		retV+='奇偶行样式$' + TR_Style_Str + '┆';  
		retV+='显示列数$' + obj.RowNum.value + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='Div样式$' + DIV_Style_Str + '┆';  
		retV+='描述内容字数$' + obj.ContentNumber.value + '┆'; 
		retV+='更多连接$' + obj.MoreLinkStr.value + '┆';
		retV+='更多连接样式$' + obj.MoreLinkCss.value + '┆';
		retV+='引用样式$' + obj.NewsStyle.value + '┆';
		retV+='弹出窗口$' + obj.OpenMode.value + '┆';     
		retV+='子类列数$' + obj.ClassRowNum.value + '┆';
		retV+='表格样式$' + obj.TableCss.value;
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






