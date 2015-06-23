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
	'session判断
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
	'获取新闻子系统自由标签下拉列表
	'================================
	Function GetNewsFreeList(SysType)
	Dim Rs,Sql
	Sql = "Select LabelID,LabelName From FS_MF_FreeLabel Where ID > 0 And SysType = '" & SysType & "'"
	Set Rs = Conn.ExeCute(Sql)
	GetNewsFreeList = "<select name=""FreeList"" id=""FreeList"">" & vbnewline
	GetNewsFreeList = GetNewsFreeList & "<option value="""">选择自由标签</option>"
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
			<td height="27"  align="Left" class="xingmu">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="20%" class="xingmu"><strong>常规标签创建</strong></td>
						<td width="80%">
							<div align="right">
								<input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
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
				<div align="center"><a href="News_C_Label.asp?type=ReadNews" target="_self">新闻浏览</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=AllCode" target="_self">通用调用</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=FlashFilt" target="_self">FLASH幻灯</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=NorFilt" target="_self">轮换幻灯</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=siteMap" target="_self">站点地图</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=TodayWord" target="_self">文字头条</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=TodayPic" target="_self">图片头条</a></div>
			</td>
			<td width="13%">
				<div align="center"><a href="News_C_Label.asp?type=Search" target="_self">搜索表单</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassNavi" target="_self">栏目导航</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SpecialNavi" target="_self">专题导航</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=RssFeed" target="_self">RSS聚合</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SpecialCode" target="_self">专题调用</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassCode" target="_self">栏目调用</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=infoStat" target="_self">信息统计</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=c_news" target="_self">相关新闻</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=c_ext" target="_self"></a><a href="News_C_Label.asp?type=OldNews" target="_self">归档标签</a></div>
			</td>
		</tr>
		<tr class="hback">
			<td>
				<div align="center"><a href="News_C_Label.asp?type=ClassInfo" target="_self">栏目信息</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=FreeLabel" target="_self">自由标签</a></div>
			</td>
			<td>
				<div align="center"><a href="News_C_Label.asp?type=SingleClass" target="_self">单页栏目</a></div>
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
				<div align="right">引用样式</div>
			</td>
			<td width="79%" class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
				<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span> </td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示日期格式</div>
			</td>
			<td class="hback">
				<input name="DateStyle" type="text" id="DateStyle" value="YY02-MM-DD HH:MI:SS" size="28">
				<span class="tx">格式:YY02代表2位的年份(如06表示2006年),YY04表示4位数的年份(2006)，MM代表月，DD代表日，HH代表小时，MI代表分，SS代表秒</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:NS=ReadNews┆';
	retV+='引用样式$' + obj.NewsStyle.value + '┆';
	retV+='日期格式$' + obj.DateStyle.value + '';
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
				<div align="right">根据条件</div>
			</td>
			<td class="hback">
				<label>
				<select name="ifelse" id="ifelse" onChange="if(this.options[this.selectedIndex].value=='3')ifelsearea.style.display=''; else ifelsearea.style.display='none';">
					<option value="0">新闻标题相关</option>
					<option value="1" selected>新闻关键字相关</option>
					<option value="2">关键字和标题相关</option>
					<option value="3">标题关联某个关键字</option>
				</select>
				<span id="ifelsearea" style="display:none">
				<select name="ifelsetype" id="ifelsetype">
					<%=PrintOption("=","==:等于,&lt;&gt;:不等于,*s:以开头,*e:以结尾,**:像")%>
				</select>
				<input type="text" maxlength="3" name="ifelsevalue" id="ifelsevalue" title="指定源新闻的关键字序号.如关键字:江苏,周庄图片,指定1表示江苏,2表示周庄图片"
		onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))" onMouseOver="this.focus();">
				</span> </label>
			</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right">显示数量</div>
			</td>
			<td width="79%" class="hback">
				<label>
				<input name="titleNumber" type="text" id="titleNumber" value="10">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示标题字数</div>
			</td>
			<td class="hback">
				<label>
				<input name="leftTitle" type="text" id="leftTitle" value="40">
				中文占2个字符</label>
			</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right">引用样式</div>
			</td>
			<td width="79%" class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
				<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span> </td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var value3='';
	if(obj.ifelse.value=='3')
	{
		if (obj.ifelsevalue.value=='')	{alert('请指定源新闻的关键字序号.');obj.ifelsevalue.focus();return false;}
		value3 = obj.ifelse.value+''+obj.ifelsetype.value+obj.ifelsevalue.value;
	}
	else
	value3 = obj.ifelse.value;
	
	var retV = '{FS:NS=c_news┆';
	retV+='根据条件$' + value3 + '┆';
	retV+='显示数量$' + obj.titleNumber.value + '┆';
	retV+='标题字数$' + obj.leftTitle.value + '┆';
	retV+='引用样式$' + obj.NewsStyle.value + '';
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
			<td class="hback">此标签没参数</td>
		</tr>
		<tr>
			<td width="21%" class="hback">
				<div align="right"></div>
			</td>
			<td width="79%" class="hback">
				<input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
function ok(obj)
{
	var retV = '{FS:NS=OldNews┆';
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
				        document.write('<option style="background:#'+color[j]+color[k]+color[i]+'" value="#'+color[j]+color[k]+color[i]+'" selected>　　</option>');
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
				<div align="right">选择栏目</div>			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName4" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID4">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx"></span>
				<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">否</option>
					<option value="1">是</option>
				</select>
				包含子类				</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">调用数量</div>			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题字数</div>			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片尺寸(高度,宽度)</div>			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				格式120,100。请正确使用格式</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">文本高度</div>			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">背景颜色</div></td>
		  <td class="hback"><div align="left"><select name="FlashBG" id="FlashBG" style="width:200px;" class="form"><script language="javascript" type="text/javascript">getColorOptions();</script></select>
		    </div></td>
	  </tr>
		<tr>
			<td class="hback">
				<div align="right"></div>			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=FlashFilt┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='数量$' + obj.NewsNumber.value + '┆';
		retV+='标题字数$' + obj.TitleNumber.value + '┆';
		retV+='图片尺寸$' + obj.p_size.value +  '┆';
		retV+='文本高度$' + obj.TextSize.value + '┆';
		retV+='包含子类$' + obj.containSubClass.value + '┆';
		retV+='背景颜色$' + obj.FlashBG.value +  '';
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
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName4" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID4">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx"></span>
				<span class="tx"></span>
				<select name="containSubClass" id="containSubClass">
					<option value="0" selected="selected">否</option>
					<option value="1">是</option>
				</select>
				包含子类
		  </td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示标题</div>
			</td>
			<td class="hback">
				<select name="ShowTitle" id="ShowTitle">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">调用数量</div>
			</td>
			<td class="hback">
				<input  name="NewsNumber" type="text" id="NewsNumber" value="5" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题字数</div>
			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="30" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片尺寸(高度,宽度)</div>
			</td>
			<td class="hback">
				<input  name="p_size" type="text" id="p_size" value="120,100" size="12">
				格式120,100。请正确使用格式</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">文本高度</div>
			</td>
			<td class="hback">
				<input  name="TextSize" type="text" id="Picsize" value="20" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>
			</td>
			<td class="hback">
				<input  name="CSS" type="text" id="CSS" size="12">
			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">窗口打开方式</div></td>
		  <td class="hback"><div align="left">
		    <select name="OpenType" id="OpenType">
		      <option value="0">原窗口</option>
		      <option value="1">新窗口</option>
	        </select>
		    </div></td>
	  </tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=NorFilt┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='数量$' + obj.NewsNumber.value + '┆';
		retV+='标题字数$' + obj.TitleNumber.value + '┆';
		retV+='图片尺寸$' + obj.p_size.value +  '┆';
		retV+='CSS样式$' + obj.CSS.value +  '┆';
		retV+='文本高度$' + obj.TextSize.value +  '┆';
		retV+='显示标题$' + obj.ShowTitle.value +  '┆';
		retV+='包含子类$' + obj.containSubClass.value +  '┆';
		retV+='窗口打开方式$' + obj.OpenType.value +  '';
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
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx"></span></td>
		</tr>
		<tr style="display:none">
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>
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
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=siteMap┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '';
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
				<div align="right">日期搜索</div>
			</td>
			<td width="78%" class="hback">
				<select name="DateShow"  id="DateShow">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">搜索栏目</div>
			</td>
			<td class="hback">
				<select name="classShow"  id="classShow">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示查询类型</div>
			</td>
			<td class="hback">
				<select name="selectShow"  id="selectShow">
					<option value="1" selected>是</option>
					<option value="0">否</option>
				</select>
				  选否则不显示按标题,内容,作者等的下拉框而采用默认	
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=Search┆';
		retV+='显示日期$' + obj.DateShow.value + '┆';
		retV+='显示栏目$' + obj.classShow.value + '┆';
		retV+='显示查询类型$' + obj.selectShow.value + '';
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
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=infoStat┆';
		retV+='显示方向$' + obj.cols.value + '';
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
				<div align="right">选择栏目</div>
			</td>
			<td class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button222" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx"></span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=TodayPic┆';
		retV+='栏目$' + obj.ClassID.value + '';
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
				<div align="right">选择栏目</div>			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx"></span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">输出格式</div>			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
				</select>			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">列数</div>			</td>
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
				显示评论
				<label>
				<select name="ShowReview" id="ShowReview">
					<option value="1">显示</option>
					<option value="0" selected>不显示</option>
				</select>
				</label>			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">调用数量</div>			</td>
			<td class="hback">
				<input  name="TitleNumber" type="text" id="TitleNumber" value="1" size="12" >
				标题字数
				<input  name="lefttitle" type="text" id="lefttitle" value="30" size="12" >			</td>
		</tr>
		<tr>
		  <td class="hback"><div align="right">对齐方式</div></td>
		  <td class="hback"><select name="Align" id="Align">
            <option value="left">左</option>
            <option value="center">中</option>
            <option value="right">右</option>
          </select>
		  (仅对普通显示方式有效，如果是DIV请用样式控制)</td>
	  </tr>
		<tr>
			<td class="hback">
				<div align="right"></div>			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=TodayWord┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='列数$' + obj.cols.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '┆';
		retV+='调用数量$' + obj.TitleNumber.value + '┆';
		retV+='标题字数$' + obj.lefttitle.value + '┆';
		retV+='显示评论$' + obj.ShowReview.value + '┆';
		retV+='对齐方式$' + obj.Align.value + '';
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
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx">如果不选择，那么在某个类就调用某个类的导航</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">输出格式</div>
			</td>
			<td class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题导航</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
		 如果要采用图片请直接填写图片地址,注意不能附带 IMG标签。</label></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=ClassNavi┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='方向$' + obj.cols.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='标题CSS$' + obj.Titlecss.value + '┆';
		retV+='标题导航$' + obj.TitleNavi.value + '';
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
				<div align="right">输出格式</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">专题CSS</div>
			</td>
			<td class="hback">
				<input  name="Titlecss" type="text" id="Titlecss" size="12" >
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题导航图片/文字</div>
			</td>
			<td class="hback">
				<label>
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
				请使用html语法</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=SpecialNavi┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='方向$' + obj.cols.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='CSS$' + obj.Titlecss.value + '┆';
		retV+='导航$' + obj.TitleNavi.value + '';
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
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button22" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx">如果不选择，那么在某个类就调用某个类的RSS</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=RssFeed┆';
		retV+='栏目$' + obj.ClassID.value + '';
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
				<div align="right">选择专题</div>
			</td>
			<td width="78%" class="hback">
				<select id="specialEName"  name="specialEName">
					<option value="">请选择专题</option>
					<% = label_special_List %>
				</select>
				<span class="tx">必须选择</span></td>
		</tr>
		<tr>
			<td width="22%" class="hback">
				<div align="right">输出格式</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示图片</div>
			</td>
			<td class="hback">
				<select name="PicTF" id="PicTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				图片高度及宽度
				<input name="PicSize" type="text" id="PicSize" value="120,100" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示专题导航内容</div>
			</td>
			<td class="hback">
				<select name="NaviTF" id="NaviTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				<input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片CSS</div>
			</td>
			<td class="hback">
				<input name="PicCSS" type="text" id="PicCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">名称CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
				内容CSS
				<input name="ContentCSS" type="text" id="ContentCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
				只对table格式有效 　　导航
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		if(obj.specialEName.value=='')
		{
		alert('请选择专题');
		obj.specialEName.focus();
		return false;
		}
		var retV = '{FS:NS=SpecialCode┆';
		retV+='专题$' + obj.specialEName.value + '┆';
		retV+='图片显示$' + obj.PicTF.value + '┆';
		retV+='图片尺寸$' + obj.PicSize.value + '┆';
		retV+='导航内容$' + obj.NaviTF.value + '┆';
		retV+='导航内容字数$' + obj.NaviNumber.value + '┆';
		retV+='专题名称CSS$' + obj.TitleCSS.value + '┆';
		retV+='导航内容CSS$' + obj.ContentCSS.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='排列方式$' + obj.cols.value + '┆';
		retV+='导航$' + obj.TitleNavi.value + '┆';
		retV+='图片css$' + obj.PicCSS.value + '';
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
				<div align="right">选择栏目</div>
			</td>
			<td width="78%" class="hback">
				<input  name="ClassName" type="text" id="ClassName" size="12" readonly>
				<input name="ClassID" type="hidden" id="ClassID">
				<input name="button223" type="button" onClick="SelectClass();" value="选择栏目">
				<span class="tx">如果不选择，则调用所有</span></td>
		</tr>
		<tr>
			<td width="22%" class="hback">
				<div align="right">输出格式</div>
			</td>
			<td width="78%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示图片</div>
			</td>
			<td class="hback">
				<select name="PicTF" id="PicTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				图片高度及宽度
				<input name="PicSize" type="text" id="PicSize" value="120,100" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片CSS</div>
			</td>
			<td class="hback">
				<input name="PicCSS" type="text" id="PicCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">显示栏目导航内容</div>
			</td>
			<td class="hback">
				<select name="NaviTF" id="NaviTF">
					<option value="1" selected>显示</option>
					<option value="0">不显示</option>
				</select>
				<input name="NaviNumber" type="text" id="NaviNumber" value="200" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">名称CSS</div>
			</td>
			<td class="hback">
				<input name="TitleCSS" type="text" id="TitleCSS" size="12">
				内容CSS
				<input name="ContentCSS2" type="text" id="ContentCSS" size="12">
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">排列方式</div>
			</td>
			<td class="hback">
				<select name="cols"  id="cols">
					<option value="0" selected>横向</option>
					<option value="1">纵向</option>
				</select>
				导航
				<input name="TitleNavi" type="text" id="TitleNavi" value="·">
				只对table格式有效</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=ClassCode┆';
		retV+='栏目$' + obj.ClassID.value + '┆';
		retV+='图片显示$' + obj.PicTF.value + '┆';
		retV+='图片尺寸$' + obj.PicSize.value + '┆';
		retV+='导航内容$' + obj.NaviTF.value + '┆';
		retV+='导航内容字数$' + obj.NaviNumber.value + '┆';
		retV+='栏目名称CSS$' + obj.TitleCSS.value + '┆';
		retV+='导航内容CSS$' + obj.ContentCSS.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='排列方式$' + obj.cols.value + '┆';
		retV+='导航$' + obj.TitleNavi.value + '┆';
		retV+='图片CSS$' + obj.PicCSS.value + '';
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
			<td class="hback">说明：此标签可以插入任何地方。如果插入栏目分类模板，就调用本栏目的数据，否则调用所有的数据</td>
		</tr>
		<tr>
			<td width="20%" class="hback">
				<div align="right">类型</div>
			</td>
			<td width="80%" class="hback">
				<label>
				<select name="NewsType" id="NewsType">
					<option value="HotNews">热点</option>
					<option value="LastNews">最新</option>
					<option value="RecNews">推荐</option>
					<option value="MarqueeNews">滚动</option>
					<option value="JcNews">精彩</option>
				</select>
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">调用数</div>
			</td>
			<td class="hback">
				<label>
				<input name="CodeNumber" type="text" id="CodeNumber" value="10">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">标题显示字数</div>
			</td>
			<td class="hback">
				<label>
				<input name="leftNumber" type="text" id="leftNumber" value="40">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">点击大于多少次</div>
			</td>
			<td class="hback">
				<label>
				<input name="ClickNum" type="text" id="ClickNum" value="0">
				此项只对热点新闻有效，如果选择为0，则不限制</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">图片新闻</div>
			</td>
			<td class="hback">
				<label>
				<select name="PicTF" id="PicTF">
					<option value="1">是</option>
					<option value="0" selected>否</option>
				</select>
				选择是，则只显示图片新闻。选择否，则显示所有图片和文字新闻</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">新闻排列列数</div>
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
				<div align="right">日期格式</div>
			</td>
			<td class="hback">
				<label>
				<input name="DateStyle" type="text" id="DateStyle" value="YY02年MM月DD日">
				</label>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">引用样式</div>
			</td>
			<td class="hback">
				<select id="NewsStyle"  name="NewsStyle" style="width:40%">
					<% = label_style_List %>
				</select>
				<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
				<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span></td>
		</tr>
		<tr class="hback" >
			<td class="hback" align="center" height="30">
				<div align="right">其他</div>
			</td>
			<td height="30"  colspan="2" align="center" class="hback">
				<div align="left"> 内容显示最多字符
					<input name="contentnumber" type="text" id="contentnumber" value="200" size="12">
					新闻导读显示最多字符
					<input name="navinumber" type="text" id="navinumber" value="200" size="12">
					(中文2个字符)</div>
			</td>
		</tr>
		<tr>
			<td width="20%" class="hback">
				<div align="right">输出格式</div>
			</td>
			<td width="80%" class="hback">
				<select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
					<option value="out_Table">普通格式</option>
					<option value="out_DIV">DIV+CSS格式</option>
					
				</select>
			</td>
		</tr>
		<tr class="hback"  id="div_id" style="font-family:宋体;display:none;" >
			<td rowspan="3"  align="center" class="hback">
				<div align="right"></div>
				<div align="right">DIV控制</div>
			</td>
			<td colspan="3" class="hback" >&lt;div id=&quot;
				<input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的ID号，请在CSS中预先定义。不能为空">
				&quot; class=&quot;
				<input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成DIV调用的Class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback" id="ul_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;ul id=&quot;
				<input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成ul调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成ul调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr class="hback"  id="li_id" style="font-family:宋体;display:none;">
			<td colspan="3" class="hback" >&lt;li id=&quot;
				<input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="前台生成li调用的ID，请在CSS中预先定义。可以为空!!">
				&quot; class=&quot;
				<input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="前台生成li调用的class名称，请在CSS中预先定义。可以为空!!">
				&quot;&gt; <span class="tx">对生成列表进行定位，样式控制,ID定义</span></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=AllCode┆';
		retV+='类型$' + obj.NewsType.value + '┆';
		retV+='调用数$' + obj.CodeNumber.value + '┆';
		retV+='标题数$' + obj.leftNumber.value + '┆';
		retV+='点击数$' + obj.ClickNum.value + '┆';
		retV+='图片新闻$' + obj.PicTF.value + '┆';
		retV+='排列数$' + obj.ColsNumber.value + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='引用样式$' + obj.NewsStyle.value + '┆';
		retV+='输出格式$' + obj.out_char.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='DivClass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='内容显示字数$' + obj.contentnumber.value + '┆';
		retV+='导航内容显示字数$' + obj.navinumber.value + '';
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
				<div align="right">调用内容</div>
			</td>
			<td class="hback">
				<select id="InfoType" name="InfoType">
					<option value="ClassName" selected>栏目名称</option>
					<option value="Keywords">栏目关键字</option>
					<option value="Description">栏目描述</option>
				</select></td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right"></div>
			</td>
			<td class="hback">
				<input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
				<input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:NS=ClassInfo┆';
		retV+='调用内容$' + obj.InfoType.value + '';
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
		  <div align="right">自由标签</div>
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
		  <input name="button2" type="button" onClick="ok(this.form);" value="确定创建此标签">
		  <input name="button2" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 ">
	    </td>
	  </tr>
	</table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:MF=FreeLabel┆';
		retV+='自由标签$' + obj.FreeList.value + '';
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
		<div align="right">引用样式</div>
	  </td>
	  <td width="79%" class="hback">
		<select id="NewsStyle"  name="NewsStyle" style="width:40%">
			<% = label_style_List %>
		</select>
		<input name="button3" type="button" id="button32" onClick="showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看">
				<span class="tx">请在各个子系统中建立前台页面新闻显示样式</span> </td>
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
		var retV = '{FS:NS=ClassPage┆';
		retV+='引用样式$' + obj.NewsStyle.value ;
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