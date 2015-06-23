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
            <td width="41%" class="xingmu"><strong>留言标签创建</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td width="19%" class="hback"><div align="right">类型</div></td>
      <td width="81%" class="hback">
	  <select name="BookPop" id="BookPop">
        <option value="0">正常</option>
        <option value="1">推荐</option>
        <option value="1">点击</option>
      </select>	  </td>
    </tr>
    <tr class="hback" id="ClassName_col"> 
      <td class="hback"  align="center"><div align="right">栏目列表</div></td>
      <td colspan="2" class="hback" ><label>
        <select name="BookClassId" id="BookClassId">
          <option value="">所有</option>
		 <%
		 dim rs
		 set rs = Conn.execute("select ClassID,ClassName From FS_WS_Class where ParentClasID='0' order by id desc")
		 do while not rs.eof
		 		response.Write "<option value="""&rs("ClassID")&""">"&rs("ClassName")&"</option>"
			 rs.movenext
		 loop
		 rs.close:set rs = nothing
		 %>
        </select>
      </label></td>
    </tr>
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
      <td class="hback"><div align="right">标题CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="10"> 
        对DIV+CSS框架无效 </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">日期CSS</div></td>
      <td class="hback"><input name="DateCSS" type="text" id="DateCSS"></td>
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
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:WS=BookCode┆';
		retV+='类型$' + obj.BookPop.value + '┆';
		retV+='栏目$' + obj.BookClassId.value + '┆';
		retV+='数量$' + obj.TitleNumber.value + '┆';
		retV+='列数$' + obj.ColsNumber.value + '┆';
		retV+='字数$' + obj.leftTitle.value + '┆';
		retV+='标题CSS$' + obj.TitleCSS.value + '┆';
		retV+='日期CSS$' + obj.DateCSS.value + '┆';
		retV+='日期格式$' + obj.DateStyle.value + '┆';
		retV+='DivID$' + obj.DivID.value + '┆';
		retV+='Divclass$' + obj.Divclass.value + '┆';
		retV+='ulid$' + obj.ulid.value + '┆';
		retV+='ulclass$' + obj.ulclass.value + '┆';
		retV+='liid$' + obj.liid.value + '┆';
		retV+='liclass$' + obj.liclass.value + '┆';
		retV+='输入格式$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
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






