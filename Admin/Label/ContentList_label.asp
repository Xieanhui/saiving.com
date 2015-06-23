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
Dim Conn,obj_label_style_Rs,label_style_List
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("MF_sPublic") then Err_Show

Set  obj_label_style_Rs = server.CreateObject(G_FS_RS)
obj_label_style_Rs.Open "Select ID,StyleName from FS_MF_Labestyle where StyleType='NS' Order by  id desc",Conn,1,3
do while Not obj_label_style_Rs.eof 
	label_style_List = label_style_List&"<option value="""& obj_label_style_Rs(0)&""">"& obj_label_style_Rs(1)&"</option>"
	obj_label_style_Rs.movenext
loop
obj_label_style_Rs.close
set obj_label_style_Rs = nothing
%>
<html>
<head>
<title>新闻标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<style type="text/css">
.lionrong {
	background-color: #EDEDED;
	border: 1px solid #000000;
	font-size:12px;
	color:#000000;
	line-height: 18px;
	padding-left:10px;
	padding-right:10px;
	padding-top:5px;
	padding-bottom:10px;
}
</style>
<base target="_self">
</head>
<body class="hback">
<FIELDSET style="width:98%;" align="center">
        <LEGEND align=left>必选参数</LEGEND>
<table width="98%" height="100" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
  <tr class="hback">
    <td width="15%">子系统</td>
    <td width="35%">
	  <select name="SubSys">
	  <%
	Dim p_Sub_List_Rs
	Set p_Sub_List_Rs	= CreateObject(G_FS_RS)
	p_Sub_List_Rs.Open "select Sub_Sys_Name,Sub_Sys_ID from FS_MF_Sub_Sys order by id asc",Conn,1,1
	do while Not p_Sub_List_Rs.Eof
	%>
	  <option value="<% = p_Sub_List_Rs("Sub_Sys_ID") %>"><% = p_Sub_List_Rs("Sub_Sys_Name") %></option>
	  <%
	  	p_Sub_List_Rs.MoveNext
	Loop
	p_Sub_List_Rs.Close
	Set p_Sub_List_Rs = Nothing
	  %>
      </select>
	</td>
    <td width="15%">标签类型</td>
    <td><select  name="select" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
      <option value="" style="background:#DEDEDE">---列表类----------</option>
      <option value="ClassNews" selected>├┄栏目新闻列表</option>
      <option value="SpecialNews">├┄专题新闻列表</option>
      <!-- <option value="ReadNews">新闻浏览(新闻页面)</option>-->
      <option value="LastNews">├┄最新新闻</option>
      <option value="HotNews">├┄热点新闻</option>
      <option value="RecNews">├┄推荐新闻</option>
      <!--<option value="FiltNews">├┄幻灯新闻</option>-->
      <option value="MarNews">├┄滚动新闻</option>
      <!-- <option value="CorrNews">├┄相关新闻</option>-->
      <!--<option value="DayNews">├┄头条新闻</option>-->
      <option value="BriNews">├┄精彩新闻</option>
      <option value="AnnNews">├┄公告新闻</option>
      <option value="ConstrNews">├┄投稿</option>
      <option value="" style="background:#DEDEDE">---终极类----------</option>
      <option value="ClassList">├┄终极新闻列表</option>
      <option value="subClassList">├┄子类新闻列表</option>
      <option value="SpecialList">├┄终极专题列表</option>
    </select></td>
  </tr>
  <tr class="hback">
    <td>显示格式</td>
    <td><select name="select2" id="select2" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
      <option value="out_Table">普通格式</option>
      <option value="out_DIV">DIV+CSS格式</option>
    </select></td>
    <td>引用样式</td>
    <td><select id="select3"  name="select3" style="width:40%">
      <% = label_style_List %>
    </select>
      <input name="button32" type="button" id="button3" onClick="showDiv(this,'aaaaaaaaaaa');//showModalDialog('News_label_styleread.asp?ID='+document.form1.NewsStyle.value,'WindowObj','dialogWidth:420pt;dialogHeight:180pt;status:yes;help:no;scroll:yes;');" value="查看"></td>
  </tr>
  <tr class="hback">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="hback">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="hback">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</FIELDSET>
<FIELDSET style="width:98%;" align="center">
        <LEGEND align=left>可选参数</LEGEND>
sdfasdfasd
</FIELDSET>
</body>
</html>
<% 
Set Conn=nothing
%>
<script language="javascript">
var LableFactory={
	Build:function(_Lables){
		var _Code='';
		for(var i=0;i<_Lables.length;i++){
			eval('_Code+=LableFactory.Build'+_Lables[i]+'();','javascript');
		}
		return _Code;
	},
	BuildTextField:function(_ID,_ClassName){
		if(_ClassName)_ClassName='class="'+_ClassName+'"';else _ClassName='';
		return '<input name="'+_ID+'" id="'+_ID+'" '+_ClassName+' type="text" />';
	},
	BuildSelectField:function(_ID,_ClassName,_Options,_onchangeHandler){
		if(_ClassName)_ClassName='class="'+_ClassName+'"';else _ClassName='';
		if(_onchangeHandler)_onchangeHandler='onchange="'+_onchangeHandler+'"';else _onchangeHandler='';
		var _Code='<select id="'+_ID+'" name="'+_ID+'" '+_ClassName+' '+_onchangeHandler+'>';
		for(var i=0;i<_Options.length;i++){
			_Code+='<option value="'+_Options[i].Value+'">'+_Options[i].Text+'</option>';
		}
		return _Code+'</select>';
	},
	BuildButtonField:function(_ID,_ClassName,_ButtonText,_onclickHandler){
		if(_ClassName)_ClassName='class="'+_ClassName+'"';else _ClassName='';
		if(_onclickHandler)_onclickHandler='onchange="'+_onclickHandler+'"';else _onclickHandler='';
		return '<input name="'+_ID+'" id="'+_ID+'" '+_ClassName+' type="button" value="'+_ButtonText+'" '+_onclickHandler+' />'
	},
	BuildField:function(_NameText,_FieldStr){
		
	}
	BuildClassNews:function(){
		return 'ClassNews';
	},
	BuildSpecialNews:function(){
		return 'SpecialNews';
	}
};
var FoosunLableMass=(function()
{
	var NSLables='ClassNews,SpecialNews';
	return{
		LableCode:function(_Sys){
			var Lables='',LableArray=null,Str='';
			switch(_Sys){
				case 'NS':
					Lables=NSLables;
					break;
				default:
					Lables='';
					break;
			}
			if(Lables!=''){
				LableArray=Lables.split(',');
				Str+=LableFactory.Build(LableArray);
			}
			return Str;
		}
	}
})();

var Options=new Array({Text:1,Value:1},{Text:2,Value:2});
alert(LableFactory.BuildSelectField('','',Options));
//alert(FoosunLableMass.LableCode('NS'));
function showDiv(obj,content)
{
    var pos = getPosition(obj)
    var objDiv = document.createElement("div");
    objDiv.className="lionrong";//For IE
    objDiv.style.position = "absolute";
	var tempheight=pos.y;
	var tempwidth1,tempheight1;
	var windowwidth=document.body.clientWidth;
	var isIE5 = (navigator.appVersion.indexOf("MSIE 5")>0) || (navigator.appVersion.indexOf("MSIE")>0 && parseInt(navigator.appVersion)> 4);
	var isIE55 = (navigator.appVersion.indexOf("MSIE 5.5")>0);
	var isIE6 = (navigator.appVersion.indexOf("MSIE 6")>0);
	var isIE7 = (navigator.appVersion.indexOf("MSIE 7")>0);

	if(isIE5||isIE55||isIE6||isIE7){var tempwidth=pos.x+305;}else{var tempwidth=pos.x+312;}
	objDiv.style.width = "300px";
    objDiv.innerHTML = content;
	if (tempwidth>windowwidth)
	{
		tempwidth1=tempwidth-windowwidth
		objDiv.style.left = (pos.x-tempwidth1) + "px";
	}
	else
	{
		if(isIE5||isIE55||isIE6||isIE7){objDiv.style.left = (pos.x + 10) + "px";}else{objDiv.style.left = (pos.x) + "px";}
	}
	if(isIE5||isIE55||isIE6||isIE7){objDiv.style.top = (pos.y+16) + "px";}else{objDiv.style.top = (pos.y+16) + "px";}

    objDiv.style.display = "";
    //document.onclick=function () { if(objDiv.style.display==""){objDiv.style.display="none";} }
    document.body.appendChild(objDiv);
}
getPosition = function(oElement)
{
    var objParent = oElement
    var oPosition = new position(0,0);
    while (objParent.tagName != "BODY")
    {
        oPosition.x += objParent.offsetLeft;
        oPosition.y += objParent.offsetTop;
        objParent = objParent.offsetParent;
    }
    return oPosition;
} 
position = function(x,y)
{
    this.x = x;
    this.y = y;
}
</script>