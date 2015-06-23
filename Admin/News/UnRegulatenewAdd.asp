<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="NF_News_Function.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF
if not MF_Check_Pop_TF("NS_UnRl") then Err_Show
dim Fs_news
set Fs_news = new Cls_News
Dim CharIndexStr
CharIndexStr=all_substring
Dim UnNewsArray,ActUrl,NewsID,RsNewsObj,StrSql
UnNewsArray = "new Array()"
ActUrl="SetUnRegulate.asp?Action=Add"
NewsID=""
IF Request.QueryString("Action")="Edit" Then
	if not MF_Check_Pop_TF("NS046") then Err_Show
	NewsID=NoSqlHack(Request.QueryString("NewsID"))
	ActUrl="SetUnRegulate.asp?Action=Edit&MainNewsID="&NewsID
	
	Set RsNewsObj = Server.CreateObject(G_FS_RS)
	StrSql = "Select MainUnregNewsID,UnregNewsName,NewsTitle,[Rows] From FS_NS_News_Unrgl,FS_NS_News where FS_NS_News.NewsID=FS_NS_News_Unrgl.MainUnregNewsID and UnRegulatedMain='"&NoSqlHack(NewsID)&"' order by FS_NS_News_Unrgl.ID ASC"
	RsNewsObj.Open StrSql,Conn,1,1
	UnNewsArray=""
	While Not RsNewsObj.Eof
		If UnNewsArray="" Then
			UnNewsArray="['"&RsNewsObj("MainUnregNewsID")&"','"&RsNewsObj("NewsTitle")&"','"&RsNewsObj("UnregNewsName")&"',"&RsNewsObj("Rows")&"]"
		Else
			UnNewsArray=UnNewsArray&",['"&RsNewsObj("MainUnregNewsID")&"','"&RsNewsObj("NewsTitle")&"','"&RsNewsObj("UnregNewsName")&"',"&RsNewsObj("Rows")&"]"
		End If
		RsNewsObj.MoveNext
	Wend
	UnNewsArray="["&UnNewsArray&"]"
End If
%>
<html>
<head>
<title>不规则新闻 规则管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="javascript" src="../../Fs_inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_INC/prototype.js"></script>
<script language="javascript">
<!--
Array.prototype.remove = function(start,deleteCount){
	if(isNaN(start)||start>this.length||deleteCount>(this.length-start)){return false;}
	this.splice(start,deleteCount);
}

String.prototype.trim=function(){
  return this.replace(/(^\s*)|(\s*$)/g,"");
}

moveStart = function (event, _sId)
{
	var oObj = $(_sId);
	oObj.onmousemove = mousemove;
	oObj.onmouseup = mouseup;
	oObj.setCapture ? oObj.setCapture() : function(){};
	oEvent = window.event ? window.event : event;
	var dragData = {x : oEvent.clientX, y : oEvent.clientY};
	var backData = {x : parseInt(oObj.style.top), y : parseInt(oObj.style.left)};
	function mousemove()
	{
		var oEvent = window.event ? window.event : event;
		var iLeft = oEvent.clientX - dragData["x"] + parseInt(oObj.style.left);
		var iTop = oEvent.clientY - dragData["y"] + parseInt(oObj.style.top);
		oObj.style.left = iLeft;
		oObj.style.top = iTop;
/*		$('dialogBoxShadow').style.left = iLeft + 6;
		$('dialogBoxShadow').style.top = iTop + 6;
		if ($('dialogIframBG'))
		{
			$('dialogIframBG').style.left = iLeft;
			$('dialogIframBG').style.top = iTop;
		}*/
		dragData = {x: oEvent.clientX, y: oEvent.clientY};

	}
	function mouseup()
	{
		var oEvent = window.event ? window.event : event;
		oObj.onmousemove = null;
		oObj.onmouseup = null;
		if(oEvent.clientX < 1 || oEvent.clientY < 1 || oEvent.clientX > document.body.clientWidth || oEvent.clientY > document.body.clientHeight){
			oObj.style.left = backData.y;
			oObj.style.top = backData.x;
/*			$('dialogBoxShadow').style.left = backData.y + 6;
			$('dialogBoxShadow').style.top = backData.x + 6;
			if ($('dialogIframBG'))
			{
				$('dialogIframBG').style.left = backData.y;
				$('dialogIframBG').style.top = backData.x;
			}*/
		}
		oObj.releaseCapture ? oObj.releaseCapture() : function(){};
	}
}


UnNewArray=<%= UnNewsArray %>;
function CheckNum(obj){
	if (isNaN(obj.value) || obj.value<=0){
		alert("您输入的不是正确的行数,\n请输入一个正整数.");
		obj.value="";
		obj.focus();
		}
}

function DisplayUnNews()
{
	var StrUnNewsList="";
	var ListLen=UnNewArray.length;
	var StrUnNewsListSub="";
	for (var i=0;i<ListLen;i++){
		StrUnNewsListSub="<div id=\"Arr"+i+"\"><input name=\"NewsID\" type=\"hidden\" id=\"NewsID_"+i+"\" value=\""+UnNewArray[i][0]+"\" /><a href=\"原新闻标题\" title=\"原新闻标题:"+UnNewArray[i][1]+"\" onclick=\"return false;\">标题</a>：<input name=\"NewsTitle"+UnNewArray[i][0]+"\" type=\"text\" id=\"NewsTitle_"+i+"\" value=\""+UnNewArray[i][2]+"\" size=\"50\" onkeyup=\"UnNewModify(this,'')\" onmousedown=\"new Form.Element.Observer('NewsTitle"+UnNewArray[i][0]+"',1,UnNewModify);\" />&nbsp;放在第<input name=\"Row"+UnNewArray[i][0]+"\" type=\"text\" id=\"Row_"+i+"\" value=\""+UnNewArray[i][3]+"\" size=\"2\" maxlength=\"2\" onkeyup=\"UnNewModify(this,'')\" onbeforepaste=\"clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''));\" onmousedown=\"new Form.Element.Observer('Row"+UnNewArray[i][0]+"',1,UnNewModify);\" />行&nbsp;<button onclick=\"UnNewDel("+i+")\">移除</button></div>";;
		if (StrUnNewsList==""){
			StrUnNewsList=StrUnNewsListSub;
		}else{
			StrUnNewsList+=StrUnNewsListSub;
		}
	}
	document.getElementById("UnNewsList").innerHTML=StrUnNewsList;
}

function UnNewModify(modobj,col){
	for (var i=0;i<UnNewArray.length;i++){
		UnNewArray[i][2]=$("NewsTitle_"+i).value;
		$("Row_"+i).value=$("Row_"+i).value.replace(/[^\d]/g,'');
		UnNewArray[i][3]=parseInt($("Row_"+i).value);
	}
	UnNewPreviewCh();
}

function UnNewDel(Row){
	if (confirm("确定移除吗？")){
		UnNewArray.remove(Row,1);
		DisplayUnNews();
		UnNewPreviewCh();
		document.DisNews.CheckUnNews();
	}
}

function DivCenter(M_div,M_width,M_height,M_zindex)
{
	var xposition=0,yposition=0;
	$(M_div).style.position='absolute';
	$(M_div).style.width=M_width.toString(10)+'px';
	$(M_div).style.height=M_height.toString(10)+'px';
	$(M_div).style.zIndex=M_zindex.toString(10);

	if (parseInt(navigator.appVersion) >= 4 )
	{
		xposition = (document.body.offsetWidth - M_width) / 2;
		yposition = (document.body.offsetHeight - M_height) / 2;
		$(M_div).style.left=xposition.toString(10)+"px";
		$(M_div).style.top=(yposition).toString(10)+"px";
	}
}

function UnNewPreviewCh(){
	if ($("preview").style.display==""){
		UnNewPreview();
	}
}
function UnNewPreview(){
	var ListLen=UnNewArray.length;
	var Maxrow=1;
	var PreviewStr="";
	var PreviewRowStr="";
	for (var i=0;i<ListLen;i++){
		if (UnNewArray[i][3]>Maxrow){
			Maxrow=UnNewArray[i][3];
		}
	}
	PreviewStr="<table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";
	for (i=1;i<=Maxrow;i++){
		FindFlag="";
		PreviewRowStr="";
		for (var j=0;j<ListLen;j++){
			if (UnNewArray[j][3]==i){
				if (FindFlag==""){
					FindFlag=j.toString(10);
				}else{
					FindFlag+=","+j;
				}
			}
		}
		
		PreviewStr+="<tr><td>";
		if (FindFlag){
			PreviewRowStr=FindFlag.split(",");
			for (var j=0;j<PreviewRowStr.length;j++){
				if (j==0){
					PreviewStr+="<a href=\"#\" onclick=\"return false;\">"+UnNewArray[PreviewRowStr[j]][2]+"</a>";
				}else{
					PreviewStr+="&nbsp;<a href=\"#\" onclick=\"return false;\">"+UnNewArray[PreviewRowStr[j]][2]+"</a>";
				}
			}
		}else{
			PreviewStr+="&nbsp;";
		}
		PreviewStr+="</td></tr>";
	}
	PreviewStr+="<tr>\
					<td align=\"center\"><button onclick=\"$('preview').style.display='none';\">关闭</button></td>\
				</tr>";
	PreviewStr+="</table>";
	
	if ($("preview").style.display=="none"){
		$("preview").style.display="";
		DivCenter("preview",680,200,100);
	}
	$("PreviewContent").innerHTML=PreviewStr;
}
function UnNewcheck(){
	var ListLen=UnNewArray.length;
	var Maxrow=1;
	var ErrStr="";
	for (var i=0;i<ListLen;i++){
		if (UnNewArray[i][3]==0){
			ErrStr=" -第 "+(i+1)+"条 存放行数不能为 0";
		}
		if (isNaN(UnNewArray[i][3])){
			ErrStr=" -第 "+(i+1)+"条 存放行数不能为空";
		}
		if (UnNewArray[i][2]==""){
			ErrStr=" -第 "+(i+1)+"条 不规则标题不能为空";
		}
		if (UnNewArray[i][3]>Maxrow){
			Maxrow=UnNewArray[i][3];
		}
	}
	var FindFlag=false;
	for (i=1;i<=Maxrow;i++){
		FindFlag=false;
		for (var j=0;j<ListLen;j++){
			if (UnNewArray[j][3]==i){
				FindFlag=true;
				break;
			}
		}
		if (!FindFlag){
			ErrStr+="\n -第 "+i+" 行中没有新闻";
		}
	}
	if (ErrStr){
		alert("发生以下错误：\n"+ErrStr);
		return false;
	}else{
		return true;
	}
}
-->
</script>
</head>
<body>
<div id="preview" style="display:none">
	<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
		<tr>
			<td align="center" class="hback_1" style="cursor:move;" onMouseDown="new moveStart(event,'preview')"><strong>预览不规则新闻(点击拖动)</strong></td>
		</tr>
		<tr>
			<td align="center" class="hback" id="PreviewContent"  style="cursor:move;" onMouseDown="new moveStart(event,'preview')"></td>
		</tr>
	</table>
</div>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="hback">
		<td class="xingmu">不规则新闻管理</td>
	</tr>
	<tr>
		<td height="26" valign="middle" class="hback">
			<table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="4%" align="center" valign="bottom" class="hback" vlign="middle"><a href="UnRegulatenewAdd.asp">添加</a></td>
					<td width="1%" align="center" valign="bottom" class="Gray" vlign="middle">|</td>
					<td width="10%" align="center" valign="bottom" class="hback" vlign="middle"><a href="DefineNews_Manage.asp">返回管理页面</a></td>
					<td width="1%" align="center" valign="bottom" class="Gray" vlign="center">|</td>
					<td width="3%" align="center" valign="bottom" class="hback"  vlign="center"><a href="../../help?Lable=NS_UnRegualNewAdd" target="_blank" style="cursor:help;">帮助</a></td>
					<form name='SearchForm' method="post" target='DisNews' action="News_Display.asp"><td width="81%"><div align="right">搜索关键字：
			          <input name="SearchKey" type="text" value="" size="20">
				搜索栏目：
				<select name="ClassID" id="ClassID">
				  <option value="">所有栏目</option>
				  <%	
			If Request.QueryString("ClassID")<>"" then
				Set DefaultRs=Conn.execute("Select ClassID,ClassName From FS_NS_NewsClass Where ClassID='"&NoSqlHack(request.QueryString("ClassID"))&"'")	
				if not DefaultRs.eof then
				%>
				  <option value=<%=DefaultRs("ClassID")%>><%=DefaultRs("ClassName")%></option>
				  <%
				end if
				Set DefaultRs=nothing
			end if
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0' and ReycleTF=0 order by AddTime DESC",Conn,1,3
			str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &""">"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
				  </select>
				不规则新闻
				<input name="UnAll" type="checkbox" value="UnNews" checked>
				<input name="submit" type="submit" value="搜 索">
				  </div></td></form>
				</tr>
			</table>		</td>
	</tr>
	<form name='SearchForm' method="post" target='DisNews' action="News_Display.asp">
	</form>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<form action="<%= ActUrl %>" method="post" name="GetNewsIDForm" id="GetNewsIDForm">
		<tr>
		  <td class="hback">
				<input type="submit" name="Submit" onClick="return UnNewcheck();" value="保存">
			
			<input name="View" type="button" id="View" onClick="if (UnNewcheck())UnNewPreview();" value="预览效果">
			<label></label></td>
		</tr>
		<tr>
			<td class="hback" id="UnNewsList"></td>
		</tr>
	</form>
</table>
<script language="JavaScript">
<!--
DisplayUnNews();
//-->
</script>
<iframe name="DisNews" src="News_Display.asp?UnAll=UnNews" width="100%" frameborder="0" height="400" scrolling="no"></iframe>
</body>
<%
Set Conn=nothing
%>
</html>






