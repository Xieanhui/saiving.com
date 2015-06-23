<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/Cls_RefreshJs.asp"-->
<!--#include file="lib/cls_js.asp"-->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,FS_NS_JS_Obj,FS_NS_JS_Sql
Dim Temp_Admin_Is_Super,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
MF_Default_Conn
'session判断
MF_Session_TF 
dim  sRootDir,str_CurrPath,str_CurrPathPic,db_NewsDir
if not MF_Check_Pop_TF("NS040") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页
'******************************************************************



if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""

if Temp_Admin_Is_Super = 1 then
	str_CurrPathPic = sRootDir &"/"&G_UP_FILES_DIR 
Else
	str_CurrPathPic = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&Temp_Admin_Name,"//","/")
End if

str_CurrPath = replace(sRootDir,"//","/") &"/"& db_NewsDir
str_CurrPath = Replace(str_CurrPath,"'","\'")

Function Get_While_Info(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp  ,int_Tmp_i,New_ClassID,ClassID,D_Value
	Str_Tmp = "ID,FileCName,NewsType,LinkCSS,NewsNum,AddTime , FileName,FileType,TitleNum,TitleCSS,RowNum,NaviPic," _
		&"RowBetween,FileSavePath,RowSpace,DateType,DateCSS,ClassName,SonClass,RightDate,MoreContent,LinkWord,PicWidth,PicHeight," _
		&"MarSpeed,MarDirection,ShowTitle,OpenMode,MarWidth,MarHeight"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_NS_Sysjs order by ID desc"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Str_Tmp="ClassName" then 
				Req_Str = NoSqlHack(Trim(request("str"&Str_Tmp)))
			else
				Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
			end if	
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "ID","NewsNum","TitleNum","RowNum","RowSpace","DateType","ClassName","SonClass","RightDate","AddTime","MoreContent","PicWidth","PicHeight","MarSpeed","ShowTitle","OpenMode","MarWidth","MarHeight"
					''数字,日期
						regxp = "|<|>|=|<=|>=|<>|"
						if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
						elseif instr(Req_Str,"*")>0 then 
							if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
							if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
						else	
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
						end if		
					case else
					''字符
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
					response.Write(New_Search_Str)
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	Str_Tmp = ""
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)	
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	'response.Write(This_Fun_Sql)
	'response.End()  
	On Error Resume Next
	Set FS_NS_JS_Obj = CreateObject(G_FS_RS)
	FS_NS_JS_Obj.Open This_Fun_Sql,Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	IF FS_NS_JS_Obj.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
	FS_NS_JS_Obj.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>FS_NS_JS_Obj.PageCount Then cPageNo=FS_NS_JS_Obj.PageCount 
	FS_NS_JS_Obj.AbsolutePage=cPageNo

	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Js_Sys_Modify.asp?FileID="&FS_NS_JS_Obj("ID")&""" class=""otherset"" title='点击修改'>"&FS_NS_JS_Obj("ID")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Js_Sys_Modify.asp?FileID="&FS_NS_JS_Obj("ID")&""" class=""otherset"" title='点击修改'>"&FS_NS_JS_Obj("FileCName")&"</a></td>" & vbcrlf
		select case FS_NS_JS_Obj("NewsType")
			case "RecNews"
			Str_Tmp = "推荐新闻"
			case "MarqueeNews"
			Str_Tmp = "滚动新闻"
			case "SBSNews"
			Str_Tmp = "并排新闻"
			case "PicNews"
			Str_Tmp = "图片新闻"
			case "NewNews"
			Str_Tmp = "最新新闻"
			case "HotNews"
			Str_Tmp = "热点新闻"
			case "WordNews"
			Str_Tmp = "文字新闻"
			case "TitleNews"
			Str_Tmp = "标题新闻"
			case "ProclaimNews"
			Str_Tmp = "公告新闻"
			case else
			Str_Tmp = "[未知]新闻"		
		end select
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("LinkCSS") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("NewsNum") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""javascript:getCode('"& FS_NS_JS_Obj("ID") &"')"">代码</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("AddTime") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox""  name=""FileID"" id=""FileID"" value="""&FS_NS_JS_Obj("ID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		FS_NS_JS_Obj.MoveNext
 		if FS_NS_JS_Obj.eof or FS_NS_JS_Obj.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(FS_NS_JS_Obj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""hidden"" name=""JsAction"" id=""JsAction"" value=""""><input type=""button"" name=""RefreshJs"" id=""RefreshJs"" value=""生成"" onclick=""FunSub('AddV')"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" name=""DleSub"" id=""DleSub"" value="" 删除 "" onclick=""FunSub('Del');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	FS_NS_JS_Obj.close
	Get_While_Info = Get_Html
End Function

Dim SysJsAction,Str_Tmp,AllJsID,RefreshJsFileName,GetjsRs,SuTF,JsArr,Js_i,Err_Str
SysJsAction = Request.Form("JsAction")
IF SysJsAction = "DelSysJS" Then
	Call Del()
ElseIf SysJsAction = "RefreshSysJs" Then
	Call Refresh()
End If

Sub Refresh()
	if not MF_Check_Pop_TF("NS042") then Err_Show
	Str_Tmp = Trim(request.form("FileID"))
	if Str_Tmp = "" then response.Redirect("lib/error.asp?ErrCodes=<li>你必须至少选择一个进行生成。</li>")
	Str_Tmp = replace(Str_Tmp," ","")
	If Instr(Str_Tmp,",") = 0 Then
		AllJsID = CintStr(Trim(Str_Tmp))
		Set GetjsRs = Conn.ExeCute("Select FileName,FileCName From FS_NS_Sysjs Where ID = " & AllJsID)
		If Not GetjsRs.Eof Then
			RefreshJsFileName = GetjsRs(0)
			SuTF = CreateSysJS(RefreshJsFileName)
		Else
			response.Redirect("lib/error.asp?ErrCodes=<li>所选的js记录已不存在</li>")
			Response.End
		End If
		GetjsRs.Close : Set GetjsRs = Nothing
	Else
		JsArr = Split(Str_Tmp,",")
		For Js_i = LBound(JsArr) To UBound(JsArr)
			AllJsID = CintStr(Trim(JsArr(Js_i)))
			Set GetjsRs = Conn.ExeCute("Select FileName,FileCName From FS_NS_Sysjs Where ID = " & AllJsID)
			If Not GetjsRs.Eof Then
				RefreshJsFileName = GetjsRs(0)
				SuTF = CreateSysJS(RefreshJsFileName)
				If SuTF = True Then
					SuTF = True
				Else
					Err_Str = Err_Str & " | 系统js-" & GetjsRs(1) & "-生成失败;"
					If Left(Err_Str,3) = " | " Then
						Err_Str = Right(Err_Str,Len(Err_Str) - 3)
					End iF	
					SuTF = Err_Str & "其他js生成成功;"
				End If	
			End If
			GetjsRs.Close : Set GetjsRs = Nothing	
		Next
	End IF	
	'Response.Write SuTF : response.End 
	Call MF_Insert_oper_Log("系统JS","批量生成了系统JS,生成ID："& Replace(Str_Tmp," ","") &"",now,session("admin_name"),"NS")
	If SuTF = true Then
		response.Redirect("lib/Success.asp?ErrorUrl=../Js_Sys_manage.asp&ErrCodes=<li>恭喜，生成成功。</li>")
	Else
		response.Redirect("lib/error.asp?ErrCodes=<li>" & SuTF & "</li>")
	End If
	Response.End	
End Sub


Sub Del()
	if not MF_Check_Pop_TF("NS042") then Err_Show
	if request.QueryString("FileID")<>"" then 
		Conn.execute("Delete from FS_NS_Sysjs where ID = "&CintStr(request.QueryString("FileID")))
	else
		Str_Tmp = FormatIntArr(request.form("FileID"))
		if Str_Tmp="" then response.Redirect("lib/error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		
		Conn.execute("Delete from FS_NS_Sysjs where ID in ("&Str_Tmp&")")
	end if
	Call MF_Insert_oper_Log("系统JS","批量删除了系统JS,删除ID："& Replace(Str_Tmp," ","") &"",now,session("admin_name"),"NS")
	response.Redirect("lib/Success.asp?ErrorUrl=../Js_Sys_manage.asp&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统JS管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script src="js/Public.js" language="JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
//点击标题排序
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;
function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&filterorderby=")==-1&&Old_Sql.indexOf("?filterorderby=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&filterorderby=" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?filterorderby=" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);
				break;	
			}
		}
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName):New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
function selectAll(f)
{
	for(i=0;i<f.length;i++)
	{
		if(f(i).type=="checkbox" && f(i)!=event.srcElement)
		{
			f(i).checked=event.srcElement.checked;
		}
	}
}
-->
</script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" ><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>系统JS管理</strong></a></td>
  </tr>
  <tr  class="hback">
    <td class="hback" ><a href="Js_Sys_manage.asp?Act=View" >管理首页</a>&nbsp;|&nbsp;
      <a href="Js_Sys_Add.asp">新增</a> &nbsp;|&nbsp;
      <a href="Js_Sys_manage.asp?Act=Search">查询</a>
	</td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
	View
	case "Search"
	Search
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ID')" class="sd"><b>〖ID号〗</b></a> <span id="Show_Oder_ID" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('FileCName')" class="sd"><b>中文名称</b></a> <span id="Show_Oder_FileCName" class="tx"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('NewsType')" class="sd"><b>类型</b></a> <span id="Show_Oder_NewsType" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('LinkCSS')" class="sd"><b>样式</b></a> <span id="Show_Oder_LinkCSS" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('NewsNum')" class="sd"><b>新闻条数</b></a> <span id="Show_Oder_NewsNum" class="tx"></span></td>
	  <td align="center" class="xingmu">获取代码</td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('AddTime')" class="sd"><b>添加时间</b></a> <span id="Show_Oder_AddTime" class="tx"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub
Sub Search()%>

  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<form action="?Act=SearchGo" method="post" name="ClassJSForm">
    <tr class="hback">
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;ID编号</td>
      <td colspan="3"> 
        <input name="ID" type="text" id="ID" size="15" maxlength="11" value=""></td>
	</tr>
	
	
    <tr class="hback">
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;选择栏目</td>
      <td width="35%"> 
		<input name="strClassName" type="text" id="strClassName" style="width:50%" value="" readonly="">
		<input name="strClassID" type="hidden" id="strClassID" value="">
		<input type="button" name="Submit" value="选择栏目"   onClick="SelectClass();">	   </td>
      <td colspan="2">不选择则调出所有     </td>
    </tr>
	
	
    <tr class="hback">
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;中文名称</td>
      <td width="35%"> 
	  	<input type="hidden" name="Act" value="SearchGo">
        <input name="FileCName" type="text" id="FileCName" style="width:90%" value=""></td>
      <td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;文件名称</td>
      <td width="35%"> 
        <input name="FileName" type="text" id="FileName" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;栏目名称</td>
      <td> 
		<input type="text" style="width:90%" name="ClassID" value="0" disabled>
	 </td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新闻类型</td>
      <td> 
        <select name="NewsType" style="width:90%" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
		  <option value="">请选择</option>
          <option value="RecNews">推荐新闻</option>
          <option value="MarqueeNews">滚动新闻</option>
          <option value="SBSNews">并排新闻</option>
          <option value="PicNews">图片新闻</option>
          <option value="NewNews">最新新闻</option>
          <option value="HotNews">热点新闻</option>
          <option value="WordNews">文字新闻</option>
          <option value="TitleNews">标题新闻</option>
          <option value="ProclaimNews">公告新闻</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;更多链接</td>
      <td> 
        <select name="MoreContent" id="MoreContent" style="width:90% " onChange="ChooseLink(this.options[this.selectedIndex].value);" disabled>
          <option value="1">是</option>
          <option value="0">否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;链接字样</td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" style="width:90%" value="" disabled></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;新闻数量</td>
      <td> 
        <input name="NewsNum" type="text" id="NewsNum" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;每行数量</td>
      <td> 
        <input name="RowNum" type="text" id="RowNum" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;链接样式</td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" style="width:90%" value="" disabled></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;标题字数</td>
      <td> 
        <input name="TitleNum" type="text" id="TitleNum" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;图片宽度</td>
      <td> 
        <input name="PicWidth" type="text" id="PicWidth" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;标题样式</td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;图片高度</td>
      <td> 
	     <input name="PicHeight" type="text" id="PicHeight" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新闻行距</td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;滚动速度</td>
      <td> 
        <input name="MarSpeed" type="text" id="MarSpeed" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;滚动方向</td>
      <td> 
        <select name="MarDirection" id="MarDirection" style="width:90% ">
		  <option value="">请选择</option>
          <option value="up">向上</option>
          <option value="down">向下</option>
          <option value="left">向左</option>
          <option value="right">向右</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;公告宽度</td>
      <td> 
        <input name="MarWidth" type="text" id="MarWidth" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;公告高度</td>
      <td> 
        <input name="MarHeight" type="text" id="MarHeight" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;显示标题</td>
      <td> 
        <select name="ShowTitle" id="ShowTitle" style="width:90%">
		  <option value="">请选择</option>
          <option value="1">是</option>
          <option value="0">否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新开窗口</td>
      <td> 
        <select name="OpenMode" id="OpenMode" style="width:90%">
		  <option value="">请选择</option>
          <option value="1">是</option>
          <option value="0">否</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;导航图片</td>
      <td> 
        <input name="NaviPic" type="text" id="NaviPic" style="width:60%" value="">
        <input type="button" name="bnt_ChoosePic_naviPic"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.NaviPic);"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;行间图片</td>
      <td> 
        <input name="RowBetween" type="text" id="RowBetween" style="width:52%" value="">
        <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.RowBetween);"></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;调用日期</td>
      <td> 
        <select name="DateType" id="DateType" style="width:90%">
		  <option value="">请选择</option>
          <option value="0">调用日期类型</option>
          <option value="1">2006-7-26</option>
          <option value="2">2006.7.26</option>
          <option value="3">2006/7/26</option>
          <option value="4">7/26/2006</option>
          <option value="5">26/7/2006</option>
          <option value="6">7-26-2006</option>
          <option value="7">7.26.2006</option>
          <option value="8">7-26</option>
          <option value="9">7/26</option>
          <option value="10">7.26</option>
          <option value="11">7月26日</option>
          <option value="12">26日14时</option>
          <option value="13">26日14点</option>
          <option value="14">14时56分</option>
          <option value="15">14:56</option>
          <option value="16">2006年7月26日</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;保存路径</td>
      <td> 
        <input name="SaveFilePath" type="text" id="SaveFilePath" style="width:52%" value="">
        <INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= str_CurrPath %>',300,250,window,document.ClassJSForm.SaveFilePath);document.ClassJSForm.SaveFilePath.focus();"></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;日期样式</td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;显示栏目</td>
      <td> 
        <select name="ClassName" id="ClassName" style="width:90%">
		  <option value="">请选择</option>
          <option value="1">显示</option>
          <option value="0">不显示</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;调用子类</td>
      <td> 
        <select name="SonClass" id="SonClass" style="width:90%" disabled>
		  <option value="">请选择</option>
          <option value="1">是</option>
          <option value="0">否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;日期右对齐</td>
      <td> 
        <select name="RightDate" id="RightDate" style="width:90%">
		  <option value="">请选择</option>
          <option value="1">是</option>
          <option value="0">否</option>
        </select></td>
    </tr>
    <tr class="hback">  
	<td colspan="10" align="center">
		<input type="submit" value=" 执行查询 ">&nbsp;&nbsp;
		<input type="reset" value=" 重置 ">
	</td>
	</tr>	
  </form>
</table><p>
<%End Sub%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
//打开后根据规则显示箭头
var Req_FildName;
if (Old_Sql.indexOf("filterorderby=")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);break;}	
		}
	}	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + "csed")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
///////////////////////////////////////////////////////// 
function getCode(jsid)
{
	if (jsid!=""&&!isNaN(jsid))
	{
		OpenWindow('lib/Frame.asp?PageTitle=获取JS调用代码&FileName=showSysJsPath.asp&JsID='+jsid,360,140,window);
	}else
	{
		alert("出现错误，请联系客服人员！")
	}
}


function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	try {
		$("strClassID").value= ReturnValue[0][0];
		$("strClassName").value= ReturnValue[1][0];
	}
	catch (ex) { }
}


//
function FunSub(Str)
{
	if (Str == 'AddV')
	{
		document.getElementById('JsAction').value = 'RefreshSysJs';
		document.form1.submit();
	}
	else
	{
		document.getElementById('JsAction').value = 'DelSysJS';
		document.form1.submit();
	}
}

-->
</script>
<%
Set FS_NS_JS_Obj=nothing
Conn.close
Set Conn=nothing
%>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 