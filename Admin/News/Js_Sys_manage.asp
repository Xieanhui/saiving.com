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
'session�ж�
MF_Session_TF 
dim  sRootDir,str_CurrPath,str_CurrPathPic,db_NewsDir
if not MF_Check_Pop_TF("NS040") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"				'βҳ
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
					''����,����
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
					''�ַ�
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
		response.Redirect("../error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	IF FS_NS_JS_Obj.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
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
		Get_Html = Get_Html & "<td align=""center""><a href=""Js_Sys_Modify.asp?FileID="&FS_NS_JS_Obj("ID")&""" class=""otherset"" title='����޸�'>"&FS_NS_JS_Obj("ID")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Js_Sys_Modify.asp?FileID="&FS_NS_JS_Obj("ID")&""" class=""otherset"" title='����޸�'>"&FS_NS_JS_Obj("FileCName")&"</a></td>" & vbcrlf
		select case FS_NS_JS_Obj("NewsType")
			case "RecNews"
			Str_Tmp = "�Ƽ�����"
			case "MarqueeNews"
			Str_Tmp = "��������"
			case "SBSNews"
			Str_Tmp = "��������"
			case "PicNews"
			Str_Tmp = "ͼƬ����"
			case "NewNews"
			Str_Tmp = "��������"
			case "HotNews"
			Str_Tmp = "�ȵ�����"
			case "WordNews"
			Str_Tmp = "��������"
			case "TitleNews"
			Str_Tmp = "��������"
			case "ProclaimNews"
			Str_Tmp = "��������"
			case else
			Str_Tmp = "[δ֪]����"		
		end select
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("LinkCSS") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("NewsNum") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""javascript:getCode('"& FS_NS_JS_Obj("ID") &"')"">����</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& FS_NS_JS_Obj("AddTime") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox""  name=""FileID"" id=""FileID"" value="""&FS_NS_JS_Obj("ID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		FS_NS_JS_Obj.MoveNext
 		if FS_NS_JS_Obj.eof or FS_NS_JS_Obj.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(FS_NS_JS_Obj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""hidden"" name=""JsAction"" id=""JsAction"" value=""""><input type=""button"" name=""RefreshJs"" id=""RefreshJs"" value=""����"" onclick=""FunSub('AddV')"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" name=""DleSub"" id=""DleSub"" value="" ɾ�� "" onclick=""FunSub('Del');""></td>"
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
	if Str_Tmp = "" then response.Redirect("lib/error.asp?ErrCodes=<li>���������ѡ��һ���������ɡ�</li>")
	Str_Tmp = replace(Str_Tmp," ","")
	If Instr(Str_Tmp,",") = 0 Then
		AllJsID = CintStr(Trim(Str_Tmp))
		Set GetjsRs = Conn.ExeCute("Select FileName,FileCName From FS_NS_Sysjs Where ID = " & AllJsID)
		If Not GetjsRs.Eof Then
			RefreshJsFileName = GetjsRs(0)
			SuTF = CreateSysJS(RefreshJsFileName)
		Else
			response.Redirect("lib/error.asp?ErrCodes=<li>��ѡ��js��¼�Ѳ�����</li>")
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
					Err_Str = Err_Str & " | ϵͳjs-" & GetjsRs(1) & "-����ʧ��;"
					If Left(Err_Str,3) = " | " Then
						Err_Str = Right(Err_Str,Len(Err_Str) - 3)
					End iF	
					SuTF = Err_Str & "����js���ɳɹ�;"
				End If	
			End If
			GetjsRs.Close : Set GetjsRs = Nothing	
		Next
	End IF	
	'Response.Write SuTF : response.End 
	Call MF_Insert_oper_Log("ϵͳJS","����������ϵͳJS,����ID��"& Replace(Str_Tmp," ","") &"",now,session("admin_name"),"NS")
	If SuTF = true Then
		response.Redirect("lib/Success.asp?ErrorUrl=../Js_Sys_manage.asp&ErrCodes=<li>��ϲ�����ɳɹ���</li>")
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
		if Str_Tmp="" then response.Redirect("lib/error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		
		Conn.execute("Delete from FS_NS_Sysjs where ID in ("&Str_Tmp&")")
	end if
	Call MF_Insert_oper_Log("ϵͳJS","����ɾ����ϵͳJS,ɾ��ID��"& Replace(Str_Tmp," ","") &"",now,session("admin_name"),"NS")
	response.Redirect("lib/Success.asp?ErrorUrl=../Js_Sys_manage.asp&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳJS����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script src="js/Public.js" language="JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
//�����������
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
    <td class="xingmu" ><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>ϵͳJS����</strong></a></td>
  </tr>
  <tr  class="hback">
    <td class="hback" ><a href="Js_Sys_manage.asp?Act=View" >������ҳ</a>&nbsp;|&nbsp;
      <a href="Js_Sys_Add.asp">����</a> &nbsp;|&nbsp;
      <a href="Js_Sys_manage.asp?Act=Search">��ѯ</a>
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
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ID')" class="sd"><b>��ID�š�</b></a> <span id="Show_Oder_ID" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('FileCName')" class="sd"><b>��������</b></a> <span id="Show_Oder_FileCName" class="tx"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('NewsType')" class="sd"><b>����</b></a> <span id="Show_Oder_NewsType" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('LinkCSS')" class="sd"><b>��ʽ</b></a> <span id="Show_Oder_LinkCSS" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('NewsNum')" class="sd"><b>��������</b></a> <span id="Show_Oder_NewsNum" class="tx"></span></td>
	  <td align="center" class="xingmu">��ȡ����</td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('AddTime')" class="sd"><b>���ʱ��</b></a> <span id="Show_Oder_AddTime" class="tx"></span></td>
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
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;ID���</td>
      <td colspan="3"> 
        <input name="ID" type="text" id="ID" size="15" maxlength="11" value=""></td>
	</tr>
	
	
    <tr class="hback">
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;ѡ����Ŀ</td>
      <td width="35%"> 
		<input name="strClassName" type="text" id="strClassName" style="width:50%" value="" readonly="">
		<input name="strClassID" type="hidden" id="strClassID" value="">
		<input type="button" name="Submit" value="ѡ����Ŀ"   onClick="SelectClass();">	   </td>
      <td colspan="2">��ѡ�����������     </td>
    </tr>
	
	
    <tr class="hback">
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td width="35%"> 
	  	<input type="hidden" name="Act" value="SearchGo">
        <input name="FileCName" type="text" id="FileCName" style="width:90%" value=""></td>
      <td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;�ļ�����</td>
      <td width="35%"> 
        <input name="FileName" type="text" id="FileName" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ����</td>
      <td> 
		<input type="text" style="width:90%" name="ClassID" value="0" disabled>
	 </td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="NewsType" style="width:90%" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
		  <option value="">��ѡ��</option>
          <option value="RecNews">�Ƽ�����</option>
          <option value="MarqueeNews">��������</option>
          <option value="SBSNews">��������</option>
          <option value="PicNews">ͼƬ����</option>
          <option value="NewNews">��������</option>
          <option value="HotNews">�ȵ�����</option>
          <option value="WordNews">��������</option>
          <option value="TitleNews">��������</option>
          <option value="ProclaimNews">��������</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="MoreContent" id="MoreContent" style="width:90% " onChange="ChooseLink(this.options[this.selectedIndex].value);" disabled>
          <option value="1">��</option>
          <option value="0">��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" style="width:90%" value="" disabled></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="NewsNum" type="text" id="NewsNum" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;ÿ������</td>
      <td> 
        <input name="RowNum" type="text" id="RowNum" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" style="width:90%" value="" disabled></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <input name="TitleNum" type="text" id="TitleNum" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ���</td>
      <td> 
        <input name="PicWidth" type="text" id="PicWidth" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ�߶�</td>
      <td> 
	     <input name="PicHeight" type="text" id="PicHeight" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�����о�</td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;�����ٶ�</td>
      <td> 
        <input name="MarSpeed" type="text" id="MarSpeed" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="MarDirection" id="MarDirection" style="width:90% ">
		  <option value="">��ѡ��</option>
          <option value="up">����</option>
          <option value="down">����</option>
          <option value="left">����</option>
          <option value="right">����</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������</td>
      <td> 
        <input name="MarWidth" type="text" id="MarWidth" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;����߶�</td>
      <td> 
        <input name="MarHeight" type="text" id="MarHeight" style="width:90%" value=""></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��ʾ����</td>
      <td> 
        <select name="ShowTitle" id="ShowTitle" style="width:90%">
		  <option value="">��ѡ��</option>
          <option value="1">��</option>
          <option value="0">��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�¿�����</td>
      <td> 
        <select name="OpenMode" id="OpenMode" style="width:90%">
		  <option value="">��ѡ��</option>
          <option value="1">��</option>
          <option value="0">��</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;����ͼƬ</td>
      <td> 
        <input name="NaviPic" type="text" id="NaviPic" style="width:60%" value="">
        <input type="button" name="bnt_ChoosePic_naviPic"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.NaviPic);"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�м�ͼƬ</td>
      <td> 
        <input name="RowBetween" type="text" id="RowBetween" style="width:52%" value="">
        <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPathPic %>',500,300,window,document.ClassJSForm.RowBetween);"></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="DateType" id="DateType" style="width:90%">
		  <option value="">��ѡ��</option>
          <option value="0">������������</option>
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
          <option value="11">7��26��</option>
          <option value="12">26��14ʱ</option>
          <option value="13">26��14��</option>
          <option value="14">14ʱ56��</option>
          <option value="15">14:56</option>
          <option value="16">2006��7��26��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;����·��</td>
      <td> 
        <input name="SaveFilePath" type="text" id="SaveFilePath" style="width:52%" value="">
        <INPUT type="button"  name="Submit4" value="ѡ��·��" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= str_CurrPath %>',300,250,window,document.ClassJSForm.SaveFilePath);document.ClassJSForm.SaveFilePath.focus();"></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;������ʽ</td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" style="width:90%" value=""></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��ʾ��Ŀ</td>
      <td> 
        <select name="ClassName" id="ClassName" style="width:90%">
		  <option value="">��ѡ��</option>
          <option value="1">��ʾ</option>
          <option value="0">����ʾ</option>
        </select></td>
    </tr>
    <tr class="hback">  
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;��������</td>
      <td> 
        <select name="SonClass" id="SonClass" style="width:90%" disabled>
		  <option value="">��ѡ��</option>
          <option value="1">��</option>
          <option value="0">��</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;�����Ҷ���</td>
      <td> 
        <select name="RightDate" id="RightDate" style="width:90%">
		  <option value="">��ѡ��</option>
          <option value="1">��</option>
          <option value="0">��</option>
        </select></td>
    </tr>
    <tr class="hback">  
	<td colspan="10" align="center">
		<input type="submit" value=" ִ�в�ѯ ">&nbsp;&nbsp;
		<input type="reset" value=" ���� ">
	</td>
	</tr>	
  </form>
</table><p>
<%End Sub%>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
<script language="JavaScript">
<!--//�жϺ���������.�ֶ���������ʾָʾ
//�򿪺���ݹ�����ʾ��ͷ
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
			eval('Show_Oder_'+Req_FildName).innerText = '��';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '��';
		}
	}	
}
///////////////////////////////////////////////////////// 
function getCode(jsid)
{
	if (jsid!=""&&!isNaN(jsid))
	{
		OpenWindow('lib/Frame.asp?PageTitle=��ȡJS���ô���&FileName=showSysJsPath.asp&JsID='+jsid,360,140,window);
	}else
	{
		alert("���ִ�������ϵ�ͷ���Ա��")
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
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 