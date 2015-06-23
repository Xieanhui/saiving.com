<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,User_Conn,VClass_Rs,VClass_Sql
Dim CheckStr,Sys_MoneyName
MF_Default_Conn
MF_User_Conn
MF_Session_TF
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

set VClass_Rs=User_Conn.execute("select top 1 MoneyName from FS_ME_SysPara")
if not VClass_Rs.eof then 
Sys_MoneyName = VClass_Rs(0)
end if
VClass_Rs.close

Function Get_Card(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "LogID,LogType,UserNumber,points,moneys,LogTime,LogContent,Logstyle"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_Log"
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Trim(request.Form("frm_"&Str_Tmp))<>"" then 
				Req_Str = NoSqlHack(Trim(request("frm_"&Str_Tmp)))
				
				if Str_Tmp="points" then
					Req_Str = Replace(request("JF1")& Req_Str,",","")
				elseif  Str_Tmp="moneys" then
					Req_Str =  Replace(request("JB1")& Req_Str,",","")
				elseif Str_Tmp="LogTime" then
					Req_Str = Replace(request("RQ")& "#"&Req_Str&"#",",","")
				end if

				select case Str_Tmp
					case "points","moneys","LogTime","Logstyle"
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
						New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '%"& Req_Str & "%'"
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	'response.End()
	On Error Resume Next
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	IF not VClass_Rs.eof THEN
	VClass_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>VClass_Rs.PageCount Then cPageNo=VClass_Rs.PageCount 
	VClass_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Integral_MoreInfo.asp?Act=Edit&LogID="&VClass_Rs("LogID")&""" class=""otherset"" title='点击修改'>"&VClass_Rs("LogID")&"</a></td>" & vbcrlf
		for ii=1 to 7  
			select case ii
				case 7 
				if VClass_Rs(ii)=1 then 
					Str_Tmp="减少"
				else
					Str_Tmp="增加"
				end if		
				case 3
				Str_Tmp = VClass_Rs(ii) & "点"
				case 4
				Str_Tmp = VClass_Rs(ii) & Sys_MoneyName
				case else
				Str_Tmp = VClass_Rs(ii)
			end select		
				Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		next
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><a href=""Integral.asp?Act=View&Add_Sql="&server.URLEncode( Encrypt("UserNumber='"&VClass_Rs("UserNumber")&"'") )&""" class=""otherset"" title='查看该用户信息'>用户信息</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""LogID"" id=""LogID"" value="""&VClass_Rs("LogID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_Card = Get_Html
End Function

Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql 传入sql语句,EquValue与数据库相同的值如果是<option>则加上selected,Get_Type=1为<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>抱歉,传入的Sql语句有问题.或表和字段不存在.</li>")
do while not This_Fun_Rs.eof 
	select case Get_Type
	  case 1
		''<option>		
		if instr(This_Fun_Sql,",") >0 then 
			Text = This_Fun_Rs(1)
		else
			Text = This_Fun_Rs(0)
		end if	
		if EquValue = This_Fun_Rs(0) then 
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&"""  style=""color:#0000FF"" selected>"&Text&"</option>"&vbNewLine
		else
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&""">"&Text&"</option>"&vbNewLine
		end if		
	  case else
		exit do : Get_FildValue_List = "Get_Type值传入错误" : exit Function
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_FildValue_List = Get_Html
End Function

Sub Del()
	Dim Str_Tmp
	if request.QueryString("LogID")<>"" then 
		User_Conn.execute("Delete from FS_ME_Log where LogID = "&CintStr(request.QueryString("LogID")))
	else
		Str_Tmp = FormatIntArr(request.form("LogID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		Str_Tmp = replace(Str_Tmp," ","")
		User_Conn.execute("Delete from FS_ME_Log where LogID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral_MoreInfo.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================

Sub Save()
	Dim Str_Tmp,Arr_Tmp,LogID
	Str_Tmp = "LogType,UserNumber,points,moneys,LogTime,LogContent,Logstyle"
	Arr_Tmp = split(Str_Tmp,",")
	LogID = NoSqlHack(request.Form("LogID"))
	if not isnumeric(LogID) or LogID = "" then LogID = 0 
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_Log where LogID="&LogID
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if LogID > 0 then 
	''修改
		On error resume next
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next	
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral_MoreInfo.asp?Act=Edit&LogID="&LogID )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	''新增
		On error resume next
		VClass_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next	
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral_MoreInfo.asp?Act=Add&LogID="&LogID )&"&ErrCodes=<li>恭喜，新增成功。</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
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
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace(oldFildName,FildName):New_Sql = Old_Sql.replace(oldFildName,FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
-->
</script>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
.style2 {color: #FF0000}
-->
</style>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >积分详细清单管理</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="Integral_MoreInfo.asp?Act=View">查看全部</a> | <a href="Integral_MoreInfo.asp?Act=Add">新增</a> 
      | <a href="Integral_MoreInfo.asp?Act=Search">查询</a> | <a href="Integral.asp?Act=View">返回首页</a></td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
	View
	case "Add","Edit"
	Add_Edit
	case "Save"
	Save
	case "Del"
	Del
	case "Search"
	Search
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="?Act=Del">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('LogID')" class="sd"><b>〖日志编号〗</b></a> <span id="Show_Oder_LogID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('LogType')" class="sd"><b>日志类型</b></a> <span id="Show_Oder_LogType"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('UserNumber')" class="sd"><b>会员编号</b></a> <span id="Show_Oder_UserNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('points')" class="sd"><b>积分</b></a> <span id="Show_Oder_points"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('moneys')" class="sd"><b>金币</b></a> <span id="Show_Oder_moneys"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('LogTime')" class="sd"><b>交易日期</b></a> <span id="Show_Oder_LogTime"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('LogContent')" class="sd"><b>交易说明</b></a> <span id="Show_Oder_LogContent"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('Logstyle')" class="sd"><b>类型</b></a> <span id="Show_Oder_Logstyle"></span></td>
	  <td align="center" class="xingmu">用户信息</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_Card( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim LogID,Bol_IsEdit
Dim UserNum,UserLog,LogCont
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	LogID = request.QueryString("LogID")
	if LogID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的LogID没有提供</li>") : response.End()
	VClass_Sql = "select LogID,LogType,UserNumber,points,moneys,LogTime,LogContent,Logstyle from FS_ME_Log where LogID="&LogID
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
	Bol_IsEdit = True
	UserNum = VClass_Rs(2) : UserLog = VClass_Rs(1) : LogCont = VClass_Rs(6)
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return Validator.Validate(this,3);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("修改用户交易详细信息<input type=""hidden"" name=""LogID"" value="""&VClass_Rs(0)&""">") else response.Write("添加用户交易详细信息") end if%></td>
	</tr>
    <tr  class="hback"> 
      <td align="right">日志类型</td>
      <td>
	 	 <input type="text" name="frm_LogType" size="40" value="<%if Bol_IsEdit then response.Write(UserLog) end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
	  	 <select style="width:120" name="select11" onChange="frm_LogType.value=this.options[this.selectedIndex].value">
		    <option value="">---请选择---</option>
		 	<%=Get_FildValue_List("select distinct LogType from FS_ME_Log",UserLog,1)%>
		 </select>		 
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">会员编号</td>
      <td>
	 	 <input type="text" name="frm_UserNumber" size="40" value="<%if Bol_IsEdit then response.Write(UserNum) end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual"><span id="usernum" style="color:#FF0000">*请填入真实会员编号,否则查看的用户信息有误</span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">交易积分</td>
      <td>
	 	 <input type="text" name="frm_points" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
	  </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">交易金额</td>
      <td>
	  <input type="text" name="frm_moneys" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(4)) end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
	   <%=Sys_MoneyName%>
	  </td>

    <tr  class="hback"> 
      <td align="right">交易日期</td>
      <td>
	  <input name="frm_LogTime" type="text" id="frm_LogTime" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(5)) end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
        &nbsp;
        <input type="button" name="chooseEndTime" value="日 期" onClick="OpenWindowAndSetValue('../../../admin/CommPages/SelectDate.asp',280,110,window,document.all.frm_LogTime);document.all.frm_LogTime.focus();">
	  </td>
    </tr>
	
    <tr  class="hback"> 
      <td align="right">交易说明</td>
      <td>
	 	 <textarea name="frm_LogContent" cols="40" datatype="Compare" msg="必须&gt;=0" to="0" operator="GreaterThanEqual"><%if Bol_IsEdit then response.Write(LogCont) end if%></textarea>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">类型</td>
      <td>
        <input type="radio" name="frm_Logstyle" value="0" <%if Bol_IsEdit then if VClass_Rs(7)=0 then response.Write(" checked ") end if else response.Write(" checked ") end if%>>
          增加		  
        <input type="radio" name="frm_Logstyle" value="1" <%if Bol_IsEdit then if VClass_Rs(7)=1 then response.Write(" checked ") end if end if%>>
          减少
	  </td>
    </tr>

    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 保存 " /> <!--<%IF request.QueryString("Act")="Put" then%> onClick="Put_CardNum_Len.to = (Put_CardAddStr.value.length+2).toString();Put_CardNum_Len.msg='长度必须大于等于'+(Put_CardAddStr.value.length+2).toString()" <%end if%>-->
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="javascript:return confirm('确定要删除所选项目吗?');">
           </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub

Sub Search()
%>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" onSubmit="return Validator.Validate(this,3)" method="post" action="?Act=SearchGo">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >查询会员消费清单</td>
    </tr>
<tr  class="hback"> 
      <td align="right">日志类型</td>
      <td>
	 	 <input type="text" name="frm_LogType" size="40" value="">
	  	 <select style="width:120" name="select11" onChange="frm_LogType.value=this.options[this.selectedIndex].value">
		    <option value="">---请选择---</option>
		 	<%=Get_FildValue_List("select distinct LogType from FS_ME_Log","",1)%>
		 </select>		 
        模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">会员编号</td>
      <td>
	 	 <input type="text" name="frm_UserNumber" size="40" value="">
        模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">交易积分</td>
      <td>
	 	 <select name="JF1" style="width:55">
	     <option value="" selected="selected"></option>
	     <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select> 
	 	 <input type="text" name="frm_points" size="30" value="" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
        数字,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">交易金额</td>
      <td>
	   <select name="JB1" style="width:55">
	     <option value="" selected="selected"></option>
	     <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select> 
	   <input type="text" name="frm_moneys" size="30" value="" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
	   <%=Sys_MoneyName%> 数字,可在开头加上简单比较符号,*号表示模糊查询 </td>

    <tr  class="hback"> 
      <td align="right">交易日期</td>
      <td>
	 	 <select name="RQ" style="width:55">
	     <option value="" selected="selected"></option>
	     <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select> 
	 	 <input type="text" name="frm_LogTime" size="17" value="" readonly>
        <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_LogTime);"> 
        日期,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
	
    <tr  class="hback"> 
      <td align="right">交易说明</td>
      <td>
	 	 <textarea name="frm_LogContent" cols="40"></textarea>
        模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">类型</td>
      <td>
        <input type="radio" name="frm_Logstyle" value="0" >
          增加		  
        <input type="radio" name="frm_Logstyle" value="1" >
          减少
	  </td>
    </tr>
	<tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 执行查询 " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " /> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>

<%End Sub%>
</body>
<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
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
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 






