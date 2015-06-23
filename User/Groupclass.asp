<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<!--#include file="lib/strlib.asp" --> 
<!--#include file="lib/UserCheck.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Dim VClass_Rs,VClass_Sql
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------分页定义
int_RPP=15 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页

set VClass_Rs=User_Conn.execute("select count(*) from FS_ME_GroupDebateClass")
if VClass_Rs(0) = 0 then response.Redirect("lib/Error.asp?ErrorUrl=../main.asp&ErrCodes=<li>抱歉，社群分类尚未创建，请联系管理员。 </li>") : response.End()
VClass_Rs.close

Function set_Def(old,Def)
	if old<>"" then 
		set_Def = old
	else
		set_Def = Def
	end if
End Function

Function Get_FValue_Html(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Dim fun_ii,fun_ClassID,fun_ClassType
	Str_Tmp = "gdID,ClassID,Title,InfoType,ClassType,hits,AddTime"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_GroupDebateManage where UserNumber = '"&session("FS_UserNumber")&"'"
	if request.QueryString("Act")="SearchGo" then 
		Str_Tmp = "gdID,Title,Content,AppointUserNumber,AppointUserGroup,InfoType,AddTime,isLock,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys,hits"
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Trim(request("frm_"&Str_Tmp))<>"" then 
				Req_Str = NoSqlHack(request("frm_"&Str_Tmp))
				select case Str_Tmp
					case "gdID","InfoType","hits","AddTime","isLock","PerPageNum","isSys"
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
		''=========================================
		'''vclass表示ClassID,Hy_vclass表示ClassType
		for fun_ii = 4 to 1 step -1			
			if request("vclass"&fun_ii)<>"" then fun_ClassID = CintStr(request("vclass"&fun_ii)) : exit for
		next
		for fun_ii = 4 to 1 step -1			
			if request("Hy_vclass"&fun_ii)<>"" then fun_ClassType = CintStr(request("Hy_vclass"&fun_ii)) : exit for
		next
		if fun_ClassID = "[ChangeToTop]" then fun_ClassID = 0
		if fun_ClassType = "[ChangeToTop]" then fun_ClassType = 0
		if fun_ClassID<>"" then New_Search_Str = and_where( New_Search_Str ) & "ClassID" &" = "& CintStr(fun_ClassID)
		if fun_ClassType<>"" then New_Search_Str = and_where( New_Search_Str ) & "ClassType" &" = "& CintStr(fun_ClassType)

		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
		'response.Write(This_Fun_Sql)
		'response.End()
	end if
	Str_Tmp = ""
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	On Error Resume Next
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("lib/error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF VClass_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""GroupClass.asp?Act=Edit&gdID="&VClass_Rs("gdID")&""" class=""otherset"" title='点击修改'>〖"&VClass_Rs("gdID")&"〗</a></td>" & vbcrlf
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_GroupDebateClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"无") ''社群分类
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("Title") & "</td>" & vbcrlf
		select case VClass_Rs("InfoType")
			case 0
			Str_Tmp = "新闻"  
			case 1
			Str_Tmp = "下载"
			case 2
			Str_Tmp = "商品"
			case 3
			Str_Tmp = "房产" 
			case 4
			Str_Tmp = "供求"
			case 5
			Str_Tmp = "求职"
			case 6
			Str_Tmp = "招聘"
			case 7
			Str_Tmp = "其它"
			case else
			Str_Tmp = "无"
		end select 
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_VocationClass where VCID="&set_Def(VClass_Rs("ClassType"),0),"无") ''行业分类
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& set_Def(VClass_Rs("hits"),0) & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("AddTime") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""gdID"" id=""gdID"" value="""&VClass_Rs("gdID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html &"</td></tr>"
	
	VClass_Rs.close
	Get_FValue_Html = Get_Html
End Function

Function Get_FildValue(This_Fun_Sql,Default)
	Dim This_Fun_Rs
	set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	if not This_Fun_Rs.eof then 
		Get_FildValue = This_Fun_Rs(0)
	else
		Get_FildValue = Default
	end if
	This_Fun_Rs.close
End Function

Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql 传入sql语句,EquValue与数据库相同的值如果是<option>则加上selected,Get_Type=1为<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("lib/Error.asp?ErrCodes=<li>抱歉,Get_FildValue_List函数传入的Sql语句有问题.或表和字段不存在.</li>")
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
''================================================================

Sub Del()
	Dim Str_Tmp
	if request.QueryString("gdID")<>"" then 
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID = "&CintStr(request.QueryString("gdID")))
	else
		Str_Tmp = FormatIntArr(request.form("gdID"))
		if Str_Tmp="" then response.Redirect("lib/Error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID in ("&Str_Tmp&")")
	end if
	response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub

Sub Save()
	'''vclass表示ClassID,Hy_vclass表示ClassType
	Dim Str_Tmp,Arr_Tmp,gdID,ii,New_ClassID,New_ClassType
	for ii = 4 to 1 step -1			
		if request.Form("vclass"&ii)<>"" then New_ClassID = NoSqlHack(request.Form("vclass"&ii)) : exit for
	next
	for ii = 4 to 1 step -1			
		if request.Form("Hy_vclass"&ii)<>"" then New_ClassType = NoSqlHack(request.Form("Hy_vclass"&ii)) : exit for
	next
	if New_ClassID = "[ChangeToTop]" then New_ClassID = 0
	if New_ClassType = "[ChangeToTop]" then New_ClassType = 0
	Str_Tmp = "ClassID,Title,Content,InfoType,ClassType,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,AddTime,isSys,isLock,hits"
	gdID = NoSqlHack(request.Form("gdID"))
	if not isnumeric(gdID) or gdID = "" then gdID = 0 
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_GroupDebateManage where gdID="&CintStr(gdID)
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	Str_Tmp = "Title,Content,InfoType,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys"
	Arr_Tmp = split(Str_Tmp,",")
	if gdID > 0 then 
	''修改
		''''''''''''''''''''''''''
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next
		if New_ClassID<>"" then VClass_Rs("ClassID") = New_ClassID
		if New_ClassType<>"" then VClass_Rs("ClassType") = New_ClassType
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=Edit&gdID="&gdID )&"&ErrCodes=<li>恭喜，修改社群成功。</li>")
	else
	''新增
	''获得群数量
'		Dim rsCount,CountSQL
'		set rsCount = Server.CreateObject(G_FS_RS)
'		CountSQL = "select gdID From FS_ME_GroupDebateManage where UserNumber='"&Fs_User.UserNumber&"'"
'		rsCount.open CountSQL,User_Conn,1,1
'		Call getGroupIDinfo()
'		if Cint(GroupDebateNum) <= rsCount.recordcount then
'			Response.Redirect("lib/Error.asp?ErrCodes=<li>您建立的社群数量已经超过最大极限，您允许建立的社群数量为："&split(GroupDebateNum,",")(0)&"个。</li>")
'			Response.end
'		end if
		VClass_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form("frm_"&Str_Tmp))&"<br>")
		next
		VClass_Rs("AddTime") = NoSqlHack(request.Form("frm_AddTime"))
		VClass_Rs("isLock") = NoSqlHack(request.Form("frm_isLock"))
		VClass_Rs("hits") = NoSqlHack(request.Form("frm_hits"))
		VClass_Rs("ClassID") = set_Def(New_ClassID,0)
		VClass_Rs("ClassType") = set_Def(New_ClassType,0)
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=Add&gdID="&gdID )&"&ErrCodes=<li>恭喜，新增社群成功。</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<%if instr(",Add,Edit,Search,",","&request.QueryString("Act")&",")>0 then%>
<script language="javascript" src="../FS_Inc/class_liandong.js" type="text/javascript"></script>
<%end if%>
<script language="JavaScript">
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

</script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;"> 
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table"> 
  <tr> 
    <td> <!--#include file="top.asp" --> </td> 
  </tr> 
</table> 
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table"> 
  <tr class="back"> 
    <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td> 
  </tr> 
  <tr class="back"> 
    <td width="18%" valign="top" class="hback"> <div align="left"> 
        <!--#include file="menu.asp" --> 
      </div></td> 
    <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td width="72%"  valign="top"> 

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
 <tr  class="hback"> 
    <td align="left" class="hback_1"><a href="GroupClass.asp?Act=Add">新建</a>	
	 | <%if request.QueryString("Act")="Edit" then
	 	response.Write("<a href=""GroupClass.asp?Act=Del&gdID="&NoSqlHack(request.QueryString("gdID"))&""" title=""确定删除这条记录吗?"">删除</a>")
	elseif request.QueryString("Act")="View" or request.QueryString("Act")="" then
		response.Write("<a href=""javascript:if (confirm('确定删除吗?')) document.form1.submit();"">删除</a>")
	else
		response.Write("删除")		
	end if%> | <a href="GroupClass.asp?Act=View">查看全部</a> | <a href="GroupClass.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("islock"))%>">已锁定</a> 
	| <a href="GroupClass.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("not islock"))%>">未锁定</a>
	 | <a href="GroupClass.asp?Act=Search">查询</a></td>
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
      <td align="center" class="hback_1"><a href="javascript:OrderByName('gdID')" class="sd"><b>〖编号〗</b></a> 
        <span id="Show_Oder_gdID"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('ClassID')" class="sd"><b>社群分类</b></a> 
        <span id="Show_Oder_ClassID"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('Title')" class="sd"><b>社群主题</b></a> 
        <span id="Show_Oder_Title"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('InfoType')" class="sd"><b>应用子类</b></a> 
        <span id="Show_Oder_InfoType"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('ClassType')" class="sd"><b>所属行业</b></a> 
        <span id="Show_Oder_ClassType"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('hits')" class="sd"><b>人气</b></a> 
        <span id="Show_Oder_hits"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('AddTime')" class="sd"><b>加入时间</b></a> 
        <span id="Show_Oder_AddTime"></span></td>
      <td width="2%" align="center" class="hback_1"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_FValue_Html( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim gdID,Bol_IsEdit,AppointUserNumber,AppointUserGroup
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	gdID = request.QueryString("gdID")
	if gdID="" then response.Redirect("lib/Error.asp?ErrorUrl=&ErrCodes=<li>必要的gdID没有提供</li>") : response.End()
	VClass_Sql = "select gdID,ClassID,Title,Content,InfoType,ClassType,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID="&NoSqlHack(gdID)
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("lib/Error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
	Bol_IsEdit = True
	AppointUserNumber = VClass_Rs(4)
	AppointUserGroup = VClass_Rs(5)
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return Validator.Validate(this,3);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" > <%if Bol_IsEdit then response.Write("修改社群信息"&vbNewLine&"<input type=""hidden"" name=""gdID"" value="""&VClass_Rs(0)&""">" ) else response.Write("添加社群信息")  end if%> </td>
    </tr>
    <tr  class="hback"<%if not Bol_IsEdit then response.Write(" style=""display:none""") end if%>> 
      <td width="25%" align="right">所属社群分类</td>
      <td><strong>
        <%if Bol_IsEdit then response.Write( Get_FildValue( "select vClassName from FS_ME_GroupDebateClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"无" ) ) end if%>
        </strong></td>
    </tr>
    <tr  class="hback"> 
      <td width="25%" align="right"><%if Bol_IsEdit then response.Write("〖变更为〗") else response.Write("所属社群分类") end if%></td>
      <td> 
        <!---联动菜单开始--->
        <select name="vclass1" id="vclass1"<%if not Bol_IsEdit then%> datatype="Require" msg="必须填写"<%end if%> onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群主题</td>
      <td> <input type="text" name="frm_Title" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(2)) end if%>" dataType="Require" msg="必须填写">
                    支持A* *B A B其它字符型同理</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群公告</td>
      <td> <textarea name="frm_Content" cols="40" rows="5" dataType="Require" msg="必须填写"><%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%></textarea> 
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">用于哪个子类</td>
      <td>
		  <select name="frm_InfoType" datatype="Require" msg="必须选择">
          <option value="0"<%if Bol_IsEdit then if VClass_Rs(4)=0 then response.Write(" selected") end if end if%>>新闻</option>
          <%if IsExist_SubSys("DS") Then%><option value="1"<%if Bol_IsEdit then if VClass_Rs(4)=1 then response.Write(" selected") end if end if%>>下载</option><%end if%>
          <%if IsExist_SubSys("MS") Then%><option value="2"<%if Bol_IsEdit then if VClass_Rs(4)=2 then response.Write(" selected") end if end if%>>商品</option><%end if%>
          <%if IsExist_SubSys("HS") Then%><option value="3"<%if Bol_IsEdit then if VClass_Rs(4)=3 then response.Write(" selected") end if end if%>>房产</option><%end if%>
          <%if IsExist_SubSys("SD") Then%><option value="4"<%if Bol_IsEdit then if VClass_Rs(4)=4 then response.Write(" selected") end if end if%>>供求</option><%end if%>
          <%if IsExist_SubSys("AP") Then%><option value="5"<%if Bol_IsEdit then if VClass_Rs(4)=5 then response.Write(" selected") end if end if%>>求职</option><%end if%>
          <%if IsExist_SubSys("AP") Then%><option value="6"<%if Bol_IsEdit then if VClass_Rs(4)=6 then response.Write(" selected") end if end if%>>招聘</option>
         <option value="7"<%if Bol_IsEdit then if VClass_Rs(4)=7 then response.Write(" selected") end if end if%>>其它</option><%end if%>
        </select> </td>
    </tr>
    <tr class="hback"<%if not Bol_IsEdit then response.Write(" style=""display:none""") end if%>> 
      <td align="right">社群所属行业</td>
      <td><strong>
        <%if Bol_IsEdit then response.Write( Get_FildValue( "select vClassName from FS_ME_VocationClass where VCID="&set_Def(VClass_Rs("ClassType"),0),"无" ) ) end if%>
        </strong></td>
    </tr>
    <tr  class="hback"> 
      <td align="right"><%if Bol_IsEdit then response.Write("〖变更为〗") else response.Write("社群所属行业") end if%></td>
      <td> 
        <!---联动菜单开始--->
        <select name="Hy_vclass1" id="select"<%if not Bol_IsEdit then%> datatype="Require" msg="必须填写"<%end if%> onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--->
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">附件地址</td>
      <td> <input type="text" name="frm_AccessFile" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(6)) end if%>"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群创始人用户编号</td>
      <td> <input name="frm_UserNumber" type="text" value="<%if Bol_IsEdit then response.Write(VClass_Rs(7)) else response.Write(session("FS_UserNumber")) end if%>" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群现在管理员用户编号</td>
      <td> <input name="frm_AdminName" type="text" value="<%if Bol_IsEdit then response.Write(VClass_Rs(8)) else response.Write(session("FS_UserNumber")) end if%>" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群的成员</td>
      <td> <textarea name="frm_ClassMember" cols="40" datatype="Require" msg="必须填写"><%if Bol_IsEdit then response.Write(VClass_Rs(9)) else response.Write(session("FS_UserNumber")) end if%></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群讨论每页显示多少数量</td>
      <td>
	   <input type="text" name="frm_PerPageNum" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(10)) end if%>" dataType="Range" msg="在1~30之间" min="0" max="31"> 
       <input type="hidden" name="frm_isSys" value="0">
	   <%if not Bol_IsEdit then%>
	   <input type="hidden" name="frm_AddTime" value="<%=now()%>">
       <input type="hidden" name="frm_isLock" value="1">
	   <input type="hidden" name="frm_hits" value="0">
	   <%end if%>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 保存 " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " /> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%End Sub
Sub Search()
%>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="?Act=SearchGo">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >查询社群信息</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ID编号</td>
      <td><input type="text" name="frm_gdID" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td width="25%" align="right">所属社群分类</td>
      <td> 
        <!---联动菜单开始--->
        <select name="vclass1" id="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群主题</td>
      <td><input type="text" name="frm_Title" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群公告</td>
      <td> <textarea name="frm_Content" cols="40" rows="5"></textarea> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">可查看的用户编号</td>
      <td> <textarea name="frm_AppointUserNumber" cols="40"></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr class="hback"> 
      <td align="right">可查看的会员组</td>
      <td> <textarea name="frm_AppointUserGroup" cols="40"></textarea> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">用于哪个子类</td>
      <td> <select name="frm_InfoType">
          <option value="">请选择</option>
          <option value="0">新闻</option>
          <option value="1">下载</option>
          <option value="2">商品</option>
          <option value="3">房产</option>
          <option value="4">供求</option>
          <option value="5">求职</option>
          <option value="6">招聘</option>
          <option value="7">其它</option>
        </select> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群所属行业</td>
      <td> 
        <!---联动菜单开始--->
        <select name="Hy_vclass1" id="select" onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">加入时间</td>
      <td> <input type="text" name="frm_AddTime" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否锁定</td>
      <td> <input type="radio" name="frm_isLock"  value="true">
        已锁定 
        <input type="radio" name="frm_isLock"  value="false">
        未锁定 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">附件地址</td>
      <td> <input type="text" name="frm_AccessFile" size="40" value=""> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群创始人用户编号</td>
      <td> <input name="frm_UserNumber" type="text" value="" size="40"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群现在管理员用户编号</td>
      <td> <input name="frm_AdminName" type="text" value="" size="40"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群的成员</td>
      <td> <textarea name="frm_ClassMember" cols="40"></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群讨论每页显示多少数量</td>
      <td> <input type="text" name="frm_PerPageNum" size="40" value="" require="false" dataType="Range" msg="在1~30之间" min="0" max="31"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">人气/点击数</td>
      <td> <input type="text" name="frm_hits" size="40">
                    支持&gt;=&lt;&lt;&gt;等符号其它数字日期型同理</td>
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

</td> 
        </tr> 
      </table></td> 
  </tr> 
  <tr class="back"> 
    <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
        <!--#include file="Copyright.asp" --> 
      </div></td> 
  </tr> 
</table> 
</BODY>
<%
MF_User_Conn
%>

<script language="javascript">
<!-- 
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
<%if instr(",Add,Edit,Search,",","&request.QueryString("Act")&",")>0 then%>
var array=new Array();
<%dim js_sql,js_rs,js_i
  js_sql="select VCID,ParentID,vClassName from FS_ME_GroupDebateClass"
  set js_rs=User_Conn.execute(js_sql)
  js_i=0
  do while not js_rs.eof
%>
array[<%=js_i%>]=new Array("<%=js_rs("VCID")%>","<%=js_rs("ParentID")%>","<%=js_rs("vClassName")%>"); 
<%
	js_rs.movenext
	js_i=js_i+1
loop
js_rs.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");

///行业

var array1=new Array();
<%dim js_sql1,js_rs1,js_i1
  js_sql1="select VCID,ParentID,vClassName from FS_ME_VocationClass"
  set js_rs1=User_Conn.execute(js_sql1)
  js_i1=0
  do while not js_rs1.eof
%>
array1[<%=js_i1%>]=new Array("<%=js_rs1("VCID")%>","<%=js_rs1("ParentID")%>","<%=js_rs1("vClassName")%>"); 
<%
	js_rs1.movenext
	js_i1=js_i1+1
loop
js_rs1.close              
%>

var liandong=new CLASS_LIANDONG_YAO(array1)
liandong.firstSelectChange("0","Hy_vclass1");
liandong.subSelectChange("Hy_vclass1","Hy_vclass2");
liandong.subSelectChange("Hy_vclass2","Hy_vclass3");
liandong.subSelectChange("Hy_vclass3","Hy_vclass4");

function RemoveChildopt(obj,StrList)
{
	var TmpArr = StrList.split(',');
	if(obj.selectedIndex<2)
	{		
		for (var i=TmpArr.length-1 ; i>=0; i--)
		{
			//alert(TmpArr[i]);
			if (TmpArr[i]!='') 
				//清除下拉内容
				for (var j=document.getElementById(TmpArr[i]).options.length-1 ; j>=0 ; j--)
				document.getElementById(TmpArr[i]).options.remove(j);				
		}	
	}
} 
<%end if%>
-->
</script>

<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
</HTML>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->