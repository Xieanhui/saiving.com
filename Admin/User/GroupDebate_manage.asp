<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<% 
'on error resume next
Dim Conn,User_Conn,VClass_Rs,VClass_Sql
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Form") then Err_Show 

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

set VClass_Rs=User_Conn.execute("select count(*) from FS_ME_GroupDebateClass")
if VClass_Rs(0) = 0 then response.Redirect("GroupDebate_Class.asp") : response.End()
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
	Str_Tmp = "gdID,ClassID,Title,InfoType,ClassType,hits,AddTime,isLock"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_GroupDebateManage"
	if Add_Sql<>"" then 
		if Add_Sql = "islock" Then
			This_Fun_Sql = and_where(This_Fun_Sql) &" "&"isLock = '1'"
		ElseIf Add_Sql = "notislock" Then
			This_Fun_Sql = and_where(This_Fun_Sql) &" "&"isLock = '0'"
		End If
	End If	
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	if request.QueryString("Act")="SearchGo" then 
		Str_Tmp = "gdID,Title,Content,AppointUserNumber,AppointUserGroup,InfoType,AddTime,isLock,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys,hits"
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Trim(request("frm_"&Str_Tmp))<>"" then 
				Req_Str = NoSqlHack(Trim(request("frm_"&Str_Tmp)))
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
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
				end select 		
			end if
		next
		''=========================================
		'''vclass表示ClassID,Hy_vclass表示ClassType
		for fun_ii = 4 to 1 step -1			
			if request.Form("vclass"&fun_ii)<>"" then fun_ClassID = request.Form("vclass"&fun_ii) : exit for
		next
		for fun_ii = 4 to 1 step -1			
			if request.Form("Hy_vclass"&fun_ii)<>"" then fun_ClassType = request.Form("Hy_vclass"&fun_ii) : exit for
		next
		if fun_ClassID = "[ChangeToTop]" then fun_ClassID = 0
		if fun_ClassType = "[ChangeToTop]" then fun_ClassType = 0
		if fun_ClassID<>"" then New_Search_Str = and_where( New_Search_Str ) & "ClassID" &" = "& fun_ClassID
		if fun_ClassType<>"" then New_Search_Str = and_where( New_Search_Str ) & "ClassType" &" = "& fun_ClassType

		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	
	Str_Tmp = ""
	'On Error Resume Next
	'response.Write(This_Fun_Sql)
	'response.End()
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF VClass_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
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
		Get_Html = Get_Html & "<td align=""center""><a href=""GroupDebate_manage.asp?Act=Edit&gdID="&VClass_Rs("gdID")&""" class=""otherset"" title='点击修改'>"&VClass_Rs("gdID")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""GroupDebate_manage.asp?Act=Edit&gdID="&VClass_Rs("gdID")&""" class=""otherset"" title='点击修改'>"& VClass_Rs("Title") & "</a></td>" & vbcrlf
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_GroupDebateClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"无") ''社群分类
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
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
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_VocationClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"无") ''行业分类
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& set_Def(VClass_Rs("hits"),0) & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("AddTime") & "</td>" & vbcrlf
		if cbool(VClass_Rs("isLock")) then 
			''锁定,需要解锁
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""锁 定"" onclick=""javascript:location='GroupDebate_manage.asp?Act=OtherEdit&EditSql="&server.URLEncode(Encrypt( "Update FS_ME_GroupDebateManage set isLock=0 where gdID="&VClass_Rs("gdID")) )&"&Red_Url='"" title=""点击解锁"" style=""color:red""></td>" & vbcrlf
		else
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""正 常"" onclick=""javascript:location='GroupDebate_manage.asp?Act=OtherEdit&EditSql="&server.URLEncode(Encrypt( "Update FS_ME_GroupDebateManage set isLock=1 where gdID="&VClass_Rs("gdID")) )&"&Red_Url='"" title=""点击锁定""></td>" & vbcrlf
		end if
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""gdID"" id=""gdID"" value="""&VClass_Rs("gdID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
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
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>抱歉,Get_FildValue_List函数传入的Sql语句有问题.或表和字段不存在.</li>")
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
End Function ''================================================================


Sub OtherEdit()
	if not MF_Check_Pop_TF("ME025") then Err_Show '权限判断
	Dim Red_Url
	Red_Url = request.QueryString("Red_Url")
	if Red_Url = "" then Red_Url = "GroupDebate_manage.asp"
	On Error Resume Next
	if request.QueryString("EditSql")<>"" then 
		User_Conn.execute( Decrypt(request.QueryString("EditSql")) )
		If Err.Number <> 0 then response.Redirect("../error.asp?ErrCodes=<li>抱歉,OtherEdit过程传入的Sql语句有问题.或表和字段不存在.</li>")
	end if
	response.Redirect(Red_Url)
End Sub

Sub Del()
	if not MF_Check_Pop_TF("ME024") then Err_Show '权限判断
	Dim Str_Tmp
	if request.QueryString("gdID")<>"" then 
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where gdID = "&CintStr(request.QueryString("gdID")))
	else
		Str_Tmp = FormatIntArr(request.form("gdID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where gdID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/GroupDebate_manage.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub

Sub Save()
	'''vclass表示ClassID,Hy_vclass表示ClassType
	Dim Str_Tmp,Arr_Tmp,gdID,ii,New_ClassID,New_ClassType,kkk
	for ii = 4 to 1 step -1			
		if request.Form("vclass"&ii)<>"" then New_ClassID = request.Form("vclass"&ii) : exit for
	next
	for ii = 4 to 1 step -1			
		if request.Form("Hy_vclass"&ii)<>"" then New_ClassType = request.Form("Hy_vclass"&ii) : exit for
	next
	if New_ClassID = "[ChangeToTop]" then New_ClassID = 0
	if New_ClassType = "[ChangeToTop]" then New_ClassType = 0
	Str_Tmp = "ClassID,Title,Content,AppointUserNumber,AppointUserGroup,InfoType,ClassType,AddTime,isLock,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys,hits"
	gdID = NoSqlHack(request.Form("gdID"))
	if not isnumeric(gdID) or gdID = "" then gdID = 0 
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_GroupDebateManage where gdID="&gdID
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	Str_Tmp = "Title,Content,AppointUserNumber,AppointUserGroup,InfoType,AddTime,isLock,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys,hits"
	Arr_Tmp = split(Str_Tmp,",")
	if gdID > 0 then 
	''修改
		''''''''''''''''''''''''''
		if not MF_Check_Pop_TF("ME024") then Err_Show '权限判断
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next
		if New_ClassID<>"" then VClass_Rs("ClassID") = New_ClassID
		if New_ClassType<>"" then VClass_Rs("ClassType") = New_ClassType
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/GroupDebate_manage.asp?Act=Edit&gdID="&gdID )&"&ErrCodes=<li>恭喜，修改社群成功。</li>")
	else
	''新增
		if not MF_Check_Pop_TF("ME022") then Err_Show '权限判断
		VClass_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next
		VClass_Rs("ClassID") = set_Def(New_ClassID,0)
		VClass_Rs("ClassType") = set_Def(New_ClassType,0)
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/GroupDebate_manage.asp?Act=Add&gdID="&gdID )&"&ErrCodes=<li>恭喜，新增社群成功。</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<%if instr(",Add,Edit,Search,",","&request.QueryString("Act")&",")>0 then%>
<script language="javascript" src="../../FS_Inc/class_liandong.js" type="text/javascript"></script>
<%end if%>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;"> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >社群总管理</td>
  </tr>
 <tr  class="hback"> 
    <td align="left"><a href="GroupDebate_manage.asp?Act=View">管理首页</a> 
      | <a href="GroupDebate_manage.asp?Act=Add">新建</a> | <a href="GroupDebate_manage.asp?Act=View&Add_Sql=islock">已锁定</a> 
      | <a href="GroupDebate_manage.asp?Act=View&Add_Sql=notislock">未锁定</a> 
      | <a href="GroupDebate_manage.asp?Act=Search">查询</a> | <a href="GroupDebate_Class.asp">社群分类管理</a></td>
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
	case "OtherEdit"
		OtherEdit
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="?Act=Del">
    <tr  class="hback">  
      <td align="center" class="xingmu">〖编号〗</td>
	  <td align="center" class="xingmu">社群主题</td>
      <td align="center" class="xingmu">社群分类</td>
      <td align="center" class="xingmu">应用子类</td>
	  <td align="center" class="xingmu">所属行业</td>
      <td align="center" class="xingmu">人气</td>
	  <td align="center" class="xingmu">加入时间</td>
      <td align="center" class="xingmu">是否锁定</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
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
	if gdID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的gdID没有提供</li>") : response.End()
	VClass_Sql = "select gdID,ClassID,Title,Content,AppointUserNumber,AppointUserGroup,InfoType,ClassType,AddTime,isLock,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum,isSys,hits from FS_ME_GroupDebateManage where gdID="&gdID
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
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
        <!---联动菜单开始  onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"--->
        <select name="vclass1" id="vclass1"<%if not Bol_IsEdit then%> datatype="Require" msg="必须填写"<%end if%> style="width:100px">
          <option></option>
       </select>
<!--    <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--- > -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群主题</td>
      <td> <input type="text" name="frm_Title" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(2)) end if%>" dataType="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群公告</td>
      <td> <textarea name="frm_Content" cols="40" rows="5"><%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%></textarea> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">可查看的用户编号</td>
      <td> <textarea name="frm_AppointUserNumber" cols="40"  datatype="Require" msg="必须填写"><%if Bol_IsEdit then response.Write(VClass_Rs(4)) end if%></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">可查看的会员组</td>
      <td> <textarea name="frm_AppointUserGroup" cols="40"  datatype="Require" msg="必须填写"><%if Bol_IsEdit then response.Write(VClass_Rs(5)) end if%></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">用于哪个子类</td>
      <td> <select name="frm_InfoType" datatype="Require" msg="必须选择">
          <option value="0"<%if Bol_IsEdit then if VClass_Rs(6)=0 then response.Write(" selected") end if end if%>>新闻</option>
          <option value="1"<%if Bol_IsEdit then if VClass_Rs(6)=1 then response.Write(" selected") end if end if%>>下载</option>
          <option value="2"<%if Bol_IsEdit then if VClass_Rs(6)=2 then response.Write(" selected") end if end if%>>商品</option>
          <option value="3"<%if Bol_IsEdit then if VClass_Rs(6)=3 then response.Write(" selected") end if end if%>>房产</option>
          <option value="4"<%if Bol_IsEdit then if VClass_Rs(6)=4 then response.Write(" selected") end if end if%>>供求</option>
          <option value="5"<%if Bol_IsEdit then if VClass_Rs(6)=5 then response.Write(" selected") end if end if%>>求职</option>
          <option value="6"<%if Bol_IsEdit then if VClass_Rs(6)=6 then response.Write(" selected") end if end if%>>招聘</option>
          <option value="7"<%if Bol_IsEdit then if VClass_Rs(6)=7 then response.Write(" selected") end if end if%>>其它</option>
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
        <!---联动菜单开始 onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"--->
        <select name="Hy_vclass1" id="select"<%if not Bol_IsEdit then%> datatype="Require" msg="必须填写"<%end if%>  style="width:100px">
          <option></option>
        </select>
<!--		<select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--- > -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">加入时间</td>
      <td> <input type="text" name="frm_AddTime" size="27" value="<%if Bol_IsEdit then response.Write(VClass_Rs(8)) else response.Write(date()) end if%>">
      <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);">      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否锁定</td>
      <td> <input type="radio" name="frm_isLock"  value="true" <%if Bol_IsEdit then if VClass_Rs(9) then response.Write(" checked ") end if end if%>>
        已锁定 
        <input type="radio" name="frm_isLock"  value="false" <%if Bol_IsEdit then if not VClass_Rs(9) then response.Write(" checked ") end if else response.Write(" checked ") end if%>>
        不锁定 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">附件地址</td>
      <td> <input type="text" name="frm_AccessFile" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(10)) end if%>" require="false" dataType="Url"   msg="非法的Url"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群创始人用户编号</td>
      <td> <input name="frm_UserNumber" type="text" onBlur="if(frm_AdminName.value=='')frm_AdminName.value=this.value;" value="<%if Bol_IsEdit then response.Write(VClass_Rs(11)) end if%>" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群现在管理员用户编号</td>
      <td> <input name="frm_AdminName" type="text"  onBlur="if(frm_ClassMember.value=='')frm_ClassMember.value=this.value;" value="<%if Bol_IsEdit then response.Write(VClass_Rs(12)) end if%>" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群的成员</td>
      <td> <textarea name="frm_ClassMember" cols="40" datatype="Require" msg="必须填写"><%if Bol_IsEdit then response.Write(VClass_Rs(13)) end if%></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群讨论每页显示多少数量</td>
      <td> <input type="text" name="frm_PerPageNum" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(14)) else response.Write("20") end if%>" dataType="Range" msg="在1~30之间" min="-1" max="31" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群是否是管理员建立</td>
      <td> <input type="radio" name="frm_isSys"  value="1" <%if Bol_IsEdit then if VClass_Rs(15)=1 then response.Write(" checked ") end if end if%>>
        是 
        <input type="radio" name="frm_isSys"  value="0" <%if Bol_IsEdit then if VClass_Rs(15)=0 then response.Write(" checked ") end if else response.Write(" checked ") end if%>>
        否 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">人气/点击数</td>
      <td> <input type="text" name="frm_hits" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(16)) else response.Write("0") end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  dataType="Range" msg="在1~10W之间" min="-1" max="100001"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 保存 " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
 			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="if(confirm('确定删除该项目吗？')) location='<%="GroupDebate_manage.asp?Act=Del&gdID="&request.QueryString("gdID")%>'">
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
        <!---联动菜单开始 onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');" --->
        <select name="vclass1" id="vclass1" style="width:100px">
          <option></option>
        </select>
<!--		 <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--- > -->
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
        <!---联动菜单开始 onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"  --->
        <select name="Hy_vclass1" id="select" style="width:100px">
          <option></option>
        </select>
<!--		<select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---联动菜单结束--- > -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">加入时间</td>
      <td> <input type="text" name="frm_AddTime" size="27" value="" readonly>
        <input name="SelectDate2" type="button" id="SelectDate2" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);"></td>
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
      <td> <input type="text" name="frm_AccessFile" size="40" value="" require="true" dataType="Url"   msg="非法的Url"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群创始人用户编号</td>
      <td> <input name="frm_UserNumber" type="text" value="" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群现在管理员用户编号</td>
      <td> <input name="frm_AdminName" type="text" value="" size="40" datatype="Require" msg="必须填写"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群的成员</td>
      <td> <textarea name="frm_ClassMember" cols="40" datatype="Require" msg="必须填写"></textarea>
        多个则用“,”分开 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">此群讨论每页显示多少数量</td>
      <td> <input type="text" name="frm_PerPageNum" size="40" value="" dataType="Range" msg="在1~30之间" min="-1" max="31" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">社群是否是管理员建立</td>
      <td> <input type="radio" name="frm_isSys"  value="1" >
        是 
        <input type="radio" name="frm_isSys"  value="0" >
        否 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">人气/点击数</td>
      <td> <input type="text" name="frm_hits" size="40"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"> </td>
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

</BODY>
<%if instr(",Add,Edit,Search,",","&request.QueryString("Act")&",")>0 then%>

<script language="javascript">
<!-- 
//awen created
//联动菜单---行业类别   最多4级   --start 
//数据格式 ID，父级ID，名称
var Hy_array=new Array();
<%dim sql,rs,i
  sql="select VCID,ParentID,vClassName from FS_ME_VocationClass"
  set rs=User_Conn.execute(sql)
  i=0
  do while not rs.eof
%>
Hy_array[<%=i%>]=new Array("<%=rs("VCID")%>","<%=rs("ParentID")%>","<%=rs("vClassName")%>"); 
<%
	rs.movenext
	i=i+1
loop
rs.close
%>

var Hy_liandong=new CLASS_LIANDONG_YAO(Hy_array)
Hy_liandong.firstSelectChange("0","Hy_vclass1");
/*
Hy_liandong.subSelectChange("Hy_vclass1","Hy_vclass2");
Hy_liandong.subSelectChange("Hy_vclass2","Hy_vclass3");
Hy_liandong.subSelectChange("Hy_vclass3","Hy_vclass4");
*/
Hy_liandong.close
//end 
//---------------------------------------------
//联动菜单---所属的社群分类     --start 
var array=new Array();
<%
  sql="select VCID,ParentID,vClassName from FS_ME_GroupDebateClass"
  set rs=User_Conn.execute(sql)
  i=0
  do while not rs.eof
%>
array[<%=i%>]=new Array("<%=rs("VCID")%>","<%=rs("ParentID")%>","<%=rs("vClassName")%>"); 
<%
	rs.movenext
	i=i+1
loop
rs.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
/*
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");
*/
liandong.close
document.getElementById('vclass1').options.remove(1);
document.getElementById('Hy_vclass1').options.remove(1);

/*
//---------------------------清除关联下拉框的内容
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
*/
//end 
-->
</script>
<%end if%>

<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
</HTML>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 






