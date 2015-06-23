<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<% 
Dim Conn,User_Conn,VClass_Rs,VClass_Sql
Dim str_Url_Add,CheckStr
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
'******************************************************************
Sub DelClass()
	Dim DelID,Str_Tmp,Str_Tmp1
	DelID = FormatIntArr(request.Form("DelID"))
	if DelID = "" then 
		response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>你必须选择一项再删除。</li>")
		response.End()
	end if
	DelID = replace(DelID," ","")
	Str_Tmp1 = DelID
	''**********************
	do while Str_Tmp1 <> "" 
		Str_Tmp1 = Get_VCID_VCID_TO_Del(Str_Tmp1)
		Str_Tmp = Str_Tmp & Str_Tmp1 
	loop		
	if right(Str_Tmp,1) = "," then 
		Str_Tmp	= Str_Tmp & DelID
	elseif Str_Tmp<>"" then 
		Str_Tmp	= Str_Tmp &","& DelID	
	else
		Str_Tmp = DelID	
	end if
	''**********************
	'response.Write(Str_Tmp)
	''+++++++++++++++++
	''删除相关的其它表内容.
	User_Conn.execute("delete from FS_ME_GroupDebateManage where ClassID in ("&FormatIntArr(Str_Tmp)&")")
	''+++++++++++++++++
	User_Conn.execute("delete from FS_ME_GroupDebateClass where VCID in ("&FormatIntArr(Str_Tmp)&")")
	response.Redirect("../Success.asp?ErrorUrl=User/GroupDebate_Class.asp&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub

Sub SaveClass()
	call Say_Limit_Class("SaveMode") 
	Dim Edit_PartID,ii,Str_Req_Child,Str_Req_Tmp,Str_Biangeng_ParentID_ParentID,Err_info
	VClass_Sql = "select VCID,vClassName,vClassName_En,ParentID from FS_ME_GroupDebateClass"
	if NoSqlHack(request.Form("ParID"))<>"" then 
	''新增
		if CheckCardCF(NoSqlHack(Trim(request.Form("vClassName"))),request.Form("ParID"))<>"" then 
			response.Redirect("../Error.asp?ErrCodes=<li>同级别的类别: "&request.Form("vClassName")&" 已经存在。</li>")
			response.End()
		end if	
		VClass_Sql = VClass_Sql & " where VCID=0"
	elseif 	NoSqlHack(request.Form("VCID"))<>"" then 
	''修改
		VClass_Sql = VClass_Sql & " where VCID=" & NoSqlHack(request.Form("VCID"))
		''**********************
		''当行业类别要变更时判断当前级别加上其子级再加上要变更的类别一共是否在4级以内。若不再这返回。
		for ii = 4 to 1 step -1
			if request.Form("vclass"&ii)<>"" then Edit_PartID = request.Form("vclass"&ii) : exit for
		next
		''''''''''''''''''''''''''
		if Edit_PartID = "[ChangeToTop]" then Edit_PartID = 0
		if Edit_PartID > 0 then
		''''''''''''''''
		''Str_Biangeng_ParentID_ParentID准备变更到的父级类别和其本身的父级类别
		Str_Req_Tmp = Edit_PartID		
		do while Str_Req_Tmp<>""
			Str_Req_Tmp = Get_ParID_ParID_TO_Save(Str_Req_Tmp)
			Str_Biangeng_ParentID_ParentID = Str_Biangeng_ParentID_ParentID & Str_Req_Tmp 
		loop		
		if right(Str_Biangeng_ParentID_ParentID,1) = "," then 
			Str_Biangeng_ParentID_ParentID	= Str_Biangeng_ParentID_ParentID & Edit_PartID
		elseif Str_Biangeng_ParentID_ParentID<>"" then 
			Str_Biangeng_ParentID_ParentID	= Str_Biangeng_ParentID_ParentID &","& Edit_PartID	
		else
			Str_Biangeng_ParentID_ParentID = Edit_PartID	
		end if
		Err_info = Err_info & "<li>欲变更到的父级和其所有父级:"&Str_Biangeng_ParentID_ParentID&"</li>"
		response.Write(Err_info)
		'response.End()
		''''''''''''''''''''''''''
		''Str_Req_Child当前的类别和其子内		
		Str_Req_Tmp = request.Form("VCID")		
		do while Str_Req_Tmp <> "" 
			Str_Req_Tmp = Get_VCID_VCID_TO_Del(Str_Req_Tmp)
			Str_Req_Child = Str_Req_Child & Str_Req_Tmp 
		loop		
		if right(Str_Req_Child,1) = "," then 
			Str_Req_Child	= Str_Req_Child & request.Form("VCID")
		elseif Str_Req_Child<>"" then 
			Str_Req_Child	= Str_Req_Child &","& request.Form("VCID")	
		else
			Str_Req_Child = request.Form("VCID")	
		end if
		Err_info = Err_info & "<li>当前类别和其所有子类:"&Str_Req_Child&"</li>"
		Err_info = Err_info & "<li>预料的总级数:"&ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1&"</li>"
		response.Write(Err_info)
		'response.End()
		if ubound(split(Str_Biangeng_ParentID_ParentID & "," & Str_Req_Child,",")) + 1 > 4 then
			Err_info = Err_info & "<li>抱歉,将要或已经超过四级!所有不能更改到该类下.</li>"
			response.Redirect("../error.asp?ErrCodes="&Err_info&"")
			response.End()
		end if
		'''''''''''''''
		end if
		''**********************
	else
		response.Redirect("../error.asp?ErrCodes=<li>必要的行业ID没有提供。</li>")	
		response.End()
	end if
	'response.Write(VClass_Sql)
	'response.End()
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if NoSqlHack(request.Form("ParID"))<>"" then 
		VClass_Rs.AddNew
		VClass_Rs("ParentID") = NoSqlHack(request.Form("ParID"))
		VClass_Rs("vClassName") = NoSqlHack(Trim(request.Form("vClassName")))
		VClass_Rs("vClassName_En") = NoSqlHack(Trim(request.Form("vClassName_En")))
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/GroupDebate_Class.asp?Act=Add&VCID="&NoSqlHack(request.Form("ParID"))&"&VCText="&NoSqlHack(request.Form("VCText")) )&"&ErrCodes=<li>保存成功。</li>")
	end if
	'''修改
	if NoSqlHack(request.Form("VCID"))<>"" then 
		'response.Write("<br>VClass_Sql:"&VClass_Sql&"<br>Edit_PartID:"&Edit_PartID)
		'response.End()
		VClass_Rs("vClassName") = NoSqlHack(Trim(request.Form("vClassName")))
		VClass_Rs("vClassName_En") = NoSqlHack(Trim(request.Form("vClassName_En")))
		if Edit_PartID<>"" then VClass_Rs("ParentID") = Edit_PartID
		VClass_Rs.update

		if Edit_PartID<>"" then 
			Dim PartID_PartID_Rs
			''取得变更后的父级ID的父级ID，并传给返回的VCID，以便显示父级ID同级别的所有类别
			set PartID_PartID_Rs = User_Conn.execute("select ParentID from FS_ME_GroupDebateClass where VCID="&Edit_PartID)
			if not PartID_PartID_Rs.eof then  Edit_PartID = PartID_PartID_Rs(0)
			PartID_PartID_Rs.close
			set PartID_PartID_Rs = nothing
		else
			Edit_PartID = VClass_Rs("ParentID")
		end if		
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/GroupDebate_Class.asp?Act=View&VCID="&Edit_PartID )&"&ErrCodes=<li>保存成功。</li>")
	end if
End Sub

Function CheckCardCF(vClassName,ParentID)
''检查录入的是否重复,重复则返回vClassName,不重复则返回""
	Dim CheckCardCF_Rs
	set CheckCardCF_Rs = User_Conn.execute( "select Count(*) from FS_ME_GroupDebateClass where vClassName='"&vClassName&"' and ParentID="&ParentID )
	if  CheckCardCF_Rs(0)>0 then 
		CheckCardCF = vClassName
	else 
		CheckCardCF = ""
	end if
	CheckCardCF_Rs.close	
End Function

Sub Say_Limit_Class(SownMode)
Dim Arr_Tmp,str_Session
str_Session = session("TopMenu_DaoHang_VCID_List")
select case SownMode
case "AddMode"
	''============================
	''行业类别最多四级的判断.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>2 then 
		''按扭
			response.Write(" disabled title=""行业分类最多四级!当前等级:第"&cstr(ubound(Arr_Tmp) + 1)&"级."" ")	
		end if
	end if
	''============================
case "SaveMode"
	''============================
	''行业类别最多四级的判断.
	if str_Session <> "" then 
		if left(str_Session,1)="," then str_Session = mid(str_Session,2,len(str_Session))
		if right(str_Session,1)="," then str_Session = mid(str_Session,1,len(str_Session) - 1)
		Arr_Tmp = split(str_Session,",")
		if ubound(Arr_Tmp)>2 then 
			response.Redirect("../error.asp?ErrCodes=<li>行业类别不能超过四级。</li>")	
			response.End()
		else
			response.Write("行业分类最多四级!当前等级:第"&cstr(ubound(Arr_Tmp) + 1)&"级.")	
		end if
	end if
	''============================
case else

end select 	
End Sub

Function Get_VClass(vcid)
''递归调用显示分类
	Dim Get_Html
	VClass_Sql = "select VCID,vClassName,vClassName_En,ParentID from FS_ME_GroupDebateClass"
	if vcid<>"" and vcid>0 then 
		VClass_Sql = VClass_Sql &" where ParentID = "&vcid
	else
		VClass_Sql = VClass_Sql &" where ParentID = 0"	
	end if
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1	
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
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("vClassName") & "</td>" & vbcrlf
'		Get_Html = Get_Html & "<td align=""center""><a href=GroupDebate_Class.asp?Act=View&VCID="&VClass_Rs("VCID")&">"& VClass_Rs("vClassName") & "</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("vClassName_En") & "</td>" & vbcrlf
'		Get_Html = Get_Html & "<td align=""center""><a href=""GroupDebate_Class.asp?Act=Add&VCID="&VClass_Rs("VCID")&"&VCText="&VClass_Rs("vClassName")&"""  class=""otherset"">新增</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""GroupDebate_Class.asp?Act=Edit&VCID="&VClass_Rs("VCID")&"&VCText="&VClass_Rs("vClassName")&"""  class=""otherset"">设置</FONT></a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""DelID"" id=""DelID"" value="""&VClass_Rs("VCID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""button"" name=""button1112"" value="" 新增同级栏目 "" onClick=""location='GroupDebate_Class.asp?Act=Add"& str_Url_Add &"'"">" & vbcrlf
	Get_Html = Get_Html & "<input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"& vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_VClass = Get_Html
End Function

Function Get_VCID_VCID_TO_Del(Req_VCID)
''循环调用，通过传入的VCID得到子级的VCID，并且组合起来，一起传给DelClass过程以删除所有相关类别。
Dim Str_Tmp,This_Fun_Sql
if Req_VCID="" or isnull(Req_VCID) or Req_VCID="," then Get_VCID_VCID_TO_Del="" : exit Function
This_Fun_Sql="select VCID from FS_ME_GroupDebateClass where ParentID in ("&FormatIntArr(Req_VCID)&")"
set VClass_Rs=User_Conn.execute(This_Fun_Sql)
do while not VClass_Rs.eof 
	Str_Tmp = Str_Tmp & VClass_Rs(0) & ","	
	VClass_Rs.movenext
loop	
VClass_Rs.close
Get_VCID_VCID_TO_Del = Str_Tmp
End Function

Function Get_ParID_ParID_TO_Save(Req_ParID)
''循环调用，通过传入的ParentID得到其上的所有ParentID，并且组合起来，一起传给SaveClass过程以判断是否超过4级.
Dim Str_Tmp,This_Fun_Sql
if Req_ParID="" or isnull(Req_ParID) or Req_ParID="," then Get_ParID_ParID_TO_Save="" : exit Function
This_Fun_Sql="select ParentID from FS_ME_GroupDebateClass where VCID in ("&FormatIntArr(Req_ParID)&")"
set VClass_Rs=User_Conn.execute(This_Fun_Sql)
do while not VClass_Rs.eof 
	if VClass_Rs(0)=0 then exit do : Get_ParID_ParID_TO_Save="" : exit Function 
	Str_Tmp = Str_Tmp & VClass_Rs(0) & ","	
	VClass_Rs.movenext
loop	
VClass_Rs.close
Get_ParID_ParID_TO_Save = Str_Tmp
End Function

Function Get_PatID_TO_View(View_ID)
''导航作用，得到父级ID以便返回父级层查看
	Dim This_Fun_Sql
	if View_ID="" then Get_PatID_TO_View=0 : exit Function
	This_Fun_Sql = "select ParentID from FS_ME_GroupDebateClass where VCID="&NoSqlHack(View_ID)
	Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
	if not VClass_Rs.eof then 
		Get_PatID_TO_View = VClass_Rs(0)
	else
		Get_PatID_TO_View = 0	
	end if
	VClass_Rs.close
End Function

Function Get_PatTxt_TO_View(View_ID)
''导航作用，得到父级类目名称以便返回父级层查看，和上面对应
Dim VClass_Rs1,This_Fun_Sql
	if View_ID="" then Get_PatTxt_TO_View="无" : exit Function
	This_Fun_Sql = "select ParentID from FS_ME_GroupDebateClass where VCID="&NoSqlHack(View_ID)
	Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
	if not VClass_Rs.eof then 
		set VClass_Rs1 = User_Conn.execute( "select vClassName from FS_ME_GroupDebateClass where VCID="&VClass_Rs(0) )
		if not VClass_Rs1.eof then 
			Get_PatTxt_TO_View = VClass_Rs1(0)
		else
			Get_PatTxt_TO_View = "无"	
		end if
		VClass_Rs1.close
		set VClass_Rs1=nothing		
	else
		Get_PatTxt_TO_View = "无"	
	end if
	VClass_Rs.close
End Function

Function TopMenu_DaoHang()
''层层导航作用 session("TopMenu_DaoHang_VCID_List") = ',1,2,3,4,'
Dim str_Req_VCID,str_Session,Sql_Session,Per_ID_Session,Str_Tmp,Str_Remove,Arr_Tmp,This_Fun_Sql
	if request.QueryString("VCID")="" or request.QueryString("VCID")="0" then 
		session("TopMenu_DaoHang_VCID_List")="" : TopMenu_DaoHang = "" : exit Function
	end if
	''**************
	str_Req_VCID = NoSqlHack(request.QueryString("VCID"))
	str_Req_VCID = ","&str_Req_VCID&","
	'if session("TopMenu_DaoHang_VCID_List") = "" then session("TopMenu_DaoHang_VCID_List") = str_Req_VCID
	str_Session = session("TopMenu_DaoHang_VCID_List")
	if str_Session = "" or isEmpty(str_Session) then str_Session = str_Req_VCID 
	if left(str_Session,1)<>"," then str_Session = ","&str_Session
	if right(str_Session,1)<>"," then str_Session = str_Session&","
	session("TopMenu_DaoHang_VCID_List") = str_Session
	'response.Write("原"&session("TopMenu_DaoHang_VCID_List"))
	',1,2,3,4,' 与 ',3,'
	if right(str_Session,len(str_Req_VCID)) <> str_Req_VCID then 
	if instr(str_Session,str_Req_VCID)>0 then
	''截取
		str_Session = mid(str_Session,1,instr(str_Session,str_Req_VCID) + len(str_Req_VCID) - 1)
		session("TopMenu_DaoHang_VCID_List") = str_Session
		'response.Write("减"&session("TopMenu_DaoHang_VCID_List"))
	else
	''添加
		str_Session = str_Session & NoSqlHack(request.QueryString("VCID")) &","
		session("TopMenu_DaoHang_VCID_List") = str_Session
		'response.Write("加"&session("TopMenu_DaoHang_VCID_List"))
	end if
	end if
	''**************
	Sql_Session = str_Session
	if left(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,2,len(Sql_Session))
	if right(Sql_Session,1)="," then Sql_Session = mid(Sql_Session,1,len(Sql_Session) - 1)
	'response.Write(session("TopMenu_DaoHang_VCID_List"))
	''**************
	''顺序真实的组合
	str_Session = ""
	Arr_Tmp = split(Sql_Session,",")	
	for each Per_ID_Session in Arr_Tmp
		This_Fun_Sql = "select VCID,vClassName from FS_ME_GroupDebateClass where VCID = "&Per_ID_Session
		Set VClass_Rs = User_Conn.execute(This_Fun_Sql)
		if not VClass_Rs.eof then 
			Str_Tmp = Str_Tmp & "<a href=""GroupDebate_Class.asp?Act=View&VCID=" &VClass_Rs(0)&""">"&VClass_Rs(1)&"<a> >> "
			str_Session = str_Session &VClass_Rs(0)& "," 
		end if
		VClass_Rs.close
	next
	session("TopMenu_DaoHang_VCID_List") = "," & str_Session	
	if right(Str_Tmp,len(" >> "))=" >> " then Str_Tmp = mid(Str_Tmp,1,len(Str_Tmp) - len(" >> "))
	TopMenu_DaoHang = Str_Tmp
	'response.Write(session("TopMenu_DaoHang_VCID_List"))
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
-->
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<%if request.QueryString("Act")="Edit" then%><script language="javascript" src="../../FS_Inc/class_liandong.js" type="text/javascript"></script><%end if%>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<%
'******************************************************************
select case request.QueryString("Act")
	case "View",""
		ViewClass
	case "Add"
		AddClass
	case "Del"
		DelClass
	case "Edit"
		EditClass
	case "Save"	
		SaveClass
end select
'******************************************************************

Sub ViewClass()
Dim View_VCID,IsOk,VClass_Rs1
IsOk = false
View_VCID = NoSqlHack(request.QueryString("VCID"))
if View_VCID<>"" then if isnumeric(View_VCID) then if View_VCID>0 then IsOk = true
if IsOk=false then 
	str_Url_Add = "&VCID=0&VCText=无"	
else	
	set VClass_Rs1 = User_Conn.execute( "select vClassName from FS_ME_GroupDebateClass where VCID="&View_VCID )
	if not VClass_Rs1.eof then 
		str_Url_Add = "&VCID="&View_VCID&"&VCText="	& VClass_Rs1(0)
	else
		str_Url_Add = "&VCID=0&VCText=无"	
	end if
	VClass_Rs1.close
	set VClass_Rs1=nothing		
end if		
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Del">
  <tr  class="hback"> 
    <td class="xingmu"  colspan="5">社群分类管理</td>
  </tr>
    <tr  class="hback"> 
      <td colspan="5"><a href="GroupDebate_Class.asp">管理首页</a> | <a href="GroupDebate_manage.asp">社群总管理</a> 
        <!-- | 分类导航：<a href="GroupDebate_Class.asp">首页</a> >> < %=TopMenu_DaoHang()%> -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="center" class="xingmu" >分类名称</td>
      <td width="35%" align="center" class="xingmu">英文名称</td>
<!--	  <td width="10%" align="center" class="xingmu">新增子栏目</td> -->
      <td width="20%" align="center" class="xingmu">设置</td>
      <td width="5%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_VClass(request.QueryString("VCID")) )
	%>
  </form>
</table>
<%End Sub
Sub AddClass()
dim View_Par_ID
if request.QueryString("VCID")<>"" then 
	View_Par_ID = request.QueryString("VCID")
else
	View_Par_ID = request.QueryString("ParID")
end if		
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td align="left" class="xingmu" colspan="5"><a href="GroupDebate_Class.asp">管理首页</a> 
        | <a href="GroupDebate_manage.asp">社群总管理</a> 
        <!-- | 分类导航：<a href="GroupDebate_Class.asp">首页</a> >> < %=TopMenu_DaoHang()%> -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >新增社群分类信息 目前最多一级</td>
    </tr>
    <tr  class="hback"> 
      <td width="200" align="right">上一级行业名称</td>
      <td width="774">
	  <input name="VCText" id="VCText" readonly type="text" size="50" value="<%=request.QueryString("VCText")%>"/>
	  <input type="hidden" name="ParID" id="ParID" value="<%=View_Par_ID%>"/>
	  </td>
    </tr>
    <tr class="hback"> 
      <td align="right">新的社群名称：</td>
      <td width="774"><input name="vClassName" type="text" id="vClassName" size="50" onBlur="if(this.value==''){Chk_vCName.innerText='该项必须填写';this.focus();} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''"><span id="Chk_vCName" class="tx"></span> 
      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">对应英文名：</td>
      <td width="774"><input name="vClassName_En" type="text" id="vClassName_En" size="50" onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
  		<tr>
		  <td align="center">
			<input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass" id="SaveClass" value=" 保存 " onClick="if(vClassName.value==''){alert('必要的参数必须填写');vClassName.focus();return false}" /> &nbsp;
			<input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
		  </td>
		  <td width="50" align="right">
			<a href="GroupDebate_Class.asp?Act=View&VCID=<%=Get_PatID_TO_View(View_Par_ID)%>">返回父级</a>
		  </td>
		  <td width="50" align="right">
			<a href="#" onClick="javascript:history.back();">返回</a>
		  </td>
		</tr>  
      </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub
Sub EditClass()
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="VCForm" id="VCForm" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td align="left" class="xingmu" colspan="5"><a href="GroupDebate_Class.asp">管理首页</a> 
        | <a href="GroupDebate_manage.asp">社群总管理</a> 
        <!-- | 分类导航：<a href="GroupDebate_Class.asp">首页</a> >> < %=TopMenu_DaoHang()%> -->
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="2" align="left" class="xingmu" >修改社群分类信息 目前最多一级</td>     
    </tr>
	<tr  class="hback" style="display:none"> 
      <td width="201" align="right">变更上级类别<br><span class="tx">若不变更请重置</span></td>
      <td width="773">
<!---联动菜单开始--->	
	<SELECT NAME="vclass1" ID="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass2" ID="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass3" ID="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
        <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass4" ID="vclass4" style="width:100px">
    	<OPTION></OPTION>
    </SELECT>
<!---联动菜单结束--->		
      </td>
    </tr>
    <tr  class="hback"> 
      <td width="100" align="right">社群名称</td>
      <td>
	  <input name="vClassName" id="vClassName" type="text" size="50" value="<%=request.QueryString("VCText")%>" onBlur="if(this.value==''){Chk_vCName.innerText='该项必须填写';this.focus()} else if(Chk_vCName.innerText!='')Chk_vCName.innerText=''"><span id="Chk_vCName" class="tx"></span>
      <input type="hidden" name="VCID" id="VCID" value="<%=request.QueryString("VCID")%>">
  	  <!--<input type="text" name="ParentID" id="ParentID" value="">-->
	  </td>
    </tr>
    <tr width="100" class="hback"> 
      <td align="right">对应英文名：</td>
      <td><input name="vClassName_En" type="text" id="vClassName_En" size="50" onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
  		<tr>
		  <td align="center">
			<input type="submit" <%call Say_Limit_Class("AddMode")%> name="SaveClass" id="SaveClass" value=" 保存 " onClick="if(vClassName.value==''){vClassName.focus();return false}" /> &nbsp;
			<input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
		  </td>
		  <td width="50" align="right">
			<a href="GroupDebate_Class.asp?Act=View&VCID=<%=Get_PatID_TO_View(request.QueryString("VCID"))%>">返回父级</a>
		  </td>
		  <td width="50" align="right">
			<a href="#" onClick="javascript:history.back();">返回</a>
		  </td>
		</tr>  
      </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub%>
</body>
<%if request.QueryString("Act")="Edit" then%>
<script language="javascript">
<!-- 
//awen created
//联动菜单---行业类别   最多4级   --start 
//数据格式 ID，父级ID，名称
var array=new Array();
<%dim sql,rs,i
  sql="select VCID,ParentID,vClassName from FS_ME_GroupDebateClass where VCID<>"&request.QueryString("VCID")
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
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");

//---------------------------清除关联下拉框的内容
function RemoveChildopt(obj,StrList)
{
/*
	if (StrList=='') 
	if (document.getElementById('ParentID')!=null) 
	{
		if (obj.value!='')
		{document.getElementById('ParentID').value=obj.value;return;}
		else
		{document.getElementById('ParentID').value=document.getElementById('vclass3').value;return;}
	}
*/	
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
/*	else
	//选择实际项目将值副给隐藏域，以便修改的时候存盘。
	{
		var Tmpstr = '';
		for (var i=TmpArr.length-1 ; i>=0; i--)
		{			
			if (TmpArr[i]!='')
			{	if (document.getElementById(TmpArr[i]).selectedIndex>1)
				Tmpstr += document.getElementById(TmpArr[i]).options[document.getElementById(TmpArr[i]).selectedIndex].value;
			}
		}	
		//alert(document.VCForm.ParentID.value);	
		if (Tmpstr=='')
		{
			//副给隐藏域				
			if (document.getElementById('ParentID')!=null) document.all.ParentID.value=obj.value;
		}
	}
*/		
} 
//end 
-->
</script>
<%end if%>
<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>

</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 






