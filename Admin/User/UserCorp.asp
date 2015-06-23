<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<% 
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_List") then Err_Show
if not MF_Check_Pop_TF("ME001") then Err_Show 

int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导c航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页


''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			Get_OtherTable_Value = Get_OtherTable_Value & This_Fun_Rs(0) &"&nbsp;"
			This_Fun_Rs.movenext
		loop
	else			
		if not This_Fun_Rs.eof then 
			Get_OtherTable_Value = This_Fun_Rs(0)
		else
			Get_OtherTable_Value = ""
		end if
	end if	
	set This_Fun_Rs=nothing 
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

Function Get_WhileData(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "A.UserNumber,UserName,C_Name,C_Tel,Email,C_Website,Integral,FS_Money,RegTime,isLock,isLockCorp"
	Str_Tmp = Str_Tmp & ",C_Name,C_ShortName,C_Province,C_City,C_Address,C_PostCode,C_ConactName,C_Tel,C_Fax,C_VocationClassID,C_Website,C_size,C_Capital,C_BankName,C_BankUserName"
	This_Fun_Sql = "select "&Str_Tmp&"  from FS_ME_Users A,FS_ME_CorpUser B where  A.UserNumber=B.UserNumber and A.IsCorporation=1"
	if Add_Sql<>"" then 
		if instr(Add_Sql,"order by")>0 then 
			This_Fun_Sql = This_Fun_Sql &"  "& replace(Add_Sql,"UserNumber","A.UserNumber")
		else
			This_Fun_Sql = and_where(This_Fun_Sql) & Add_Sql
		end if		
	end if
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Trim(request.Form("frm_"&Str_Tmp))<>"" then 
				Req_Str = NoSqlHack(Trim(request.Form("frm_"&Str_Tmp)))
				select case Str_Tmp
					case "Integral","FS_Money","RegTime","Certificate","C_VocationClassID","isLock"
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
		'response.Write(This_Fun_Sql)
		'response.End()
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	Set GetUserDataObj_Rs = CreateObject(G_FS_RS)
	GetUserDataObj_Rs.Open This_Fun_Sql,User_Conn,1,1
	IF not GetUserDataObj_Rs.eof THEN
	
	GetUserDataObj_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>GetUserDataObj_Rs.PageCount Then cPageNo=GetUserDataObj_Rs.PageCount 
	GetUserDataObj_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""#"" onclick=""javascript:if(TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display=='') TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display='none'; else TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display='';"" class=""otherset"" title='点击查看更多信息'>"&GetUserDataObj_Rs("UserNumber")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""UserCorp.asp?Act=Edit&UserNumber="&GetUserDataObj_Rs("UserNumber")&""" class=""otherset"" title='点击修改公司信息'>"&GetUserDataObj_Rs("C_Name")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""UserCorp.asp?Act=Edit&UserNumber="&GetUserDataObj_Rs("UserNumber")&""" class=""otherset"" title='点击修改公司信息'>"&GetUserDataObj_Rs("UserName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""mailto:"&GetUserDataObj_Rs("Email")&""" title=""发邮件给他"">"& GetUserDataObj_Rs("Email") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("RegTime") & "</td>" & vbcrlf
		if GetUserDataObj_Rs("isLockCorp")&"" <> "0" then 
			''锁定,需要解锁
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""锁 定"" onclick=""javascript:location='UserCorp.asp?Act=OtherEdit&EditSql="&server.URLEncode( "isLock=0" )&"&UserNumber="&GetUserDataObj_Rs("UserNumber")&"'"" title=""点击解锁"" style=""color:red""></td>" & vbcrlf
		else
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""正 常"" onclick=""javascript:location='UserCorp.asp?Act=OtherEdit&EditSql="&server.URLEncode( "isLock=1" )&"&UserNumber="&GetUserDataObj_Rs("UserNumber")&"'"" title=""点击锁定""></td>" & vbcrlf
		end if
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""frm_UserNumber"" id=""frm_UserNumber"" value="""&GetUserDataObj_Rs("UserNumber")&""" /><input type=""hidden"" name=""frm_UserName"" id=""frm_UserName"" value="""&GetUserDataObj_Rs("UserName")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++点开用户编号时显示详细信息。
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& GetUserDataObj_Rs("UserNumber") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td>" & GetUserDataObj_Rs("C_Province") &" | "& GetUserDataObj_Rs("C_City") &" | "& GetUserDataObj_Rs("C_Address") &" | " & GetUserDataObj_Rs("C_size") &" | "& GetUserDataObj_Rs("C_Capital") & "</td>" & vbcrlf
		Get_Html = Get_Html & "</tr></table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		Str_Tmp = ""
		GetUserDataObj_Rs.MoveNext
 		if GetUserDataObj_Rs.eof or GetUserDataObj_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(GetUserDataObj_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	GetUserDataObj_Rs.close
	Get_WhileData = Get_Html
End Function
Sub OtherEdit()
	Dim islock,islockCorp
	islock = request.QueryString("EditSql")
	islockCorp = replace(islock,"isLock","isLockCorp")
	User_Conn.execute("Update FS_ME_Users set "&islock&" where UserNumber='"&NoSqlHack(request.QueryString("UserNumber"))&"'")	
	User_Conn.execute("Update FS_ME_CorpUser set "&islockCorp&" where UserNumber='"&NoSqlHack(request.QueryString("UserNumber"))&"'")	
	response.Redirect("UserCorp.asp?Act=View")
End Sub
''================================================================
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
var Old_Sql = document.URL;

function OrderByName(FildName)
{
	//alert(document.URL);	
	var New_Sql;
	if(Old_Sql.indexOf('Add_Sql')<0)
	{
		if(Old_Sql.indexOf('?')<0)
			New_Sql = Old_Sql + "?Add_Sql=order by " + FildName;	
		else
			New_Sql = Old_Sql + "&Add_Sql=order by " + FildName;	
	}
	else
	{
		if(Old_Sql.indexOf("Add_Sql=order by " + FildName + " desc")>-1)
		{
			New_Sql = Old_Sql.substring(0,Old_Sql.indexOf("Add_Sql=")) + "Add_Sql=order by " + FildName;
		}
		else
		{
			New_Sql = Old_Sql.substring(0,Old_Sql.indexOf("Add_Sql=")) + "Add_Sql=order by " + FildName + " desc";	
		}	
	}
	//alert(New_Sql);	
	location = New_Sql;
}
-->
</script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr  class="hback">
		<td class="xingmu" colspan=20>企业用户信息管理</td>
	</tr>
	<tr  class="hback">
		<td><a href="UserCorp.asp?Act=View">管理首页</a>
			<!-- | <a href="UserCorp.asp?Act=Add"><b>新增</b></a>-->
			| <a href="UserCorp.asp?Act=Search">查询</a></td>
	</tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
	View
	case "Edit","Search","Add_BaseData"
	Add_Edit_Search
	case "Add"
	response.Write("后台添加公司信息已屏蔽。若需开通请与我们联系。")
	response.End()
	case "OtherEdit"
	OtherEdit
	case else
	response.Write(request.QueryString("Act")&"参数传递错误！")
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="UserCorp_DataAction.asp?Act=Del">
		<tr  class="hback">
			<td align="center" class="xingmu" ><a href="javascript:OrderByName('UserNumber')" class="sd"><b>〖用户编号〗</b></a> <span id="Show_Oder_UserNumber"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('C_Name')" class="sd"><b>企业名称</b></a> <span id="Show_Oder_C_Name"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('UserName')" class="sd"><b>用户名</b></a> <span id="Show_Oder_UserName"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('Email')" class="sd"><b>Email</b></a> <span id="Show_Oder_Email"></span></td>
			<td width="10%" align="center" class="xingmu"><a href="javascript:OrderByName('RegTime')" class="sd"><b>注册日期</b></a> <span id="Show_Oder_RegTime"></span></td>
			<td width="10%" align="center" class="xingmu"><a href="javascript:OrderByName('isLock')" class="sd"><b>是否锁定</b></a> <span id="Show_Oder_isLock"></span></td>
			<td width="2%" align="center" class="xingmu">
				<input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" />
			</td>
		</tr>
		<%
		response.Write( Get_WhileData( request.QueryString("Add_Sql") ) )
	%>
	</form>
</table>
<%End Sub

Sub Add_Edit_Search()
''添加删除查询共用。
Dim UserNumber,Bol_IsEdit
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	UserNumber = NoSqlHack(Trim(request.QueryString("UserNumber")))
	if UserNumber="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的UserName没有提供</li>") : response.End()
	UserSql = "select UserID,UserName,UserPassword,HeadPic,HeadPicSize,PassQuestion,PassAnswer,safeCode,tel,Mobile,isMessage,Email," _
			  &"HomePage,QQ,MSN,Corner,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Certificate,CertificateCode,IsCorporation,PopList," _
			  &"PopList,Integral,FS_Money,RegTime,CloseTime,TempLastLoginTime,TempLastLoginTime_1,IsMarray,SelfIntro,isOpen,GroupID,LastLoginIP," _
			  &"ConNumber,ConNumberNews,isLock,UserFavor,MySkin,UserLoginCode,OnlyLogin,hits" _

			  &",C_Name,C_ShortName,C_Province,C_City,C_Address,C_PostCode,C_ConactName,C_Tel,C_Fax,C_VocationClassID,C_Website,C_size,C_Capital,C_BankName,C_BankUserName" _
			  &" from FS_ME_Users A,FS_ME_CorpUser B where  A.UserNumber=B.UserNumber and A.UserNumber= '"& NoSqlHack(UserNumber) &"'"
	Set GetUserDataObj_Rs	= CreateObject(G_FS_RS)
	GetUserDataObj_Rs.Open UserSql,User_Conn,1,1
	if GetUserDataObj_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
	Bol_IsEdit = True
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="hback">
		<td width="140" class="xingmu" colspan="3">
			<%if request.QueryString("Act")<>"Search" then response.Write("会员系统参数设置") else response.Write("查询会员") end if %>
		</td>
	</tr>
	<tr class="hback">
		<td width="33%"  id="Lab_Base">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>第一步：企业注册信息</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">企业注册信息查询模式</a>") 
	else
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">企业注册信息设置</a>") 
	end if%>
			</div>
		</td>
		<td width="33%" height="19" class="xingmu" id="Lab_Other">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>第二步：企业扩展信息</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">企业扩展信息查询模式</a>") 
	else 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">企业扩展信息设置</a>") 
	end if%>
			</div>
		</td>
		<td height="19" class="xingmu" id="Lab_Three">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>第三步：公司相关信息</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(3)"">企业相关信息查询模式</a>") 
	else 
		response.Write("<a href=""#"" onClick=""showDataPanel(3)"">企业相关信息设置</a>") 
	end if%>
			</div>
		</td>
	</tr>
	<tr class="hback">
		<td align="right"  colspan="3">
			<!---基础数据开始-->
			<div id="Layer1" style="position:relative; z-index:1; left: 0px; top: 0px;">
				<table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
					<form name="UserForm" method="post" <%if request.QueryString("Act")="Add" then 
			   response.Write(" action=""?Act=Add_BaseData""  onsubmit=""return CheckForm(this);""") 
			  elseif request.QueryString("Act")="Search" then 
			  	response.Write(" action=""?Act=SearchGo""")
			  else
			   response.Write(" action=""UserCorp_DataAction.asp?Act=BaseData""  onsubmit=""return CheckForm(this);""") 
			  end if%>>
						<%if request.QueryString("Act")<>"Search" then%>
						<tr class="hback">
							<td height="20" colspan="3" class="xingmu">请填写您的基本资料<span class="tx">(以下项目为空的不修改则请留空)</span></td>
						</tr>
						<%end if%>
						<tr class="hback">
							<td width="15%" height="65">
								<div align="right">用户名</div>
							</td>
							<td width="29%">
								<input name="frm_UserName" type="text" id="frm_UserName" style="width:90%" value="<%if Bol_IsEdit then response.Write""&GetUserDataObj_Rs("UserName")&"" end if%>" <%if Bol_IsEdit then response.write "readonly"%> >
								<%if request.QueryString("Act")<>"Search" then%>
								<a href="javascript:CheckName('../../user/lib/CheckName.asp')">是否被占用</a>
								<%end if%>
							</td>
							<td width="56%">
								<%if request.QueryString("Act")<>"Search" then%>
								户名由a～z的英文字母(不区分大小写)、0～9的数字、点、减号或下划线及中文组成，长度为3～18个字符，只能以数字或字母开头和结尾,例如:coolls1980。
								<%else%>
								模糊查询
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16" colspan="3" class="xingmu">请填写安全设置：（安全设置用于验证帐号和找回密码）</td>
						</tr>
						<%If p_isValidate = 0 and request.QueryString("Act")<>"Search" then%>
						<tr class="hback">
							<td height="16">
								<div align="right">密码</div>
							</td>
							<td>
								<input name="frm_UserPassword" type="password" id="frm_UserPassword" style="width:90%" maxlength="50">
							</td>
							<td rowspan="2">密码长度为<%=p_LenPassworMin%>～<%=p_LenPassworMax%>位，区分字母大小写。登录密码可以由字母、数字、特殊字符组成。</td>
						</tr>
						<tr class="hback">
							<td height="24">
								<div align="right">确认密码</div>
							</td>
							<td>
								<input name="frm_cUserPassword" type="password" id="frm_cUserPassword" style="width:90%" maxlength="50">
							</td>
						</tr>
						<%End if
				if request.QueryString("Act") <> "Search" then %>
						<tr class="hback">
							<td height="16">
								<div align="right">密码提示问题</div>
							</td>
							<td>
								<input name="frm_PassQuestion" type="text" id="frm_PassQuestion" style="width:90%" maxlength="30">
							</td>
							<td rowspan="2">当您忘记密码时可由此找回密码。例如，问题是“我的哥哥是谁？”，答案为&quot;coolls8&quot;。问题长度不大于36个字符，一个汉字占两个字符。答案长度在6～30位之间，区分大小写。</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">密码答案</div>
							</td>
							<td>
								<input name="frm_PassAnswer" type="text" id="frm_PassAnswer" style="width:90%" maxlength="50">
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">安全码</div>
							</td>
							<td>
								<input name="frm_SafeCode" type="password" id="frm_SafeCode" style="width:90%" maxlength="30">
							</td>
							<td rowspan="2">全码是您找回密码的重要途径，安全码长度为6～20位，区分字母大小写，由字母、数字、特殊字符组成。<br>
								<Span class="tx">特别提醒：安全码一旦设定，将不可自行修改.</Span></td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">确认安全码</div>
							</td>
							<td>
								<input name="frm_cSafeCode" type="password" id="frm_cSafeCode" style="width:90%" maxlength="30">
							</td>
						</tr>
						<%end if%>
						<tr class="hback">
							<td height="16">
								<div align="right">电子邮件</div>
							</td>
							<td>
								<input name="frm_Email" type="text" id="frm_Email" style="width:90%" maxlength="100" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Email")) end if%>">
								<%if request.QueryString("Act")<>"Search" then%>
								<br>
								<a href="javascript:CheckEmail('../../user/lib/Checkemail.asp')">是否被占用</a>
								<%end if%>
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								您的注册电子邮件。<Span class="tx">注册成功后，将不能修改</span>
								<%end if%>
							</td>
						</tr>
						<!--会员类型。-->
						<input type="hidden" name="frm_UserNumber_Edit1" value="<%=UserNumber%>">
						<tr class="hback">
							<td height="39" colspan="3">
								<div align="center">
									<input type="submit" name="Submit" value="<%if request.QueryString("Act")="Search" then response.Write(" 执行查询 ") else response.Write(" 保存企业基本信息 ") end if%>" style="CURSOR:hand">
									<input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
								</div>
							</td>
						</tr>
					</form>
				</table>
			</div>
			<!---基础数据结束-->
			<div id="Layer2" style="position:relative; z-index:1; left: 0px; top: 0px; width: 889px; height: 942px;">
				<table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
					<form name="UserForm" method="post"<%if request.QueryString("Act")="Add_BaseData" then 
			   response.Write(" action=""UserCorp_DataAction.asp?Act=Add_OtherData""  onsubmit=""return CheckForm_Other(this);""") 
			  elseif request.QueryString("Act")="Search" then 
			  	response.Write(" action=""?Act=SearchGo""")
			  else
			   response.Write(" action=""UserCorp_DataAction.asp?Act=OtherData""  onsubmit=""return CheckForm_Other(this);""") 
			  end if%>>
						<tr class="hback">
							<td height="27">
								<div align="right"><span class="tx">*</span>昵称</div>
							</td>
							<td>
								<input name="frm_NickName" type="text" id="frm_NickName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("NickName")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								请填写您对外的昵称。可以为中文
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td width="15%" height="27">
								<div align="right">姓名</div>
							</td>
							<td width="29%">
								<input name="frm_RealName" type="text" id="frm_RealName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("RealName")) end if%>">
							</td>
							<td width="56%">
								<%if request.QueryString("Act")<>"Search" then%>
								请填写您的真实姓名。
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right"><span class="tx">*</span>性别</div>
							</td>
							<td>
								<input type="radio" name="frm_Sex" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=0 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								男
								<input type="radio" name="frm_Sex" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=1 then response.Write("checked") end if end if%>>
								女 </td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								请您选择性别。
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">生日</td>
							<td>
								<input type="text" name="frm_BothYear" style="width:60%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("BothYear")) end if%>" readonly>
								<input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_BothYear);">
							</td>
							<td>
								<%if request.QueryString("Act")="Search" then response.Write("支持简单比较运算符，*123，123*，123的模糊查询。") else response.Write("请填写您的真实生日，该项用于取回密码。") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">证件类别</div>
							</td>
							<td>
								<select name=frm_Certificate  id="frm_Certificate">
									<%if request.QueryString("Act")="Search" then response.Write("<option value="""">请选择</option>") end if%>
									<option value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=0 then response.Write("selected") end if else if request.QueryString("Act")<>"Search" then response.Write("selected") end if end if%>>身份证</option>
									<option value="2" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=2 then response.Write("selected") end if end if%>>学生证</option>
									<option value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=1 then response.Write("selected") end if end if%>>驾驶证</option>
									<option value="3" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=3 then response.Write("selected") end if end if%>>军人证</option>
									<option value="4" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=4 then response.Write("selected") end if end if%>>护照</option>
								</select>
							</td>
							<td rowspan="2">
								<%if request.QueryString("Act")="Search" then response.Write("支持简单比较运算符，*123，123*，123的模糊查询。") else response.Write("有效证件作为取回帐号的最后手段，用以核实帐号的合法身份，请您务必如实填写。<br> <span class=""tx"">特别提醒：有效证件一旦设定，不可更改</span>") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">证件号码</div>
							</td>
							<td>
								<input name="frm_CerTificateCode" type="text" id="frm_CerTificateCode" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CerTificateCode")) end if%>">
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">您现在所在的省份</td>
							<td>
								<input type="text" name="frm_Province" readonly="" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Province")) end if%>">
							</td>
							<td>
								<select name="select111" size=1 onChange="javascript:frm_Province.value=this.options[this.selectedIndex].value">
									<option value="">请选择</option>
									<option value="四川">四川</option>
									<option value="北京">北京</option>
									<option value="上海">上海</option>
									<option value="天津">天津</option>
									<option value="重庆">重庆</option>
									<option value="安徽">安徽</option>
									<option value="甘肃">甘肃</option>
									<option value="广东">广东</option>
									<option value="广西">广西</option>
									<option value="贵州">贵州</option>
									<option value="福建">福建</option>
									<option value="海南">海南</option>
									<option value="河北">河北</option>
									<option value="河南">河南</option>
									<option value="黑龙江">黑龙江</option>
									<option value="湖北">湖北</option>
									<option value="湖南">湖南</option>
									<option value="吉林">吉林</option>
									<option value="江苏">江苏</option>
									<option value="江西">江西</option>
									<option value="辽宁">辽宁</option>
									<option value="内蒙古">内蒙古</option>
									<option value="宁夏">宁夏</option>
									<option value="青海">青海</option>
									<option value="山东">山东</option>
									<option value="山西">山西</option>
									<option value="陕西">陕西</option>
									<option value="西藏">西藏</option>
									<option value="新疆">新疆</option>
									<option value="云南">云南</option>
									<option value="浙江">浙江</option>
									<option value="港澳台">港澳台</option>
									<option value="海外">海外</option>
									<option value="其它">其它</option>
								</select>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">城市</div>
							</td>
							<td height="16">
								<input name="frm_City" type="text" id="frm_City" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("City")) end if%>">
							</td>
							<td height="16">您现在所在的城市</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">联系地址</div>
							</td>
							<td height="16">
								<input name="frm_Address" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Address")) end if%>">
							</td>
							<td height="16">您的联系地址</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">邮政编码</div>
							</td>
							<td height="16">
								<input name="frm_PostCode" type="text"  size="6" maxlength="6" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("PostCode")) end if%>">
							</td>
							<td height="16">邮政编码</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">头像地址</div>
							</td>
							<td height="16">
								<input name="frm_HeadPic" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPic")) end if%>">
							</td>
							<td height="16">您的头像地址</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">头像尺寸</div>
							</td>
							<td height="16">
								<input name="frm_HeadPicSize" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPicSize")) end if%>">
							</td>
							<td height="16">格式：[宽,高]如60,60 80,80 120,140</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">私人电话</div>
							</td>
							<td height="16">
								<input name="frm_tel" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("tel")) end if%>">
							</td>
							<td height="16">您的电话</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">手机</div>
							</td>
							<td height="16">
								<input name="frm_Mobile" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Mobile")) end if%>">
							</td>
							<td height="16">您的手机</td>
						</tr>
						<tr class="hback" id="tr_isMessage">
							<td height="16">
								<div align="right">短信验证手机</div>
							</td>
							<td height="16">
								<input type="checkbox" name="frm_isMessage" value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("isMessage")=1 then response.Write(" checked") end if end if%>>
							</td>
							<td height="16">是否通过短信验证手机,并捆绑 如果选择是,需要通信网关</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">企业网站</div>
							</td>
							<td height="16">
								<input name="frm_HomePage" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HomePage")) end if%>">
							</td>
							<td height="16">您企业的网站</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">QQ</div>
							</td>
							<td height="16">
								<input name="frm_QQ" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("QQ")) end if%>">
							</td>
							<td height="16">您常用的腾讯QQ号码</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">MSN</div>
							</td>
							<td height="16">
								<input name="frm_MSN" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("MSN")) end if%>">
							</td>
							<td height="16">您常用的MSN帐户</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">您现在的职业</div>
							</td>
							<td height="16">
								<input name="frm_Vocation" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Vocation")) end if%>">
							</td>
							<td height="16">您现在所从事的职业</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">积分</div>
							</td>
							<td height="16">
								<input name="frm_Integral" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Integral")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
							</td>
							<td height="16"><a href="Integral.asp">[点这里查看详细积分管理]</a></td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">金币</div>
							</td>
							<td height="16">
								<input name="frm_FS_Money" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("FS_Money")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
							</td>
							<td height="16">您的金币和当地金钱等价</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">临时登陆时间</div>
							</td>
							<td height="16">
								<input name="frm_TempLastLoginTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime")) end if%>">
							</td>
							<td height="16">记录某天内登陆的第一次登陆时间，以方便结算金钱</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">临时登陆时间</div>
							</td>
							<td height="16">
								<input name="frm_TempLastLoginTime_1" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime_1")) end if%>">
							</td>
							<td height="16">以方便记录积分</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">会员到期日期</div>
							</td>
							<td height="16">
								<input name="frm_CloseTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CloseTime")) else if request.QueryString("Act")<>"Search" then response.Write("3000-1-1") end if end if%>">
							</td>
							<td height="16">格式：2006-6-4,如果为3000-1-1,表示不过期</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">婚否</div>
							</td>
							<td height="16">
								<select name="frm_IsMarray">
									<%if request.QueryString("Act")="Search" then %>
									<option>请选择</option>
									<%end if%>
									<option value="2"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=2 then response.Write(" selected") end if else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>未婚</option>
									<option value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=1 then response.Write(" selected") end if end if%>>已婚</option>
								</select>
							</td>
							<td height="16">&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">自我介绍</div>
							</td>
							<td height="16">
								<textarea name="frm_SelfIntro" cols="30" rows="6"><%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("SelfIntro")) end if%>
</textarea>
							</td>
							<td height="16">您的自我介绍</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">您的爱好</div>
							</td>
							<td height="16">
								<input name="frm_UserFavor" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("UserFavor")) end if%>">
							</td>
							<td height="16">您的爱好</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">所属会员组</div>
							</td>
							<td height="16">
								<select name="frm_GroupID">
									<%if request.QueryString("Act")="Search" then %>
									<option>请选择</option>
									<%end if%>
									<option value="0" style="color:#FF0000"<%if request.QueryString("Act")<>"Search" then response.Write(" selected") end if%>>不分组</option>
									<%if Bol_IsEdit then 
					response.Write(Get_FildValue_List("select GroupID,GroupName from FS_ME_Group",GetUserDataObj_Rs("GroupID"),1))
				else
					response.Write(Get_FildValue_List("select GroupID,GroupName from FS_ME_Group","",1))
				end if		
				%>
								</select>
							</td>
							<td height="16">&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">是否开启资料</div>
							</td>
							<td height="16">
								<input type="radio" name="frm_isOpen" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=0 then response.Write("checked") end if end if%>>
								关闭
								<input type="radio" name="frm_isOpen" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								开放 </td>
							<td height="16">外人是否可见</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">是否锁定该用户</div>
							</td>
							<td height="16">
								<input type="radio" name="frm_isLock" value="0" <%if Bol_IsEdit then
					 if GetUserDataObj_Rs("isLock")=0 then
						 response.Write("checked")
					 end if 
				 else
					 if request.QueryString("Act")<>"Search" then
						if p_RegisterCheck = 0 then
							response.Write("checked") 
						end if 
					 end if 
				end if%>>
								不锁定
								<input type="radio" name="frm_isLock" value="1" <%if Bol_IsEdit then
					 if GetUserDataObj_Rs("isLock")=1 then
						 response.Write("checked")
					 end if 
				 else
					 if request.QueryString("Act")<>"Search" then
						if p_RegisterCheck = 1 then
							response.Write("checked") 
						end if 
					 end if 
				end if%>>
								锁定 </td>
							<td height="16">锁定后该用户将无法登陆</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">是否允许多人登陆</div>
							</td>
							<td height="16">
								<input type="radio" name="frm_OnlyLogin" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=0 then response.Write("checked") end if end if%>>
								不允许
								<input type="radio" name="frm_OnlyLogin" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								允许 </td>
							<td height="16">如果不选定则表示不允许</td>
						</tr>
						<tr class="hback">
							<td colspan="3" align="center">
								<%if request.QueryString("Act")="Add_BaseData" then %>
								<input name="frm_UserNumber_Edit2" type="hidden" value="<% = request.Form("frm_UserNumber") %>">
								<input name="frm_UserName" type="hidden" value="<% = request.Form("frm_UserName") %>">
								<input name="frm_UserPassword" type="hidden"  value="<% = request.Form("frm_UserPassword") %>">
								<input name="frm_PassQuestion" type="hidden" value="<% = request.Form("frm_PassQuestion") %>">
								<input name="frm_PassAnswer" type="hidden" value="<% = request.Form("frm_PassAnswer") %>">
								<input name="frm_SafeCode" type="hidden" value="<% = request.Form("frm_SafeCode") %>">
								<input name="frm_Email" type="hidden" value="<% = request.Form("frm_Email") %>">
								<%else%>
								<input name="frm_UserNumber_Edit2" type="hidden" value="<% = UserNumber %>">
								<%end if%>
								<input type="submit" name="OtherSubmitButtont" value="<%if request.QueryString("Act")="Search" then response.Write(" 执行查询 ") else response.Write(" 保存企业扩展信息 ") end if%>" />
								<input type="reset" name="Submit2" value=" 重置 " />
							</td>
						</tr>
					</form>
				</table>
			</div>
			<div id="Layer3" style="position:relative;z-index:1; left: 0px; top: 0px;">
				<table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
					<form name="UserForm" method="post"<%if request.QueryString("Act")="Add_OtherData" then 
			   response.Write(" action=""UserCorp_DataAction.asp?Act=Add_AllData""  onsubmit=""return CheckForm_Three(this);""") 
			  elseif request.QueryString("Act")="Search" then 
			  	response.Write(" action=""?Act=SearchGo""")
			  else
			   response.Write(" action=""UserCorp_DataAction.asp?Act=ThreeData""  onsubmit=""return CheckForm_Three(this);""") 
			  end if%>>
						<tr class="hback">
							<td height="16">
								<div align="right"><span class="tx">*</span>公司名称</div>
							</td>
							<td>
								<input name="frm_C_Name" type="text" id="frm_C_Name" style="width:90%" maxlength="50" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Name")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")="Search" then response.Write("模糊查询") else response.Write("请填写您公司的全称") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">公司简称</div>
							</td>
							<td>
								<input name="frm_C_ShortName" type="text" id="frm_C_ShortName" style="width:90%" maxlength="30" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_ShortName")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								请填写您公司的简单称呼
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">请填写您公司所在的省份</td>
							<td>
								<input type="text" name="frm_C_Province" readonly="" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Province")) end if%>">
							</td>
							<td>
								<select name="select111" size=1 onChange="javascript:frm_C_Province.value=this.options[this.selectedIndex].value">
									<option value="">请选择</option>
									<option value="四川">四川</option>
									<option value="北京">北京</option>
									<option value="上海">上海</option>
									<option value="天津">天津</option>
									<option value="重庆">重庆</option>
									<option value="安徽">安徽</option>
									<option value="甘肃">甘肃</option>
									<option value="广东">广东</option>
									<option value="广西">广西</option>
									<option value="贵州">贵州</option>
									<option value="福建">福建</option>
									<option value="海南">海南</option>
									<option value="河北">河北</option>
									<option value="河南">河南</option>
									<option value="黑龙江">黑龙江</option>
									<option value="湖北">湖北</option>
									<option value="湖南">湖南</option>
									<option value="吉林">吉林</option>
									<option value="江苏">江苏</option>
									<option value="江西">江西</option>
									<option value="辽宁">辽宁</option>
									<option value="内蒙古">内蒙古</option>
									<option value="宁夏">宁夏</option>
									<option value="青海">青海</option>
									<option value="山东">山东</option>
									<option value="山西">山西</option>
									<option value="陕西">陕西</option>
									<option value="西藏">西藏</option>
									<option value="新疆">新疆</option>
									<option value="云南">云南</option>
									<option value="浙江">浙江</option>
									<option value="港澳台">港澳台</option>
									<option value="海外">海外</option>
									<option value="其它">其它</option>
								</select>
							</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right">公司所在城市</div>
							</td>
							<td>
								<input name="frm_C_City" type="text" id="frm_C_City" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_City")) end if%>">
							</td>
							<td>您公司所在的城市</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>公司地址</div>
							</td>
							<td>
								<input name="frm_C_Address" type="text" id="frm_C_Address" style="width:90%" maxlength="100" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Address")) end if%>">
							</td>
							<td>您公司的地址</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>邮政编码</div>
							</td>
							<td>
								<input name="frm_C_PostCode" type="text" id="frm_C_PostCode" style="width:90%" maxlength="6" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_PostCode")) end if%>">
							</td>
							<td>您公司的邮政编码</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>公司联系人</div>
							</td>
							<td>
								<input name="frm_C_ConactName" type="text" id="frm_C_ConactName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_ConactName")) end if%>">
							</td>
							<td>公司联系人</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>公司联系电话</div>
							</td>
							<td>
								<input name="frm_C_Tel" type="text" id="frm_C_Tel" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Tel")) end if%>">
							</td>
							<td>公司联系电话。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
						</tr>
						<tr class="hback">
							<td height="1">
								<div align="right">公司传真</div>
							</td>
							<td>
								<input name="frm_C_Fax" type="text" id="frm_C_Fax" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Fax")) end if%>">
							</td>
							<td>公司传真。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
						</tr>
						<tr class="hback">
							<td height="3">
								<div align="right"><span class="tx">*</span>行业</div>
							</td>
							<td>
								<input name="frm_C_VocationClassName" type="text" id="frm_C_VocationClassName" style="width:90%" readonly value="<%if Bol_IsEdit then response.Write(Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&GetUserDataObj_Rs("C_VocationClassID")&"")) end if%>">
								<input type="hidden" name="frm_C_VocationClassID" id="frm_C_VocationClassID" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_VocationClassID")) end if%>">
							</td>
							<td>
								<input type="button" name="Submit3" value="选择行业" onClick="SelectClass();">
								公司所在的行业</td>
						</tr>
						<tr class="hback">
							<td height="8">
								<div align="right">公司网站</div>
							</td>
							<td>
								<input name="frm_C_Website" type="text" id="frm_C_Website" style="width:90%" maxlength="200" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Website")) end if%>">
							</td>
							<td>公司所在的企业站点</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">公司规模</div>
							</td>
							<td>
								<select name="frm_C_size" id="frm_C_size">
									<%if request.QueryString("Act")="Search" then response.Write("<option value="""">请选择</option>") end if%>
									<option value="1-20人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="1-20人" then response.Write(" selected") else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>1-20人</option>
									<option value="21-50人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="21-50人" then response.Write(" selected") end if%>>21-50人</option>
									<option value="51-100人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="51-100人" then response.Write(" selected") end if%>>51-100人</option>
									<option value="101-200人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="101-200人" then response.Write(" selected") end if%>>101-200人</option>
									<option value="201-500人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="201-500人" then response.Write(" selected") end if%>>201-500人</option>
									<option value="501-1000人"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="501-1000人" then response.Write(" selected") end if%>>501-1000人</option>
									<option value="1000人以上"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="1000人以上" then response.Write(" selected") end if%>>1000人以上</option>
								</select>
							</td>
							<td>&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="1">
								<div align="right">公司注册资本</div>
							</td>
							<td>
								<select name="frm_C_Capital" id="frm_C_Capital">
									<option value="10万以下"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="10万以下" then response.Write(" selected") end if%>>10万以下</option>
									<option value="10万-19万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="10万-19万" then response.Write(" selected") end if%>>10万-19万</option>
									<option value="20万-49万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="20万-49万" then response.Write(" selected") end if%>>20万-49万</option>
									<option value="50万-99万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="50万-99万" then response.Write(" selected") else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>50万-99万</option>
									<option value="100万-199万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="100万-199万" then response.Write(" selected") end if%>>100万-199万</option>
									<option value="200万-499万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="200万-499万" then response.Write(" selected") end if%>>200万-499万</option>
									<option value="500万-999万"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="500万-999万" then response.Write(" selected") end if%>>500万-999万</option>
									<option value="1000万以上"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="1000万以上" then response.Write(" selected") end if%>>1000万以上</option>
								</select>
							</td>
							<td>&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="3">
								<div align="right">开户银行</div>
							</td>
							<td>
								<input name="frm_C_BankName" type="text" id="frm_C_BankName" style="width:90%" maxlength="50" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_BankName")) end if%>">
							</td>
							<td rowspan="2">
								<p>公司银行帐户，以方便放在您的联系资料中。<br>
									开户银行例子：中国工商银行成都分行双楠分理处<br>
									银行帐户名：</p>
							</td>
						</tr>
						<tr class="hback">
							<td height="8">
								<div align="right">银行帐号及帐户名</div>
							</td>
							<td>
								<textarea name="frm_C_BankUserName" cols="30" rows="4" id="frm_C_BankUserName"><%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_BankUserName")) end if%>
</textarea>
							</td>
						<tr class="hback">
							<td colspan="3" align="center">
								<%if request.QueryString("Act")="Add_OtherData" then %>
								<input name="frm_UserNumber_Edit3" type="hidden" value="<% = request.Form("frm_UserNumber") %>">
								<input name="frm_UserName" type="hidden" value="<% = request.Form("frm_UserName") %>">
								<input name="frm_UserPassword" type="hidden"  value="<% = request.Form("frm_UserPassword") %>">
								<input name="frm_PassQuestion" type="hidden" value="<% = request.Form("frm_PassQuestion") %>">
								<input name="frm_PassAnswer" type="hidden" value="<% = request.Form("frm_PassAnswer") %>">
								<input name="frm_SafeCode" type="hidden" value="<% = request.Form("frm_SafeCode") %>">
								<input name="frm_Email" type="hidden" value="<% = request.Form("frm_Email") %>">
								<!--扩展-->
								<input name="frm_HeadPic" type="hidden" value="<% = request.Form("frm_HeadPic") %>">
								<input name="frm_HeadPicSize" type="hidden" value="<% = request.Form("frm_HeadPicSize") %>">
								<input name="frm_tel" type="hidden" value="<% = request.Form("frm_tel") %>">
								<input name="frm_Mobile" type="hidden" value="<% = request.Form("frm_Mobile") %>">
								<input name="frm_isMessage" type="hidden" value="<% = request.Form("frm_isMessage") %>">
								<input name="frm_HomePage" type="hidden" value="<% = request.Form("frm_HomePage") %>">
								<input name="frm_QQ" type="hidden" value="<% = request.Form("frm_QQ") %>">
								<input name="frm_MSN" type="hidden" value="<% = request.Form("frm_MSN") %>">
								<input name="frm_Address" type="hidden" value="<% = request.Form("frm_Address") %>">
								<input name="frm_PostCode" type="hidden" value="<% = request.Form("frm_PostCode") %>">
								<input name="frm_Vocation" type="hidden" value="<% = request.Form("frm_Vocation") %>">
								<input name="frm_Integral" type="hidden" value="<% = request.Form("frm_Integral") %>">
								<input name="frm_FS_Money"  type="hidden" value="<% = request.Form("frm_FS_Money") %>">
								<input name="frm_TempLastLoginTime" type="hidden" value="<% = request.Form("frm_TempLastLoginTime") %>">
								<input name="frm_TempLastLoginTime_1" type="hidden" value="<% = request.Form("frm_TempLastLoginTime_1") %>">
								<input name="frm_CloseTime" type="hidden" value="<% = request.Form("frm_CloseTime") %>">
								<input name="frm_IsMarray" type="hidden" value="<% = request.Form("frm_IsMarray") %>">
								<input name="frm_isOpen" type="hidden" value="<% = request.Form("frm_isOpen") %>">
								<input name="frm_GroupID" type="hidden" value="<% = request.Form("frm_GroupID") %>">
								<input name="frm_isLock" type="hidden" value="<% = request.Form("frm_isLock") %>">
								<input name="frm_UserFavor" type="hidden" value="<% = request.Form("frm_UserFavor") %>">
								<input name="frm_OnlyLogin" type="hidden" value="<% = request.Form("frm_OnlyLogin") %>">
								<%else%>
								<input name="frm_UserNumber_Edit3" type="hidden" value="<% = UserNumber %>">
								<%end if%>
								<input type=hidden name="frm_isLockCorp" value="">
								<input type="submit" name="OtherSubmitButtont" onClick="frm_isLockCorp.value=frm_isLock.value;" value="<%if request.QueryString("Act")="Search" then response.Write(" 执行查询 ") else response.Write(" 保存企业相关信息 ") end if%>" />
								<input type="reset" name="Submit2" value=" 重置 " />
							</td>
						</tr>
					</form>
				</table>
			</div>
		</td>
	</tr>
</table>
<%End Sub%>
</body>
<%
Set GetUserDataObj_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
var Req_FildName;
var New_FildName='';
if (Old_Sql.indexOf("Add_Sql=order by ")>-1)
{
	if(Old_Sql.indexOf(" desc")>-1)
		Req_FildName = Old_Sql.substring(Old_Sql.indexOf("Add_Sql=order by ") + "Add_Sql=order by ".length , Old_Sql.indexOf(" desc"));
	else
		Req_FildName = Old_Sql.substring(Old_Sql.indexOf("Add_Sql=order by ") + "Add_Sql=order by ".length , Old_Sql.length);	
	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + " desc")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../../<%=G_USER_DIR%>/lib/SelectClassFrame.asp',400,300,window);
	alert(ReturnValue);
	try {
		document.getElementById('frm_C_VocationClassID').value = ReturnValue.split('***')[0];
		document.getElementById('frm_C_VocationClassName').value = ReturnValue.split('***')[1];
	}
	catch (ex) { }
}

<%if instr(",Add,Edit,Search,Add_BaseData,",","&request.QueryString("Act")&",")>0 then%>
//展开层
var selected="Lab_Base";

<%if request.QueryString("Act")="Add_BaseData" then%>
showDataPanel(2); //添加时提交基础数据后，显示第二个。
<%else%>
showDataPanel(1); //加载时显示第一个.
<%end if%>
function showDataPanel(Data)
{
	switch(Data)
	{
		case 1:
		document.getElementById("Layer1").style.display="block";
		document.getElementById("Layer2").style.display="none";	
		document.getElementById("Layer3").style.display="none";	
		document.getElementById("Lab_Base").className ="";
		if(selected!="Lab_Base")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Base";
		break;
		case 2:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="block";
		document.getElementById("Layer3").style.display="none";
		document.getElementById("Lab_Other").className="";
		if(selected!="Lab_Other")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Other";
		break;
		case 3:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="none";
		document.getElementById("Layer3").style.display="block";
		document.getElementById("Lab_Three").className="";
		if(selected!="Lab_Three")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Three";
		break;
	}
}
<%end if%>
function CheckForm(obj)
{
	<%if p_AllowChineseName = 0 then%>

	obj.frm_UserName.dataType='LimitB';
	obj.frm_UserName.min='<%=p_NumLenMin%>';
	obj.frm_UserName.max='<%=p_NumLenMax%>';
	obj.frm_UserName.msg='用户名必须在[<%=p_NumLenMin%>-<%=p_NumLenMax%>]个字节之间。';

	if( strlen2(obj.frm_UserName.value) ) {
	alert("您的用户名不能有非法字符,或者中文字符")
	obj.frm_UserName.focus();
	return false;
	}
	<%else%>
	obj.frm_UserName.dataType='Limit';
	obj.frm_UserName.min='<%=p_NumLenMin%>';
	obj.frm_UserName.max='<%=p_NumLenMax%>';
	obj.frm_UserName.msg='用户名必须在[<%=p_NumLenMin%>-<%=p_NumLenMax%>]个字符之间。';
	<%End if%>

	<%if p_isValidate = 0 then%>
		 
	obj.frm_UserPassword.dataType="LimitB";
	obj.frm_UserPassword.require="false";
	obj.frm_UserPassword.min='<%=p_LenPassworMin%>';
	obj.frm_UserPassword.max='<%=p_LenPassworMax%>';
	obj.frm_UserPassword.msg='密码必须在[<%=p_LenPassworMin%>-<%=p_LenPassworMax%>]个字节之间。';
	
	obj.frm_cUserPassword.dataType="Repeat";
	obj.frm_cUserPassword.require="false";
	obj.frm_cUserPassword.to="frm_UserPassword";
	obj.frm_cUserPassword.msg="两次输入的密码不一致";		 
	<%End if%>
		
	obj.frm_PassQuestion.dataType="Limit";
	obj.frm_PassQuestion.require="false";
	obj.frm_PassQuestion.min="1";
	obj.frm_PassQuestion.max="36";
	obj.frm_PassQuestion.msg="密码提示问题不能为空并且不能超过36个字符.";		 

	obj.frm_SafeCode.dataType="LimitB";
	obj.frm_SafeCode.require="false";
	obj.frm_SafeCode.min='6';
	obj.frm_SafeCode.max='20';
	obj.frm_SafeCode.msg='安全码必须在[6-20]个字节之间。';

	obj.frm_cSafeCode.dataType="Repeat";
	obj.frm_cSafeCode.require="false";
	obj.frm_cSafeCode.to="frm_SafeCode";
	obj.frm_cSafeCode.msg="两次输入的安全码不一致";		 
		
	obj.frm_Email.dataType="Email";
	obj.frm_Email.msg='信箱格式不正确';
	
	<%if p_AllowChineseName = 0 then%>
	function strlen2(str){		
		var len;
		var i;
		len = 0;
		for (i=0;i<str.length;i++){
			if (str.charCodeAt(i)>255) return true;
		}
		return false;
	}
	<%End if%>

//开始验证
	if ( Validator.Validate(obj,2) )
		{
		if( obj.frm_cUserPassword.value != obj.frm_UserPassword.value)
			{
				alert("重复密码不一致")
				obj.frm_cUserPassword.focus();
				return false;
			}
			if( obj.frm_cSafeCode.value != obj.frm_SafeCode.value )
			{
				alert("重复安全码不一致")
				obj.frm_SafeCode.focus();
				return false;
			}
			if( obj.frm_PassQuestion.value != '' && obj.frm_PassAnswer.value=='')
			{
				alert("提问填写后必须填写回答。")
				obj.frm_PassAnswer.focus();
				return false;
			}
		 <%if request.QueryString("Act")="Add" then%>
		if( obj.frm_UserPassword.value=='' || obj.frm_SafeCode.value=='' || obj.frm_PassQuestion.value=='')
			{
				alert("密码，安全码，提问回答必须填写。")
				return false;
			}		 
		 <%end if%>		
		} 
	else
		return false;
}

//-------------------------------------end
function CheckName(gotoURL) {
   var ssn=document.all.frm_UserName.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=document.all.frm_Email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}

function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}

//-------------------------------------------
function CheckForm_Other(obj)
{
	obj.frm_NickName.dataType="Require";
	obj.frm_NickName.msg="您的昵称不能为空";

	if(obj.frm_Certificate.value=='0')
	{
		obj.frm_CerTificateCode.dataType="IdCard";
		obj.frm_CerTificateCode.msg="身份证号码不正确";
	}
	
	obj.frm_Province.dataType="Require";
	obj.frm_Province.msg="您的省份不能为空";
	
	obj.frm_CloseTime.dataType="date";
	obj.frm_CloseTime.format="ymd";
	obj.frm_CloseTime.msg="到期日期格式不正确";
	
	obj.frm_tel.require="false";
	obj.frm_tel.dataType="Phone";
	obj.frm_tel.msg="电话号码不正确";
	
	obj.frm_Mobile.require="false";
	obj.frm_Mobile.dataType="Mobile";
	obj.frm_Mobile.msg="手机号码不正确";
	
	obj.frm_HomePage.require="false";
	obj.frm_HomePage.dataType="Url";
	obj.frm_HomePage.msg="个人网站格式不正确";
	
	obj.frm_QQ.require="false";
	obj.frm_QQ.dataType="QQ";
	obj.frm_QQ.msg="QQ号码不正确";
	
	obj.frm_PostCode.require="false";
	obj.frm_PostCode.dataType="Zip";
	obj.frm_PostCode.msg="邮政编码不存在";

	obj.frm_Integral.require="false";
	obj.frm_Integral.dataType="Double";
	obj.frm_Integral.msg="积分必须是数字";
	
	obj.frm_FS_Money.require="false";
	obj.frm_FS_Money.dataType="Double";
	obj.frm_FS_Money.msg="金币必须是数字";
	
//开始验证
	return Validator.Validate(obj,2);
}
function CheckForm_Three(obj)
{
	obj.frm_C_Name.dataType="Require";
	obj.frm_C_Name.msg="公司名称不能为空";
	
	obj.frm_C_Address.dataType="Require";
	obj.frm_C_Address.msg="公司地址不能为空";
	
	obj.frm_C_PostCode.require="false";
	obj.frm_C_PostCode.dataType="Zip";
	obj.frm_C_PostCode.msg="公司邮政编码不存在";

	obj.frm_C_Tel.dataType="Phone";
	obj.frm_C_Tel.msg="公司电话格式不正确";
	
	obj.frm_C_VocationClassName.dataType="Require";
	obj.frm_C_VocationClassName.msg="公司所属行业不能为空";
	
//开始验证
	return Validator.Validate(obj,2);
}
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






