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

int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=10 '���ֵ�c����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"				'βҳ


''�õ���ر��ֵ��
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
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>��Ǹ,�����Sql���������.�����ֶβ�����.</li>")
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
		exit do : Get_FildValue_List = "Get_Typeֵ�������" : exit Function 
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
		Get_Html = Get_Html & "<td align=""center""><a href=""#"" onclick=""javascript:if(TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display=='') TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display='none'; else TD_U_"&GetUserDataObj_Rs("UserNumber")&".style.display='';"" class=""otherset"" title='����鿴������Ϣ'>"&GetUserDataObj_Rs("UserNumber")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""UserCorp.asp?Act=Edit&UserNumber="&GetUserDataObj_Rs("UserNumber")&""" class=""otherset"" title='����޸Ĺ�˾��Ϣ'>"&GetUserDataObj_Rs("C_Name")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""UserCorp.asp?Act=Edit&UserNumber="&GetUserDataObj_Rs("UserNumber")&""" class=""otherset"" title='����޸Ĺ�˾��Ϣ'>"&GetUserDataObj_Rs("UserName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""mailto:"&GetUserDataObj_Rs("Email")&""" title=""���ʼ�����"">"& GetUserDataObj_Rs("Email") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("RegTime") & "</td>" & vbcrlf
		if GetUserDataObj_Rs("isLockCorp")&"" <> "0" then 
			''����,��Ҫ����
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""�� ��"" onclick=""javascript:location='UserCorp.asp?Act=OtherEdit&EditSql="&server.URLEncode( "isLock=0" )&"&UserNumber="&GetUserDataObj_Rs("UserNumber")&"'"" title=""�������"" style=""color:red""></td>" & vbcrlf
		else
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""�� ��"" onclick=""javascript:location='UserCorp.asp?Act=OtherEdit&EditSql="&server.URLEncode( "isLock=1" )&"&UserNumber="&GetUserDataObj_Rs("UserNumber")&"'"" title=""�������""></td>" & vbcrlf
		end if
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""frm_UserNumber"" id=""frm_UserNumber"" value="""&GetUserDataObj_Rs("UserNumber")&""" /><input type=""hidden"" name=""frm_UserName"" id=""frm_UserName"" value="""&GetUserDataObj_Rs("UserName")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪�û����ʱ��ʾ��ϸ��Ϣ��
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
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
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
//�����������
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
		<td class="xingmu" colspan=20>��ҵ�û���Ϣ����</td>
	</tr>
	<tr  class="hback">
		<td><a href="UserCorp.asp?Act=View">������ҳ</a>
			<!-- | <a href="UserCorp.asp?Act=Add"><b>����</b></a>-->
			| <a href="UserCorp.asp?Act=Search">��ѯ</a></td>
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
	response.Write("��̨��ӹ�˾��Ϣ�����Ρ����迪ͨ����������ϵ��")
	response.End()
	case "OtherEdit"
	OtherEdit
	case else
	response.Write(request.QueryString("Act")&"�������ݴ���")
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="UserCorp_DataAction.asp?Act=Del">
		<tr  class="hback">
			<td align="center" class="xingmu" ><a href="javascript:OrderByName('UserNumber')" class="sd"><b>���û���š�</b></a> <span id="Show_Oder_UserNumber"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('C_Name')" class="sd"><b>��ҵ����</b></a> <span id="Show_Oder_C_Name"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('UserName')" class="sd"><b>�û���</b></a> <span id="Show_Oder_UserName"></span></td>
			<td align="center" class="xingmu"><a href="javascript:OrderByName('Email')" class="sd"><b>Email</b></a> <span id="Show_Oder_Email"></span></td>
			<td width="10%" align="center" class="xingmu"><a href="javascript:OrderByName('RegTime')" class="sd"><b>ע������</b></a> <span id="Show_Oder_RegTime"></span></td>
			<td width="10%" align="center" class="xingmu"><a href="javascript:OrderByName('isLock')" class="sd"><b>�Ƿ�����</b></a> <span id="Show_Oder_isLock"></span></td>
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
''���ɾ����ѯ���á�
Dim UserNumber,Bol_IsEdit
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	UserNumber = NoSqlHack(Trim(request.QueryString("UserNumber")))
	if UserNumber="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��UserNameû���ṩ</li>") : response.End()
	UserSql = "select UserID,UserName,UserPassword,HeadPic,HeadPicSize,PassQuestion,PassAnswer,safeCode,tel,Mobile,isMessage,Email," _
			  &"HomePage,QQ,MSN,Corner,Province,City,Address,PostCode,NickName,RealName,Vocation,Sex,BothYear,Certificate,CertificateCode,IsCorporation,PopList," _
			  &"PopList,Integral,FS_Money,RegTime,CloseTime,TempLastLoginTime,TempLastLoginTime_1,IsMarray,SelfIntro,isOpen,GroupID,LastLoginIP," _
			  &"ConNumber,ConNumberNews,isLock,UserFavor,MySkin,UserLoginCode,OnlyLogin,hits" _

			  &",C_Name,C_ShortName,C_Province,C_City,C_Address,C_PostCode,C_ConactName,C_Tel,C_Fax,C_VocationClassID,C_Website,C_size,C_Capital,C_BankName,C_BankUserName" _
			  &" from FS_ME_Users A,FS_ME_CorpUser B where  A.UserNumber=B.UserNumber and A.UserNumber= '"& NoSqlHack(UserNumber) &"'"
	Set GetUserDataObj_Rs	= CreateObject(G_FS_RS)
	GetUserDataObj_Rs.Open UserSql,User_Conn,1,1
	if GetUserDataObj_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="hback">
		<td width="140" class="xingmu" colspan="3">
			<%if request.QueryString("Act")<>"Search" then response.Write("��Աϵͳ��������") else response.Write("��ѯ��Ա") end if %>
		</td>
	</tr>
	<tr class="hback">
		<td width="33%"  id="Lab_Base">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>��һ������ҵע����Ϣ</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">��ҵע����Ϣ��ѯģʽ</a>") 
	else
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">��ҵע����Ϣ����</a>") 
	end if%>
			</div>
		</td>
		<td width="33%" height="19" class="xingmu" id="Lab_Other">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>�ڶ�������ҵ��չ��Ϣ</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">��ҵ��չ��Ϣ��ѯģʽ</a>") 
	else 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">��ҵ��չ��Ϣ����</a>") 
	end if%>
			</div>
		</td>
		<td height="19" class="xingmu" id="Lab_Three">
			<div align="center">
				<%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>����������˾�����Ϣ</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(3)"">��ҵ�����Ϣ��ѯģʽ</a>") 
	else 
		response.Write("<a href=""#"" onClick=""showDataPanel(3)"">��ҵ�����Ϣ����</a>") 
	end if%>
			</div>
		</td>
	</tr>
	<tr class="hback">
		<td align="right"  colspan="3">
			<!---�������ݿ�ʼ-->
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
							<td height="20" colspan="3" class="xingmu">����д���Ļ�������<span class="tx">(������ĿΪ�յĲ��޸���������)</span></td>
						</tr>
						<%end if%>
						<tr class="hback">
							<td width="15%" height="65">
								<div align="right">�û���</div>
							</td>
							<td width="29%">
								<input name="frm_UserName" type="text" id="frm_UserName" style="width:90%" value="<%if Bol_IsEdit then response.Write""&GetUserDataObj_Rs("UserName")&"" end if%>" <%if Bol_IsEdit then response.write "readonly"%> >
								<%if request.QueryString("Act")<>"Search" then%>
								<a href="javascript:CheckName('../../user/lib/CheckName.asp')">�Ƿ�ռ��</a>
								<%end if%>
							</td>
							<td width="56%">
								<%if request.QueryString("Act")<>"Search" then%>
								������a��z��Ӣ����ĸ(�����ִ�Сд)��0��9�����֡��㡢���Ż��»��߼�������ɣ�����Ϊ3��18���ַ���ֻ�������ֻ���ĸ��ͷ�ͽ�β,����:coolls1980��
								<%else%>
								ģ����ѯ
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16" colspan="3" class="xingmu">����д��ȫ���ã�����ȫ����������֤�ʺź��һ����룩</td>
						</tr>
						<%If p_isValidate = 0 and request.QueryString("Act")<>"Search" then%>
						<tr class="hback">
							<td height="16">
								<div align="right">����</div>
							</td>
							<td>
								<input name="frm_UserPassword" type="password" id="frm_UserPassword" style="width:90%" maxlength="50">
							</td>
							<td rowspan="2">���볤��Ϊ<%=p_LenPassworMin%>��<%=p_LenPassworMax%>λ��������ĸ��Сд����¼�����������ĸ�����֡������ַ���ɡ�</td>
						</tr>
						<tr class="hback">
							<td height="24">
								<div align="right">ȷ������</div>
							</td>
							<td>
								<input name="frm_cUserPassword" type="password" id="frm_cUserPassword" style="width:90%" maxlength="50">
							</td>
						</tr>
						<%End if
				if request.QueryString("Act") <> "Search" then %>
						<tr class="hback">
							<td height="16">
								<div align="right">������ʾ����</div>
							</td>
							<td>
								<input name="frm_PassQuestion" type="text" id="frm_PassQuestion" style="width:90%" maxlength="30">
							</td>
							<td rowspan="2">������������ʱ���ɴ��һ����롣���磬�����ǡ��ҵĸ����˭��������Ϊ&quot;coolls8&quot;�����ⳤ�Ȳ�����36���ַ���һ������ռ�����ַ����𰸳�����6��30λ֮�䣬���ִ�Сд��</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">�����</div>
							</td>
							<td>
								<input name="frm_PassAnswer" type="text" id="frm_PassAnswer" style="width:90%" maxlength="50">
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��ȫ��</div>
							</td>
							<td>
								<input name="frm_SafeCode" type="password" id="frm_SafeCode" style="width:90%" maxlength="30">
							</td>
							<td rowspan="2">ȫ�������һ��������Ҫ;������ȫ�볤��Ϊ6��20λ��������ĸ��Сд������ĸ�����֡������ַ���ɡ�<br>
								<Span class="tx">�ر����ѣ���ȫ��һ���趨�������������޸�.</Span></td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">ȷ�ϰ�ȫ��</div>
							</td>
							<td>
								<input name="frm_cSafeCode" type="password" id="frm_cSafeCode" style="width:90%" maxlength="30">
							</td>
						</tr>
						<%end if%>
						<tr class="hback">
							<td height="16">
								<div align="right">�����ʼ�</div>
							</td>
							<td>
								<input name="frm_Email" type="text" id="frm_Email" style="width:90%" maxlength="100" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Email")) end if%>">
								<%if request.QueryString("Act")<>"Search" then%>
								<br>
								<a href="javascript:CheckEmail('../../user/lib/Checkemail.asp')">�Ƿ�ռ��</a>
								<%end if%>
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								����ע������ʼ���<Span class="tx">ע��ɹ��󣬽������޸�</span>
								<%end if%>
							</td>
						</tr>
						<!--��Ա���͡�-->
						<input type="hidden" name="frm_UserNumber_Edit1" value="<%=UserNumber%>">
						<tr class="hback">
							<td height="39" colspan="3">
								<div align="center">
									<input type="submit" name="Submit" value="<%if request.QueryString("Act")="Search" then response.Write(" ִ�в�ѯ ") else response.Write(" ������ҵ������Ϣ ") end if%>" style="CURSOR:hand">
									<input type="reset" name="ReSet" id="ReSet" value=" ���� " />
								</div>
							</td>
						</tr>
					</form>
				</table>
			</div>
			<!---�������ݽ���-->
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
								<div align="right"><span class="tx">*</span>�ǳ�</div>
							</td>
							<td>
								<input name="frm_NickName" type="text" id="frm_NickName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("NickName")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								����д��������ǳơ�����Ϊ����
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td width="15%" height="27">
								<div align="right">����</div>
							</td>
							<td width="29%">
								<input name="frm_RealName" type="text" id="frm_RealName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("RealName")) end if%>">
							</td>
							<td width="56%">
								<%if request.QueryString("Act")<>"Search" then%>
								����д������ʵ������
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right"><span class="tx">*</span>�Ա�</div>
							</td>
							<td>
								<input type="radio" name="frm_Sex" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=0 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								��
								<input type="radio" name="frm_Sex" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=1 then response.Write("checked") end if end if%>>
								Ů </td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								����ѡ���Ա�
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">����</td>
							<td>
								<input type="text" name="frm_BothYear" style="width:60%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("BothYear")) end if%>" readonly>
								<input name="SelectDate" type="button" id="SelectDate" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_BothYear);">
							</td>
							<td>
								<%if request.QueryString("Act")="Search" then response.Write("֧�ּ򵥱Ƚ��������*123��123*��123��ģ����ѯ��") else response.Write("����д������ʵ���գ���������ȡ�����롣") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">֤�����</div>
							</td>
							<td>
								<select name=frm_Certificate  id="frm_Certificate">
									<%if request.QueryString("Act")="Search" then response.Write("<option value="""">��ѡ��</option>") end if%>
									<option value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=0 then response.Write("selected") end if else if request.QueryString("Act")<>"Search" then response.Write("selected") end if end if%>>���֤</option>
									<option value="2" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=2 then response.Write("selected") end if end if%>>ѧ��֤</option>
									<option value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=1 then response.Write("selected") end if end if%>>��ʻ֤</option>
									<option value="3" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=3 then response.Write("selected") end if end if%>>����֤</option>
									<option value="4" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=4 then response.Write("selected") end if end if%>>����</option>
								</select>
							</td>
							<td rowspan="2">
								<%if request.QueryString("Act")="Search" then response.Write("֧�ּ򵥱Ƚ��������*123��123*��123��ģ����ѯ��") else response.Write("��Ч֤����Ϊȡ���ʺŵ�����ֶΣ����Ժ�ʵ�ʺŵĺϷ���ݣ����������ʵ��д��<br> <span class=""tx"">�ر����ѣ���Ч֤��һ���趨�����ɸ���</span>") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">֤������</div>
							</td>
							<td>
								<input name="frm_CerTificateCode" type="text" id="frm_CerTificateCode" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CerTificateCode")) end if%>">
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">���������ڵ�ʡ��</td>
							<td>
								<input type="text" name="frm_Province" readonly="" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Province")) end if%>">
							</td>
							<td>
								<select name="select111" size=1 onChange="javascript:frm_Province.value=this.options[this.selectedIndex].value">
									<option value="">��ѡ��</option>
									<option value="�Ĵ�">�Ĵ�</option>
									<option value="����">����</option>
									<option value="�Ϻ�">�Ϻ�</option>
									<option value="���">���</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�㶫">�㶫</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�ӱ�">�ӱ�</option>
									<option value="����">����</option>
									<option value="������">������</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="���ɹ�">���ɹ�</option>
									<option value="����">����</option>
									<option value="�ຣ">�ຣ</option>
									<option value="ɽ��">ɽ��</option>
									<option value="ɽ��">ɽ��</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�½�">�½�</option>
									<option value="����">����</option>
									<option value="�㽭">�㽭</option>
									<option value="�۰�̨">�۰�̨</option>
									<option value="����">����</option>
									<option value="����">����</option>
								</select>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">����</div>
							</td>
							<td height="16">
								<input name="frm_City" type="text" id="frm_City" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("City")) end if%>">
							</td>
							<td height="16">���������ڵĳ���</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��ϵ��ַ</div>
							</td>
							<td height="16">
								<input name="frm_Address" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Address")) end if%>">
							</td>
							<td height="16">������ϵ��ַ</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��������</div>
							</td>
							<td height="16">
								<input name="frm_PostCode" type="text"  size="6" maxlength="6" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("PostCode")) end if%>">
							</td>
							<td height="16">��������</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">ͷ���ַ</div>
							</td>
							<td height="16">
								<input name="frm_HeadPic" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPic")) end if%>">
							</td>
							<td height="16">����ͷ���ַ</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">ͷ��ߴ�</div>
							</td>
							<td height="16">
								<input name="frm_HeadPicSize" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPicSize")) end if%>">
							</td>
							<td height="16">��ʽ��[��,��]��60,60 80,80 120,140</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">˽�˵绰</div>
							</td>
							<td height="16">
								<input name="frm_tel" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("tel")) end if%>">
							</td>
							<td height="16">���ĵ绰</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">�ֻ�</div>
							</td>
							<td height="16">
								<input name="frm_Mobile" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Mobile")) end if%>">
							</td>
							<td height="16">�����ֻ�</td>
						</tr>
						<tr class="hback" id="tr_isMessage">
							<td height="16">
								<div align="right">������֤�ֻ�</div>
							</td>
							<td height="16">
								<input type="checkbox" name="frm_isMessage" value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("isMessage")=1 then response.Write(" checked") end if end if%>>
							</td>
							<td height="16">�Ƿ�ͨ��������֤�ֻ�,������ ���ѡ����,��Ҫͨ������</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��ҵ��վ</div>
							</td>
							<td height="16">
								<input name="frm_HomePage" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HomePage")) end if%>">
							</td>
							<td height="16">����ҵ����վ</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">QQ</div>
							</td>
							<td height="16">
								<input name="frm_QQ" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("QQ")) end if%>">
							</td>
							<td height="16">�����õ���ѶQQ����</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">MSN</div>
							</td>
							<td height="16">
								<input name="frm_MSN" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("MSN")) end if%>">
							</td>
							<td height="16">�����õ�MSN�ʻ�</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">�����ڵ�ְҵ</div>
							</td>
							<td height="16">
								<input name="frm_Vocation" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Vocation")) end if%>">
							</td>
							<td height="16">�����������µ�ְҵ</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">����</div>
							</td>
							<td height="16">
								<input name="frm_Integral" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Integral")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
							</td>
							<td height="16"><a href="Integral.asp">[������鿴��ϸ���ֹ���]</a></td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">���</div>
							</td>
							<td height="16">
								<input name="frm_FS_Money" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("FS_Money")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
							</td>
							<td height="16">���Ľ�Һ͵��ؽ�Ǯ�ȼ�</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��ʱ��½ʱ��</div>
							</td>
							<td height="16">
								<input name="frm_TempLastLoginTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime")) end if%>">
							</td>
							<td height="16">��¼ĳ���ڵ�½�ĵ�һ�ε�½ʱ�䣬�Է�������Ǯ</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��ʱ��½ʱ��</div>
							</td>
							<td height="16">
								<input name="frm_TempLastLoginTime_1" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime_1")) end if%>">
							</td>
							<td height="16">�Է����¼����</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��Ա��������</div>
							</td>
							<td height="16">
								<input name="frm_CloseTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CloseTime")) else if request.QueryString("Act")<>"Search" then response.Write("3000-1-1") end if end if%>">
							</td>
							<td height="16">��ʽ��2006-6-4,���Ϊ3000-1-1,��ʾ������</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">���</div>
							</td>
							<td height="16">
								<select name="frm_IsMarray">
									<%if request.QueryString("Act")="Search" then %>
									<option>��ѡ��</option>
									<%end if%>
									<option value="2"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=2 then response.Write(" selected") end if else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>δ��</option>
									<option value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=1 then response.Write(" selected") end if end if%>>�ѻ�</option>
								</select>
							</td>
							<td height="16">&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">���ҽ���</div>
							</td>
							<td height="16">
								<textarea name="frm_SelfIntro" cols="30" rows="6"><%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("SelfIntro")) end if%>
</textarea>
							</td>
							<td height="16">�������ҽ���</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">���İ���</div>
							</td>
							<td height="16">
								<input name="frm_UserFavor" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("UserFavor")) end if%>">
							</td>
							<td height="16">���İ���</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">������Ա��</div>
							</td>
							<td height="16">
								<select name="frm_GroupID">
									<%if request.QueryString("Act")="Search" then %>
									<option>��ѡ��</option>
									<%end if%>
									<option value="0" style="color:#FF0000"<%if request.QueryString("Act")<>"Search" then response.Write(" selected") end if%>>������</option>
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
								<div align="right">�Ƿ�������</div>
							</td>
							<td height="16">
								<input type="radio" name="frm_isOpen" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=0 then response.Write("checked") end if end if%>>
								�ر�
								<input type="radio" name="frm_isOpen" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								���� </td>
							<td height="16">�����Ƿ�ɼ�</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">�Ƿ��������û�</div>
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
								������
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
								���� </td>
							<td height="16">��������û����޷���½</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">�Ƿ�������˵�½</div>
							</td>
							<td height="16">
								<input type="radio" name="frm_OnlyLogin" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=0 then response.Write("checked") end if end if%>>
								������
								<input type="radio" name="frm_OnlyLogin" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
								���� </td>
							<td height="16">�����ѡ�����ʾ������</td>
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
								<input type="submit" name="OtherSubmitButtont" value="<%if request.QueryString("Act")="Search" then response.Write(" ִ�в�ѯ ") else response.Write(" ������ҵ��չ��Ϣ ") end if%>" />
								<input type="reset" name="Submit2" value=" ���� " />
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
								<div align="right"><span class="tx">*</span>��˾����</div>
							</td>
							<td>
								<input name="frm_C_Name" type="text" id="frm_C_Name" style="width:90%" maxlength="50" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Name")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")="Search" then response.Write("ģ����ѯ") else response.Write("����д����˾��ȫ��") end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��˾���</div>
							</td>
							<td>
								<input name="frm_C_ShortName" type="text" id="frm_C_ShortName" style="width:90%" maxlength="30" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_ShortName")) end if%>">
							</td>
							<td>
								<%if request.QueryString("Act")<>"Search" then%>
								����д����˾�ļ򵥳ƺ�
								<%end if%>
							</td>
						</tr>
						<tr class="hback">
							<td height="24" align="right">����д����˾���ڵ�ʡ��</td>
							<td>
								<input type="text" name="frm_C_Province" readonly="" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Province")) end if%>">
							</td>
							<td>
								<select name="select111" size=1 onChange="javascript:frm_C_Province.value=this.options[this.selectedIndex].value">
									<option value="">��ѡ��</option>
									<option value="�Ĵ�">�Ĵ�</option>
									<option value="����">����</option>
									<option value="�Ϻ�">�Ϻ�</option>
									<option value="���">���</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�㶫">�㶫</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�ӱ�">�ӱ�</option>
									<option value="����">����</option>
									<option value="������">������</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="���ɹ�">���ɹ�</option>
									<option value="����">����</option>
									<option value="�ຣ">�ຣ</option>
									<option value="ɽ��">ɽ��</option>
									<option value="ɽ��">ɽ��</option>
									<option value="����">����</option>
									<option value="����">����</option>
									<option value="�½�">�½�</option>
									<option value="����">����</option>
									<option value="�㽭">�㽭</option>
									<option value="�۰�̨">�۰�̨</option>
									<option value="����">����</option>
									<option value="����">����</option>
								</select>
							</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right">��˾���ڳ���</div>
							</td>
							<td>
								<input name="frm_C_City" type="text" id="frm_C_City" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_City")) end if%>">
							</td>
							<td>����˾���ڵĳ���</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>��˾��ַ</div>
							</td>
							<td>
								<input name="frm_C_Address" type="text" id="frm_C_Address" style="width:90%" maxlength="100" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Address")) end if%>">
							</td>
							<td>����˾�ĵ�ַ</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>��������</div>
							</td>
							<td>
								<input name="frm_C_PostCode" type="text" id="frm_C_PostCode" style="width:90%" maxlength="6" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_PostCode")) end if%>">
							</td>
							<td>����˾����������</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>��˾��ϵ��</div>
							</td>
							<td>
								<input name="frm_C_ConactName" type="text" id="frm_C_ConactName" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_ConactName")) end if%>">
							</td>
							<td>��˾��ϵ��</td>
						</tr>
						<tr class="hback">
							<td height="0">
								<div align="right"><span class="tx">*</span>��˾��ϵ�绰</div>
							</td>
							<td>
								<input name="frm_C_Tel" type="text" id="frm_C_Tel" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Tel")) end if%>">
							</td>
							<td>��˾��ϵ�绰���зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
						</tr>
						<tr class="hback">
							<td height="1">
								<div align="right">��˾����</div>
							</td>
							<td>
								<input name="frm_C_Fax" type="text" id="frm_C_Fax" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Fax")) end if%>">
							</td>
							<td>��˾���档�зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
						</tr>
						<tr class="hback">
							<td height="3">
								<div align="right"><span class="tx">*</span>��ҵ</div>
							</td>
							<td>
								<input name="frm_C_VocationClassName" type="text" id="frm_C_VocationClassName" style="width:90%" readonly value="<%if Bol_IsEdit then response.Write(Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&GetUserDataObj_Rs("C_VocationClassID")&"")) end if%>">
								<input type="hidden" name="frm_C_VocationClassID" id="frm_C_VocationClassID" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_VocationClassID")) end if%>">
							</td>
							<td>
								<input type="button" name="Submit3" value="ѡ����ҵ" onClick="SelectClass();">
								��˾���ڵ���ҵ</td>
						</tr>
						<tr class="hback">
							<td height="8">
								<div align="right">��˾��վ</div>
							</td>
							<td>
								<input name="frm_C_Website" type="text" id="frm_C_Website" style="width:90%" maxlength="200" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_Website")) end if%>">
							</td>
							<td>��˾���ڵ���ҵվ��</td>
						</tr>
						<tr class="hback">
							<td height="16">
								<div align="right">��˾��ģ</div>
							</td>
							<td>
								<select name="frm_C_size" id="frm_C_size">
									<%if request.QueryString("Act")="Search" then response.Write("<option value="""">��ѡ��</option>") end if%>
									<option value="1-20��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="1-20��" then response.Write(" selected") else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>1-20��</option>
									<option value="21-50��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="21-50��" then response.Write(" selected") end if%>>21-50��</option>
									<option value="51-100��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="51-100��" then response.Write(" selected") end if%>>51-100��</option>
									<option value="101-200��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="101-200��" then response.Write(" selected") end if%>>101-200��</option>
									<option value="201-500��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="201-500��" then response.Write(" selected") end if%>>201-500��</option>
									<option value="501-1000��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="501-1000��" then response.Write(" selected") end if%>>501-1000��</option>
									<option value="1000������"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_size")="1000������" then response.Write(" selected") end if%>>1000������</option>
								</select>
							</td>
							<td>&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="1">
								<div align="right">��˾ע���ʱ�</div>
							</td>
							<td>
								<select name="frm_C_Capital" id="frm_C_Capital">
									<option value="10������"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="10������" then response.Write(" selected") end if%>>10������</option>
									<option value="10��-19��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="10��-19��" then response.Write(" selected") end if%>>10��-19��</option>
									<option value="20��-49��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="20��-49��" then response.Write(" selected") end if%>>20��-49��</option>
									<option value="50��-99��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="50��-99��" then response.Write(" selected") else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>50��-99��</option>
									<option value="100��-199��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="100��-199��" then response.Write(" selected") end if%>>100��-199��</option>
									<option value="200��-499��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="200��-499��" then response.Write(" selected") end if%>>200��-499��</option>
									<option value="500��-999��"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="500��-999��" then response.Write(" selected") end if%>>500��-999��</option>
									<option value="1000������"<%if Bol_IsEdit then if GetUserDataObj_Rs("C_Capital")="1000������" then response.Write(" selected") end if%>>1000������</option>
								</select>
							</td>
							<td>&nbsp;</td>
						</tr>
						<tr class="hback">
							<td height="3">
								<div align="right">��������</div>
							</td>
							<td>
								<input name="frm_C_BankName" type="text" id="frm_C_BankName" style="width:90%" maxlength="50" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("C_BankName")) end if%>">
							</td>
							<td rowspan="2">
								<p>��˾�����ʻ����Է������������ϵ�����С�<br>
									�����������ӣ��й��������гɶ�����˫骷���<br>
									�����ʻ�����</p>
							</td>
						</tr>
						<tr class="hback">
							<td height="8">
								<div align="right">�����ʺż��ʻ���</div>
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
								<!--��չ-->
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
								<input type="submit" name="OtherSubmitButtont" onClick="frm_isLockCorp.value=frm_isLock.value;" value="<%if request.QueryString("Act")="Search" then response.Write(" ִ�в�ѯ ") else response.Write(" ������ҵ�����Ϣ ") end if%>" />
								<input type="reset" name="Submit2" value=" ���� " />
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
<!--//�жϺ���������.�ֶ���������ʾָʾ
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
			eval('Show_Oder_'+Req_FildName).innerText = '��';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '��';
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
//չ����
var selected="Lab_Base";

<%if request.QueryString("Act")="Add_BaseData" then%>
showDataPanel(2); //���ʱ�ύ�������ݺ���ʾ�ڶ�����
<%else%>
showDataPanel(1); //����ʱ��ʾ��һ��.
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
	obj.frm_UserName.msg='�û���������[<%=p_NumLenMin%>-<%=p_NumLenMax%>]���ֽ�֮�䡣';

	if( strlen2(obj.frm_UserName.value) ) {
	alert("�����û��������зǷ��ַ�,���������ַ�")
	obj.frm_UserName.focus();
	return false;
	}
	<%else%>
	obj.frm_UserName.dataType='Limit';
	obj.frm_UserName.min='<%=p_NumLenMin%>';
	obj.frm_UserName.max='<%=p_NumLenMax%>';
	obj.frm_UserName.msg='�û���������[<%=p_NumLenMin%>-<%=p_NumLenMax%>]���ַ�֮�䡣';
	<%End if%>

	<%if p_isValidate = 0 then%>
		 
	obj.frm_UserPassword.dataType="LimitB";
	obj.frm_UserPassword.require="false";
	obj.frm_UserPassword.min='<%=p_LenPassworMin%>';
	obj.frm_UserPassword.max='<%=p_LenPassworMax%>';
	obj.frm_UserPassword.msg='���������[<%=p_LenPassworMin%>-<%=p_LenPassworMax%>]���ֽ�֮�䡣';
	
	obj.frm_cUserPassword.dataType="Repeat";
	obj.frm_cUserPassword.require="false";
	obj.frm_cUserPassword.to="frm_UserPassword";
	obj.frm_cUserPassword.msg="������������벻һ��";		 
	<%End if%>
		
	obj.frm_PassQuestion.dataType="Limit";
	obj.frm_PassQuestion.require="false";
	obj.frm_PassQuestion.min="1";
	obj.frm_PassQuestion.max="36";
	obj.frm_PassQuestion.msg="������ʾ���ⲻ��Ϊ�ղ��Ҳ��ܳ���36���ַ�.";		 

	obj.frm_SafeCode.dataType="LimitB";
	obj.frm_SafeCode.require="false";
	obj.frm_SafeCode.min='6';
	obj.frm_SafeCode.max='20';
	obj.frm_SafeCode.msg='��ȫ�������[6-20]���ֽ�֮�䡣';

	obj.frm_cSafeCode.dataType="Repeat";
	obj.frm_cSafeCode.require="false";
	obj.frm_cSafeCode.to="frm_SafeCode";
	obj.frm_cSafeCode.msg="��������İ�ȫ�벻һ��";		 
		
	obj.frm_Email.dataType="Email";
	obj.frm_Email.msg='�����ʽ����ȷ';
	
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

//��ʼ��֤
	if ( Validator.Validate(obj,2) )
		{
		if( obj.frm_cUserPassword.value != obj.frm_UserPassword.value)
			{
				alert("�ظ����벻һ��")
				obj.frm_cUserPassword.focus();
				return false;
			}
			if( obj.frm_cSafeCode.value != obj.frm_SafeCode.value )
			{
				alert("�ظ���ȫ�벻һ��")
				obj.frm_SafeCode.focus();
				return false;
			}
			if( obj.frm_PassQuestion.value != '' && obj.frm_PassAnswer.value=='')
			{
				alert("������д�������д�ش�")
				obj.frm_PassAnswer.focus();
				return false;
			}
		 <%if request.QueryString("Act")="Add" then%>
		if( obj.frm_UserPassword.value=='' || obj.frm_SafeCode.value=='' || obj.frm_PassQuestion.value=='')
			{
				alert("���룬��ȫ�룬���ʻش������д��")
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
	obj.frm_NickName.msg="�����ǳƲ���Ϊ��";

	if(obj.frm_Certificate.value=='0')
	{
		obj.frm_CerTificateCode.dataType="IdCard";
		obj.frm_CerTificateCode.msg="���֤���벻��ȷ";
	}
	
	obj.frm_Province.dataType="Require";
	obj.frm_Province.msg="����ʡ�ݲ���Ϊ��";
	
	obj.frm_CloseTime.dataType="date";
	obj.frm_CloseTime.format="ymd";
	obj.frm_CloseTime.msg="�������ڸ�ʽ����ȷ";
	
	obj.frm_tel.require="false";
	obj.frm_tel.dataType="Phone";
	obj.frm_tel.msg="�绰���벻��ȷ";
	
	obj.frm_Mobile.require="false";
	obj.frm_Mobile.dataType="Mobile";
	obj.frm_Mobile.msg="�ֻ����벻��ȷ";
	
	obj.frm_HomePage.require="false";
	obj.frm_HomePage.dataType="Url";
	obj.frm_HomePage.msg="������վ��ʽ����ȷ";
	
	obj.frm_QQ.require="false";
	obj.frm_QQ.dataType="QQ";
	obj.frm_QQ.msg="QQ���벻��ȷ";
	
	obj.frm_PostCode.require="false";
	obj.frm_PostCode.dataType="Zip";
	obj.frm_PostCode.msg="�������벻����";

	obj.frm_Integral.require="false";
	obj.frm_Integral.dataType="Double";
	obj.frm_Integral.msg="���ֱ���������";
	
	obj.frm_FS_Money.require="false";
	obj.frm_FS_Money.dataType="Double";
	obj.frm_FS_Money.msg="��ұ���������";
	
//��ʼ��֤
	return Validator.Validate(obj,2);
}
function CheckForm_Three(obj)
{
	obj.frm_C_Name.dataType="Require";
	obj.frm_C_Name.msg="��˾���Ʋ���Ϊ��";
	
	obj.frm_C_Address.dataType="Require";
	obj.frm_C_Address.msg="��˾��ַ����Ϊ��";
	
	obj.frm_C_PostCode.require="false";
	obj.frm_C_PostCode.dataType="Zip";
	obj.frm_C_PostCode.msg="��˾�������벻����";

	obj.frm_C_Tel.dataType="Phone";
	obj.frm_C_Tel.msg="��˾�绰��ʽ����ȷ";
	
	obj.frm_C_VocationClassName.dataType="Require";
	obj.frm_C_VocationClassName.msg="��˾������ҵ����Ϊ��";
	
//��ʼ��֤
	return Validator.Validate(obj,2);
}
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






