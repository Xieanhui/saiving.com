<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../../FS_Inc/Md5.asp" -->
<!--#include file="../../API/Cls_PassportApi.asp" -->
<%
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_List") then Err_Show
if not MF_Check_Pop_TF("ME001") then Err_Show 

User_GetParm
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

Function and_where(sql)
	if instr(lcase(sql)," where ")>0 then 
		and_where = sql & " and "
	else
		and_where = sql & " where "	
	end if
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
	Str_Tmp = "UserNumber,UserName,Email,Sex,Integral,FS_Money,RegTime,LoginNum,hits,isLock"
	Str_Tmp = Str_Tmp & ",NickName,RealName,BothYear,Certificate,CerTificateCode,Province,City"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_Users where IsCorporation=0"
	if Add_Sql<>"" then 
		if instr(Add_Sql,"order by")>0 then 
			This_Fun_Sql = This_Fun_Sql &"  "& Add_Sql
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
					case "Sex","Integral","FS_Money","RegTime","LoginNum","hits","BothYear","Certificate","isLock"
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
		Get_Html = Get_Html & "<td align=""center""><a href=""User_manage.asp?Act=Edit&UserNumber="&GetUserDataObj_Rs("UserNumber")&""" class=""otherset"" title='����޸�'>"&GetUserDataObj_Rs("UserName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""mailto:"&GetUserDataObj_Rs("Email")&""" title=""���ʼ�����"">"& GetUserDataObj_Rs("Email") & "</td>" & vbcrlf
		if GetUserDataObj_Rs("Sex")=0 then 
			Str_Tmp = "��"
		else
			Str_Tmp = "Ů"
		end if
		Get_Html = Get_Html & "<td align=""center"">" & Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("Integral") & "[��]</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("FS_Money") &"["&p_MoneyName&"]</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("RegTime") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">" & GetUserDataObj_Rs("LoginNum") & "[��]</td>" & vbcrlf
		if cbool(GetUserDataObj_Rs("isLock")) then 
			''����,��Ҫ����
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""�� ��"" onclick=""javascript:location='User_manage.asp?Act=OtherEdit&EditSql=false&UserName="&GetUserDataObj_Rs("UserName")&"';"" alt=""�������"" style=""color:red""></td>" & vbcrlf
		else
			Get_Html = Get_Html & "<td align=""center""><input type=button value=""�� ��"" onclick=""javascript:location='User_manage.asp?Act=OtherEdit&EditSql=true&UserName="&GetUserDataObj_Rs("UserName")&"';"" alt=""�������"" ></td>" & vbcrlf
		end if
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""frm_UserName"" id=""frm_UserName"" value="""&GetUserDataObj_Rs("UserName")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪�û����ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& GetUserDataObj_Rs("UserNumber") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td>�ǳ�:" & GetUserDataObj_Rs("NickName")&" | ����:" & GetUserDataObj_Rs("RealName") &" | ����:"& GetUserDataObj_Rs("BothYear") &" | ʡ:"& GetUserDataObj_Rs("Province") &" | ��:" & GetUserDataObj_Rs("City") &" ����:[" & GetUserDataObj_Rs("hits")&"]" & "</td>" & vbcrlf
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
	Dim UserStatus
	If request.QueryString("EditSql")="true" Then
		UserStatus = 1
	Else
		UserStatus = 0
	End If
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim API_Obj,API_SaveCookie,SysKey
	If API_Enable Then
		SysKey = Md5(request.QueryString("UserName")&API_SysKey,16)
		Set API_Obj = New PassportApi
			API_Obj.NodeValue "syskey",SysKey,0,False
			API_Obj.NodeValue "action","lock",0,False
			API_Obj.NodeValue "username",request.QueryString("UserName"),1,False
			API_Obj.NodeValue "userstatus",UserStatus,1,False
			API_Obj.SendHttpData
		Set API_Obj = Nothing
	End If
	'-----------------------------------------------------------------
	User_Conn.execute("Update FS_ME_Users set isLock="&CintStr(UserStatus)&" where UserName='"&NoSqlHack(request.QueryString("UserName"))&"'")
	response.Redirect("User_manage.asp?Act=View")
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
    <td class="xingmu" colspan=20>�����û���Ϣ����</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="User_manage.asp?Act=View">������ҳ</a>
	 | <a href="User_manage.asp?Act=Add">����</a>
	 | <a href="User_manage.asp?Act=Search">��ѯ</a>
	 | <a href="javascript:history.back();">������һ��</a></td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
	View
	case "Add","Edit","Search","Add_BaseData"
	Add_Edit_Search
	case "OtherEdit"
	OtherEdit
	case else
	response.Write(request.QueryString("Act")&"�������ݴ���")
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="User_DataAction.asp?Act=Del">
    <tr  class="hback">
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('UserNumber')" class="sd"><b>�û����</b></a> <span id="Show_Oder_UserNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserName')" class="sd"><b>�û���</b></a> <span id="Show_Oder_UserName"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('Email')" class="sd"><b>Email</b></a> <span id="Show_Oder_Email"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('Sex')" class="sd"><b>�Ա�</b></a> <span id="Show_Oder_Sex"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('Integral')" class="sd"><b>����</b></a> <span id="Show_Oder_Integral"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('FS_Money')" class="sd"><b>���</b></a> <span id="Show_Oder_FS_Money"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('RegTime')" class="sd"><b>ע������</b></a> <span id="Show_Oder_RegTime"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('LoginNum')" class="sd"><b>��½����</b></a> <span id="Show_Oder_LoginNum"></span></td>
	  <td align="center" class="xingmu">�Ƿ�����</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
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
			  &" from FS_ME_Users where UserNumber= '"& NoSqlHack(UserNumber) &"'"
	Set GetUserDataObj_Rs	= CreateObject(G_FS_RS)
	GetUserDataObj_Rs.Open UserSql,User_Conn,1,1
	if GetUserDataObj_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
end if	
%>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="140" class="xingmu" colspan="3"><%if request.QueryString("Act")<>"Search" then response.Write("��Աϵͳ��������") else response.Write("��ѯ��Ա") end if %></td>
    </tr>
	<tr class="hback"> 
	<td width="50%"  id="Lab_Base"><div align="center"><%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>��һ��������ע����Ϣ</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">����������ѯģʽ</a>") 
	else
		response.Write("<a href=""#"" onClick=""showDataPanel(1)"">������������</a>") 
	end if%></div></td>
	<td width="50%" height="19" class="xingmu" id="Lab_Other"> <div align="center"><%if left(request.QueryString("Act"),3)="Add" then 
		response.Write("<span class=tx>�ڶ�����������ϵ��Ϣ</span>") 
	elseif request.QueryString("Act")="Search" then 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">����������ѯģʽ</a>") 
	else 
		response.Write("<a href=""#"" onClick=""showDataPanel(2)"">������������</a>") 
	end if%></div></td>
	</tr>
    <tr class="hback">
      <td align="right"  colspan="3">
<!---�������ݿ�ʼ-->        
        
      <div id="Layer1" style="position:relative; z-index:1; left: 0px; top: 0px;"> 
        <table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
          <form name="UserForm" method="post"<%if request.QueryString("Act")="Add" then 
			   response.Write(" action=""?Act=Add_BaseData""  onsubmit=""return CheckForm(this);""") 
			  elseif request.QueryString("Act")="Search" then 
			  	response.Write(" action=""?Act=SearchGo""")
			  else
			   response.Write(" action=""User_DataAction.asp?Act=BaseData""  onsubmit=""return CheckForm(this);""") 
			  end if%>>
            <%if request.QueryString("Act")<>"Search" then%>
            <tr class="hback"> 
              <td height="20" colspan="3" class="xingmu">����д���Ļ�������<span class="tx">(������ĿΪ�յĲ��޸���������)</span></td>
            </tr>
            <%end if%>
            <tr class="hback"> 
              <td width="15%" height="65"> <div align="right">�û���</div></td>
              <td width="29%"><input name="frm_UserName" type="text" style="width:90%" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("UserName")) end if%>"> 
                <%if request.QueryString("Act")<>"Search" then%> <a href="javascript:CheckName('../../user/lib/CheckName.asp')">�Ƿ�ռ��</a> 
                <%end if%> </td>
              <td width="56%"> <%if request.QueryString("Act")<>"Search" then%>
                ������a��z��Ӣ����ĸ(�����ִ�Сд)��0��9�����֡��㡢���Ż��»��߼�������ɣ�����Ϊ3��18���ַ���ֻ�������ֻ���ĸ��ͷ�ͽ�β,����:coolls1980�� 
                <%else%>
                ģ����ѯ 
                <%end if%> </td>
            </tr>
            <tr class="hback"> 
              <td height="16" colspan="3" class="xingmu">����д��ȫ���ã�����ȫ����������֤�ʺź��һ����룩</td>
            </tr>
            <%If request.QueryString("Act")<>"Search" then%>
            <tr class="hback"> 
              <td height="16"><div align="right">����</div></td>
              <td><input name="frm_UserPassword" type="password" style="width:90%" maxlength="50"></td>
              <td rowspan="2">���볤��Ϊ<%=p_LenPassworMin%>��<%=p_LenPassworMax%>λ��������ĸ��Сд����¼�����������ĸ�����֡������ַ���ɡ�
			  <span class="tx"><b>�����ϵ�����̳�򲩿��������ͬ�����ǵ������޸ġ�</b></span></td>
            </tr>
            <tr class="hback"> 
              <td height="24"> <div align="right">ȷ������</div></td>
              <td><input name="frm_cUserPassword" type="password" style="width:90%" maxlength="50"></td>
            </tr>
            <%End if
				if request.QueryString("Act") <> "Search" then %>
            <tr class="hback"> 
              <td height="16"><div align="right">������ʾ����</div></td>
              <td><input name="frm_PassQuestion" type="text" style="width:90%" maxlength="30"></td>
              <td rowspan="2">������������ʱ���ɴ��һ����롣���磬�����ǡ��ҵĸ����˭��������Ϊ&quot;coolls8&quot;�����ⳤ�Ȳ�����36���ַ���һ������ռ�����ַ����𰸳�����6��30λ֮�䣬���ִ�Сд��</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">�����</div></td>
              <td><input name="frm_PassAnswer" type="text" style="width:90%" maxlength="50"></td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��ȫ��</div></td>
              <td><input name="frm_SafeCode" type="password" style="width:90%" maxlength="30"></td>
              <td rowspan="2">ȫ�������һ��������Ҫ;������ȫ�볤��Ϊ6��20λ��������ĸ��Сд������ĸ�����֡������ַ���ɡ�<br> 
                <Span class="tx">�ر����ѣ���ȫ��һ���趨�������������޸�.</Span></td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">ȷ�ϰ�ȫ��</div></td>
              <td><input name="frm_cSafeCode" type="password" style="width:90%" maxlength="30"></td>
            </tr>
            <%end if%>
            <tr class="hback"> 
              <td height="16"><div align="right">�����ʼ�</div></td>
              <td><input name="frm_Email" type="text" style="width:90%" maxlength="100" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Email")) end if%>"> 
                <%if request.QueryString("Act")<>"Search" then%> <br> <a href="javascript:CheckEmail('../../user/lib/Checkemail.asp')">�Ƿ�ռ��</a> 
                <%end if%> </td>
              <td> <%if request.QueryString("Act")<>"Search" then%>
                ����ע������ʼ���<Span class="tx">ע��ɹ��󣬽������޸�</span> <%end if%> </td>
            </tr>
            <!--��Ա���͡�-->
            <input type="hidden" name="frm_UserNumber_Edit1" value="<%=UserNumber%>">
            <tr class="hback"> 
              <td height="39" colspan="3"> <div align="center"> 
                  <input type="submit" name="Submit" value="<%if request.QueryString("Act")="Search" then response.Write(" ִ�в�ѯ ") else response.Write(" �����Ա������Ϣ ") end if%>" style="CURSOR:hand">
                  <input type="reset" name="ReSet" value=" ���� " />
                </div></td>
            </tr>
          </form>
        </table>
      </div>	  
<!---�������ݽ���-->        
      <div id="Layer2" style="position:relative; z-index:1; left: 0px; top: 0px; width: 889px; height: 942px;"> 
        <table width="96%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
          <form name="UserForm" id="UserForm" method="post"<%if request.QueryString("Act")="Add_BaseData" then 
			   response.Write(" action=""User_DataAction.asp?Act=Add_AllData""  onsubmit=""return CheckForm_Other(this);""") 
			  elseif request.QueryString("Act")="Search" then 
			  	response.Write(" action=""?Act=SearchGo""")
			  else
			   response.Write(" action=""User_DataAction.asp?Act=OtherData""  onsubmit=""return CheckForm_Other(this);""") 
			  end if%>>
            <tr class="hback"> 
              <td height="27"><div align="right"><span class="tx">*</span>�ǳ�</div></td>
              <td><input name="frm_NickName" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("NickName")) end if%>"></td>
              <td> <%if request.QueryString("Act")<>"Search" then%>
                ����д��������ǳơ�����Ϊ���� 
                <%end if%> </td>
            </tr>
            <tr class="hback"> 
              <td width="15%" height="27"> <div align="right">����</div></td>
              <td width="29%"><input name="frm_RealName" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("RealName")) end if%>"> 
              </td>
              <td width="56%"> <%if request.QueryString("Act")<>"Search" then%>
                ����д������ʵ������ 
                <%end if%> </td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right"><span class="tx">*</span>�Ա�</div></td>
              <td> <input type="radio" name="frm_Sex" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=0 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
                �� 
                <input type="radio" name="frm_Sex" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Sex")=1 then response.Write("checked") end if end if%>>
                Ů </td>
              <td> <%if request.QueryString("Act")<>"Search" then%>
                ����ѡ���Ա� 
                <%end if%> </td>
            </tr>
            <tr class="hback"> 
              <td height="24" align="right">����</td>
              <td> <input type="text" name="frm_BothYear" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("BothYear")) end if%>" readonly></td>
              <td><input name="SelectDate" type="button" id="SelectDate" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_BothYear);" ><%if request.QueryString("Act")="Search" then response.Write("֧�ּ򵥱Ƚ��������*123��123*��123��ģ����ѯ��") else response.Write("����д������ʵ���գ���������ȡ�����롣") end if%></td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">֤�����</div></td>
              <td> <select name=frm_Certificate  id="frm_Certificate">
                  <%if request.QueryString("Act")="Search" then response.Write("<option value="""">��ѡ��</option>") end if%>
                  <option value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=0 then response.Write("selected") end if else if request.QueryString("Act")<>"Search" then response.Write("selected") end if end if%>>���֤</option>
                  <option value="2" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=2 then response.Write("selected") end if end if%>>ѧ��֤</option>
                  <option value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=1 then response.Write("selected") end if end if%>>��ʻ֤</option>
                  <option value="3" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=3 then response.Write("selected") end if end if%>>����֤</option>
                  <option value="4" <%if Bol_IsEdit then if GetUserDataObj_Rs("Certificate")=4 then response.Write("selected") end if end if%>>����</option>
                </select> </td>
              <td rowspan="2"> <%if request.QueryString("Act")="Search" then response.Write("֧�ּ򵥱Ƚ��������*123��123*��123��ģ����ѯ��") else response.Write("��Ч֤����Ϊȡ���ʺŵ�����ֶΣ����Ժ�ʵ�ʺŵĺϷ���ݣ����������ʵ��д��<br> <span class=""tx"">�ر����ѣ���Ч֤��һ���趨�����ɸ���</span>") end if%> </td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">֤������</div></td>
              <td><input name="frm_CerTificateCode" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CerTificateCode")) end if%>"></td>
            </tr>
            <tr class="hback"> 
              <td height="24" align="right"><span class="tx"*</span>���������ڵ�ʡ��</td>
              <td> <input type="text" name="frm_Province" readonly="" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Province")) end if%>"></td>
              <td><select name="select111" size=1 onChange="javascript:frm_Province.value=this.options[this.selectedIndex].value">
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
                </select></td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">����</div></td>
              <td height="16"><input name="frm_City" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("City")) end if%>"></td>
              <td height="16">���������ڵĳ���</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��ϵ��ַ</div></td>
              <td height="16"><input name="frm_Address" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Address")) end if%>"></td>
              <td height="16">������ϵ��ַ</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��������</div></td>
              <td height="16"><input name="frm_PostCode" type="text"  size="6" maxlength="6" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("PostCode")) end if%>"></td>
              <td height="16">��������</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">ͷ���ַ</div></td>
              <td height="16"><input name="frm_HeadPic" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPic")) end if%>"></td>
              <td height="16">����ͷ���ַ</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">ͷ��ߴ�</div></td>
              <td height="16"><input name="frm_HeadPicSize" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HeadPicSize")) end if%>"></td>
              <td height="16">��ʽ��[��,��]��60,60 80,80 120,140</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">˽�˵绰</div></td>
              <td height="16"><input name="frm_tel" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("tel")) end if%>"></td>
              <td height="16">���ĵ绰</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">�ֻ�</div></td>
              <td height="16"><input name="frm_Mobile" type="text" style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Mobile")) end if%>"></td>
              <td height="16">�����ֻ�</td>
            </tr>
            <tr class="hback" id="tr_isMessage"> 
              <td height="16"><div align="right">������֤�ֻ�</div></td>
              <td height="16"> <input type="checkbox" name="frm_isMessage" value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("isMessage")=1 then response.Write(" checked") end if end if%>> 
              </td>
              <td height="16">�Ƿ�ͨ��������֤�ֻ�,������ ���ѡ����,��Ҫͨ������</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">������վ</div></td>
              <td height="16"><input name="frm_HomePage" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("HomePage")) end if%>"></td>
              <td height="16">���ĸ�����վ</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">QQ</div></td>
              <td height="16"><input name="frm_QQ" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("QQ")) end if%>"></td>
              <td height="16">�����õ���ѶQQ����</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">MSN</div></td>
              <td height="16"><input name="frm_MSN" type="text"  style="width:90%" maxlength="50" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("MSN")) end if%>"></td>
              <td height="16">�����õ�MSN�ʻ�</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">�����ڵ�ְҵ</div></td>
              <td height="16"><input name="frm_Vocation" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Vocation")) end if%>"></td>
              <td height="16">�����������µ�ְҵ</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">����</div></td>
              <td height="16"><input name="frm_Integral" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("Integral")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"></td>
              <td height="16"><a href="Integral.asp">[������鿴��ϸ���ֹ���]</a></td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">���</div></td>
              <td height="16"><input name="frm_FS_Money" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("FS_Money")) end if%>" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')"></td>
              <td height="16">���Ľ�Һ͵��ؽ�Ǯ�ȼ�</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��ʱ��½ʱ��</div></td>
              <td height="16"><input name="frm_TempLastLoginTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime")) end if%>"></td>
              <td height="16">��¼ĳ���ڵ�½�ĵ�һ�ε�½ʱ�䣬�Է�������Ǯ</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��ʱ��½ʱ��</div></td>
              <td height="16"><input name="frm_TempLastLoginTime_1" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("TempLastLoginTime_1")) end if%>"></td>
              <td height="16">�Է����¼����</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">��Ա��������</div></td>
              <td height="16"><input name="frm_CloseTime" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("CloseTime")) else if request.QueryString("Act")<>"Search" then response.Write("3000-1-1") end if end if%>"></td>
              <td height="16">��ʽ��2006-6-4,���Ϊ3000-1-1,��ʾ������</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">���</div></td>
              <td height="16"> <select name="frm_IsMarray">
                  <%if request.QueryString("Act")="Search" then %>
                  <option>��ѡ��</option>
                  <%end if%>
                  <option value="2"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=2 then response.Write(" selected") end if else if request.QueryString("Act")<>"Search" then response.Write(" selected") end if end if%>>δ��</option>
                  <option value="1"<%if Bol_IsEdit then if GetUserDataObj_Rs("IsMarray")=1 then response.Write(" selected") end if end if%>>�ѻ�</option>
                </select> </td>
              <td height="16">&nbsp;</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">���ҽ���</div></td>
              <td height="16"><textarea name="frm_SelfIntro" cols="30" rows="6"><%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("SelfIntro")) end if%></textarea></td>
              <td height="16">�������ҽ���</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">���İ���</div></td>
              <td height="16"><input name="frm_UserFavor" type="text"  style="width:90%" maxlength="20" value="<%if Bol_IsEdit then response.Write(GetUserDataObj_Rs("UserFavor")) end if%>"></td>
              <td height="16">���İ���</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">������Ա��</div></td>
              <td height="16"> <select name="frm_GroupID">
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
                </select> </td>
              <td height="16">&nbsp;</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">�Ƿ�������</div></td>
              <td height="16"> <input type="radio" name="frm_isOpen" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=0 then response.Write("checked") end if end if%>>
                �ر� 
                <input type="radio" name="frm_isOpen" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("isOpen")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
                ���� </td>
              <td height="16">�����Ƿ�ɼ�</td>
            </tr>
            <tr class="hback"> 
              <td height="16"><div align="right">�Ƿ��������û�</div></td>
              <td height="16"> <input type="radio" name="frm_isLock" value="0" <%if Bol_IsEdit then
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
              <td height="16"><div align="right">�Ƿ�������˵�½</div></td>
              <td height="16"> <input type="radio" name="frm_OnlyLogin" value="0" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=0 then response.Write("checked") end if end if%>>
                ������ 
                <input type="radio" name="frm_OnlyLogin" value="1" <%if Bol_IsEdit then if GetUserDataObj_Rs("OnlyLogin")=1 then response.Write("checked") end if else if request.QueryString("Act")<>"Search" then response.Write("checked") end if end if%>>
                ���� </td>
              <td height="16">�����ѡ�����ʾ������</td>
            </tr>
            <tr class="hback"> 
              <td colspan="3" align="center"> <%if request.QueryString("Act")="Add_BaseData" then %> <input name="frm_UserNumber_Edit2" type="hidden" value="<% = NoSqlHack(request.Form("frm_UserNumber")) %>"> 
                <input name="frm_UserName" type="hidden" value="<% = NoSqlHack(request.Form("frm_UserName")) %>"> 
                <input name="frm_UserPassword" type="hidden"  value="<% = NoSqlHack(request.Form("frm_UserPassword")) %>"> 
                <input name="frm_PassQuestion" type="hidden" value="<% = NoSqlHack(request.Form("frm_PassQuestion")) %>"> 
                <input name="frm_PassAnswer" type="hidden" value="<% = NoSqlHack(request.Form("frm_PassAnswer")) %>"> 
                <input name="frm_SafeCode" type="hidden" value="<% = NoSqlHack(request.Form("frm_SafeCode")) %>"> 
                <input name="frm_Email" type="hidden" value="<% = NoSqlHack(request.Form("frm_Email")) %>"> 
                <%else%> <input name="frm_UserNumber_Edit2" type="hidden" value="<% = UserNumber %>"> 
                <%end if%> <input type="submit" name="OtherSubmitButtont" value="<%if request.QueryString("Act")="Search" then response.Write(" ִ�в�ѯ ") else response.Write(" �����Ա��չ��Ϣ ") end if%>" /> 
                <input type="reset" name="Submit2" value=" ���� " /> </td>
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
			document.getElementById("Lab_Base").className ="";
			if(selected!="Lab_Base")
			document.getElementById(selected).className ="xingmu";
			selected="Lab_Base";
			break;
			case 2:
			document.getElementById("Layer1").style.display="none";
			document.getElementById("Layer2").style.display="block";
			document.getElementById("Lab_Other").className="";
			if(selected!="Lab_Other")
			document.getElementById(selected).className ="xingmu";
			selected="Lab_Other";
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

	<%if request.QueryString("Act")="Add" then%>
		 
	obj.frm_UserPassword.dataType="LimitB";
	obj.frm_UserPassword.min='<%=p_LenPassworMin%>';
	obj.frm_UserPassword.max='<%=p_LenPassworMax%>';
	obj.frm_UserPassword.msg='���������[<%=p_LenPassworMin%>-<%=p_LenPassworMax%>]���ֽ�֮�䡣';
	
	obj.frm_cUserPassword.dataType="Repeat";
	obj.frm_cUserPassword.to="frm_UserPassword";
	obj.frm_cUserPassword.msg="������������벻һ��";		 

	<%end if%>		
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
function CheckForm_Other(obj)
{
	obj.frm_NickName.dataType="Require";
	obj.frm_NickName.msg="�����ǳƲ���Ϊ��";

	if(obj.frm_Certificate.value=='0')
	{
		obj.frm_CerTificateCode.require="false";
		obj.frm_CerTificateCode.dataType="IdCard";
		obj.frm_CerTificateCode.msg="���֤���벻��ȷ";
	}
	
	//obj.frm_Province.dataType="Require";
	//obj.frm_Province.msg="����ʡ�ݲ���Ϊ��";
	
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
	obj.frm_Integral.dataType="Number";
	obj.frm_Integral.msg="���ֱ���������";
	
	obj.frm_FS_Money.require="false";
	obj.frm_FS_Money.dataType="Number";
	obj.frm_FS_Money.msg="��ұ���������";
	
//��ʼ��֤
	return Validator.Validate(obj,2);
}
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 






