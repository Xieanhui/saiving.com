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
'---------------------------------��ҳ����
int_RPP=15 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ

set VClass_Rs=User_Conn.execute("select count(*) from FS_ME_GroupDebateClass")
if VClass_Rs(0) = 0 then response.Redirect("lib/Error.asp?ErrorUrl=../main.asp&ErrCodes=<li>��Ǹ����Ⱥ������δ����������ϵ����Ա�� </li>") : response.End()
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
		''=========================================
		'''vclass��ʾClassID,Hy_vclass��ʾClassType
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
		response.Redirect("lib/error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF VClass_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
	else	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""GroupClass.asp?Act=Edit&gdID="&VClass_Rs("gdID")&""" class=""otherset"" title='����޸�'>��"&VClass_Rs("gdID")&"��</a></td>" & vbcrlf
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_GroupDebateClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"��") ''��Ⱥ����
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& VClass_Rs("Title") & "</td>" & vbcrlf
		select case VClass_Rs("InfoType")
			case 0
			Str_Tmp = "����"  
			case 1
			Str_Tmp = "����"
			case 2
			Str_Tmp = "��Ʒ"
			case 3
			Str_Tmp = "����" 
			case 4
			Str_Tmp = "����"
			case 5
			Str_Tmp = "��ְ"
			case 6
			Str_Tmp = "��Ƹ"
			case 7
			Str_Tmp = "����"
			case else
			Str_Tmp = "��"
		end select 
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Str_Tmp = Get_FildValue("select vClassName from FS_ME_VocationClass where VCID="&set_Def(VClass_Rs("ClassType"),0),"��") ''��ҵ����
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
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("lib/Error.asp?ErrCodes=<li>��Ǹ,Get_FildValue_List���������Sql���������.�����ֶβ�����.</li>")
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
''================================================================

Sub Del()
	Dim Str_Tmp
	if request.QueryString("gdID")<>"" then 
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID = "&CintStr(request.QueryString("gdID")))
	else
		Str_Tmp = FormatIntArr(request.form("gdID"))
		if Str_Tmp="" then response.Redirect("lib/Error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		
		User_Conn.execute("Delete from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID in ("&Str_Tmp&")")
	end if
	response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub

Sub Save()
	'''vclass��ʾClassID,Hy_vclass��ʾClassType
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
	''�޸�
		''''''''''''''''''''''''''
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next
		if New_ClassID<>"" then VClass_Rs("ClassID") = New_ClassID
		if New_ClassType<>"" then VClass_Rs("ClassType") = New_ClassType
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=Edit&gdID="&gdID )&"&ErrCodes=<li>��ϲ���޸���Ⱥ�ɹ���</li>")
	else
	''����
	''���Ⱥ����
'		Dim rsCount,CountSQL
'		set rsCount = Server.CreateObject(G_FS_RS)
'		CountSQL = "select gdID From FS_ME_GroupDebateManage where UserNumber='"&Fs_User.UserNumber&"'"
'		rsCount.open CountSQL,User_Conn,1,1
'		Call getGroupIDinfo()
'		if Cint(GroupDebateNum) <= rsCount.recordcount then
'			Response.Redirect("lib/Error.asp?ErrCodes=<li>����������Ⱥ�����Ѿ���������ޣ�������������Ⱥ����Ϊ��"&split(GroupDebateNum,",")(0)&"����</li>")
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
		response.Redirect("lib/Success.asp?ErrorUrl="&server.URLEncode( "../GroupClass.asp?Act=Add&gdID="&gdID )&"&ErrCodes=<li>��ϲ��������Ⱥ�ɹ���</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<%if instr(",Add,Edit,Search,",","&request.QueryString("Act")&",")>0 then%>
<script language="javascript" src="../FS_Inc/class_liandong.js" type="text/javascript"></script>
<%end if%>
<script language="JavaScript">
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
    <td align="left" class="hback_1"><a href="GroupClass.asp?Act=Add">�½�</a>	
	 | <%if request.QueryString("Act")="Edit" then
	 	response.Write("<a href=""GroupClass.asp?Act=Del&gdID="&NoSqlHack(request.QueryString("gdID"))&""" title=""ȷ��ɾ��������¼��?"">ɾ��</a>")
	elseif request.QueryString("Act")="View" or request.QueryString("Act")="" then
		response.Write("<a href=""javascript:if (confirm('ȷ��ɾ����?')) document.form1.submit();"">ɾ��</a>")
	else
		response.Write("ɾ��")		
	end if%> | <a href="GroupClass.asp?Act=View">�鿴ȫ��</a> | <a href="GroupClass.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("islock"))%>">������</a> 
	| <a href="GroupClass.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("not islock"))%>">δ����</a>
	 | <a href="GroupClass.asp?Act=Search">��ѯ</a></td>
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
      <td align="center" class="hback_1"><a href="javascript:OrderByName('gdID')" class="sd"><b>����š�</b></a> 
        <span id="Show_Oder_gdID"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('ClassID')" class="sd"><b>��Ⱥ����</b></a> 
        <span id="Show_Oder_ClassID"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('Title')" class="sd"><b>��Ⱥ����</b></a> 
        <span id="Show_Oder_Title"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('InfoType')" class="sd"><b>Ӧ������</b></a> 
        <span id="Show_Oder_InfoType"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('ClassType')" class="sd"><b>������ҵ</b></a> 
        <span id="Show_Oder_ClassType"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('hits')" class="sd"><b>����</b></a> 
        <span id="Show_Oder_hits"></span></td>
      <td align="center" class="hback_1"><a href="javascript:OrderByName('AddTime')" class="sd"><b>����ʱ��</b></a> 
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
	if gdID="" then response.Redirect("lib/Error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��gdIDû���ṩ</li>") : response.End()
	VClass_Sql = "select gdID,ClassID,Title,Content,InfoType,ClassType,AccessFile,UserNumber,AdminName,ClassMember,PerPageNum from FS_ME_GroupDebateManage where (UserNumber = '" & Fs_User.UserNumber & "' OR AdminName = '" & Fs_User.UserNumber & "') And gdID="&NoSqlHack(gdID)
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("lib/Error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
	AppointUserNumber = VClass_Rs(4)
	AppointUserGroup = VClass_Rs(5)
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return Validator.Validate(this,3);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" > <%if Bol_IsEdit then response.Write("�޸���Ⱥ��Ϣ"&vbNewLine&"<input type=""hidden"" name=""gdID"" value="""&VClass_Rs(0)&""">" ) else response.Write("�����Ⱥ��Ϣ")  end if%> </td>
    </tr>
    <tr  class="hback"<%if not Bol_IsEdit then response.Write(" style=""display:none""") end if%>> 
      <td width="25%" align="right">������Ⱥ����</td>
      <td><strong>
        <%if Bol_IsEdit then response.Write( Get_FildValue( "select vClassName from FS_ME_GroupDebateClass where VCID="&set_Def(VClass_Rs("ClassID"),0),"��" ) ) end if%>
        </strong></td>
    </tr>
    <tr  class="hback"> 
      <td width="25%" align="right"><%if Bol_IsEdit then response.Write("�����Ϊ��") else response.Write("������Ⱥ����") end if%></td>
      <td> 
        <!---�����˵���ʼ--->
        <select name="vclass1" id="vclass1"<%if not Bol_IsEdit then%> datatype="Require" msg="������д"<%end if%> onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---�����˵�����--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����</td>
      <td> <input type="text" name="frm_Title" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(2)) end if%>" dataType="Require" msg="������д">
                    ֧��A* *B A B�����ַ���ͬ��</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����</td>
      <td> <textarea name="frm_Content" cols="40" rows="5" dataType="Require" msg="������д"><%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%></textarea> 
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">�����ĸ�����</td>
      <td>
		  <select name="frm_InfoType" datatype="Require" msg="����ѡ��">
          <option value="0"<%if Bol_IsEdit then if VClass_Rs(4)=0 then response.Write(" selected") end if end if%>>����</option>
          <%if IsExist_SubSys("DS") Then%><option value="1"<%if Bol_IsEdit then if VClass_Rs(4)=1 then response.Write(" selected") end if end if%>>����</option><%end if%>
          <%if IsExist_SubSys("MS") Then%><option value="2"<%if Bol_IsEdit then if VClass_Rs(4)=2 then response.Write(" selected") end if end if%>>��Ʒ</option><%end if%>
          <%if IsExist_SubSys("HS") Then%><option value="3"<%if Bol_IsEdit then if VClass_Rs(4)=3 then response.Write(" selected") end if end if%>>����</option><%end if%>
          <%if IsExist_SubSys("SD") Then%><option value="4"<%if Bol_IsEdit then if VClass_Rs(4)=4 then response.Write(" selected") end if end if%>>����</option><%end if%>
          <%if IsExist_SubSys("AP") Then%><option value="5"<%if Bol_IsEdit then if VClass_Rs(4)=5 then response.Write(" selected") end if end if%>>��ְ</option><%end if%>
          <%if IsExist_SubSys("AP") Then%><option value="6"<%if Bol_IsEdit then if VClass_Rs(4)=6 then response.Write(" selected") end if end if%>>��Ƹ</option>
         <option value="7"<%if Bol_IsEdit then if VClass_Rs(4)=7 then response.Write(" selected") end if end if%>>����</option><%end if%>
        </select> </td>
    </tr>
    <tr class="hback"<%if not Bol_IsEdit then response.Write(" style=""display:none""") end if%>> 
      <td align="right">��Ⱥ������ҵ</td>
      <td><strong>
        <%if Bol_IsEdit then response.Write( Get_FildValue( "select vClassName from FS_ME_VocationClass where VCID="&set_Def(VClass_Rs("ClassType"),0),"��" ) ) end if%>
        </strong></td>
    </tr>
    <tr  class="hback"> 
      <td align="right"><%if Bol_IsEdit then response.Write("�����Ϊ��") else response.Write("��Ⱥ������ҵ") end if%></td>
      <td> 
        <!---�����˵���ʼ--->
        <select name="Hy_vclass1" id="select"<%if not Bol_IsEdit then%> datatype="Require" msg="������д"<%end if%> onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---�����˵�����--->
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">������ַ</td>
      <td> <input type="text" name="frm_AccessFile" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(6)) end if%>"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ��ʼ���û����</td>
      <td> <input name="frm_UserNumber" type="text" value="<%if Bol_IsEdit then response.Write(VClass_Rs(7)) else response.Write(session("FS_UserNumber")) end if%>" size="40" datatype="Require" msg="������д"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ���ڹ���Ա�û����</td>
      <td> <input name="frm_AdminName" type="text" value="<%if Bol_IsEdit then response.Write(VClass_Rs(8)) else response.Write(session("FS_UserNumber")) end if%>" size="40" datatype="Require" msg="������д"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ�ĳ�Ա</td>
      <td> <textarea name="frm_ClassMember" cols="40" datatype="Require" msg="������д"><%if Bol_IsEdit then response.Write(VClass_Rs(9)) else response.Write(session("FS_UserNumber")) end if%></textarea>
        ������á�,���ֿ� </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����ÿҳ��ʾ��������</td>
      <td>
	   <input type="text" name="frm_PerPageNum" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(10)) end if%>" dataType="Range" msg="��1~30֮��" min="0" max="31"> 
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
            <td align="center"> <input type="submit" name="submit" value=" ���� " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " /> 
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
      <td colspan="3" align="left" class="xingmu" >��ѯ��Ⱥ��Ϣ</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ID���</td>
      <td><input type="text" name="frm_gdID" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td width="25%" align="right">������Ⱥ����</td>
      <td> 
        <!---�����˵���ʼ--->
        <select name="vclass1" id="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="vclass2" id="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass3" id="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
          <option></option>
        </select> <select name="vclass4" id="vclass4" style="width:100px">
          <option></option>
        </select> 
        <!---�����˵�����--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����</td>
      <td><input type="text" name="frm_Title" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����</td>
      <td> <textarea name="frm_Content" cols="40" rows="5"></textarea> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ɲ鿴���û����</td>
      <td> <textarea name="frm_AppointUserNumber" cols="40"></textarea>
        ������á�,���ֿ� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�ɲ鿴�Ļ�Ա��</td>
      <td> <textarea name="frm_AppointUserGroup" cols="40"></textarea> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�����ĸ�����</td>
      <td> <select name="frm_InfoType">
          <option value="">��ѡ��</option>
          <option value="0">����</option>
          <option value="1">����</option>
          <option value="2">��Ʒ</option>
          <option value="3">����</option>
          <option value="4">����</option>
          <option value="5">��ְ</option>
          <option value="6">��Ƹ</option>
          <option value="7">����</option>
        </select> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ������ҵ</td>
      <td> 
        <!---�����˵���ʼ--->
        <select name="Hy_vclass1" id="select" onBlur="javascript:RemoveChildopt(this,'Hy_vclass2,Hy_vclass3,Hy_vclass4');"  style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass2" id="select2" onBlur="javascript:RemoveChildopt(this,'Hy_vclass3,Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass3" id="select3" onBlur="javascript:RemoveChildopt(this,'Hy_vclass4');" style="width:100px">
          <option></option>
        </select> <select name="Hy_vclass4" id="select4" style="width:100px">
          <option></option>
        </select> 
        <!---�����˵�����--->
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">����ʱ��</td>
      <td> <input type="text" name="frm_AddTime" size="40" value=""> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ƿ�����</td>
      <td> <input type="radio" name="frm_isLock"  value="true">
        ������ 
        <input type="radio" name="frm_isLock"  value="false">
        δ���� </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">������ַ</td>
      <td> <input type="text" name="frm_AccessFile" size="40" value=""> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ��ʼ���û����</td>
      <td> <input name="frm_UserNumber" type="text" value="" size="40"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ���ڹ���Ա�û����</td>
      <td> <input name="frm_AdminName" type="text" value="" size="40"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ�ĳ�Ա</td>
      <td> <textarea name="frm_ClassMember" cols="40"></textarea>
        ������á�,���ֿ� </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ⱥ����ÿҳ��ʾ��������</td>
      <td> <input type="text" name="frm_PerPageNum" size="40" value="" require="false" dataType="Range" msg="��1~30֮��" min="0" max="31"> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">����/�����</td>
      <td> <input type="text" name="frm_hits" size="40">
                    ֧��&gt;=&lt;&lt;&gt;�ȷ�����������������ͬ��</td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" ִ�в�ѯ " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " /> 
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

///��ҵ

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
				//�����������
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->