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
if not MF_Check_Pop_TF("ME_Intergel") then Err_Show

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

set VClass_Rs=User_Conn.execute("select top 1 MoneyName from FS_ME_SysPara")
if not VClass_Rs.eof then 
Sys_MoneyName = VClass_Rs(0)
end if
VClass_Rs.close

Function Get_Card(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "UserID,UserNumber,UserName,Sex,Integral,FS_Money"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_Users"
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if Trim(request("frm_"&Str_Tmp))<>"" then 
				Req_Str = NoSqlHack(Trim(request("frm_"&Str_Tmp)))
				if Str_Tmp="FS_Money" then
				Req_Str = request("jb1")& Req_Str
				elseif  Str_Tmp="Integral" then
				Req_Str = request("jf1")& Req_Str
				end if
				select case Str_Tmp
					case "UserID","Integral","FS_Money"
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
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
		'response.End()
	end if
	Str_Tmp = ""
	On Error Resume Next
	Set VClass_Rs = CreateObject(G_FS_RS)
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1	
	IF VClass_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
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
		Get_Html = Get_Html & "<td align=""center""><a href=""Integral.asp?Act=Edit&UserID="&VClass_Rs("UserID")&""" class=""otherset"" title='����޸�'>"&VClass_Rs("UserNumber")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Integral.asp?Act=Edit&UserID="&VClass_Rs("UserID")&""" class=""otherset"" title='����޸�'>"&VClass_Rs("UserName")&"</a></td>" & vbcrlf
		for ii=3 to 5
			select case ii
				case 3 
				if VClass_Rs(ii)=1 then 
					Str_Tmp="Ů"
				else
					Str_Tmp="��"
				end if		
				case 4
				Str_Tmp = VClass_Rs(ii) & "��"
				case 5
				Str_Tmp = VClass_Rs(ii) & Sys_MoneyName
				case else
				Str_Tmp = VClass_Rs(ii)
			end select		
				Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		next
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><a href=""Integral_MoreInfo.asp?Act=View&Add_Sql="&server.URLEncode( Encrypt("UserNumber='"&VClass_Rs("UserNumber")&"'") )&"&UserID="&VClass_Rs("UserID")&""" class=""otherset"" title='�鿴��ϸ�嵥'>�鿴��ϸ</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""UserID"" id=""UserID"" value="""&VClass_Rs("UserID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ��Ա���ֳ� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	VClass_Rs.close
	Get_Card = Get_Html
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
Sub Del()
	if not MF_Check_Pop_TF("ME019") then Err_Show 'Ȩ���ж�
	Dim Str_Tmp
	if request.QueryString("UserID")<>"" then 
		User_Conn.execute("Update FS_ME_Users Set Integral='0' where UserID = "&CintStr(request.QueryString("UserID")))'Crazy
	else
		Str_Tmp = FormatIntArr(request.form("UserID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		Str_Tmp = replace(Str_Tmp," ","")
		User_Conn.execute("Update FS_ME_Users Set Integral='0' where UserID in ("&Str_Tmp&")")'Crazy
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================

Sub Save()
	Dim Str_Tmp,Arr_Tmp,UserID
	Str_Tmp = "Integral,FS_Money"
	Arr_Tmp = split(Str_Tmp,",")
	UserID = NoSqlHack(request.Form("UserID"))
	if not isnumeric(UserID) or UserID = "" then UserID = 0 
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_Users where UserID="&UserID
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if UserID > 0 then 
	''�޸�
		if not MF_Check_Pop_TF("ME002") then Err_Show 'Ȩ���ж�
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next	
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral.asp?Act=Edit&UserID="&UserID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		if not MF_Check_Pop_TF("ME002") then Err_Show 'Ȩ���ж�
		VClass_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next	
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Integral.asp?Act=Add" ) &"ErrCodes=<li>��ϲ�������ɹ���</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

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
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace(oldFildName,FildName):New_Sql = Old_Sql.replace(oldFildName,FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
-->
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >�����ܹ���</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="Integral.asp?Act=View">������ҳ</a> | <a href="Integral.asp?Act=Search">��ѯ</a> | <a href="javascript:history.back();">������һ��</a></td>
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
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('UserID')" class="sd"><b>���û���š�</b></a> <span id="Show_Oder_UserID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserName')" class="sd"><b>�û���</b></a> <span id="Show_Oder_UserName"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('Sex')" class="sd"><b>�Ա�</b></a> <span id="Show_Oder_Sex"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('Integral')" class="sd"><b>����</b></a> <span id="Show_Oder_Integral"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('FS_Money')" class="sd"><b>���</b></a> <span id="Show_Oder_FS_Money"></span></td>
	  <td align="center" class="xingmu">�鿴��ϸ</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_Card( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim UserID,Bol_IsEdit
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	UserID = request.QueryString("UserID")
	if UserID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��UserIDû���ṩ</li>") : response.End()
	VClass_Sql = "select UserID,UserNumber,UserName,Sex,Integral,FS_Money from FS_ME_Users where UserID="&UserID
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return Validator.Validate(this,3);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("�޸��û����ֵ���Ϣ<input type=""hidden"" name=""UserID"" value="""&VClass_Rs(0)&""">") else response.Write("�û����ֵ���Ϣ") end if%></td>
	</tr>
    <tr  class="hback"> 
      <td align="right">�û����</td>
      <td>
		<%if Bol_IsEdit then response.Write(VClass_Rs(1)) end if%>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�û���</td>
      <td>
		<%if Bol_IsEdit then response.Write(VClass_Rs(2)) end if%>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ա�</td>
      <td>
		<%if Bol_IsEdit then if VClass_Rs(3)=0 then response.Write("��") else response.Write("Ů") end if end if%>
	  </td>
    </tr>


    <tr  class="hback"> 
      <td align="right">����</td>
      <td>
	  <input type="text" name="frm_Integral" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(4)) end if%>" dataType="Compare" msg="����>=0" to="0" operator="GreaterThanEqual" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">���</td>
      <td>
	  <input type="text" name="frm_FS_Money" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(5)) end if%>" dataType="Compare" msg="����>=0" to="0" operator="GreaterThanEqual" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
	   <%=Sys_MoneyName%>
	  </td>

    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" ���� " /> <!--<%IF request.QueryString("Act")="Put" then%> onClick="Put_CardNum_Len.to = (Put_CardAddStr.value.length+2).toString();Put_CardNum_Len.msg='���ȱ�����ڵ���'+(Put_CardAddStr.value.length+2).toString()" <%end if%>-->
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("Integral.asp?Act=Del&UserID="&UserID)%>'">
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
      <td colspan="3" align="left" class="xingmu" >��ѯ��Ա����</td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">�û����</td>
      <td> <input type="text" name="frm_UserNumber" size="40" value="">
        ģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">�û���</td>
      <td> <input type="text" name="frm_UserName" size="40" value="">
        ģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">�Ա�</td>
      <td> 
         <input type="radio" name="frm_sex" value="0">��
         <input type="radio" name="frm_sex" value="1">Ů
 </td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">����</td>
      <td> <select name="JF1" style="width:55">
        <option value="" selected="selected"></option>
		<option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
      <input type="text" name="frm_Integral" size="30" value="" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
      ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">���</td>
      <td> <select name="JB1" style="width:55">
	     <option value="" selected="selected"></option>
	     <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_FS_Money" size="30" value="" onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
      <%=Sys_MoneyName%> ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ</td>
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
</body>
<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
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
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 






