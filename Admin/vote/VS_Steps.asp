<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/VS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,VS_Rs,VS_Sql
Dim AutoDelete,Months
MF_Default_Conn 
MF_Session_TF
if not MF_Check_Pop_TF("VS003") then Err_Show

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
''�õ���ر��ֵ��
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		Get_OtherTable_Value = This_Fun_Rs(0)
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Type&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>"&This_Fun_Sql&"��Ǹ,�����Sql���������.�����ֶβ�����.</li>")
do while not This_Fun_Rs.eof 
	select case Get_Type
	  case 1
		''<option>		
		if instr(This_Fun_Sql,",") >0 then 
			Text = This_Fun_Rs(1)
		else
			Text = This_Fun_Rs(0)
		end if	
		if trim(EquValue) = trim(This_Fun_Rs(0)) then 
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

Function Get_While_Info(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,db_ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "SID,TID,Steps,QuoteID"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_VS_Steps"
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case ""
					''�ַ�
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
					case else
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
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	On Error Resume Next
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open This_Fun_Sql,Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	IF  VS_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
	else	
	VS_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>VS_Rs.PageCount Then cPageNo=VS_Rs.PageCount 
	VS_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Steps.asp?Act=Edit&SID="&VS_Rs("SID")&""" title=""����޸Ļ�鿴��ϸ"">"&VS_Rs("SID")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Get_OtherTable_Value("select ItemName from FS_VS_Items where TID= "&VS_Rs("TID"))&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Steps.asp?Act=Edit&SID="&VS_Rs("SID")&""" title=""����޸Ļ�鿴��ϸ"">��"&VS_Rs("Steps")&"��</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Get_OtherTable_Value("select ItemName from FS_VS_Items where TID= "&VS_Rs("QuoteID"))&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""SID"" id=""SID"" value="""&VS_Rs("SID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		VS_Rs.MoveNext
 		if VS_Rs.eof or VS_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VS_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	VS_Rs.close
	Get_While_Info = Get_Html
End Function

Sub Del()
	if not MF_Check_Pop_TF("VS002") then Err_Show'Ȩ���ж�
	Dim Str_Tmp
	if request.QueryString("SID")<>"" then 
		Conn.execute("Delete from FS_VS_Steps where SID = "&NoSqlHack(request.QueryString("SID")))
	else
		Str_Tmp = FormatIntArr(request.form("SID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>"):response.End()
		Str_Tmp = replace(Str_Tmp," ","")
		Conn.execute("Delete from FS_VS_Steps where SID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Steps.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
Sub Save()
	if not MF_Check_Pop_TF("VS002") then Err_Show
	Dim Str_Tmp,Arr_Tmp,SID,MaxNum
	Str_Tmp = "TID,Steps,QuoteID"
	Arr_Tmp = split(Str_Tmp,",")	
	SID = NoSqlHack(request.Form("SID"))
	if not isnumeric(SID) or not SID<>"" then SID = 0
	VS_Sql = "select "&Str_Tmp&"  from FS_VS_Steps  where SID = "&SID
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,3,3
	if not VS_Rs.eof then 
	''�޸�
		for each Str_Tmp in Arr_Tmp
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Steps.asp?Act=Edit&SID="&SID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		if Conn.execute("Select Count(*) from FS_VS_Steps where TID="&NoSqlHack(request.Form("TID"))&" and  QuoteID="&NoSqlHack(request.Form("QuoteID")))(0)>0 then 
			response.Redirect("../error.asp?ErrCodes=<li>���ύ�������Ѿ����ڣ������ظ��ύ��������ؼ��֡�</li>"):response.End()
		end if
		VS_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		'response.End()
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Steps.asp?Act=Add&TID="&request.Form("TID")&"&Steps="&request.Form("Steps") ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
	end if
End Sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript">
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
-->
</script>
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            
    <td colspan="10" align="left" class="xingmu" >�ಽͶƱ����</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="VS_Steps.asp">������ҳ</a> | <a href="VS_Steps.asp?Act=Add">����</a> | <a href="VS_Steps.asp?Act=Search">��ѯ</a>	
	</td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "Add","Edit","Search"
		Add_Edit_Search
	case "View","SearchGo",""
		View
	case "Save"
		Save
	case "Del"
		Del
	case "OtherSet"
		OtherSet(request.QueryString("Sql"))
	case else
	response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>����Ĳ������ݡ�</li>") : response.End()
end select
'******************************************************************
Sub View()
if not MF_Check_Pop_TF("VS003") then Err_Show
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="?Act=Del">
   <tr  class="hback"> 
      <td align="center" class="xingmu"><a href="javascript:OrderByName('SID')" class="sd"><b>[ID]</b></a> 
        <span id="Show_Oder_SID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('TID')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_TID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('Steps')" class="sd"><b>˳���</b></a> 
        <span id="Show_Oder_Steps"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('QuoteID')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_QuoteID"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
   </form>	
</table>
<%End Sub
Sub Add_Edit_Search()
if not MF_Check_Pop_TF("VS003") then Err_Show
Dim Bol_IsEdit,SID,TID,Steps,QuoteID
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then
	SID = request.QueryString("SID")
	if SID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��SIDû���ṩ��</li>") : response.End()
	VS_Sql = "select SID,TID,Steps,QuoteID from FS_VS_Steps where SID = "& CintStr(SID)
	Set VS_Rs	= CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,1,1
	if not VS_Rs.eof then 
		Bol_IsEdit = True
		TID = VS_Rs("TID")
		Steps = VS_Rs("Steps")
		QuoteID = VS_Rs("QuoteID")
	end if
elseif request.QueryString("Act") = "Add" then 
	TID = request.QueryString("TID")
	Steps = request.QueryString("Steps")
	if Steps = "" then 
		Steps = 1
	else
		if isnumeric(Steps) then 
			Steps = cint(Steps)+1
		else
			Steps = 1
		end if		
	end if		
	QuoteID = ""
end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" <%if request.QueryString("Act")<>"Search" then response.Write("action=""?Act=Save"" onsubmit=""return chkinput();""") else response.Write("action=""?Act=SearchGo""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >ͶƱ������Ϣ<%if Bol_IsEdit then	 response.Write("<input type=""hidden"" name=""SID"" id=""SID"" value="""&VS_Rs("SID")&""">") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>

    <tr class="hback"> 
      <td width="100" align="right">�Զ����</td>
      <td>
	  	<input type="text" name="SID" id="SID" size="11" maxlength="11">
      </td>
    </tr>
<%end if%>
    <tr  class="hback"> 
      <td align="right">��������</td>
      <td>
		<select name="TID" id="TID" onChange="Do.these('TID',function(){return isEmpty('TID','TID_Alt')})">
		<option value="">��ѡ��</option>
		<%=Get_FildValue_List("select TID,Theme from FS_VS_Theme",TID,1)%>
		</select>
		<span id="TID_Alt"></span>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">˳���</td>
      <td>
		<input type="text" name="Steps" id="Steps" size="30" maxlength="3" onFocus="Do.these('Steps',function(){return isEmpty('Steps','Steps_Alt')&&isNumber('Steps','Steps_Alt','����������',true)})" onKeyUp="Do.these('Steps',function(){return isEmpty('Steps','Steps_Alt')&&isNumber('Steps','Steps_Alt','����������',true)})" value="<%=Steps%>">
		<span id="Steps_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��������</td>
      <td>
		<select name="QuoteID" id="QuoteID" onChange="Do.these('QuoteID',function(){return isEmpty('QuoteID','QuoteID_Alt')})">
		<option value="">��ѡ��</option>
		<%=Get_FildValue_List("select TID,Theme from FS_VS_Theme",QuoteID,1)%>
		</select>
		<span id="QuoteID_Alt"></span>
	  </td>
    </tr>
   <tr class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" ȷ���ύ "/> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%="VS_Steps.asp?Act=Del&SID="&SID%>'">
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%
End Sub
set VS_Rs = Nothing
Conn.close
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
function chkinput()
{
	return isEmpty('TID','TID_Alt') && isEmpty('Steps','Steps_Alt') && isNumber('Steps','Steps_Alt','����������',true) && isEmpty('QuoteID','QuoteID_Alt');
}
function SearchAdd()
{
	if(document.all.StartDate.value) if (document.all.StartDate.value.indexOf('>=')<0) {document.all.StartDate.value='>=#'+document.all.StartDate.value+'#'};
	if(document.all.EndDate.value) if (document.all.EndDate.value.indexOf('<=')<0) {document.all.EndDate.value='<=#'+document.all.EndDate.value+'#'};
}
-->
</script>


<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





