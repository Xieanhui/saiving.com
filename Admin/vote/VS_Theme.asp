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
	Str_Tmp = "TID,CID,Theme,Type,MaxNum,DisMode,StartDate,EndDate,ItemMOde"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_VS_Theme"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "Theme"
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
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
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
	IF VS_Rs.eof THEN
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
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Theme.asp?Act=Edit&TID="&VS_Rs("TID")&""" title=""����޸Ļ�鿴��ϸ"">"&VS_Rs("Theme")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Get_OtherTable_Value("select ClassName from FS_VS_Class where CID= "&VS_Rs("CID"))&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("Type"),"1:��ѡ,2:��ѡ")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("MaxNum")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("DisMode"),"1:ֱ��ͼ,2:��ͼ,3:����ͼ")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("StartDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("EndDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""javascript:getCode('"&VS_Rs("TID")&"')"">����</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""TID"" id=""TID"" value="""&VS_Rs("TID")&""" /></td>" & vbcrlf
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
	if not MF_Check_Pop_TF("VS002") then Err_Show
	Dim Str_Tmp
	if request.QueryString("TID")<>"" then 
		Conn.execute("Delete from FS_VS_Theme where TID = "&CintStr(request.QueryString("TID")))
		Conn.execute("Delete from FS_VS_Items where TID = "&CintStr(request.QueryString("TID")))
		Conn.execute("Delete from FS_VS_Items_Result where TID = "&CintStr(request.QueryString("TID")))
		Conn.execute("Delete from FS_VS_Steps where TID = "&CintStr(request.QueryString("TID")))
	else
		Str_Tmp = FormatIntArr(request.form("TID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>"):response.End()
		Str_Tmp = replace(Str_Tmp," ","")
		Conn.execute("Delete from FS_VS_Theme where TID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Items where TID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Items_Result where TID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Steps where TID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Theme.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
Sub Save()
	if not MF_Check_Pop_TF("VS002") then Err_Show
	Dim Str_Tmp,Arr_Tmp,TID,MaxNum
	Str_Tmp = "CID,Theme,Type,DisMode,StartDate,EndDate,ItemMOde"
	Arr_Tmp = split(Str_Tmp,",")	
	TID = NoSqlHack(request.Form("TID"))
	MaxNum = NoSqlHack(request.Form("MaxNum"))
	if not isnumeric(TID) or not TID<>"" then TID = 0
	if not isnumeric(MaxNum) or not MaxNum<>"" then MaxNum = 1
	VS_Sql = "select MaxNum,"&Str_Tmp&"  from FS_VS_Theme  where TID = "&TID
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,3,3
	if not VS_Rs.eof then 
	''�޸�
		VS_Rs("MaxNum") = MaxNum
		for each Str_Tmp in Arr_Tmp
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Theme.asp?Act=Edit&TID="&TID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		if Conn.execute("Select Count(*) from FS_VS_Theme where Theme='"&NoSqlHack(request.Form("Theme"))&"'")(0)>0 then 
			response.Redirect("../error.asp?ErrCodes=<li>���ύ�������Ѿ����ڣ������ظ��ύ��������ؼ��֡�</li>"):response.End()
		end if
		VS_Rs.addnew
		VS_Rs("MaxNum") = MaxNum
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		'response.End()
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Theme.asp?Act=Add&Theme="&request.Form("Theme") ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
	end if
End Sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
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
            
    <td colspan="10" align="left" class="xingmu" >ͶƱ�������</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="VS_Theme.asp">������ҳ</a> | <a href="VS_Theme.asp?Act=Add">����</a> | <a href="VS_Theme.asp?Act=Search">��ѯ</a>	
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
      <td align="center" class="xingmu"><a href="javascript:OrderByName('Theme')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_Theme"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('CID')" class="sd"><b>�������</b></a> 
        <span id="Show_Oder_CID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('Type')" class="sd"><b>��Ŀ����</b></a> 
        <span id="Show_Oder_Type"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('MaxNum')" class="sd"><b>���ѡ��</b></a> 
        <span id="Show_Oder_MaxNum"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('DisMode')" class="sd"><b>��ʾ��ʽ</b></a> 
        <span id="Show_Oder_DisMode"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('StartDate')" class="sd"><b>��ʼʱ��</b></a> 
        <span id="Show_Oder_StartDate"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('EndDate')" class="sd"><b>����ʱ��</b></a> 
        <span id="Show_Oder_EndDate"></span></td>
      <td align="center" class="xingmu" ><b>JS����</b></td>
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
Dim Bol_IsEdit,TID,CID,DisMode,sType,StartDate,EndDate,ItemMOde,MaxNum
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then
	TID = request.QueryString("TID")
	if TID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��TIDû���ṩ��</li>") : response.End()
	VS_Sql = "select TID,CID,Theme,Type,MaxNum,DisMode,StartDate,EndDate,ItemMOde from FS_VS_Theme where TID = "& TID
	Set VS_Rs	= CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,1,1
	if not VS_Rs.eof then 
		Bol_IsEdit = True
		CID = VS_Rs("CID")
		sType = VS_Rs("Type")
		DisMode = VS_Rs("DisMode")
		StartDate = VS_Rs("StartDate")
		EndDate = VS_Rs("EndDate")
		ItemMOde = VS_Rs("ItemMOde")
		MaxNum = VS_Rs("MaxNum")
	end if
elseif request.QueryString("Act") = "Add" then 
	sType = 1
	DisMode = 1 
	StartDate = formatdatetime(now,0)
	EndDate = formatdatetime(dateadd("m",1,now),0) ''һ����
	ItemMOde = 1
	MaxNum = 1
end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" <%if request.QueryString("Act")<>"Search" then response.Write("action=""?Act=Save"" onsubmit=""return chkinput();""") else response.Write("action=""?Act=SearchGo"" onsubmit=""SearchAdd();""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >ͶƱ������Ϣ<%if Bol_IsEdit then	 response.Write("<input type=""hidden"" name=""TID"" id=""TID"" value="""&VS_Rs("TID")&""">") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>

    <tr class="hback"> 
      <td width="100" align="right">�Զ����</td>
      <td>
	  	<input type="text" name="TID" id="TID" size="11" maxlength="11">
      </td>
    </tr>
<%end if%>
    <tr  class="hback"> 
      <td align="right">�������</td>
      <td>
		<select name="CID" id="CID" onChange="Do.these('CID',function(){return isEmpty('CID','CID_Alt')})">
		<option value="">��ѡ��</option>
		<%=Get_FildValue_List("select CID,ClassName from FS_VS_Class",CID,1)%>
		</select>
		<span id="CID_Alt"></span>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��������</td>
      <td>
		<input type="text" name="Theme" id="Theme" size="30" maxlength="50" onFocus="Do.these('Theme',function(){return isEmpty('Theme','Theme_Alt')})" onKeyUp="Do.these('Theme',function(){return isEmpty('Theme','Theme_Alt')})" value="<%if Bol_IsEdit then response.Write(VS_Rs("Theme")) end if%>">
		<span id="Theme_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ŀ����</td>
      <td>
		<select name="Type" id="Type" onChange="Do.these('Type',function(){return isEmpty('Type','Type_Alt')});this.options[this.selectedIndex].value=='2'?MaxNum.disabled=false:MaxNum.disabled=true;">
		<%=PrintOption(sType,":��ѡ��,1:��ѡ,2:��ѡ")%>
		</select>
		<span id="Type_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">���ѡ�����</td>
      <td>
		<input type="text" name="MaxNum" id="MaxNum" size="5" maxlength="4" onFocus="Do.these('MaxNum',function(){return isNumber('MaxNum','MaxNum_Alt','������������',true)})" onKeyUp="Do.these('MaxNum',function(){return isNumber('MaxNum','MaxNum_Alt','������������',true)})" value="<%=MaxNum%>">
		<span class="tx">���ѡ�����,ֻ��Զ�ѡ</span>&nbsp;<span id="MaxNum_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��ʾ��ʽ</td>
      <td> <select name="DisMode" id="DisMode" onChange="Do.these('DisMode',function(){return isEmpty('DisMode','DisMode_Alt')});">
          <%=PrintOption(DisMode,":��ѡ��,1:ֱ��ͼ,"",""")%> 
        </select>          <!--<PrintOption(DisMode,":��ѡ��,1:ֱ��ͼ,2:��ͼ,3:����ͼ")%> -->
        <span id="DisMode_Alt"></span> ��ֻ֧��ֱ��ͼ</td>
    </tr>
     <tr  class="hback"> 
      <td align="right">��ʼʱ��</td>
      <td>
	  <input name="StartDate" type="text" id="StartDate" style="WIDTH: 150px; HEIGHT: 22px"  onfocus="setday(this)" value="<%=StartDate%>" readonly="" maskType="longDate">
	  <IMG onClick="StartDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
	  </td>
    </tr>
   <tr  class="hback"> 
      <td align="right">����ʱ��</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate" readonly=""  onfocus="setday(this)" style="WIDTH: 150px; HEIGHT: 22px" maskType="longDate" value="<%=EndDate%>">
	  <IMG onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom">
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ѡ�����з�ʽ</td>
      <td>
		<select name="ItemMOde" id="ItemMOde" onChange="Do.these('ItemMOde',function(){return isEmpty('ItemMOde','ItemMOde_Alt')});">
		 <%=PrintOption(ItemMOde,":��ѡ��,0:��������,1:1ѡ��/��(����),2:2ѡ��/��,3:3ѡ��/��,4:4ѡ��/��,5:5ѡ��/��,6:6ѡ��/��,7:7ѡ��/��,8:8ѡ��/��,9:9ѡ��/��,10:10ѡ��/��,11:11ѡ��/��,12:12ѡ��/��")%>
		</select>
		<span id="ItemMOde_Alt"></span>		
	  </td>
    </tr>
   <tr class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" ȷ���ύ "/> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%="VS_Theme.asp?Act=Del&TID="&TID%>'">
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
	if (document.all.Type.value!='2' ) document.all.MaxNum.value='1';
	return isEmpty('Type','Type_Alt') && isNumber('MaxNum','MaxNum_Alt','������������',true) && isEmpty('Theme','Theme_Alt') && isEmpty('DisMode','DisMode_Alt') && isEmpty('CID','CID_Alt');
}
function SearchAdd()
{
	if(document.all.StartDate.value) if (document.all.StartDate.value.indexOf('>=')<0) {document.all.StartDate.value='>=#'+document.all.StartDate.value+'#'};
	if(document.all.EndDate.value) if (document.all.EndDate.value.indexOf('<=')<0) {document.all.EndDate.value='<=#'+document.all.EndDate.value+'#'};
}

function getCode(jsid)
{
	if (jsid!=""&&!isNaN(jsid))
	{
		OpenWindow('Frame.asp?PageTitle=��ȡJS���ô���&FileName=showJsPath.asp&JsID='+jsid,360,180,window);
	}else
	{
		alert("���ִ�������ϵ�ͷ���Ա��")
	}
	
}

-->
</script>


<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





