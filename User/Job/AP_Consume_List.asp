<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if not Session("FS_UserNumber")<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>����δ��½,�����.</li>&ErrorUrl=../login.asp") : response.End()

Dim Ap_Rs
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
''�õ���ر��ֵ��
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
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
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
  
Function Get_While_Info(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp,PublicDate,EndDate
	PublicDate = NoSqlHack(request.QueryString("PublicDate"))
	EndDate = NoSqlHack(request.QueryString("EndDate"))
	Str_Tmp = "FS_AP_Consume,ConsumeDate,SearchUId,SearchUName,LeftCount"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_AP_Consume where UserNumber = '"&Session("FS_UserNumber")&"'"
	if 1=2 then
	Arr_Tmp = split(Str_Tmp,",")
	for each Str_Tmp in Arr_Tmp
		Req_Str = NoSqlHack(Trim(request.QueryString(Str_Tmp)))
		if Req_Str<>"" then 				
			select case Str_Tmp
			case "FS_AP_Consume","ConsumeDate"
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
				if left(Req_Str,1)="=" then
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &left(Req_Str,1)&"'"& mid(Req_Str,2) &"'"
				elseif left(Req_Str,2)="<>" then
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" not like '%"& mid(Req_Str,3) &"%'"
				elseif instr(Req_Str,"*")>0 then 
					if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
					if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
				else	
					New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '%"& Req_Str &"%'"
				end if		
			end select 		
		end if
	next
	end if
	if G_IS_SQL_DB = 1 then PublicDate = replace(PublicDate,"#","'"):EndDate = replace(EndDate,"#","'")
	if PublicDate<>"" then New_Search_Str = and_where(New_Search_Str) & "ConsumeDate "&PublicDate
	if EndDate<>"" then New_Search_Str = and_where(New_Search_Str) & "ConsumeDate "&EndDate
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	if instr(Add_Sql,"order by")>0 then 
		This_Fun_Sql = This_Fun_Sql &"  "& Add_Sql
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	Set Ap_Rs = CreateObject(G_FS_RS)
	Ap_Rs.Open This_Fun_Sql,Conn,1,1	
	IF not Ap_Rs.eof THEN

	Ap_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>Ap_Rs.PageCount Then cPageNo=Ap_Rs.PageCount 
	Ap_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("FS_AP_Consume")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("ConsumeDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("SearchUId")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("SearchUName")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("LeftCount")&"</td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		Ap_Rs.MoveNext
 		if Ap_Rs.eof or Ap_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(Ap_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	Ap_Rs.close
	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��ֵ��¼��ѯ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
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
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="job_applications.asp">��Ƹ��ҳ</a>����ֵ��¼��ѯ</td>
          </tr>
        </table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            <td colspan="10" align="left" class="xingmu" >��ֵ��¼��ѯ</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="AP_Payment_List.asp">��ֵ��¼��ѯ</a> | <a href="AP_Consume_List.asp">���Ѽ�¼��ѯ</a>	
	</td>
  </tr>
</table>
<%
'******************************************************************
if request.QueryString("Act")="SearchGo" then 
	Call View
else
	Call Search
end if	
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
   <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('FS_AP_Consume')" class="sd"><b>���</b></a> 
        <span id="Show_Oder_FS_AP_Consume"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('ConsumeDate')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_ConsumeDate"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('SearchUId')" class="sd"><b>��ѯ����ϵ�˱��</b></a> 
        <span id="Show_Oder_SearchUId"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('SearchUName')" class="sd"><b>��ѯ����ϵ������</b></a> 
        <span id="Show_Oder_SearchUName"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('LeftCount')" class="sd"><b>ʣ�����</b></a> 
        <span id="Show_Oder_LeftCount"></span></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql") ) )
	%>
</table>
<%End Sub
Sub Search()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="get" onSubmit="return chkinput();">
	<tr  class="hback"> 
            <td width="200" align="right">����������ʼ����</td>
      <td>
	  <input type="hidden" name="Act" value="SearchGo">
	  <input type="text" name="PublicDate" id="PublicDate" readonly="" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate">
	  <IMG id="img3" onClick="PublicDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="PublicDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�������ڽ�ֹ����</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate" readonly=""  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%=date+1%>">
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EndDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <input type="submit" value=" �顤ѯ " /> 
              &nbsp; <input type="reset" value=" ���� " />
      </td>
    </tr>	
</form>
</table>
<%End Sub%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Ap_Rs=nothing
Set Fs_User = Nothing
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
function chkinput()
{
	if(document.all.PublicDate.value) if (document.all.PublicDate.value.indexOf('>=')<0) {document.all.PublicDate.value='>=#'+document.all.PublicDate.value+'#'};
	if(document.all.EndDate.value) if (document.all.EndDate.value.indexOf('<=')<0) {document.all.EndDate.value='<=#'+document.all.EndDate.value+'#'};
}
-->
</script>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





