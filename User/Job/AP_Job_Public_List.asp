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
	if Err<>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
  
Function Get_While_Info(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "PID,JobName,JobDescription,ResumeLang,WorkCity,PublicDate,EndDate,NeedNum,EducateExp,Sex,WorkAge,Age,JobType,OtherJobDes,MoneyMonth,FreeMoney,OtherMoneyDes,HolleType"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_AP_Job_Public where UserNumber = '"&Session("FS_UserNumber")&"'"
	Arr_Tmp = split(Str_Tmp,",")
	for each Str_Tmp in Arr_Tmp
		Req_Str = NoSqlHack(Trim(request.QueryString(Str_Tmp)))
		if Req_Str<>"" then 				
			select case Str_Tmp
			case "PID","PublicDate","EndDate","NeedNum"
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
				if left(Req_Str,1)="" then
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
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	if instr(Add_Sql,"order by")>0 then 
		This_Fun_Sql = This_Fun_Sql &"  "& Add_Sql
	end if
	Str_Tmp = ""
	On Error Resume Next
	Set Ap_Rs = CreateObject(G_FS_RS)
	Ap_Rs.Open This_Fun_Sql,Conn,1,1	

	if Err<>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>��ȷ������д�����������Ƿ�ƥ��</li>") : response.End()
	end if

	IF  Ap_Rs.eof THEN
		Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 height=30 align=""center"" class=""ischeck"">δ��ѯ��������������Ϣ"
		Get_Html = Get_Html &"</td></tr>"
	else
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
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&Ap_Rs("PID")&".style.display=='') TD_U_"&Ap_Rs("PID")&".style.display='none'; else TD_U_"&Ap_Rs("PID")&".style.display='';"" title='����鿴������Ϣ'>"&Ap_Rs("PID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("JobName")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("PublicDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Ap_Rs("EndDate")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""AP_Job_Public_AddUpdate.asp?Act=Edit&PID="&Ap_Rs("PID")&""">�༭</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""PID"" id=""PID"" value="""&Ap_Rs("PID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& Ap_Rs("PID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td  colspan=10>ְλ����:"&Ap_Rs("JobDescription")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>���ܼ�������:"&Ap_Rs("ResumeLang") & "</td><td>�����ص�:"&Ap_Rs("WorkCity") & "</td><td>��Ƹ����:"&Ap_Rs("NeedNum")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		Ap_Rs.MoveNext
 		if Ap_Rs.eof or Ap_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(Ap_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	Ap_Rs.close
	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>�ѷ�������Ƹ---��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
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
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="job_applications.asp">��Ƹ��ҳ</a>���ѷ�������Ƹ</td>
          </tr>
        </table>

<%
'******************************************************************
Call View
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="AP_Job_Public_Action.asp?Act=Del">
     <tr  class="hback"> 
      <td colspan="10" align="left" class="xingmu" >�ѷ�������Ϣ</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="AP_Job_Public_List.asp">�ѷ�������Ϣ</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">�����Ƹ��Ϣ</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">��ѯ��Ƹ��Ϣ</a>	
	</td>
  </tr>

   <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('PID')" class="sd">���</a> 
        <span id="Show_Oder_PID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('JobName')" class="sd">ְλ����</a> 
        <span id="Show_Oder_JobName"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('PublicDate')" class="sd">��������</a> 
        <span id="Show_Oder_PublicDate"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('EndDate')" class="sd">��Ч����</a> 
        <span id="Show_Oder_EndDate"></span></td>
      <td align="center" class="xingmu">�༭</td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql") ) )
	%>
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
-->
</script>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





