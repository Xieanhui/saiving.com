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

Dim Hs_Rs
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
	Add_Sql = Decrypt(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "ID,HouseName,PubDate,Audited,Position,Direction,Class,OpenDate,PreSaleNumber,IssueDate,PreSaleRange,Status,Price,PubDate,Tel,Click,UserNumber,Audited"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_HS_Quotation"
	Arr_Tmp = split(Str_Tmp,",")
	for each Str_Tmp in Arr_Tmp
		Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
		if Req_Str<>"" then 				
			select case Str_Tmp
			case "ID","OpenDate","IssueDate","Status","Price","PubDate","Click","Audited","PicNumber","isRecyle"
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
	if instr(Add_Sql,"filterorderby")>0 then 
		This_Fun_Sql = This_Fun_Sql &"  "& replace(replace(Add_Sql,"csed"," Desc"),"filterorderby","Order By ")
	elseif Add_Sql<>"" then 
		This_Fun_Sql = and_where(This_Fun_Sql) &" "& Add_Sql	
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	'response.End()
	On Error Resume Next
	Set Hs_Rs = CreateObject(G_FS_RS)
	Hs_Rs.Open This_Fun_Sql,Conn,1,1
	if Err<>0 then 
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	IF not Hs_Rs.eof THEN

	Hs_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>Hs_Rs.PageCount Then cPageNo=Hs_Rs.PageCount 
	Hs_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" onmouseover=""this.style.cursor = 'hand'"" onclick=""javascript:if(TD_U_"&Hs_Rs("ID")&".style.display=='') TD_U_"&Hs_Rs("ID")&".style.display='none'; else TD_U_"&Hs_Rs("ID")&".style.display='';"" title='����鿴������Ϣ'>"&Hs_Rs("ID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Hs_Rs("HouseName")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Hs_Rs("PubDate")&"</td>" & vbcrlf
		if Hs_Rs("Audited")=0 then 
			Str_Tmp = "δ���"
		elseif Hs_Rs("Audited")=1 then 
			Str_Tmp = "�����"
		else
			Str_Tmp = "δ�����״̬?"	
		end if
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		if Hs_Rs("UserNumber")=0 then 
			Str_Tmp = "ϵͳ����Ա"
		elseif Hs_Rs("UserNumber")<>"" then 
			Str_Tmp = Get_OtherTable_Value("select C_Name+' [��ַ:'+C_Address+' �绰:'+C_Tel+' ����:'+C_Fax+']' as tmpfild from FS_ME_CorpUser where UserNumber='"&Hs_Rs("UserNumber")&"'")
			if Str_Tmp<>"" then 
				Str_Tmp = "<span  onmouseover=""this.style.cursor = 'hand'"" id="""&Hs_Rs("UserNumber")&""" title="""&Str_Tmp&""">������˾:"&Hs_Rs("UserNumber")&"</span>"
			else
				Str_Tmp = "δ֪���û�"
			end if
		else
			Str_Tmp = "������˾"	
		end if
		Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& Hs_Rs("ID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>¥��λ��:"&Hs_Rs("Position") & "</td><td>¥�̷�λ:"&Hs_Rs("Direction")&"</td><td>��Ŀ���:"&Hs_Rs("Class") & "</td><td>��������:"&Hs_Rs("OpenDate")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>Ԥ�����֤:"&Hs_Rs("PreSaleNumber") & "</td><td>��֤����:"&Hs_Rs("IssueDate")&"</td><td>����Ԥ�۷�Χ:"&Hs_Rs("PreSaleRange") & "</td><td>����״��:"&Hs_Rs("Status")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>����:"&Hs_Rs("Price") & "</td><td>��������:"&Hs_Rs("PubDate")&"</td><td>��ϵ�绰:"&Hs_Rs("Tel") & "</td><td>���:"&Hs_Rs("Click")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		Hs_Rs.MoveNext
 		if Hs_Rs.eof or Hs_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(Hs_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	Hs_Rs.close
	Get_While_Info = Get_Html
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-¥������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
//�����������
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;
function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&Add_Sql=")==-1&&Old_Sql.indexOf("?Add_Sql=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&Add_Sql=filterorderby" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?Add_Sql=filterorderby" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].length);
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
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="default.asp">����</a>��¥������</td>
          </tr>
        </table>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >¥����Ϣ����</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="HS_Quotation_Search.asp">��ҳ</a>
      | <a href="HS_Quotation_Search.asp">¥����Ϣ</a> | 
      <a href="HS_Second_Search.asp">���ַ���Ϣ</a> | 
	  <a href="HS_Tenancy_Search.asp">������Ϣ</a></td>
  </tr>
</table>
<%
'******************************************************************
Call View
'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="?Act=Del&TableName=<%=request.QueryString("TableName")%>">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ID')" class="sd"><b>��ID��</b></a> 
        <span id="Show_Oder_ID"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('HouseName')" class="sd"><b>��Դ����</b></a> 
        <span id="Show_Oder_HouseName"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('PubDate')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_PubDate"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('Audited')" class="sd"><b>�Ƿ����</b></a> 
        <span id="Show_Oder_Audited"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserNumber')" class="sd"><b>��Ա���</b></a> 
        <span id="Show_Oder_UserNumber"></span></td>
    </tr>
    <%
		response.Write( Get_While_Info( Encrypt(request.QueryString("Add_Sql")) ) ) 
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
Set Hs_Rs=nothing
Set Fs_User = Nothing
%>
<script language="JavaScript">
<!--//�жϺ���������.�ֶ���������ʾָʾ
//�򿪺���ݹ�����ʾ��ͷ
var Req_FildName;
if (Old_Sql.indexOf("filterorderby")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby") + "filterorderby".length , tmp_arr_[ii].length);break;}	
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

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





