<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<% 
Dim FS_MS_WithDraw_Rs,This_Sub_Sql
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
Function Get_While_Info(Add_Sql)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "ID,WithDrawNumber,OrderNumber,Descryption,State,Dealer,DealTime,ReturnInfo,RealName,Sex,Address,PostCode,Mobile,Tel"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_MS_WithDraw where UserNumber='"&session("FS_UserNumber")&"'"
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
					case "ID","State","DealTime"
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
	Set FS_MS_WithDraw_Rs = CreateObject(G_FS_RS)
	FS_MS_WithDraw_Rs.Open This_Fun_Sql,Conn,1,1	
	IF not FS_MS_WithDraw_Rs.eof THEN

	FS_MS_WithDraw_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>FS_MS_WithDraw_Rs.PageCount Then cPageNo=FS_MS_WithDraw_Rs.PageCount 
	FS_MS_WithDraw_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&FS_MS_WithDraw_Rs("ID")&".style.display=='') TD_U_"&FS_MS_WithDraw_Rs("ID")&".style.display='none'; else TD_U_"&FS_MS_WithDraw_Rs("ID")&".style.display='';"" title='����鿴�������'>"&FS_MS_WithDraw_Rs("ID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_SH_"&FS_MS_WithDraw_Rs("ID")&".style.display=='') TD_SH_"&FS_MS_WithDraw_Rs("ID")&".style.display='none'; else TD_SH_"&FS_MS_WithDraw_Rs("ID")&".style.display='';"" title='����鿴�ջ������'>"&FS_MS_WithDraw_Rs("RealName")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""WithDraw_Apply.asp?Act=Edit&ID="&FS_MS_WithDraw_Rs("ID")&""" class=""otherset"" title='����鿴��ϸ/�޸�'>"&FS_MS_WithDraw_Rs("WithDrawNumber")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""../order.asp?OrderNumber="&FS_MS_WithDraw_Rs("OrderNumber")&""" title=""�鿴�ö�����Ϣ"" target=_blank>"&FS_MS_WithDraw_Rs("OrderNumber")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& left(FS_MS_WithDraw_Rs("Descryption"),50) & "...</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"& Replacestr(FS_MS_WithDraw_Rs("State"),"1:�Ѵ���,0:<span class=tx>δ����</span>,else:δ֪��״̬") & "</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""ID"" id=""ID"" value="""&FS_MS_WithDraw_Rs("ID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf 
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_SH_"& FS_MS_WithDraw_Rs("ID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>�ջ�������:"&FS_MS_WithDraw_Rs("RealName")& " �ֻ�����:"&FS_MS_WithDraw_Rs("Mobile")& " ��ϵ�绰:"&FS_MS_WithDraw_Rs("Tel")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>�����ַ:"&FS_MS_WithDraw_Rs("Address") & " ��������:"&FS_MS_WithDraw_Rs("PostCode") & "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& FS_MS_WithDraw_Rs("ID") &""" style=""display:none""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td  colspan=10>������:<br />"&FS_MS_WithDraw_Rs("ReturnInfo")& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>����������:"&FS_MS_WithDraw_Rs("Dealer") & "</td><td>����ʱ��:"&FS_MS_WithDraw_Rs("DealTime") & "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		''+++++++++++++++++++++++++++++++++++++++		
		FS_MS_WithDraw_Rs.MoveNext
 		if FS_MS_WithDraw_Rs.eof or FS_MS_WithDraw_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(FS_MS_WithDraw_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	FS_MS_WithDraw_Rs.close
	Get_While_Info = Get_Html
End Function

Sub Del()
	Dim Str_Tmp
	if request.QueryString("ID")<>"" then 
		Conn.execute("Delete from FS_MS_WithDraw where ID = "&CintStr(request.QueryString("ID")))
	else
		Str_Tmp = request.form("ID")
		if Str_Tmp="" then response.Redirect("../lib/error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		Str_Tmp = replace(Str_Tmp," ","")
		Conn.execute("Delete from FS_MS_WithDraw where ID in ("&FormatIntArr(Str_Tmp)&")")
	end if
	response.Redirect("../lib/success.asp?ErrorUrl="&server.URLEncode( "../Mall/WithDraw_Apply.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
Function Get_WithDrawNumber()
	Dim this_Str
	this_Str = formatedate(year(now))&formatedate(month(now))&formatedate(day(now))&formatedate(hour(now))&formatedate(minute(now))&formatedate(second(now))
	this_Str = this_Str & GetRamCode(3)
	Get_WithDrawNumber = this_Str
	this_Str = ""
End Function
Function formatedate(str)
	select case len(str)
		case 3,4
			formatedate = right(str,2)
		case 1
			formatedate = "0"&str
	end select
	str = ""
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

Sub Save()
	Dim Str_Tmp,Arr_Tmp,ID,UserNumber
	Str_Tmp = "OrderNumber,Descryption,RealName,Sex,Address,PostCode,Mobile,Tel"
	Arr_Tmp = split(Str_Tmp,",") 
	ID = NoSqlHack(request.Form("ID"))
	if not isnumeric(ID) or ID = "" then ID = 0 
	This_Sub_Sql = "select WithDrawNumber,UserNumber,State,"&Str_Tmp&" from FS_MS_WithDraw where ID="&CintStr(ID)&" and  UserNumber='"&session("FS_UserNumber")&"'"
	Set FS_MS_WithDraw_Rs = CreateObject(G_FS_RS)
	FS_MS_WithDraw_Rs.Open This_Sub_Sql,Conn,3,3
	if ID > 0 then 
	''�޸�
		for each Str_Tmp in Arr_Tmp
			FS_MS_WithDraw_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
			'response.Write(Str_Tmp &":"&NoSqlHack(request.Form("frm_"&Str_Tmp)))&"<br>"
		next	
		FS_MS_WithDraw_Rs.update
		FS_MS_WithDraw_Rs.close
		response.Redirect("../lib/success.asp?ErrorUrl="&server.URLEncode( "../Mall/WithDraw_Apply.asp?Act=Edit&ID="&ID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		FS_MS_WithDraw_Rs.addnew
		FS_MS_WithDraw_Rs("State") = 0
		FS_MS_WithDraw_Rs("UserNumber") = session("FS_UserNumber")
		FS_MS_WithDraw_Rs("WithDrawNumber") = Get_WithDrawNumber()
		for each Str_Tmp in Arr_Tmp
			FS_MS_WithDraw_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
		next	
		FS_MS_WithDraw_Rs.update
		FS_MS_WithDraw_Rs.close
		response.Redirect("../lib/success.asp?ErrorUrl="&server.URLEncode( "../Mall/WithDraw_Apply.asp" ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�˻��������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js"></script>
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
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
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
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="../order.asp">��������</a>���˻��������</td>
          </tr>
        </table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >�˻��������</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="WithDraw_Apply.asp?Act=View">������ҳ</a>
	 | <a href="WithDraw_Apply.asp?Act=Search">��ѯ</a> | 
	 (<a href="WithDraw_Apply.asp?Act=View&Add_Sql=<%=server.URLEncode("state=0")%>">δ����</a> | <a href="WithDraw_Apply.asp?Act=View&Add_Sql=<%=server.URLEncode("state=1")%>">�Ѵ���</a>)</td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
	View
	case  "Add","Edit","Search"
	Add_Edit_Search
	case "Save"
	Save
	case "Del"
	Del
	case "OtherEdit"
	OtherEdit
	case else
	response.Write("��ȷ�������Ƿ���ȷ��")
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="?Act=Del">
    <tr  class="hback"> 
      <td align="center" class="xingmu"><a href="javascript:OrderByName('ID')" class="sd"><b>��ID��</b></a> <span id="Show_Oder_ID"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('RealName')" class="sd"><b>�ջ�������</b></a> <span id="Show_Oder_RealName"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('WithDrawNumber')" class="sd"><b>�˻����</b></a> <span id="Show_Oder_WithDrawNumber"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('OrderNumber')" class="sd"><b>�������</b></a> <span id="Show_Oder_OrderNumber"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('Descryption')" class="sd"><b>�˻�ԭ��</b></a> <span id="Show_Oder_Descryption"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('State')" class="sd"><b>�˻�״̬</b></a> <span id="Show_Oder_State"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit_Search()
Dim ID,Bol_IsEdit,OrderNumber,State,RealName,Sex,Address,PostCode,Mobile,Tel
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	ID = CintStr(request.QueryString("ID"))
	if not isnumeric(ID) or ID="" then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��IDû���ṩ</li>") : response.End()
	This_Sub_Sql = "select ID,OrderNumber,Descryption,State,Dealer,DealTime,ReturnInfo,RealName,Sex,Address,PostCode,Mobile,Tel from FS_MS_WithDraw where ID="&CintStr(ID)&" and  UserNumber='"&session("FS_UserNumber")&"'"
	Set FS_MS_WithDraw_Rs	= CreateObject(G_FS_RS)
	FS_MS_WithDraw_Rs.Open This_Sub_Sql,Conn,1,1
	if FS_MS_WithDraw_Rs.eof then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
	OrderNumber = FS_MS_WithDraw_Rs("OrderNumber")
	State  = FS_MS_WithDraw_Rs("State")
	RealName = FS_MS_WithDraw_Rs("RealName")
	Sex = FS_MS_WithDraw_Rs("Sex")
	Address = FS_MS_WithDraw_Rs("Address")
	PostCode = FS_MS_WithDraw_Rs("PostCode")
	Mobile = FS_MS_WithDraw_Rs("Mobile")
	Tel = FS_MS_WithDraw_Rs("Tel")
elseif request.QueryString("Act")="Add" then 
	State = 0
	OrderNumber = NoSqlHack(request.QueryString("OrderNumber"))
	set FS_MS_WithDraw_Rs = User_Conn.execute("select M_UserName,M_Address,M_Tel,M_PostCode,M_Mobile,M_Sex from FS_ME_Order where OrderNumber='"&NoSqlHack(OrderNumber)&"' and UserNumber='"&session("FS_UserNumber")&"'")
	if FS_MS_WithDraw_Rs.eof then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()

	RealName = FS_MS_WithDraw_Rs("M_UserName")
	Sex = FS_MS_WithDraw_Rs("M_Sex")
	Address = FS_MS_WithDraw_Rs("M_Address")
	PostCode = FS_MS_WithDraw_Rs("M_PostCode")
	Mobile = FS_MS_WithDraw_Rs("M_Mobile")
	Tel = FS_MS_WithDraw_Rs("M_Tel")
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" onSubmit="return Validator.Validate(this,3);" method="post" action="<%if request.QueryString("Act")="Search" then response.Write("?Act=SearchGo") else response.Write("?Act=Save") end if%>">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("�޸��˻���Ϣ<input type=""hidden"" name=""ID"" value="""&FS_MS_WithDraw_Rs(0)&""">") else response.Write("�˻���Ϣ") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>
    <tr  class="hback"> 
       <td width="20%" align="right">�˻�ID</td>
      <td>
		<input name="frm_ID" type="text" value="" size="50" maxlength="11">
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�˻����</td>
      <td>
		<input name="frm_WithDrawNumber" type="text" value="" size="50" maxlength="20">
      </td>
    </tr>
<%end if%>	
    <tr  class="hback"> 
      <td align="right">��Ӧ�������</td>
      <td>
	  <%if request.QueryString("Act")="Search" then %>
	    <input name="frm_OrderNumber" type="text" value="" size="50" maxlength="20">
	  <%else%>
		<input name="frm_OrderNumber" type="hidden" value="<%=OrderNumber%>">
	  <%end if%>	
		<%response.Write("<a href=""../order.asp?OrderNumber="&OrderNumber&""" title=""�鿴��Ӧ�Ķ������"" target=_blank>"&OrderNumber&"</a>")%>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�˻�ԭ��</td>
      <td>
		<textarea name="frm_Descryption" cols="50" rows="10" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Require"" msg=""�˻�ԭ�������д""") end if%>><%if Bol_IsEdit then response.Write(FS_MS_WithDraw_Rs("Descryption")) end if%></textarea>
      </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">�ջ�������</td>
      <td>
		<input type="text" name="frm_RealName" size="50" maxlength="30" value="<%=RealName%>" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Require"" msg=""�ջ�������������д""") end if%>>
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">�ջ����Ա�</td>
      <td>
		<select name="frm_Sex" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Require"" msg=""�ջ�������������д""") end if%>>
		 <%=PrintOption(State,":��ѡ��,1:Ů,0:��")%>
		</select>		 
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">�ջ����ֻ�����</td>
      <td>
		<input type="text" name="frm_Mobile" size="50"  maxlength="20" value="<%=Mobile%>" <%if request.querystring("act")<>"Search" then response.Write("require=false datatype=""Mobile"" msg=""�ջ����ֻ������ʽ����ȷ""") end if%>>
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">�ջ�����ϵ�绰</td>
      <td>
		<input type="text" name="frm_Tel" size="50"  maxlength="20" value="<%=Tel%>" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Require"" msg=""�ջ�����ϵ�绰������д""") end if%>>
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">�ջ���ַ</td>
      <td>
		<input type="text" name="frm_Address" size="50"  maxlength="100" value="<%=Address%>" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Require"" msg=""�ջ��˵�ַ������д""") end if%>>
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">��������</td>
      <td>
		<input type="text" name="frm_PostCode" size="50"  maxlength="6" value="<%=PostCode%>" <%if request.querystring("act")<>"Search" then response.Write("datatype=""Zip"" msg=""���������ʽ����ȷ""") end if%>>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">����״̬</td>
      <td>
		<%if Bol_IsEdit then 
			response.Write(Replacestr(State,"1:�Ѵ���,else:<span class=tx>δ����</span>"))
		elseif request.QueryString("Act")="Search" then%>
			<select name="frm_State">
			 <%=PrintOption(State,":��ѡ��,1:�Ѵ���,0:δ����")%>
			</select>
		<%end if%>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">������</td>
      <td>	  	
		<%if Bol_IsEdit then 
			response.Write(FS_MS_WithDraw_Rs("Dealer"))
		elseif request.QueryString("Act")="Search" then%>
			<input type="text" name="frm_Dealer" size="50">
		<%end if%>	
	  </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">����ʱ��</td>
      <td>
		<%if Bol_IsEdit then 
			response.Write(FS_MS_WithDraw_Rs("DealTime"))
		elseif request.QueryString("Act")="Search" then%>
			<input type="text" name="frm_DealTime" size="50">
		<%end if%>
	  </td>
    </tr>
   <tr  class="hback"> 
      <td align="right">�����������</td>
      <td>
		<%if Bol_IsEdit then 
			response.Write(FS_MS_WithDraw_Rs("ReturnInfo"))
		elseif request.QueryString("Act")="Search" then%>
			<textarea name="frm_ReturnInfo" cols="50" rows="10"></textarea>
		<%end if%>	
	  </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" value=" <%if request.QueryString("Act")="Search" then response.Write("��ѯ") else response.Write("����") end if%> " <%if Bol_IsEdit then response.Write(Replacestr(State,"1:disabled title=""�Ѿ��������˻������ܱ༭"",else:")) end if%> /> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" value=" ɾ�� " <%if request.QueryString("Act")<>"Edit" then response.Write("disabled") end if%> onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("WithDraw_Apply.asp?Act=Del&ID="&ID)%>'">
            </td>
          </tr>
        </table>
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
</body>
<%
Set FS_MS_WithDraw_Rs=nothing
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
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 






