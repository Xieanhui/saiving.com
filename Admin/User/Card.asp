<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,User_Conn,VClass_Rs,VClass_Sql,sErrStr
Dim CheckStr,Need_Do_Pwd_Str
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Card") then Err_Show

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

Function Get_Card(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "CardID,CardNumber,CardMoney,CardDateNumber,CardPoint,isBuy,CardOverDueTime,UserNumber,UserTime  ,AddTime,IsUse,CardPasswords"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_Card"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(request(Str_Tmp))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "CardMoney","CardDateNumber","CardPoint","IsUse","isBuy"   ,"CardOverDueTime","UserTime","AddTime"
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
	end if
	Str_Tmp = "":ii=0
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	On Error Resume Next
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
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
		Get_Html = Get_Html & "<td align=""center""><a href=""Card.asp?Act=Edit&CardID="&VClass_Rs("CardID")&""" class=""otherset"" title='����޸�'>"&VClass_Rs("CardNumber")&"</a></td>" & vbcrlf
		for ii=2 to 8
			select case ii
				case 2
				Str_Tmp = VClass_Rs(ii) & "Ԫ"
				case 5 
				if VClass_Rs(ii)=1 then 
					Str_Tmp="���۳�"
				else
					Str_Tmp="δ�۳�"
				end if		
				case else
				Str_Tmp = VClass_Rs(ii)
			end select		
				Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		next	
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""DelID"" id=""DelID"" value="""&VClass_Rs("CardID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_Card = Get_Html
End Function

Sub Del()
	if not MF_Check_Pop_TF("ME019") then Err_Show 'Ȩ���ж�
	Dim Str_Tmp
	if request.QueryString("CardID")<>"" then 
		User_Conn.execute("Delete from FS_ME_Card where CardID = "&CintStr(request.QueryString("CardID")))
	else
		Str_Tmp = FormatIntArr(request.form("DelID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		
		User_Conn.execute("Delete from FS_ME_Card where CardID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
Function CheckCardCF(CardNumber)
''���¼��Ŀ��Ƿ��ظ�,�ظ��򷵻ؿ���,���ظ��򷵻�""
	Dim CheckCardCF_Rs
	Set CheckCardCF_Rs = CreateObject(G_FS_RS)
	CheckCardCF_Rs.Open "select Count(*) from FS_ME_Card where CardNumber='"&NoSqlHack(CardNumber)&"'",User_Conn,1,1
	if  CheckCardCF_Rs(0)>0 then 
		CheckCardCF = CardNumber
	else 
		CheckCardCF = ""
	end if
	CheckCardCF_Rs.close	
End Function

Sub Save()
	Dim CardID,IsOk,Arr_Tmp,Str_Tmp,Arr_Tmp1,Str_Tmp1,Arr_Tmp2,Str_Tmp2,ErrInfo,tmpi
	Dim PutNum,CardAddStr,CardNum_Len,CardPwd_Len,Put_i,Put_Rs,Str_Tmp3,Put_Pwd_Type
    Dim Randchar,Randchararr,iR
	CardID = NoSqlHack(request.Form("CardID"))
	if not isnumeric(CardID) or CardID = "" then CardID = 0
	Str_Tmp = "CardNumber,CardPasswords,CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,isBuy,UserNumber,AddTime,UserTime"
	Arr_Tmp = split(Str_Tmp,",")
	for tmpi = 2 to 7 
		if trim(request.Form("frm_"&Arr_Tmp(tmpi))) = "" then sErrStr = Arr_Tmp(tmpi) : exit for
	next
	if sErrStr <>"" then response.Redirect("../error.asp?ErrCodes=<li>��Ҫ�Ĳ��� ["&sErrStr&"] ������д��������</li>")  : response.End()
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_Card where CardID="&CardID
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if CardID >0 then 
	''�޸�
	if not MF_Check_Pop_TF("ME018") then Err_Show 'Ȩ���ж�
		for each Str_Tmp in Arr_Tmp
			if Str_Tmp = "CardPasswords" then 
				VClass_Rs(Str_Tmp) = Encrypt(NoSqlHack(request.Form("frm_"&Str_Tmp))) ''����
			else
				if request.Form("frm_"&Str_Tmp)<>"" then 
					VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					VClass_Rs(Str_Tmp) = null
				end if	
			end if	
		next
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp?Act=Edit&CardID="&CardID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
	  if not MF_Check_Pop_TF("ME017") then Err_Show 'Ȩ���ж�
	  if request.Form("ActPut")=1 then 
	  ''��������
	  	   Put_Pwd_Type = NoSqlHack(request.Form("Put_Pwd_Type"))
		   Put_Pwd_Type = replace(Put_Pwd_Type," ","")
		   ''�������빹�ɷ�ʽ�����֣���ĸ		   
		   PutNum=NoSqlHack(request.Form("Put_PutNum"))
		   if not isnumeric(PutNum) then PutNum = 1
		   CardAddStr=NoSqlHack(Trim(request.Form("Put_CardAddStr")))
		   CardNum_Len=NoSqlHack(request.Form("Put_CardNum_Len"))
		   if not isnumeric(CardNum_Len) then CardNum_Len = 12
		   CardPwd_Len=NoSqlHack(request.Form("Put_CardPwd_Len"))
		   if not isnumeric(CardPwd_Len) then CardPwd_Len = 8
		   ''���ǰ׺���ȱȳ�ֵ������-2����,���ֵ������=��ֵ������+ǰ׺���ȣ���
		   if len(CardAddStr) > CardNum_Len-2 then CardNum_Len = len(CardAddStr) + CardNum_Len + 2
		 On Error Resume Next
		   if User_Conn.execute("select Count(*) from FS_ME_CardPut")(0)=0 then 
		   	User_Conn.execute("insert into FS_ME_CardPut (PutNum,CardAddStr,CardNum_Len,CardPwd_Len) VALUES ("&PutNum&",'"&CardAddStr&"',"&CardNum_Len&","&CardPwd_Len&")")
		   else
		   	User_Conn.execute("update FS_ME_CardPut set PutNum="&PutNum&",CardAddStr='"&CardAddStr&"',CardNum_Len="&CardNum_Len&",CardPwd_Len="&CardPwd_Len&"")
		   end if
		   If Err.Number<>0 Then Err.Clear : response.Redirect("../Error.asp?ErrCodes=<li>�������ɿ��������ô���ʧ��.����ƴд�Ƿ���ȷ.</li>") : response.End()
		   Str_Tmp = "CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,AddTime,isBuy"
		   Arr_Tmp = split(Str_Tmp,",")	 
		 for Put_i = 1 to PutNum
		   Str_Tmp1 = "" :   Str_Tmp2 = "" :   Str_Tmp3 = ""	
		   ''���ɿ��ţ����벢���
		   ''''''''''''''��ӵ�FS_ME_CardPut��
		   Str_Tmp2=Str_Tmp2&GetRamCode(CardPwd_Len)
		   'response.Write("����"&Put_i&"�Σ�"&Str_Tmp2&"<br>")
		   ''''''''''''''''''''''''''''''''''
		   ''���ż��ܣ�ȡ��06-07-05
		   ''��һ��
		   Str_Tmp1 = Str_Tmp1 &GetRamCode( CardNum_Len - len(CardAddStr) )
		   '''''''''''''''''''''''���ɺ��ټ�����ݿ�,��������������һ��
		   do while CheckCardCF(CardAddStr & Str_Tmp1)<>""		   
			   Str_Tmp1 = GetRamCode( CardNum_Len - len(CardAddStr) )			 
		   loop
		   'response.Write("����"&Put_i&"�Σ�"&Str_Tmp1&"<br>")
		   '''''''''''''''''''''''
		   On Error Resume Next
		   VClass_Rs.AddNew		   
		   ''''''''''''''��ӵ�FS_ME_Card��
		   Str_Tmp3 = ""
		   VClass_Rs("CardNumber") = CardAddStr & Str_Tmp1
		   VClass_Rs("CardPasswords") = Encrypt(Str_Tmp2)
		   for each Str_Tmp3 in Arr_Tmp
				if request.Form("frm_"&Str_Tmp3)<>"" then 
					VClass_Rs(Str_Tmp3) = NoSqlHack(request.Form("frm_"&Str_Tmp3))
				else
					VClass_Rs(Str_Tmp3) = null
				end if	
				'response.Write(Str_Tmp3&":"&NoSqlHack(request.Form("frm_"&Str_Tmp3))&"<br>")				
		   next
		   VClass_Rs.update
		   'If Err.Number<>0 Then Err.Clear : response.Redirect("../Error.asp?ErrCodes=<li>�����ݴ���ʧ��.����ƴд�Ƿ���ȷ.</li>") : response.End()
		   'response.Write("Str_Tmp1:"&Str_Tmp1&"  "&"Str_Tmp2:"&Str_Tmp2&"  "&"Str_Tmp3:"&Str_Tmp3&"<br>")
	     next
 	     VClass_Rs.close 
		 'response.End()
		 response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp" )&"&ErrCodes=<li>��ϲ���������ɳɹ���</li>")		   
	  else
	  '''''''''''''''''''''  		
		if request.Form("AddMode")=1 then 
		  ''����
			if CheckCardCF(NoSqlHack(request.Form("frm_CardNumber")))<>"" then 
				response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>����:"&request.Form("frm_CardNumber")&"�Ѵ��ڡ�</li>")
				response.End()
			end if
		  VClass_Rs.AddNew
		  for each Str_Tmp in Arr_Tmp
			if Str_Tmp = "CardPasswords" then 
				VClass_Rs(Str_Tmp) = Encrypt(NoSqlHack(request.Form("frm_"&Str_Tmp))) ''����
			else
				if request.Form("frm_"&Str_Tmp)<>"" then 
					VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					VClass_Rs(Str_Tmp) = null
				end if	
			end if	
			'response.Write(NoSqlHack(request.Form("frm_"&Str_Tmp))&":"&Str_Tmp&"<br>")
    	  next
		  VClass_Rs.update
		  Put_i=1
		else
		  ''������Ӷ���
		  Str_Tmp1 = Trim(request.form("More_Mode_Area"))
		  if Str_Tmp1="" or instr(Str_Tmp1,"|")=0 or len(Str_Tmp1)<3 then 
		  	response.Redirect("../Error.asp?ErrCodes=�������ʱ�����ź����벻��Ϊ�ջ��ʽ����ȷ�򳤶Ȳ�����")	: response.End()	
		  else
		   ErrInfo = "" : Put_i=0 : iR=0  : Str_Tmp2 = ""
		   Str_Tmp = "CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,AddTime,isBuy"
		   Arr_Tmp = split(Str_Tmp,",")
		   if instr(Str_Tmp1,"&nbsp;") then Str_Tmp1 = left(Str_Tmp1,len(Str_Tmp1) - len("&nbsp;"))
		   Str_Tmp1 = replace(replace(Str_Tmp1,"&nbsp;",vbcr)," ","")
		   Arr_Tmp1 = split(Str_Tmp1,vbcr) ''ÿһ������
		   ErrInfo = ErrInfo & "<li>���ύ��: "&ubound(Arr_Tmp1)+1&" �ſ�.</li>" 
		   '''''''''''''''''''''�жϺ������ַ���
		   for iR = lbound(Arr_Tmp1) to ubound(Arr_Tmp1)
		  	  Arr_Tmp2=split(Arr_Tmp1(iR),"|") ''��ֿ��ź�����
			  if ubound(Arr_Tmp2)<>1 then 
				ErrInfo = ErrInfo & "<li>��������У�"&asc(Arr_Tmp1(iR))&"δ��⣬ԭ�򣺷ָ����|����������ȷ��</li>" 
			  else
				if CheckCardCF( Trim(Arr_Tmp2(0)) ) <> "" then 
					ErrInfo = ErrInfo & "<li>����: "&Trim(Arr_Tmp2(0))&" �ڿ����Ѵ���.δ��⡣</li>"
				else
					Str_Tmp2 = Str_Tmp2 & Trim(Arr_Tmp2(0)) & "|" & Arr_Tmp2(1) & vbcr
				end if
			  end if
		   next
		   
 		   Str_Tmp1 = Str_Tmp2  	      
		   
		  if Str_Tmp2 <> "" then 
		  
		   Arr_Tmp1 = split(Str_Tmp1,vbcr) ''�жϺ��ÿ������
		   for iR = lbound(Arr_Tmp1) to ubound(Arr_Tmp1) - 1
			   	Arr_Tmp2=split(Arr_Tmp1(iR),"|") ''��ֿ��ź�����
				''+=========================
				VClass_Rs.Addnew
				VClass_Rs("CardNumber") = Trim(Arr_Tmp2(0)) : VClass_Rs("CardPasswords") = Encrypt( Arr_Tmp2(1) )
		 		for each Str_Tmp in Arr_Tmp
					if request.Form("frm_"&Str_Tmp)<>"" then 
						VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
					else
						VClass_Rs(Str_Tmp) = null
					end if	
				next
				VClass_Rs.update
				''+=========================
				Put_i = Put_i + 1
		    next
			Str_Tmp2 = replace(Str_Tmp2,vbcr,"<br />&nbsp;&nbsp;&nbsp;")
			ErrInfo = ErrInfo & "<li>����� "&Put_i&" �ſ�.</li>" 
		   end if
		  end if
		end if
 	    VClass_Rs.close
		'response.End()
		if Put_i>0 then 
			response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode("User/Card.asp?Act=Add" )&"&ErrCodes="&server.URLEncode("<li>��ϲ�������ɹ���</li>��ϸ��Ϣ:"&ErrInfo))
		else
			response.Redirect("../Error.asp?ErrCodes="&server.URLEncode(ErrInfo))
		end if		
	  end if
	'''''''''''''''''''''  
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
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName):New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName+"csed");
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
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >��ֵ������</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="Card.asp?Act=View">������ҳ</a> | <a href="Card.asp?Act=Add">�½�</a> 
      | <a href="Card.asp?Act=Put">��������</a> | <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("IsUse=0"))%>">δʹ��</a> 
      | <a href="Card.asp?Act=View&Add_Sql=<%=Encrypt("IsUse=1")%>">��ʹ��</a> 
	  | 
	  <% If G_IS_SQL_User_DB=1 then%>
	  <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("datediff(s,CardOverdueTime,'"&DateValue(Now())&"')>0"))%>">�ѹ���</a> 
	  <%Else%>
	  <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("datediff('s',CardOverdueTime,'"&DateValue(Now())&"')>0"))%>">�ѹ���</a>
	  <%End if%>
      | <a href="Card.asp?Act=Search">��ѯ</a></td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
		View
	case "Add","Edit","Put"
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
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('CardNumber')" class="sd"><b>�����š�</b></a> <span id="Show_Oder_CardNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardMoney')" class="sd"><b>��ֵ</b></a> <span id="Show_Oder_CardMoney"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('CardDateNumber')" class="sd"><b>����</b></a> <span id="Show_Oder_CardDateNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardPoint')" class="sd"><b>����</b></a> <span id="Show_Oder_CardPoint"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('isBuy')" class="sd"><b>״̬</b></a> <span id="Show_Oder_isBuy"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardOverDueTime')" class="sd"><b>����ʱ��</b></a> <span id="Show_Oder_CardOverDueTime"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('UserNumber')" class="sd"><b>ʹ����</b></a> <span id="Show_Oder_UserNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserTime')" class="sd"><b>��ֵʱ��</b></a> <span id="Show_Oder_UserTime"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_Card( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim CardID,Bol_IsEdit,isuse,UserTime
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	CardID = request.QueryString("CardID")
	if CardID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��CardIDû���ṩ</li>") : response.End()
	VClass_Sql = "select CardID,CardNumber,CardPasswords,CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,UserTime,AddTime,isBuy from FS_ME_Card where CardID="&CardID
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
	isuse = VClass_Rs("IsUse")
	UserTime = VClass_Rs("UserTime")
	if UserTime<>"" then if isdate(UserTime) then UserTime = formatdatetime(UserTime,2)
else
	isuse = 0	
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return chkinput(this);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("�޸ĵ㿨��Ϣ<input type=""hidden"" name=""CardID"" value="""&VClass_Rs(0)&""">") else response.Write("�����㿨") end if%></td>
	</tr>
<!---------->
	<tr  class="hback" <%if Bol_IsEdit = True or request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>> 
      <td width="20%" align="right">��ӷ�ʽ</td>
      <td> 
        <input name="AddMode" type="radio" onClick="CNum.style.display='';CPwd.style.display='';More_Mode.style.display='none';frm_CardNumber.require='true';frm_CardPasswords.require='true';More_Mode_Area.require='false';" value="1" checked>
          ���ų�ֵ��		  
        <input type="radio" name="AddMode" value="0" onClick="CNum.style.display='none';CPwd.style.display='none';More_Mode.style.display='';frm_CardNumber.require='false';frm_CardPasswords.require='false';More_Mode_Area.require='true';"> 
          ������ӳ�ֵ��
      </td>
    </tr>
<!--����-->
	<tr  class="hback" id="More_Mode" style="display:none;"> 
      <td width="20%" align="right">��ʽ�ı�</td>
      <td> 
	  	<textarea name="More_Mode_Area" cols="45" rows="10" dataType="Require" min="5" max="10000" msg="���ȱ�����5-10000֮��"></textarea>
        <br>
        �ָ���Ϊ<span class="tx">|</span>���밴��ÿ��һ�ſ����س����У�ÿ�ſ���ʽΪ��<span class="tx">����|����</span>��
        <br>
		<span class="tx">�Զ�����:
        <input type="text" name="PutCardNum" size="10" value="10" maxlength="8" title="ֻ������" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
		<input type="button" value="Go" onClick="try{GetCodes(PutCardNum.value)} catch(e){return};"></span>
		</td>
    </tr>
<!--����-->
	<tr  class="hback" id="CNum" <%if request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>> 
      <td width="20%" align="right">����</td>
      <td> 
        <input type="text" name="frm_CardNumber" size="40" value="<%if Bol_IsEdit then response.Write( VClass_Rs(1) ) else if request.QueryString("Act")<>"Search" then response.Write(GetRamCode(14)) end if end if%>"<%if request.QueryString("Act")="Put" then response.Write(" require=""false"" ") else response.Write(" require=""true"" ") end if%> dataType="LimitB" min="2" max="30" msg="���ȱ�����2-30֮��">
        <span class="tx">�������Զ�����</span>
	  </td>
    </tr>
    <tr  class="hback" id="CPwd" <%if request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>>
      <td align="right">��������</td>
      <td>
	  <input type="text" name="frm_CardPasswords" size="40" value="<%if Bol_IsEdit then response.Write(Decrypt( VClass_Rs(2) )) else if request.QueryString("Act")<>"Search" then response.Write(GetRamCode(6)) end if end if%>"<%if request.QueryString("Act")="Put" then response.Write(" require=""false"" ") else response.Write(" require=""true"" ") end if%>dataType="LimitB" min="2" max="30" msg="���ȱ�����2-30֮��">
      </td>
    </tr>
<!---->	
<!----��������ʱ��ʾ------>
<%if request.QueryString("Act")="Put" then
Dim Put_Rs,Put_Sql,Bol_Put_IsEdit
Bol_Put_IsEdit = false
Put_Sql = "select top 1 PutNum,CardAddStr,CardNum_Len,CardPwd_Len from FS_ME_CardPut order by PutID desc"
''ȡ���һ�ε�����
set Put_Rs=User_Conn.execute(Put_Sql)
if not Put_Rs.eof then Bol_Put_IsEdit = true
%>
	<tr  class="hback"> 
      <td width="20%" align="right">��������</td>
      <td> 
	    <input type="hidden" name="ActPut" value="1">
		<input type="text" name="Put_PutNum" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(0)) else response.Write("1") end if%>" dataType="Compare" msg="����>=0" to="0" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">��ֵ����ǰ׺</td>
      <td> 
		<input type="text" name="Put_CardAddStr" onKeyUp="value=value.replace(/[^a-zA-Z0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9]/g,''))" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(1)) else response.Write("FS2006") end if%>" datatype="LimitB" min="-1" max="20" msg="ǰ׺���ܳ���20���ַ�">
        Ĭ��ΪFS2006��Ϊ��</td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">��ֵ���ų���</td>
      <td> 
		<input type="text" name="Put_CardNum_Len" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(2)) else response.Write("12") end if%>" dataType="Compare" to="1" msg="����������[ǰ׺��+2,30]����" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">��ֵ�����볤��</td>
      <td> 
		<input type="text" name="Put_CardPwd_Len" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(3)) else response.Write("8") end if%>" dataType="Compare" msg="����>=0" to="0" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">�����빹�ɷ�ʽ</td>
      <td> 
        <input name="Put_Pwd_Type" type="checkbox" value="1" checked<%'if Bol_Put_IsEdit then if instr(Put_Rs(4),"1")>0 then response.Write(" checked") end if else response.Write(" checked") end if%>>
        ����
        <input type="checkbox" name="Put_Pwd_Type" value="2" checked<%'if Bol_Put_IsEdit then if instr(Put_Rs(4),"2")>0 then response.Write(" checked") end if end if%>>
        ��ĸ
	  </td>
    </tr>
<%
Put_Rs.close
set Put_Rs=nothing
end if%>	
<!----��������ʱ��ʾ����------>
    <tr  class="hback"> 
      <td align="right">�㿨��ֵ</td>
      <td>
	  <input type="text" name="frm_CardMoney" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
        Ԫ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨����</td>
      <td>
	  <input type="text" name="frm_CardDateNumber" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(4)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨����</td>
      <td>
	  <input type="text" name="frm_CardPoint" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(5)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right"><span class="tx"><strong>�㿨����ʱ��</strong></span></td>
      <td>
	  <input type="text" class="tx" name="frm_CardOverDueTime" size="27" value="<%if Bol_IsEdit then response.Write(VClass_Rs(6)) else response.Write( dateadd("d",30,date()) ) end if%>" readonly> <input name="SelectDate" type="button" id="SelectDate" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_CardOverDueTime);">	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨�Ƿ���ʹ��</td>
      <td>
	  <select name="frm_IsUse" id="frm_IsUse" onChange="if(this.options[this.selectedIndex].value='0'){frm_UserNumber.value='';frm_UserTime.value=''}">
		<%=PrintOption(isuse,"1:��ʹ��,0:δʹ��")%>
      </select>
        ���ֻ��������,�뱣��δʹ��</td>
    </tr>

    <tr  class="hback"> 
      <td align="right">�㿨ӵ���߱��</td>
      <td>
	  <input type="text" name="frm_UserNumber" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(8)) end if%>">
        ���ֻ��������,������</td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">��ֵ����</td>
      <td>
	  <input type="text" name="frm_UserTime" size="40" value="<%=UserTime%>">
        ���ֻ��������,������</td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨�������</td>
      <td>
	  <input type="text" name="frm_AddTime" size="27" value="<%if Bol_IsEdit then response.Write(VClass_Rs(10)) else response.Write(date()) end if%>" DataType="Require" msg="������ڱ�����д" readonly> <input name="SelectDate5" type="button" id="SelectDate5" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);"></td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ƿ�����</td>
      <td>
        <input type="radio" name="frm_isBuy" value="1" <%if Bol_IsEdit then if VClass_Rs(11)=1 then response.Write(" checked ") end if end if%>>
          ������		  
        <input type="radio" name="frm_isBuy" value="0" <%if Bol_IsEdit then if VClass_Rs(11)=0 then response.Write(" checked ") end if else response.Write(" checked ") end if%>>
          δ����
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" ���� " /> <!--<%IF request.QueryString("Act")="Put" then%> onClick="Put_CardNum_Len.to = (Put_CardAddStr.value.length+2).toString();Put_CardNum_Len.msg='���ȱ�����ڵ���'+(Put_CardAddStr.value.length+2).toString()" <%end if%>-->
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("Card.asp?Act=Del&CardID="&CardID)%>'">
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
  <form name="form1" method="post" action="?Act=SearchGo">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >��ѯ�㿨</td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">����</td>
      <td> <input type="text" name="frm_CardNumber" size="40" value="">
        ģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��������</td>
      <td> <input type="password" name="frm_CardPasswords" size="40" value="">
        ׼ȷ��ѯ </td>
    <tr  class="hback"> 
      <td align="right">�㿨��ֵ</td>
      <td> <select name="CardMoneyDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardMoney" size="30" value="">
      ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨����</td>
      <td> <select name="CardDateNumberDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardDateNumber" size="30" value="">
      ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨����</td>
      <td> <select name="CardPointDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardPoint" size="30" value="">
      ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨����ʱ��</td>
      <td>
        <!-- dataType="Date" msg="���ڸ�ʽ����" require="false" -->
         <select name="CardOverDueTimeDate" style="width:55">
           <option value="" selected="selected"></option>
           <option value="*">*</option>
           <option value="&gt;">&gt;</option>
           <option value="&lt;">&lt;</option>
           <option value="=">=</option>
           <option value="&gt;=">&gt;=</option>
           <option value="&lt;=">&lt;=</option>
           <option value="&lt;&gt;">&lt;&gt;</option>
         </select>
         <input type="text" name="frm_CardOverDueTime" size="17" value="" readonly>
        <input name="SelectDate4" type="button" id="SelectDate4" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_CardOverDueTime);">
        ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨�Ƿ���ʹ��</td>
      <td> <input type="radio" name="frm_IsUse"  value="1">
        ��ʹ�� 
        <input type="radio" name="frm_IsUse"  value="0">
        δʹ�� </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨ӵ���߱��</td>
      <td> <input type="text" name="frm_UserNumber" size="40" value="">
        ģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��ֵ����</td>
      <td>  <select name="UserTimeDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_UserTime" size="17" value="" readonly=>
        <input name="SelectDate2" type="button" id="SelectDate2" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_UserTime);">
        ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�㿨�������</td>
      <td> <select name="AddTimeDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_AddTime" size="17" value="" readonly>
        <input name="SelectDate3" type="button" id="SelectDate3" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);">
        ����,���ڿ�ͷ���ϼ򵥱ȽϷ���,*�ű�ʾģ����ѯ </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ƿ�����</td>
      <td> <input type="radio" name="frm_isBuy" value="1">
        ������ 
        <input type="radio" name="frm_isBuy" value="0">
        δ���� </td>
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
function GetCodes(number)
{
	(isNaN(number))?10:number;
	new Ajax.Updater('More_Mode_Area', 'Card_ajax.asp?no-cache='+Math.random() , {method: 'get',parameters:'number='+number});	
}

function chkinput(form)
{
	var value_=form.frm_CardMoney.value+form.frm_CardDateNumber.value+form.frm_CardPoint.value;
	if (value_=='')
		{alert('�㿨��ֵ,�㿨����,�㿨��������������дһ��ұ��������֡�');form.frm_CardMoney.focus();return false;}
	else
		return Validator.Validate(form,3); 
}
-->
</script>

</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 






