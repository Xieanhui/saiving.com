<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/NS_Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<!--#include file="../FS_Inc/Cls_SysConfig.asp"-->
<!--#include file="../FS_Inc/md5.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,NS_DefineTable_RS_obj,VClass_Sql,sRootDir,str_CurrPath
Dim DefineAllRs,AllDefineNum
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn
'session�ж�
MF_Session_TF
if not MF_Check_Pop_TF("MF_Define") then Err_Show
'---2007-02-12 By Ken
Dim MaxDefineNum,GetSysConfigObj
Set GetSysConfigObj = New Cls_SysConfig
GetSysConfigObj.getSysParam()
MaxDefineNum = Clng(GetSysConfigObj.Define_MaxNum)
'-----------

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
'******************************************************************
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

Function morestr(str,length)
	if isnull(str) then morestr="":exit function
	if len(str)>length then 
		morestr = left(str,length)&" ..."
	else
		morestr = str
	end if
End Function
  
Function Get_While_Info(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp  ,int_Tmp_i,New_ClassID,ClassID,D_Value
	Str_Tmp = "DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_MF_DefineTable"
	if NoSqlHack(request.QueryString("Act"))="SearchGo" then 
	for int_Tmp_i = 4 to 1 step -1	
		New_ClassID = NoSqlHack(request.Form("vclass"&int_Tmp_i))
		if New_ClassID<>"" then exit for
	next
	if New_ClassID="[ChangeToTop]" then New_ClassID=0
	if New_ClassID<>"" then 
		ClassID = NoSqlHack(New_ClassID)
	else
		ClassID = NoSqlHack(request.Form("frm_ClassID"))	
	end if
	if ClassID<>"" then New_Search_Str = and_where( New_Search_Str ) &" ClassID = "&ClassID
	
	D_Value = request.Form("frm_D_Value_1")
	if D_Value="" then D_Value = NoSqlHack(request.Form("frm_D_Value"))
	if D_Value<>"" then New_Search_Str = and_where( New_Search_Str ) &" D_Value = '"&D_Value&"'"
	
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			if NoSqlHack(Trim(request("frm_"&Str_Tmp)))<>"" then 
				Req_Str = NoSqlHack(Trim(request("frm_"&Str_Tmp)))
				select case Str_Tmp
					case "DefineID","ClassID","D_Type","D_isNull"
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
	if Add_Sql<>"" then
		Add_Sql = Decrypt(Add_Sql)
		 This_Fun_Sql = and_where(This_Fun_Sql) &" "& Add_Sql
	end if 
	if orderby<>"" then 
		This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	else
		This_Fun_Sql = This_Fun_Sql &"  Order By DefineID desc"
	end if		
	'response.Write(This_Fun_Sql)
	'response.End()
	Str_Tmp = ""
	On Error Resume Next
	Set NS_DefineTable_RS_obj = CreateObject(G_FS_RS)
	NS_DefineTable_RS_obj.Open This_Fun_Sql,Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("error.asp?ErrCodes=<li>��ѯ����"&Err.Description&"</li><li>�����ֶ������Ƿ�ƥ��.</li>")
		response.End()
	end if
	IF  NS_DefineTable_RS_obj.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>��������.</td></tr>") 
	else	
	NS_DefineTable_RS_obj.PageSize=int_RPP
	cPageNo=CintStr(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>NS_DefineTable_RS_obj.PageCount Then cPageNo=NS_DefineTable_RS_obj.PageCount 
	NS_DefineTable_RS_obj.AbsolutePage=cPageNo

	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""DefineTable_Info_Manage.asp?Act=Edit&DefineID="&NS_DefineTable_RS_obj("DefineID")&""" class=""otherset"" title='����޸�'>"&NS_DefineTable_RS_obj("D_Name")&"</a></td>" & vbcrlf
		for ii=3 to 8
			select case ii
				case 4 
					select case NS_DefineTable_RS_obj(ii) 
						case 0
						Str_Tmp="�����ı�"
						case 1
						Str_Tmp="�����ı�"
						case 2
						Str_Tmp="��ѡ"
						case 3
						Str_Tmp="��ѡ"
						case 4
						Str_Tmp="�����б�"
						case 5
						Str_Tmp="��������"
						case 6
						Str_Tmp="����"
						case 7
						Str_Tmp="ͼƬ����"
						case 8
						Str_Tmp="��������"
						case 9
						Str_Tmp="�����ʼ�"
						case 10
						Str_Tmp="�����ı�(���༭��)"
					end select
					Str_Tmp = "<a href=""DefineTable_Info_Manage.asp?Act=View&Add_Sql="& server.URLEncode(Encrypt( "D_Type="&NS_DefineTable_RS_obj(ii) ))&""" title=""����鿴ͬ��"">"&Str_Tmp&"</a>"
				case 5
					if NS_DefineTable_RS_obj(ii)=1 then 
						Str_Tmp = "<input  type=""checkbox"" name=""Other_D_isNull"" title=""�����Ϊ�ɿ�"" checked onclick=""javascript:location='DefineTable_Info_Manage.asp?Act=OtherSet&SetSql="& server.URLEncode(Encrypt( "D_isNull=0" ))&"&DefineID="&NS_DefineTable_RS_obj("DefineID") &"'"">"
					else
						Str_Tmp = "<input  type=""checkbox"" name=""Other_D_isNull"" title=""�����Ϊ����"" onclick=""javascript:location='DefineTable_Info_Manage.asp?Act=OtherSet&SetSql="& server.URLEncode( Encrypt("D_isNull=1") )&"&DefineID="&NS_DefineTable_RS_obj("DefineID") &"'"">"
					end if
				case 6
					Str_Tmp = morestr(NS_DefineTable_RS_obj(ii),30)
					if Str_Tmp<>"" then Str_Tmp = "<span style=""cursor:help"" title="""&NS_DefineTable_RS_obj(ii)&""">"&server.HTMLEncode( replace(Str_Tmp,vbcrlf,"<br />"))&"</span>"
				case 8
					select case NS_DefineTable_RS_obj(ii) 
						case "NS"
						Str_Tmp="����"
						case "MS"
						Str_Tmp="�̳�"
						case "DS"
						Str_Tmp="����"
						'case "SD"
'						Str_Tmp="����"
'						case "HS"
'						Str_Tmp="����"
'						case "AP"
'						Str_Tmp="�˲�"
						case else
						Str_Tmp="[δ֪]"
					end select
					Str_Tmp = "<a href=""DefineTable_Info_Manage.asp?Act=View&Add_Sql="& server.URLEncode(Encrypt( "D_SubType='"&NS_DefineTable_RS_obj(ii)&"'" ))&""" title=""����鿴ͬ��"">"&Str_Tmp&"</a>"
				case else
					Str_Tmp = NS_DefineTable_RS_obj(ii)
			end select		
			Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		next
		Str_Tmp = NS_Fun_ExecSql("select DefineName from FS_MF_DefineTableClass where DefineID = "&NS_DefineTable_RS_obj("ClassID"),"[δ����]")
		Str_Tmp = "<a href=""DefineTable_Info_Manage.asp?Act=View&Add_Sql="& server.URLEncode(Encrypt( "ClassID="&NS_DefineTable_RS_obj("ClassID") ))&""" title=""����鿴ͬ��"">"&Str_Tmp&"</a>"
		Get_Html = Get_Html & "<td align=""center"">"&Str_Tmp&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox""  name=""DefineID"" id=""DefineID"" value="""&NS_DefineTable_RS_obj("DefineID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		NS_DefineTable_RS_obj.MoveNext
 		if NS_DefineTable_RS_obj.eof or NS_DefineTable_RS_obj.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(NS_DefineTable_RS_obj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" ɾ�� "" onclick=""javascript:return confirm('ȷ��Ҫɾ����ѡ��Ŀ��?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	NS_DefineTable_RS_obj.close
	Get_While_Info = Get_Html
End Function

Function NS_Fun_ExecSql(This_Fun_Sql,Def_Info)
	Dim Ns_This_Fun_Rs
	Set Ns_This_Fun_Rs = Conn.execute(This_Fun_Sql)
	If not Ns_This_Fun_Rs.eof then 
		NS_Fun_ExecSql = Ns_This_Fun_Rs(0)
	Else
		NS_Fun_ExecSql = Def_Info
	End if
	Ns_This_Fun_Rs.close
	set Ns_This_Fun_Rs=nothing
End Function

Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("error.asp?ErrCodes=<li>��Ǹ,�����Sql���������.�����ֶβ�����.</li>")
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

Sub OtherSet()
	Dim SetSql
	SetSql = NoSqlHack(trim(Decrypt(request.QueryString("SetSql"))))
	if SetSql<>"" then 
		SetSql = "Update FS_MF_DefineTable set "& SetSql &" where DefineID='"&NoSqlHack(request.QueryString("DefineID"))&"'"
		Conn.execute( SetSql )
		response.Redirect("DefineTable_Info_Manage.asp")	
	end if
End Sub

Sub Del()
	if not MF_Check_Pop_TF("MF022") then Err_Show
	Dim Str_Tmp
	if request.QueryString("DefineID")<>"" then 
		Conn.execute("Delete from FS_MF_DefineTable where DefineID = "&NoSqlHack(request.QueryString("DefineID")))
	else
		Str_Tmp = FormatIntArr(request.form("DefineID"))
		if Str_Tmp="" then response.Redirect("error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>")
		
		Conn.execute("Delete from FS_MF_DefineTable where DefineID in ("&FormatIntArr(Str_Tmp)&")")
	end if
	response.Redirect("Success.asp?ErrorUrl="&server.URLEncode( "DefineTable_Info_Manage.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================

Sub Save()
	Dim Str_Tmp,Arr_Tmp,DefineID,int_Tmp_i,ClassID,New_ClassID,D_Value,ArrTmp1
	Str_Tmp = "D_Name,D_Coul,D_Type,D_isNull,D_Content,D_SubType"
	Arr_Tmp = split(Str_Tmp,",")
	DefineID = NoSqlHack(request.Form("DefineID"))
	if request.Form("frm_D_Name")="" or request.Form("frm_D_Coul")="" or request.Form("frm_D_Content")="" then 
		response.Redirect("error.asp?ErrCodes=<li>�ֶ���������Ӣ�������ֶ�˵������Ϊ�գ���</li>") 
		response.End()					
	end if
	for int_Tmp_i = 4 to 1 step -1	
		New_ClassID = NoSqlHack(request.Form("vclass"&int_Tmp_i))
		if New_ClassID<>"" then exit for
	next
	if New_ClassID="[ChangeToTop]" then New_ClassID=0
	if New_ClassID<>"" then 
		ClassID = New_ClassID
	else
		ClassID = NoSqlHack(request.Form("frm_ClassID"))	
	end if
	if ClassID = "" then ClassID = 0	
	if Cint(Request.Form("frm_D_Type")) = 4 then 
		D_Value = NoSqlHack(request.Form("frm_D_Value_1"))
	  if D_Value="" then 
	  	D_Value = "����"&vbcrlf&"Ĭ��ֵ"
	  else	
		if instr(D_Value,vbcrlf)=0 or len(D_Value)<4 then 
			response.Redirect("error.asp?ErrCodes=<li>��ѡ������ʱ�����������лس�,�����ַ�����4�����ϡ�</li>") 
			response.End()
		else
			ArrTmp1 = split(D_Value,vbcrlf)
			for int_Tmp_i = lbound(ArrTmp1) to ubound(ArrTmp1)
				if isnull(ArrTmp1(int_Tmp_i)) or trim(ArrTmp1(int_Tmp_i)) = "" then 
					response.Redirect("error.asp?ErrCodes=<li>��ѡ������ʱ�����������лس�,����ÿһ�е����ݱ�����Ч.</li><li>��"&int_Tmp_i+1&"��������Ч.��Ϊ���лس�����Ϊ��!</li>") 
					response.End()					
					exit for 
				end if 
			next
		end if
	  end if	
	elseif CintStr(request.Form("frm_D_Type"))=5 Then
		D_Value = NoSqlHack(request.Form("frm_D_Value"))
		if D_Value = "" then 
			D_Value = 0
		else
			if not isnumeric(D_Value) then 
				response.Redirect("error.asp?ErrCodes=<li>Ĭ��ֵ���������֡�</li>") 
				response.End()			
			end if 
		end if
	elseif CintStr(request.Form("frm_D_Type")) = 6 then 
		D_Value = NoSqlHack(request.Form("frm_D_Value"))
		if D_Value = "" then 
			D_Value = now()
		else		
			if not isdate(D_Value) then 
				response.Redirect("error.asp?ErrCodes=<li>Ĭ��ֵ�����������͡�</li>") 
				response.End()			
			end if 
		end if
	else
		D_Value = NoSqlHack(request.Form("frm_D_Value"))					
	end if
	if D_Value="" then D_Value  = "����"
	if not isnumeric(DefineID) or DefineID = "" then DefineID = 0 
	'--------------------
	Set DefineAllRs = Server.CreateObject(G_FS_RS)
	DefineAllRs.Open "select * from FS_MF_DefineTable",Conn,1,1
	If DefineAllRs.Eof Then
		AllDefineNum = 0
	Else
		AllDefineNum = Clng(DefineAllRs.RecordCount)
	End If	
	'----------
	VClass_Sql = "select ClassID,D_Value,"&Str_Tmp&"  from FS_MF_DefineTable where DefineID="&DefineID
	Set NS_DefineTable_RS_obj = CreateObject(G_FS_RS)
	NS_DefineTable_RS_obj.Open VClass_Sql,Conn,3,3
	if DefineID > 0 then 
	''�޸�
		NS_DefineTable_RS_obj("ClassID") = ClassID
		NS_DefineTable_RS_obj("D_Value") = D_Value
		for each Str_Tmp in Arr_Tmp
			If Str_Tmp="D_isNull" And NoSqlHack(request.Form("frm_"&Str_Tmp))="" Then
				NS_DefineTable_RS_obj(Str_Tmp) = 0
			Else
				NS_DefineTable_RS_obj(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
			End If
		next
		NS_DefineTable_RS_obj.update
		NS_DefineTable_RS_obj.close
		response.Redirect("Success.asp?ErrorUrl="&server.URLEncode( "DefineTable_Info_Manage.asp?Act=Edit&DefineID="&DefineID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		If  Clng(AllDefineNum) >= Clng(MaxDefineNum) Then
			response.Redirect("error.asp?ErrCodes=<li>�Զ����ֶ����������ܳ���<strong style=""color:#FF0000;"">"&MaxDefineNum&"</strong>����</li>") 
			response.End()
		Else
			if Conn.execute("select count(*) from FS_MF_DefineTable where ClassID="&ClassID&" and D_Coul = '"&NoSqlHack(request.Form("frm_D_Coul"))&"'  and D_SubType = '"&NoSqlHack(request.Form("frm_D_SubType"))&"'")(0)>0 then 
				response.Redirect("error.asp?ErrCodes=<li>�Զ����ֶ���(Ӣ����)�����ظ���</li>") 
				response.End()
			end if
			NS_DefineTable_RS_obj.addnew
			NS_DefineTable_RS_obj("ClassID") = ClassID
			if D_Value<>"" then NS_DefineTable_RS_obj("D_Value") = D_Value
			for each Str_Tmp in Arr_Tmp
				If Str_Tmp="D_isNull" And NoSqlHack(request.Form("frm_"&Str_Tmp))="" Then
					NS_DefineTable_RS_obj(Str_Tmp) = 0
				Else
					NS_DefineTable_RS_obj(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				End If
			next	
			NS_DefineTable_RS_obj.update
			NS_DefineTable_RS_obj.close
			response.Redirect("Success.asp?ErrorUrl="&server.URLEncode( "DefineTable_Info_Manage.asp?Act=Add&ClassID="&ClassID ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
		End If	
	end if
End Sub
''=========================================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����ֶι���___Powered by foosun Inc.</title>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<%if instr(",Edit,Add,Search,",","&request.QueryString("Act")&",")>0 then%>
<script language="javascript" src="../FS_Inc/class_liandong.js" type="text/javascript"></script>
<%end if%>
<script language="JavaScript" src="../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/GetLettersByChinese.js"></script>
<script language="JavaScript" type="text/JavaScript">
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
function selectAll(f,NoSelName)
{
	for(i=0;i<f.length;i++)
	{
		if(f(i).type=="checkbox" && f(i)!=event.srcElement && f(i).name!=NoSelName)
		{
			f(i).checked=event.srcElement.checked;
		}
	}
}
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
-->
</script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" ><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br>  <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong>�Զ����ֶ�</strong></a></td>
  </tr>
  <tr  class="hback">
    <td class="hback" ><a href="DefineTable_Info_Manage.asp?Act=View" >������ҳ</a>&nbsp;|&nbsp;
      <a href="DefineTable_Info_Manage.asp?Act=Add" >����</a> &nbsp;|&nbsp;
      <a href="DefineTable_Info_Manage.asp?Act=Search">��ѯ</a> &nbsp;|&nbsp;
	  <select name="select_D_SubType" onChange='var jumpvalue = this.options[this.selectedIndex].value;location="DefineTable_Info_Manage.asp?Act=View&Add_Sql="+jumpvalue;'>
	  <option value="">�鿴</option>
	  <option value="">����</option>
	  <option value="<%=server.URLEncode(Encrypt("D_SubType = 'NS'"))%>">����</option>
	  <option value="<%=server.URLEncode(Encrypt("D_SubType = 'MS'"))%>">�̳�</option>
	  <option value="<%=server.URLEncode(Encrypt("D_SubType = 'DS'"))%>">����</option>
	 <!-- <option value="<%=server.URLEncode(Encrypt("D_SubType = 'SD'"))%>">����</option>
	  <option value="<%=server.URLEncode(Encrypt("D_SubType = 'HS'"))%>">����</option>
	  <option value="<%=server.URLEncode(Encrypt("D_SubType = 'AP'"))%>">�˲�</option>-->
	  </select> &nbsp;|&nbsp;	
      <a href="DefineTable_Manage.asp">�ֶη������</a></td>
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
	case "OtherSet"
	OtherSet
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="?Act=Del">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('D_Name')" class="sd"><b>���������ơ�</b></a> <span id="Show_Oder_D_Name" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('D_Coul')" class="sd"><b>�ֶ�����</b></a> <span id="Show_Oder_D_Coul" class="tx"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('D_Type')" class="sd"><b>�ֶ�����</b></a> <span id="Show_Oder_D_Type" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('D_isNull')" class="sd"><b>�Ƿ�ɿ�</b></a> <span id="Show_Oder_D_isNull" class="tx"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('D_Value')" class="sd"><b>Ĭ��ֵ</b></a> <span id="Show_Oder_D_Value" class="tx"></span></td>
	  <td align="center" class="xingmu">˵��</td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('D_SubType')" class="sd"><b>����ϵͳ</b></a> <span id="Show_Oder_D_SubType" class="tx"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('ClassID')" class="sd"><b>��������</b></a> <span id="Show_Oder_ClassID" class="tx"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form,'Other_D_isNull')" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim DefineID,Bol_IsEdit,Edit_ClassID,D_Type
Bol_IsEdit = false : Edit_ClassID =""
if request.QueryString("Act")="Edit" then 
	if not MF_Check_Pop_TF("MF021") then Err_Show
	DefineID = request.QueryString("DefineID")
	if DefineID="" then response.Redirect("error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��DefineIDû���ṩ</li>") : response.End()
	VClass_Sql = "select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from FS_MF_DefineTable where DefineID="&DefineID
	Set NS_DefineTable_RS_obj	= CreateObject(G_FS_RS)
	NS_DefineTable_RS_obj.Open VClass_Sql,Conn,1,1
	if NS_DefineTable_RS_obj.eof then response.Redirect("error.asp?ErrorUrl=&ErrCodes=<li>û����ص�����,��������Ѳ�����.</li>") : response.End()
	Bol_IsEdit = True
	Edit_ClassID=NS_DefineTable_RS_obj(1)
	D_Type=NS_DefineTable_RS_obj(4)
else
	if not MF_Check_Pop_TF("MF021") then Err_Show
	Edit_ClassID=request.QueryString("ClassID")
	D_Type=0
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" method="post" action="?Act=Save" onSubmit="return checkinput(this);">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("�޸��Զ����ֶ���Ϣ<input type=""hidden"" name=""DefineID"" value="""&NS_DefineTable_RS_obj(0)&""">") else response.Write("�����Զ����ֶ���Ϣ (�Զ����ֶ���಻�ܳ���<strong style=""color:#FF0000;"">" & MaxDefineNum & "</strong>��)") end if%></td>
	</tr>
    <tr  class="hback"> 
      <td align="right">�����ĸ�ϵͳ</td>
      <td>	  
	  <select name="frm_D_SubType" datatype="require" msg="�����ĸ�ϵͳ����ѡ��">
	  <option value="NS"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="NS" then response.Write(" selected") end if end if%>>����</option>
	  <option value="MS"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="MS" then response.Write(" selected") end if end if%>>�̳�</option>
	  <option value="DS"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="DS" then response.Write(" selected") end if end if%>>����</option>
	<!--  <option value="SD"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="SD" then response.Write(" selected") end if end if%>>����</option>
	  <option value="HS"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="HS" then response.Write(" selected") end if end if%>>����</option>
	   <option value="AP"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(8)="AP" then response.Write(" selected") end if end if%>>�˲�</option>-->
	  </select>	
	  <span class="tx">�����ֶ�Ӧ�����ĸ���ϵͳ�С�</span></td>
    </tr>

    <tr  class="hback"<%if Edit_ClassID="" then response.Write("style=display:none") end if%>> 
      <td align="right">�������</td>
      <td>
	  	<%if Edit_ClassID<>"" then 
			response.Write( NS_Fun_ExecSql("select DefineName from FS_MF_DefineTableClass where DefineID = "&Edit_ClassID,"[δ����]") )%>	
		<input type="hidden" name="frm_ClassID" id="frm_ClassID" value="<%=Edit_ClassID%>">
		<input type="hidden" name="frm_ClassName" id="frm_ClassName" value="<%=NS_Fun_ExecSql("select DefineName from FS_MF_DefineTableClass where DefineID = "&Edit_ClassID,"[δ����]")%>">
		<%end if%>
		<span class="tx">������������������������н���ѡ��</span>
	  </td>
    </tr>
	<tr class="hback"> 
      <td align="right">ѡ�����</td>
      <td width="596">
	<SELECT NAME="vclass1" ID="vclass1" style="width:100px" <%if request.QueryString("Act")="Add" then%> msg="������ѡ��"<%end if%>>
    	<OPTION></OPTION>
    </Select>
<!--		  
<!---�����˵���ʼ--- >	
	<SELECT NAME="vclass1" ID="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass2" ID="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass3" ID="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
        <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass4" ID="vclass4" style="width:100px">
    	<OPTION></OPTION>
    </SELECT>
<!---�����˵�����--- > -->		
  <span class="tx" id="vclass_Alt" style="color:#FF0000">ѡ����������ֶε����</span>		
     </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ���������</td>
      <td>
	  <input onBlur="<% if request.QueryString("Act")="Add" then %>SetClassEName(this.value,document.form_Save.frm_D_Coul);<% end if %>" type="text" name="frm_D_Name" size="40" value="<%if Bol_IsEdit then response.Write(NS_DefineTable_RS_obj(2)) end if%>" datatype="Require" msg="������д">
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�����Ӣ������</td>
      <td>
	  <input type="text" name="frm_D_Coul" size="40" value="<%if Bol_IsEdit then response.Write(NS_DefineTable_RS_obj(3)) end if%>" onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))"  datatype="Require" msg="������д">
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�����</td>
      <td>
		<select name="frm_D_Type" onChange="javascript:ChangeDefValueAreaType();">
			<%=PrintOption(D_Type,"0:�����ı�,1:�����ı�,4:�����б�,5:��������,6:��������,7:ͼƬ����,8:��������,9:�ʼ�����,10:�����ı����༭��")%>
		</select>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ƿ�����Ϊ��</td>
      <td>
		<input name="frm_D_isNull" type="checkbox" value="1"<%if Bol_IsEdit then if NS_DefineTable_RS_obj(5)=1 then response.Write("  checked") end if else response.Write("  checked") end if%>>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�Ĭ��ֵ</td>
      <td id="td_D_Value_Def" style="display:''">	  	
	  <input type="text" name="frm_D_Value" size="40" value="<%if Bol_IsEdit then response.Write(NS_DefineTable_RS_obj("D_Value")) end if%>">
	  <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�..." onClick="SelectFile();">
     </td>
      <td id="td_D_Value_1" style="display:'none'">	  	
 	  <textarea name="frm_D_Value_1" rows="10" cols="40"><%if Bol_IsEdit then response.Write(NS_DefineTable_RS_obj("D_Value")) end if%></textarea>
      <input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ���ļ�..." onClick="SelectFile();">	
	 </td> 
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ���˵��</td>
      <td>
	  <textarea name="frm_D_Content" cols="40" datatype="LimitB" require="true" min="1" max="200" msg="������[1-200]���ֽ��ڡ�"><%if Bol_IsEdit then response.Write(NS_DefineTable_RS_obj(7)) end if%></textarea>
	  <span class="tx">�ֶ���˵����ʹ�ù���˵����</span></td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" ���� "> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� ">
              &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��?')) location = '?Act=Del&DefineID=<%=request.QueryString("DefineID")%>'">
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
      <td colspan="3" align="left" class="xingmu" >�Զ����ֶ���Ϣ��ѯ</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ID��</td>
      <td> <input type="text" name="frm_DefineID" size="40" value=""> 
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">�����ĸ�ϵͳ</td>
      <td>	  
	  <select name="frm_D_SubType">
	  <option value="">����</option>
	  <option value="NS">����</option>
	  <option value="MS">�̳�</option>
	  <option value="DS">����</option>
	 <!-- <option value="SD">����</option>
	  <option value="HS">����</option>
	  <option value="AP">�˲�</option>-->
	  </select>	
	  <span class="tx">�����ֶ�Ӧ�����ĸ���ϵͳ�С�</span></td>
    </tr>

    <tr class="hback"> 
      <td align="right">ѡ�����</td>
      <td width="596"> 
	<SELECT NAME="vclass1" ID="vclass1" style="width:100px">
    	<OPTION></OPTION>
    </SELECT>
<!--		  
<!---�����˵���ʼ--- >	
	<SELECT NAME="vclass1" ID="vclass1" onBlur="javascript:RemoveChildopt(this,'vclass2,vclass3,vclass4');"  style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass2" ID="vclass2" onBlur="javascript:RemoveChildopt(this,'vclass3,vclass4');" style="width:100px">
         <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass3" ID="vclass3" onBlur="javascript:RemoveChildopt(this,'vclass4');" style="width:100px">
        <OPTION></OPTION>
    </SELECT>
	<SELECT NAME="vclass4" ID="vclass4" style="width:100px">
    	<OPTION></OPTION>
    </SELECT>
<!---�����˵�����--- > -->		
        </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ���������</td>
      <td> <input type="text" name="frm_D_Name" size="40" value=""> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�����Ӣ������</td>
      <td> <input type="text" name="frm_D_Coul" size="40" value=""> 
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�����</td>
      <td> <select name="frm_D_Type" onChange="javascript:ChangeDefValueAreaType();">
	      <option value="">��ѡ��</option>
          <option value="0">�����ı�</option>
          <option value="1">�����ı�</option>
          <option value="4">�����б�</option>
          <option value="5">��������</option>
          <option value="6">����</option>
          <option value="7">ͼƬ����</option>
          <option value="8">��������</option>
          <option value="9">�����ʼ�����</option>
          <option value="10">�����ı�(���༭��)</option>
        </select> </td>
    </tr>
    
    <tr  class="hback"> 
      <td align="right">�Ƿ�����Ϊ��</td>
      <td> <input name="frm_D_isNull" type="checkbox" value="1"> </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ�Ĭ��ֵ</td>
      <td id="td_D_Value_Def" style="display:''">	  	
	  <input type="text" name="frm_D_Value" size="40" value="">
     </td>
      <td id="td_D_Value_1" style="display:'none'">	  	
 	  <textarea name="frm_D_Value_1" rows="10" cols="40"></textarea>
     </td> 
    </tr>
    <tr  class="hback"> 
      <td align="right">�ֶ���˵��</td>
      <td> <textarea name="frm_D_Content" cols="40"></textarea> 
        <span class="tx">����Ϊ200���ַ�</span> </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" ִ�в�ѯ " /> 
              &nbsp; <input type="reset" name="ReSet" value=" ���� " /> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%End Sub%>
</body>
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

function ChangeDefValueAreaType()
{
	//����ѡ���ֵ�ı�Ĭ��ֵ��������
	//0:�����ı�,1:�����ı�,4:�����б�,5:��������,6:��������,7:ͼƬ����,8:��������,9:�ʼ�����,10:�����ı����༭��
	if(document.all.frm_D_Type.value=='4')
	{
		td_D_Value_Def.style.display='none';
		td_D_Value_1.style.display='';
	}
	else
	{
		td_D_Value_Def.style.display='';
		td_D_Value_1.style.display='none';
	}	
	if(document.all.frm_D_Type.value=='5')	
	{ //����
		document.all.frm_D_Value.require="false";
		document.all.frm_D_Value.dataType="Range";
		document.all.frm_D_Value.min="-32760";
		document.all.frm_D_Value.max="32760";
		document.all.frm_D_Value.msg="���ֱ�����-32760~32760֮�䡣";
	}	
	if(document.all.frm_D_Type.value=='6')	
	{ //����
		document.all.frm_D_Value.require="false";
		document.all.frm_D_Value.dataType="Date";
		document.all.frm_D_Value.format="ymd";
		document.all.frm_D_Value.msg="���ڸ�ʽ����ȷ��";
	}
	if(document.all.frm_D_Type.value=='0' || document.all.frm_D_Type.value=='1'|| document.all.frm_D_Type.value=='10')		
	{
		document.all.frm_D_Value.require="false";
		document.all.frm_D_Value.dataType="Require";
		document.all.frm_D_Value.msg="";
	}
}
function checkinput(obj)
{
	if (Validator.Validate(obj,3) == true)
	{
		if (obj.frm_D_Type.value=='4')
		///�����ı��������˵�,��Ҫһ��һ����Ŀ,�س����С�
		{
			var txt=obj.frm_D_Value_1.value;
			if  (txt=='') return true;
			if (!controlrow(txt))
			{
				alert('��ѡ�����,����ʱ,Ĭ��ֵ������д,����һ��һ����¼.�س����С�\n�����������û�лس�����,�򳬹�50��.���ַ�̫��.');
				obj.frm_D_Value_1.focus();
				return false;
			}
		}
	}
	else
		return Validator.Validate(obj,3);	
}

function controlrow(txt)   
{
	  var count=0;   
	  var index=txt.indexOf("\r");   
	  while(index!=-1)   
	  {   	      
		  count++; 
		  index=txt.indexOf("\r",index+1);   
	  }   
	  if(count<1 || txt.length<4 || count>50)   
		  return false;
	  else
	  	  return true;	  
}   

function SelectFile()     
{
//
 var returnvalue = OpenWindow('CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window);
 if (returnvalue!='')
 {
 	var obj=event.srcElement.parentNode.firstChild;
	(obj.name.indexOf('_1')>-1)?obj.value+=returnvalue:obj.value=returnvalue;
 }
}

-->
</script>
<%if instr(",Edit,Add,Search,",","&request.QueryString("Act")&",")>0 then%>
<script language="javascript">
<!-- 

ChangeDefValueAreaType();

//awen created
//�����˵�---�Զ����ֶ����   ���4��   --start 
//���ݸ�ʽ ID������ID������
var array=new Array();
<%dim NS_JS_Sql,NS_JS_RS,NS_JS_i
  if request.QueryString("ClassID")<>"" then 
  	NS_JS_Sql="select DefineID,ParentID,DefineName from FS_MF_DefineTableClass where DefineID<>"&NoSqlHack(request.QueryString("ClassID"))
  else
    NS_JS_Sql="select DefineID,ParentID,DefineName from FS_MF_DefineTableClass "	
  end if	
  set NS_JS_RS=Conn.execute(NS_JS_Sql)
  NS_JS_i=0
  do while not NS_JS_RS.eof
%>
array[<%=NS_JS_i%>]=new Array("<%=NS_JS_RS("DefineID")%>","<%=NS_JS_RS("ParentID")%>","<%=NS_JS_RS("DefineName")%>"); 
<%
	NS_JS_RS.movenext
	NS_JS_i=NS_JS_i+1
loop
NS_JS_RS.close
%>

var liandong=new CLASS_LIANDONG_YAO(array)
liandong.firstSelectChange("0","vclass1");
liandong.close

document.getElementById('vclass1').options.remove(1);
/*
liandong.subSelectChange("vclass1","vclass2");
liandong.subSelectChange("vclass2","vclass3");
liandong.subSelectChange("vclass3","vclass4");

//---------------------------������������������
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
*/
//end 
-->
</script>
<%end if%>
<%
Set NS_DefineTable_RS_obj=nothing
Conn.close
Set Conn=nothing
%>
</html>