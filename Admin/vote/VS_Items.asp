<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/VS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,VS_Rs,VS_Sql,VS_Rs1 ,sRootDir,str_CurrPath
Dim AutoDelete,Months
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn 
MF_Session_TF
if not MF_Check_Pop_TF("VS003") then Err_Show

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
	Str_Tmp = "IID,TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_VS_Items"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "IID","TID","ItemMode","VoteCount"
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
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
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
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Items.asp?Act=Edit&IID="&VS_Rs("IID")&""" title=""����޸Ļ�鿴��ϸ"">"&VS_Rs("ItemName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&VS_Rs("IID")&".style.display=='') TD_U_"&VS_Rs("IID")&".style.display='none'; else {TD_U_"&VS_Rs("IID")&".style.display='';ReImgSize('TD_Img_"&VS_Rs("IID")&"');}"" title='����鿴��ϸ���'>"&Get_OtherTable_Value("select Theme from FS_VS_Theme where TID ="&VS_Rs("TID"))&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("ItemMode"),"1:��������ģʽ,2:<span class=tx>������дģʽ</span>,3:<b>ͼƬģʽ</b>")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("PicSrc")&"</td>" & vbcrlf
		Get_Html = Get_Html & Replacestr(VS_Rs("DisColor"),":<td align=""center"">��</td>,else:<td align=""center"" bgcolor="""&VS_Rs("DisColor")&""">"&VS_Rs("DisColor")&"</td>") & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("VoteCount")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("ItemDetail"),":��,else:"&left(VS_Rs("ItemDetail"),80)&"...")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""IID"" id=""IID"" value="""&VS_Rs("IID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++�㿪ʱ��ʾ��ϸ��Ϣ��
		set VS_Rs1 = Conn.execute("select TID,CID,Theme,Type,DisMode,StartDate,EndDate,ItemMOde from FS_VS_Theme where TID ="&VS_Rs("TID"))
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& VS_Rs("IID") &""" style=""display:'none'""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td colspan=3>ͶƱ����:"&Get_OtherTable_Value("select ClassName from FS_VS_Class where CID ="&VS_Rs1("CID"))& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td colspan=3>��������:"&Get_OtherTable_Value("select Description from FS_VS_Class where CID ="&VS_Rs1("CID"))& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>��������:"&VS_Rs1("Theme")&"</td><td>��Ŀ����:"&Replacestr(VS_Rs1("Type"),"1:��ѡ,2:��ѡ,3:�ಽ")&"</td><td>��ʾ��ʽ:"&Replacestr(VS_Rs1("DisMode"),"1:ֱ��ͼ,2:��ͼ,3:����ͼ")&"</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>���з�ʽ:"&Replacestr(VS_Rs1("ItemMOde"),":��,0:��������,1:1ѡ��/��(����),2:2ѡ��/��,3:3ѡ��/��,4:4ѡ��/��,5:5ѡ��/��,6:6ѡ��/��,7:7ѡ��/��,8:8ѡ��/��,9:9ѡ��/��,10:10ѡ��/��,11:11ѡ��/��,12:12ѡ��/��")&"</td><td>��ʼʱ��:"&VS_Rs1("StartDate")&"</td><td>����ʱ��:"&VS_Rs1("EndDate")&"</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		VS_Rs1.close
		''+++++++++++++++++++++++++++++++++++++++		
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
	if request.QueryString("IID")<>"" then 
		Conn.execute("Delete from FS_VS_Items where IID = "&CintStr(request.QueryString("IID")))
		Conn.execute("Delete from FS_VS_Items_Result where IID = "&CintStr(request.QueryString("IID")))
	else
		Str_Tmp = FormatIntArr(request.form("IID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>���������ѡ��һ������ɾ����</li>"):response.End()
		
		Conn.execute("Delete from FS_VS_Items where IID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Items_Result where IID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=View" )&"&ErrCodes=<li>��ϲ��ɾ���ɹ���</li>")
End Sub
''================================================================
Sub Save()
	Dim Str_Tmp,Arr_Tmp,IID
	Str_Tmp = "TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail"
	Arr_Tmp = split(Str_Tmp,",")
	IID = NoSqlHack(request.Form("IID"))	
	if not isnumeric(IID) or not IID<>"" then IID = 0
	VS_Sql = "select "&Str_Tmp&"  from FS_VS_Items  where IID = "&CintStr(IID)
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,3,3
	if not VS_Rs.eof then 
	''�޸�
		for each Str_Tmp in Arr_Tmp
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=Edit&IID="&IID )&"&ErrCodes=<li>��ϲ���޸ĳɹ���</li>")
	else
	''����
		if Conn.execute("Select Count(*) from FS_VS_Items where ItemName='"&NoSqlHack(request.Form("ItemName"))&"' and TID = "&NoSqlHack(request.Form("TID")))(0)>0 then 
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
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=Add&VoteCount="&request.form("VoteCount")&"&TID="&request.form("TID")&"&ItemValue="&request.form("ItemValue") ) &"&ErrCodes=<li>��ϲ�������ɹ���</li>")
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
<iframe width="260" height="165" id="colorPalette" src="../CommPages/selcolor.htm" style="visibility:hidden; position: absolute;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            
    <td colspan="10" align="left" class="xingmu" >ͶƱѡ�����</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="VS_Items.asp">������ҳ</a> | <a href="VS_Items.asp?Act=Add">����</a> | <a href="VS_Items.asp?Act=Search" title="���ֺ������͵��ֶ�,֧��<=<>=><>�ȵ����������:���������>2 ; ��������֧�� A B ,A* *B ,*A* *B* ,AB��ģʽ.">��ѯ</a>	
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
Sub View
if not MF_Check_Pop_TF("VS003") then Err_Show
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="?Act=Del">
   <tr  class="hback"> 
      <td align="center" class="xingmu"><a href="javascript:OrderByName('ItemName')" class="sd"><b>ѡ������</b></a> 
        <span id="Show_Oder_ItemName"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('TID')" class="sd"><b>��������</b></a> 
        <span id="Show_Oder_TID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ItemMode')" class="sd"><b>ѡ��ģʽ</b></a> 
        <span id="Show_Oder_ItemMode"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('PicSrc')" class="sd"><b>ͼƬλ��</b></a> 
        <span id="Show_Oder_PicSrc"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('DisColor')" class="sd"><b>��ʾ��ɫ</b></a> 
        <span id="Show_Oder_DisColor"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('VoteCount')" class="sd"><b>Ʊ��</b></a> 
        <span id="Show_Oder_VoteCount"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ItemDetail')" class="sd"><b>ѡ��˵��</b></a> 
        <span id="Show_Oder_ItemDetail"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
   </form>	
</table>
<%End Sub
Sub Add_Edit_Search()
Dim Bol_IsEdit,IID,TID,ItemValue,ItemMode,DisColor,VoteCount
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then
if not MF_Check_Pop_TF("VS002") then Err_Show
	IID = request.QueryString("IID")
	if IID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>��Ҫ��IIDû���ṩ��</li>") : response.End()
	VS_Sql = "select IID,TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail from FS_VS_Items where IID = "& CintStr(IID)
	Set VS_Rs	= CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,1,1
	if not VS_Rs.eof then 
		Bol_IsEdit = True
		TID = VS_Rs("TID")
		ItemValue = VS_Rs("ItemValue")
		ItemMode = VS_Rs("ItemMode")
		DisColor = VS_Rs("DisColor")
		VoteCount = VS_Rs("VoteCount")
	end if
elseif request.QueryString("Act") = "Add" then 
	if not MF_Check_Pop_TF("VS002") then Err_Show
	TID = NoSqlHack(request.QueryString("TID"))
	ItemValue = NoSqlHack(request.QueryString("ItemValue"))
	if ItemValue = "" then	ItemValue = "1-9"
	ItemMode = 1
	DisColor = ""
	VoteCount = NoSqlHack(request.QueryString("VoteCount"))
	if VoteCount = "" then 
		randomize		
		VoteCount = CStr(Int((99* Rnd) + 1))
	end if	
end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" <%if request.QueryString("Act")<>"Search" then response.Write("action=""?Act=Save"" onsubmit=""return chkinput();""") else response.Write("action=""?Act=SearchGo""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >ͶƱѡ����Ϣ<%if Bol_IsEdit then	 response.Write("<input type=""hidden"" name=""IID"" id=""IID"" value="""&VS_Rs("IID")&""">") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>

    <tr class="hback"> 
      <td width="100" align="right">�Զ����</td>
      <td>
	  	<input type="text" name="IID" id="IID" size="11" maxlength="11">
      </td>
    </tr>
<%end if%>
    <tr  class="hback"> 
      <td align="right">����ͶƱ</td>
      <td>
		<select name="TID" id="TID" onChange="Do.these('TID',function(){return isEmpty('TID','TID_Alt')})">
		<option value="">��ѡ��</option>
		<%=Get_FildValue_List("select TID,'����:'+ClassName+'--����:'+Theme from FS_VS_Theme A,FS_VS_Class B where A.CID=B.CID",NoSqlHack(TID),1)%>
		</select>
		<span id="TID_Alt"></span>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ѡ������</td>
      <td>
		<input type="text" name="ItemName" id="ItemName" size="50" maxlength="100" onFocus="Do.these('ItemName',function(){return isEmpty('ItemName','ItemName_Alt')})" onKeyUp="Do.these('ItemName',function(){return isEmpty('ItemName','ItemName_Alt')})" value="<%if Bol_IsEdit then response.Write(VS_Rs("ItemName")) end if%>">
		<span id="ItemName_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��Ŀ����</td>
      <td>
		<select name="ItemValue" id="ItemValue">
		<%=PrintOption(ItemValue,":��ѡ��,A-Z:A-Z,a-z:a-z,1-9:1-9,��:��,else:"&ItemValue)%>
		</select>
		<span  class="tx">A-Z,a-z,1-9�������������ķ���&nbsp;</span>
		<span id="ItemValue_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ѡ��ģʽ</td>
      <td> <select name="ItemMode" id="ItemMode" onChange="Do.these('ItemMode',function(){return isEmpty('ItemMode','ItemMode_Alt')}); this.options[this.selectedIndex].value=='3'?PicSrc.disabled=false:PicSrc.disabled=true;">
          <%=PrintOption(ItemMode,":��ѡ��,1:��������ģʽ,2:������дģʽ,3:ͼƬģʽ")%> 
        </select>
		<span  class="tx">ѡ��������дģʽ,���ֺ���Զ��¼���,����ѡ��&nbsp;</span>
        <span id="ItemMode_Alt"></span></td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ͼƬλ��</td>
      <td>
		<input type="text" name="PicSrc" id="PicSrc" readonly="" size="50" maxlength="200" value="<%if Bol_IsEdit then response.Write(VS_Rs("PicSrc")) end if%>">
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.form1.PicSrc);">
		<span  class="tx">ͼƬλ��(���ͼƬģʽ����)&nbsp;</span>
		<span id="PicSrc_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��ʾ��ɫ</td>
      <td>
		<input type="text" name="DisColor" id="DisColor" size="15" maxlength="7" <%if DisColor<>"" then response.Write("style=""background-color:"&DisColor&"""") end if%> value="<%=DisColor%>">
        <img src="../Images/rectNoColor.gif" width="18" height="17" border=0 align="absmiddle" id="TitleFontColor_Show" style="cursor:pointer;background-color:;" title="ѡȡ��ɫ!" onClick="GetColor(this,'DisColor');"> 
        <span  class="tx">ͳ��ʱ��ʾ��ɫ��#FF0000&nbsp;</span> <span id="DisColor_Alt"></span>	
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��ǰƱ��</td>
      <td>
		<input type="text" name="VoteCount" id="VoteCount" size="15" maxlength="5" onFocus="Do.these('VoteCount',function(){return isEmpty('VoteCount','VoteCount_Alt')&&isNumber('VoteCount','VoteCount_Alt','��������',false)})" onKeyUp="Do.these('VoteCount',function(){return isEmpty('VoteCount','VoteCount_Alt')&&isNumber('VoteCount','VoteCount_Alt','��������',false)})" value="<%=VoteCount%>">
		<span id="VoteCount_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ѡ����ϸ˵��</td>
      <td>
		<textarea name="ItemDetail" cols="50" rows="15" id="ItemDetail"><%if Bol_IsEdit then response.Write(VS_Rs("ItemDetail")) end if%></textarea>
		<span id="ItemDetail_Alt"></span>		
	  </td>
    </tr>
   <tr class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" ȷ���ύ " onClick="ItemDetail.value=ItemDetail.value.substring(0,300);" /> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ɾ�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%="VS_Items.asp?Act=Del&IID="&IID%>'">
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
	return isEmpty('TID','TID_Alt') && isEmpty('ItemName','ItemName_Alt') && isEmpty('ItemMode','ItemMode_Alt') && isEmpty('VoteCount','VoteCount_Alt');
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function GetColor(img_val,input_val)
{
	var PaletteLeft,PaletteTop
	var obj = document.getElementById("colorPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);	
	if (obj){
		PaletteLeft = getOffsetLeft(ColorImg)
		PaletteTop = (getOffsetTop(ColorImg) + ColorImg.offsetHeight)
		if (PaletteLeft+150 > parseInt(document.body.clientWidth)) PaletteLeft = parseInt(event.clientX)-260;
		obj.style.left = PaletteLeft + "px";
		obj.style.top = PaletteTop + "px";
		if (obj.style.visibility=="hidden")
		{
			obj.style.visibility="visible";
		}else {
			obj.style.visibility="hidden";
		}
	}
}
function setColor(color)
{
	if(ColorImg.id=="FontColorShow"&&color=="#") color='#000000';
	if(ColorImg.id=="FontBgColorShow"&&color=="#") color='#FFFFFF';
	if (ColorValue){ColorValue.value = color.substr(1);}
	if (ColorImg && color.length>1){
		ColorImg.src='../Images/Rect.gif';
		ColorImg.style.backgroundColor = color;
	}else if(color=='#'){ ColorImg.src='../Images/rectNoColor.gif';}
	document.getElementById("colorPalette").style.visibility="hidden";
}
-->
</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





