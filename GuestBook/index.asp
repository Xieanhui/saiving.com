<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp"-->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
MF_Default_Conn
Dim Configobj,Topic,isUser,PageS,Style,Conn,sql
Set Configobj= server.CreateObject(G_FS_RS)
sql="select ID,Title,IsUser,IsAut,PageSize,Style From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	Topic=configobj("Title")
	PageS=configobj("PageSize")
	IsUser=configobj("IsUser")
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
set configobj=nothing
Dim int_Start,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=PageS '����ÿҳ��ʾ��Ŀ
toF_="<font face=webdings>9</font>"   			'��ҳ 
str_nonLinkColor_="#999999" '����������ɫ
int_RPP =10
int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
toF_="<font face=webdings>9</font>"   			'��ҳ 
toP10_=" <font face=webdings>7</font>"			'��ʮ
toP1_=" <font face=webdings>3</font>"			'��һ
toN1_=" <font face=webdings>4</font>"			'��һ
toN10_=" <font face=webdings>8</font>"			'��ʮ
toL_="<font face=webdings>:</font>"
dim CountRs,TodayNotes,obj_count_rs,TodayBbs,YesterdayNote,All,sqlDate,ClassAll,TodayClassAll,ClassAllRs,ClassSqlcount,PersRs,PersN,PersB
Set ClassAllRs= Server.CreateObject(G_FS_RS)
Set obj_count_rs= Server.CreateObject(G_FS_RS)
if G_IS_SQL_DB=0 then
	sqlDate="#"&datevalue(now())-1&"#"
else
	sqlDate="'"&datevalue(now())-1&"'"
end If
if G_IS_SQL_DB=0 then
	obj_count_rs.open "select id from FS_WS_BBS where ParentID<>'0' and datevalue(AddDate)>"&sqlDate&"",conn,1,1
Else
	obj_count_rs.open "select id from FS_WS_BBS where ParentID<>'0' and convert(nvarchar(10),AddDate,120)>"&sqlDate&"",conn,1,1
End if
TodayBbs = obj_count_rs.recordcount
obj_count_rs.close
obj_count_rs.open"select id from FS_WS_BBS where ParentID='0' and AddDate>"&sqlDate&"",conn,1,1
TodayNotes=obj_count_rs.recordcount
obj_count_rs.close
if G_IS_SQL_DB=0 then
	obj_count_rs.open "select id from FS_WS_BBS where ParentID='0' and AddDate>#"&datevalue(now())-2&"# and AddDate<"&sqlDate&"",conn,1,1
else
	obj_count_rs.open "select id from FS_WS_BBS where ParentID='0' and AddDate>'"&datevalue(now())-2&"' and AddDate<"&sqlDate&"",conn,1,1
end if
YesterdayNote=obj_count_rs.recordcount
obj_count_rs.close
obj_count_rs.open "select id from FS_WS_BBS",conn,1,1
All=obj_count_rs.recordcount
obj_count_rs.close
Set obj_count_rs=nothing
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">

function opencat(i)
{
  if(document.getElementById("Note"+i).style.display=="none"){
     document.getElementById("Note"+i).style.display="";
	 document.getElementById("Img"+i).src="images/nofollow.gif";
  } else {
     document.getElementById("Note"+i).style.display="none"; 
	 document.getElementById("Img"+i).src="images/plus.gif";
  }
}
</script>
<body>
<!--����-->
<!--#include file="Tell.asp"-->
<!--��¼-->
<!--��ҳ��-->
<%
Dim ClassRs,ClassSql,NoteRs,NoteSql,MsRs,MsSql,NoteAct,NoteSqlEnd,i,SelectClassID
i=0
Set ClassRs=Server.CreateObject(G_FS_RS)
Set MsRs=Server.CreateObject(G_FS_RS)
ClassRs.open "Select ID,ClassID,ClassName,ClassExp,Pid,Author,AddDate from FS_WS_Class order by Pid desc",Conn,1,1
If not ClassRs.eof then
	ClassRs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Or Not IsNumeric(cPageNo) Then 
		cPageNo = 1
	End if
	cPageNo = Clng(cPageNo)
	If cPageNo < 1 Then
		cPageNo = 1
	End If	
	If cPageNo > ClassRs.PageCount Then 
		cPageNo = ClassRs.PageCount 
	End IF
	ClassRs.AbsolutePage=cPageNo
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <%
  if session("FS_UserName")="" then
  	Response.write("<tr><td width=""100%"" class=""hback"" height=""30"">�������ǵĻ�Ա��?�뵥������:<a href=""../User/login.asp?UrlAddress=../GuestBook/index.asp""><font color=""red"">��½</font></a></td></tr>")
  else
  'ͳ�Ƹ�����Ϣ
	Set PersRs=Server.CreateObject(G_FS_RS)
	PersRs.open "Select ID from FS_WS_BBS where User='"&session("FS_UserName")&"' and ParentID='0'",Conn,1,1
	PersN=PersRs.recordcount
	PersRs.close
	PersRs.open "Select ID from FS_WS_BBS where User='"&session("FS_UserName")&"' and ParentID<>'0'",Conn,1,1
	PersB=PersRs.recordcount
	PersRs.close
	Set PersRs=nothing
  %>
  <tr class="hback" height="30">
    <td>
		��ӭ��:<%=session("FS_UserName")%>�� <a href="../<% = G_USER_DIR %>/LoginOut.asp?sUrl=../GuestBook/index.asp">�˳�</a> ����������<font color="red"><%=PersN%></font>ƪ�����ظ�<font color="red"><%=PersB%></font>ƪ����������ͳ�ƣ�������<font color="red"><%=TodayNotes%></font>ƪ����������<font color="red"><%=TodayBbs%></font>ƪ��������<font color="red"><%=YesterdayNote%></font>ƪ��������<font color="red"><%=All%></font>ƪ</td>
  </tr>
  	<%end if%>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
<%
 FOR int_Start=1 TO int_RPP 
%>
  <tr class="xingmu">
    <td height="30" colspan="3" align="center" valign="middle" ><div align="left"><a href="DefNoteList.asp?ClassID=<%=ClassRs("ClassID")%>" class="Top_Navi"><strong><%=ClassRs("ClassName")%></strong></a></div></td>
  </tr>
  <tr class="hback">
    <td width="33%" rowspan="2" class="tdhback"><img src="images/nofollow.gif" name="Img" id="Img">&nbsp;<a href="DefNoteList.asp?ClassID=<%=ClassRs("ClassID")%>"><%=ClassRs("ClassExp")%></a></td>
	<td width="28%"  class="tdhback">����ʱ��:<font color="#FF0000"><%=ClassRs("AddDate")%></font></td>
    <td width="39%" rowspan="2"  class="tdhback"><%
	dim rs,str_topic,pub_topbic,pub_date,str_id
	set rs = Conn.execute("select top 1 Id,Topic,AddDate,LastUpdateDate,LastUpdateUser,user From FS_WS_BBS where State='0' and ClassID='"&NoSqlHack(ClassRs("ClassId"))&"' order by LastUpdateDate desc,id desc")
	if rs.eof then
		str_topic = "��"
		pub_topbic = "��"
		pub_date = "��"
		str_id = ""
		rs.close:set rs = nothing
	else
		str_topic = rs("Topic")
		if rs("LastUpdateUser")<>"�ο�" and rs("LastUpdateUser")<>"����" then
			pub_topbic = "<a href=../"&G_USER_DIR&"/ShowUser.asp?UserName="&rs("LastUpdateUser")&" target=""_blank"">"&rs("LastUpdateUser")&"</a>"
		else
			pub_topbic = rs("LastUpdateUser")
		end if
		pub_date = rs("LastUpdateDate")
		str_id = rs("id")
		rs.close:set rs = nothing
	end if
	%>
    <table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td>���⣺
	<%if str_id<>"" then%>
      <a href="ShowNote.asp?NoteID=<%=str_id%>&ClassName=<%=ClassRs("ClassName")%>&ClassID=<%=ClassRs("ClassId")%>"><% = str_topic %></a></td>
  	<%else%>
      <% = str_topic %></td>
	<%end if%>
  </tr>
  <tr>
    <td>������
      <% = pub_topbic %></td>
  </tr>
  <tr>
    <td>���ڣ�
      <% = pub_date %></td>
  </tr>
</table>
    </td>
  </tr>  
  <tr>
  <%
  ClassAllRs.open "Select ID From FS_WS_BBS where ClassID='"&ClassRs("ClassID")&"'",conn,1,1  
  ClassAll=0
  TodayClassAll=0
  if not ClassAllRs.eof then
  	ClassAll=ClassAllRs.recordcount
  end if
  ClassAllRs.close
  if G_IS_SQL_DB=0 then
	ClassSqlcount="Select ID From FS_WS_BBS where ClassID='"&ClassRs("ClassID")&"' and AddDate>#"&	datevalue(now())-1&"#"
else
		ClassSqlcount="Select ID From FS_WS_BBS where ClassID='"&ClassRs("ClassID")&"' and AddDate>'"&datevalue(now())-1&"'"
  end if
  ClassAllRs.open ClassSqlcount,conn,1,1
	if not ClassAllRs.eof then
  	TodayClassAll=ClassAllRs.recordcount
  end if
  ClassAllRs.close
  %>
  <td  class="hback" >����:<font color="red"><%=ClassAll%></font>ƪ&nbsp;&nbsp;������:<font color="red"><%=TodayClassAll%></font>ƪ&nbsp;</td>
  </tr>
  <%  
    ClassRs.MoveNext
	if ClassRs.eof or ClassRs.bof then exit for
    NEXT
	response.Write "<tr><td colspan=""4"" class=""hback_1"" align=""right"">"&  fPageCount(ClassRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td></tr>"
  %>
</table>
<%
Else
	Response.Write("��������")
End If
Set ClassRs=nothing
Set Conn=nothing
%>
</body>
</html>






