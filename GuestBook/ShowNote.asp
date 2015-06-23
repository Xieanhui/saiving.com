<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_InterFace/ns_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
Dim Configobj,PageS,sql,MSTitle,ShowIP,IsUser,tmp_IsUser,s_reUserMember,style
MF_Default_Conn
MF_User_Conn
Set Configobj= server.CreateObject (G_FS_RS)
sql="select ID,Title,IPShow,IsUser,IsAut,PageSize,Style,RepUserMember From FS_WS_Config"
configobj.open sql,Conn,1,1
if not configobj.eof then
	PageS=configobj("PageSize")
	MSTitle=configobj("Title")
	ShowIP=configobj("IPShow")
	IsUser=configobj("isUser")
	s_reUserMember =configobj("RepUserMember")
	if s_reUserMember="" or not isnumeric(s_reUserMember) then
		s_reUserMember = 0
	else
		s_reUserMember = s_reUserMember
	end if
	Style = configobj("Style")
	if Style<>"" then
		Style = Style
	else
		Style = "3"
	end if
end if
response.Cookies("FoosunUserCookies")("UserLogin_Style_Num") = Style
set configobj=nothing
if IsUser="0" then
	tmp_IsUser = true
else
	if session("Fs_UserName")<>"" then
		tmp_IsUser = true
	else
		tmp_IsUser = false
	end if
end if
'---分页
Dim int_Start,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=PageS '设置每页显示数目
toF_="<font face=webdings>9</font>"   			'首页 
str_nonLinkColor_="#999999" '非热链接颜色
'int_RPP = 30
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"
%>
<html>
<HEAD>
<TITLE><%=GetGuestBookTitle%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="../<% = G_USER_DIR %>/images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript">
function AddNote(ClassID,ClassName,User,IsUser)
{
if(User==""){
	if(IsUser=="1")
		{
		alert("对不起,还没开发非会员留言权限!");
		return false;
		}
		else
		{		
		location='AddnewNotes.asp?ClassID='+ClassID+'&ClassName='+ClassName;
		return true;
		}
	}
else
   {
		location='AddnewNotes.asp?ClassID='+ClassID+'&ClassName='+ClassName;
		return true;
	}
}
function showRep(cat,User,isUser)
{
if(User==""){
	if(isUser=="1"){
		alert("对不起,还没开放非会员留言权限!");
		}
	else{
		 cat.style.display="";
  		 document.showBbs.Content.focus();	
	}
 }
else{
  cat.style.display="";
  document.showBbs.Content.focus();
 }
}
function ShowNote(NoteID,ClassName,ClassID)
{
alert(ClassID);
location="ShowNote.asp?NoteID="+NoteID+"&ClassName="+ClassName+"&ClassID="+ClassID;
}

</script>
<body>
<%
Dim ID,NoteID,NoteRs,ClassName,ClassID,BbsRs,Topic,UserName,Face,Content,i
Set NoteRs=Server.CreateObject(G_FS_RS)
Set BbsRs=Server.CreateObject(G_FS_RS)
if NoSqlHack(Request.QueryString("NoteID"))<>"" then
	ID=NoSqlHack(Request.QueryString("NoteID"))
	ClassName=NoSqlHack(Request.QueryString("ClassName"))
	ClassID=NoSqlHack(Request.queryString("ClassID"))
	Conn.execute("Update FS_WS_BBS set Hit=(Hit+1) Where ID="&CintStr(ID)&"")
	NoteRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP  From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,1,1
	if not NoteRs.eof then
	%>
 <form name="showBbs" id="showBbs" action="?Act=AddRe" method="post">
      <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr> 
    <td height="36" colspan="2" align="left" class="xingmu"><img src="images/Forum_nav.gif"> <a href="index.asp" class="Top_Navi"><b><%=MSTitle%></b></a> -> <a href="DefNoteList.asp?ClassID=<%=ClassID%>" class="Top_Navi"><b><%=ClassName%></b></a>-&gt;<%=NoteRs("Topic")%><input type="hidden" name="ClassName" value="<%=ClassName%>"></td> 
  </tr>
  <tr>
  	<td class="hback" colspan="2" >
		<img src="images/postnew.gif" alt="发表贴子" width="85" height="26" style="CURSOR: hand" onMouseUp="return AddNote('<%=ClassID%>','<%=ClassName%>','<%=session("FS_UserName")%>','<%=IsUser%>')">&nbsp;&nbsp;&nbsp;&nbsp;<%if tmp_IsUser = true then%><img src="images/mreply.gif" alt="发表贴子" width="85" height="26" style="CURSOR: hand" onMouseUp="showRep(REP,'<%=session("FS_UserName")%>','<%=IsUser%>')"><%end if%> </td>
  </tr>
  <tr>
  	<td class="hback" width="12%" rowspan="2">
	<%if NoteRs("User")<>"游客" and NoteRs("User")<>"过客" then%>
	<a href="../<%=G_USER_DIR%>/ShowUser.asp?UserName=<%=NoteRs("User")%>" target="_blank"><b><%=NoteRs("User")%></b></a><br>
	<%else%>
	<b><%=NoteRs("User")%></b><br>
	<%
	end if
	MF_User_Conn
	dim  h_rs
	set h_rs = User_Conn.execute("select HeadPic,UserName,HeadPicSize From FS_ME_Users where UserName = '"& NoteRs("User")&"'")
	if not h_rs.eof then
		if trim(h_rs("HeadPic"))<>"" then
			if instr(h_rs("HeadPicSize"),",")>0 then
				response.Write "<img src = """& h_rs("HeadPic")&""" height="""& split(h_rs("HeadPicSize"),",")(1)&""" width="""&split(h_rs("HeadPicSize"),",")(0)&""" />"
			else
				response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
			end if
		else
				response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
		end if
		h_rs.close:set h_rs= nothing
	else
		response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
		h_rs.close:set h_rs= nothing
	end if
	%>
	<br><img src="images/ip.gif">
	<%
	if ShowIP="0" then 
	Response.Write(NoteRs("IP")&"<br>") 
	end if
	%>
	<%=NoteRs("AddDate")%>	</td>
	<td width="88%" height="2" class="hback"><strong><img src="Images/face<% = NoteRs("Face")%>.gif" width="22" height="22"><%=NoteRs("Topic")%></strong></td>
  </tr>
  <tr>
  	<td height="74" class="hback">&nbsp;<%=NoteRs("Body")%><br>
  	  <%
	if session("FS_UserName")=NoteRs("User") then
	%>
      <div align="right"><a href="EditBBS.asp?Act=Edit&BBSID=<%=NoteRs("ID")%>&ClassName=<%=ClassName%>&NoteID=<%=NoteRs("ID")%>&Page=<%=cPageNo%>&ClassID=<%=NoteRs("ClassID")%>&NoteTilte=<%=NoteRs("Topic")%>">[编辑此贴子]</a>&nbsp;&nbsp;<a href="BBsDel.asp?Act=SinglDel&BBSID=<%=NoteRs("ID")%>&ClassName=<%=ClassName%>&NoteID=<%=NoteRs("ID")%>&Page=<%=cPageNo%>&ClassID=<%=NoteRs("ClassID")%>&NoteTilte=<%=NoteRs("Topic")%>" onClick="{if(confirm('确定要删除吗?')){this.document.inbox.submit();return true;}return false;}">[删除此帖子]</a>&nbsp;&nbsp;</div>
      <%end if%></td>
  </tr>
    <%
  	BbsRs.open "Select ID,ClassID,[User],ParentID,Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP  From FS_WS_BBS Where ParentID='"&NoSqlHack(ID)&"' order by AddDate",Conn,1,1
	if not BbsRs.eof then
		BbsRs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then 
			cPageNo = 1
		End if
		If not isnumeric(cPageNo) Then 
			cPageNo = 1
		End If
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then 
			cPageNo=1
		End If
		If cPageNo>BbsRs.PageCount Then 
			cPageNo=BbsRs.PageCount 
		End IF
		BbsRs.AbsolutePage=cPageNo
	i=1
	FOR int_Start=1 TO int_RPP
	i=i+1 
	%>	
  <tr>
  	<td class="hback"
 width="14%" rowspan="2">
	<%if BbsRs("User")<>"游客" and BbsRs("User")<>"过客" then%>
	<a href="../<%=G_USER_DIR%>/ShowUser.asp?UserName=<%=BbsRs("User")%>" target="_blank"><b><%=BbsRs("User")%></b></a><br>
	<%else%>
	<b><%=BbsRs("User")%></b><br>
	<%
	end if
	set h_rs = User_Conn.execute("select HeadPic,UserName,HeadPicSize From FS_ME_Users where UserName = '"& BbsRs("User")&"'")
	if not h_rs.eof then
		if trim(h_rs("HeadPic"))<>"" then
			if instr(h_rs("HeadPicSize"),",")>0 then
				response.Write "<img src = """& h_rs("HeadPic")&""" height="""& split(h_rs("HeadPicSize"),",")(1)&""" width="""&split(h_rs("HeadPicSize"),",")(0)&""" />"
			else
				response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
			end if
		else
				response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
		end if
		h_rs.close:set h_rs= nothing
	else
		response.Write "<img src = ""../sys_images/nopic_supply.gif"" border=""0"" />"
		h_rs.close:set h_rs= nothing
	end if
	%>
<br>
	<img src="images/ip.gif"><%
	if ShowIP="0" then 
	Response.Write(NoteRs("IP")&"<br>") 
	end if
	%>
	<%=BbsRs("AddDate")%>	</td>
	<td class="hback"
 width="86%" height="5"><div align="left">第<%=int_Start%>楼</div>
	  </iv></td>
  </tr>
  <tr>
  	<td class="hback"
><img src="Images/face<%=BbsRs("Face")%>.gif" width="22" height="22"><%=BbsRs("Body")%><br>
	<%
	if session("FS_UserName")=BbsRs("User") then
	%>
	<div align="right"><a href="EditBBS.asp?Act=Edit&BBSID=<%=BbsRs("ID")%>&ClassName=<%=ClassName%>&NoteID=<%=NoteRs("ID")%>&Page=<%=cPageNo%>&ClassID=<%=NoteRs("ClassID")%>&NoteTilte=<%=NoteRs("Topic")%>">[编辑此贴子]</a>&nbsp;&nbsp;<a href="BBsDel.asp?Act=SinglDel&BBSID=<%=BbsRs("ID")%>&ClassName=<%=ClassName%>&NoteID=<%=NoteRs("ID")%>&Page=<%=cPageNo%>&ClassID=<%=NoteRs("ClassID")%>&NoteTilte=<%=NoteRs("Topic")%>" onClick="{if(confirm('确定要删除吗?')){this.document.inbox.submit();return true;}return false;}">[删除此帖子]</a>&nbsp;&nbsp;</div>
	<%end if%> 
	</td>
  </tr>
	<%		
	BbsRs.MoveNext
	if BbsRs.eof or BbsRs.bof then exit for
    NEXT
	Response.Write("<tr><td class=""hback"" colspan=""2"" align=""right"">"&fPageCount(BbsRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf&"</td></tr>")
	end if
   	BbsRs.close	
   %>
  	<td class="hback" colspan="2" ><img src="images/postnew.gif" alt="发表贴子" width="85" height="26" style="CURSOR: hand" onMouseUp="return AddNote('<%=ClassID%>','<%=ClassName%>','<%=session("FS_UserName")%>','<%=IsUser%>')">&nbsp;&nbsp;&nbsp;&nbsp;<%if tmp_IsUser = true then%><img src="images/mreply.gif" alt="发表贴子" width="85" height="26" style="CURSOR: hand" onMouseUp="showRep(REP,'<%=session("FS_UserName")%>','<%=IsUser%>')"><%end if%> </td>
  </tr>
  <%if tmp_IsUser = true then%>
  <tr ID="REP"  style="display:">
  	<td colspan="2" class="hback"> 
		 <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="0"> 			
		   <tr>
  			<td class="tdtitle" colspan="2" width="100%">快速回复帖子</td>
		   </tr>
			<tr>
			<td align="right" width="15%">用户名</td>
			<td><input type="text" id="UserName" name="UserName" size="50" maxlength="80" <% 
			if session("FS_UserName")="" then 
				Response.write ("value=""游客""") 
			else 
				response.write("value="&session("FS_UserName")&"")
			end if
			%>
			 readonly></td>
			</tr>
				<td  align="right" height="25">表情</td>
				<td  ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td> <input name="FaceNum" type="radio" value="1" checked> 
                      <img src="Images/face1.gif" width="22" height="22"> </td>
                    <td> <input type="radio" name="FaceNum" value="2"> <img src="Images/face2.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="3"> <img src="Images/face3.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="4"> <img src="Images/face4.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="5"> <img src="Images/face5.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="6"> <img src="Images/face6.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="7"> <img src="Images/face7.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="8"> <img src="Images/face8.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="9"> <img src="Images/face9.gif" width="22" height="22"></td>
                  </tr>
                  <tr> 
                    <td> <input type="radio" name="FaceNum" value="10"> <img src="Images/face10.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="11"> <img src="Images/face11.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="12"> <img src="Images/face12.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="13"> <img src="Images/face13.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="14"> <img src="Images/face14.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="15"> <img src="Images/face15.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="16"> <img src="Images/face16.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="17"> <img src="Images/face17.gif" width="22" height="22">                    </td>
                    <td> <input type="radio" name="FaceNum" value="18"> <img src="Images/face18.gif" width="22" height="22">                    </td>
                  </tr>
                </table>
				</td>
			<tr>
				<td  align="right">内容</td>
				<td  valign="top" ><textarea name="Content" cols="78" rows="8" id="Content"></textarea></td>
			</tr>
			<tr>
				<td  >&nbsp;<input type="hidden" name="ClassID" value="<%=ClassID%>"><input type="hidden" name="NoteID" value="<%=NoteRs("ID")%>"><input type="hidden" name="Topic" value="<%=NoteRs("Topic")%>"><input type="hidden" name="ID" value="<%=ID%>"></td>
				 <td ><input type="submit" name="submit" value="回复帖子">&nbsp;&nbsp;
	    <input type="reset" name="reset" value=" 清  空 "></td>
			</tr>
	    </table>
	</td>
  </tr>
  <%end if%>
  </table>
</form>
	<%		
	end if
end if
Set NoteRs=nothing
if NoSqlHack(Request("Act"))="AddRe" then
	ID=NoSqlHack(Request.form("ID"))
	NoteID=NoSqlHack(Request.form("NoteID"))
	ClassID=NoSqlHack(Request.form("ClassID"))
	Topic=NoSqlHack(Request.form("Topic"))
	UserName=NoSqlHack(Request.form("UserName"))
	Face=NoSqlHack(Request.form("FaceNum"))
	Content=replace(NoHtmlHackInput(NoSqlHack(Request.form("Content"))),chr(13)&chr(10),"<br>")
	ClassName=NoSqlHack(Request.form("ClassName"))
	if NoteID="" or ClassID="" or ID="" or Topic="" or UserName="" or Face="" or ClassName="" then
		Response.write ("<script>alert('参数出错!');history.back();</script>")
		response.end
	end if
	if Content="" then 
		Response.write ("<script>alert('留言内容不能为空!');history.back();</script>")
		response.end
	end if
	BbsRs.open "Select ID,ClassID,[User],ParentID,Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP  From FS_WS_BBS Where 1=2",Conn,3,3
	BbsRs.Addnew
		BbsRs("ClassID")=ClassID
		BbsRs("User")=UserName
		BbsRs("ParentID")=ID
		BbsRs("topic")=Topic
		BbsRs("Body")=Content
		BbsRs("AddDate")=now()
		BbsRs("LastUpdateDate")=now()
		BbsRs("LastUpdateUser")=UserName
		BbsRs("Face")=Face
		BbsRs("IP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
	BbsRs.update
	Set BbsRs=nothing
	Conn.execute("Update FS_WS_BBS set Answer=(Answer+1),LastUpdateDate='"&now()&"',LastUpdateUser='"&NoSqlHack(UserName)&"' Where ID="&CintStr(ID)&"")
	'更新会员积分
	if session("FS_UserName")<>"" then
		User_Conn.execute("Update FS_ME_Users set Integral=Integral+"& CintStr(s_reUserMember) &" where UserName='"& session("FS_UserName")&"'")
		if s_reUserMember<>0 then
			dim f_AddlogObj
			Set f_AddlogObj = server.CreateObject(G_FS_RS)
			f_AddlogObj.open "select  * From FS_ME_Log where 1=0",User_Conn,1,3
			f_AddlogObj.addnew
			f_AddlogObj("LogType")="其他"
			f_AddlogObj("UserNumber")=GetFriendNumber(session("FS_UserName"))
			f_AddlogObj("points")=NoSqlHack(s_reUserMember)
			f_AddlogObj("moneys")=0
			f_AddlogObj("LogTime")=Now
			f_AddlogObj("LogContent")="发表帖子增加积分"
			f_AddlogObj("Logstyle")=0
			f_AddlogObj.update
			f_AddlogObj.close
			set f_AddlogObj = nothing
		end if 
	end if
	Response.Redirect("ShowNote.asp?NoteID="&NoSqlHack(NoteID)&"&ClassName="&ClassName&"&ClassID="&ClassID&"")
end if
Function GetFriendNumber(f_strNumber)
	Dim RsGetFriendNumber
	Set RsGetFriendNumber = User_Conn.Execute("Select UserNumber From FS_ME_Users Where UserName = '"& NoSqlHack(f_strNumber) &"'")
	If  Not RsGetFriendNumber.eof  Then 
		GetFriendNumber = RsGetFriendNumber("UserNumber")
	End If 
	set RsGetFriendNumber = nothing
End Function 
Set Conn=nothing
set User_Conn = nothing
%>
</body>
</html>






