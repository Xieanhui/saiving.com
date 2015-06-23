<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,sRootDir,str_CurrPath
MF_Default_Conn
'session判断
MF_Session_TF 
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
if not MF_Check_Pop_TF("WS001") then Err_Show
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
'---分页
Dim int_Start,int_RPP,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=10 '设置每页显示数目
toF_="<font face=webdings>9</font>"   			'首页 
str_nonLinkColor_="#999999" '非热链接颜色
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
<head>
	<title>FoosunCMS留言系统</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css"
	rel="stylesheet" type="text/css">

<script type="text/javascript" src="../../FS_Inc/PublicJS.js"></script>

<script type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>

<script type="text/javascript" src="../../FS_Inc/Prototype.js"></script>

<script type="text/javascript">
	function AddNote(ClassID) {
		location = "AddNewNote.asp?ClassID=" + ClassID;
	}
	function showRep(cat) {
		cat.style.display = "";
		document.getElementById('NewsContent').src = '../Editer/AdminEditer.asp?id=Content';
		cat.focus();
	}
	function ShowNote(NoteID, ClassName, ClassID) {
		alert(ClassID);
		location = "ShowNote.asp?NoteID=" + NoteID + "&ClassName=" + ClassName + "&ClassID=" + ClassID;
	}

</script>

<body>
	<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
		<tr>
			<td align="left" class="xingmu">
				留言管理&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=Message"
					target="_blank" style="cursor: help; '" class="sd"><img src="../Images/_help.gif"
						border="0"></a>
			</td>
		</tr>
	</table>
	<%
Dim ID,NoteID,NoteRs,ClassName,ClassID,BbsRs,Topic,UserName,Face,Content,i
Set NoteRs=Server.CreateObject(G_FS_RS)
Set BbsRs=Server.CreateObject(G_FS_RS)
if Request.QueryString("NoteID")<>"" then
	ID=NoSqlHack(Request.QueryString("NoteID"))
	ClassName=NoSqlHack(Request.QueryString("ClassName"))
	ClassID=NoSqlHack(Request.queryString("ClassID"))
	NoteRs.open "Select ID,ClassID,[User],Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP  From FS_WS_BBS Where ID="&CintStr(ID)&"",Conn,1,1
	if not NoteRs.eof then
	%>
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
		<form name="showBbs" id="showBbs" action="?Act=AddRe" method="post">
		<tr>
			<td align="left" class="xingmu" colspan="2">
				<img src="images/Forum_nav.gif"><a href="ClassMessageManager.asp" class="sd">留言管理</a>-><a
					href="ClassMessageManager.asp?Act=all" class="sd"><%=ClassName%></a>-&gt;<%=NoteRs("Topic")%>
				<input type="hidden" name="ClassName" value="<%=ClassName%>">
			</td>
		</tr>
		<tr>
			<td class="hback" colspan="2">
				<img src="images/postnew.gif" alt="发表贴子" width="85" height="26" style="cursor: hand"
					onmouseup="AddNote('<%=ClassID%>')">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/mreply.gif"
						alt="回复贴子" width="85" height="26" style="cursor: hand" onmouseup="showRep(REP)">
			</td>
		</tr>
		<tr>
			<td class="hback" width="14%" rowspan="2">
				<%=NoteRs("User")%>&nbsp;&nbsp;&nbsp;&nbsp;<br>
				<img src="images/noHeadPic.jpg" alt=" " width="91" height="84"><br>
				<img src="images/ip.gif"><%=NoteRs("IP")%><br>
				<%=NoteRs("AddDate")%>
			</td>
			<td class="hback" width="86%" height="2%">
				<img src="Images/face<% = NoteRs("Face")%>.gif" width="22" height="22"><%=NoteRs("Topic")%>
			</td>
		</tr>
		<tr>
			<td class="hback">
				&nbsp;&nbsp;<%=NoteRs("Body")%><br>
				<div align="right">
					<a href="NoteEdit.asp?Act=NoteEdit&ID=<%=NoteRs("ID")%>">[编辑此贴子]</a>&nbsp;&nbsp;<a
						href="NoteDel.asp?ID=<%=NoteRs("ID")%>&Act=single" onclick="{if(confirm('如果删除该话题,那么相关的评论都将被删除,确定要删除吗?')){return true;}return false;}">[删除此帖子]</a>&nbsp;&nbsp;</div>
			</td>
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
			cPageNo = Clng(cPageNo)
		End If
		If cPageNo<=0 Then 
			cPageNo=1
		End If
		If cPageNo>BbsRs.PageCount Then 
			cPageNo=BbsRs.PageCount 
			BbsRs.AbsolutePage=cPageNo
		End IF
	i=1
	FOR int_Start=1 TO int_RPP
	i=i+1 
		%>
		<tr>
			<td class="hback" width="15%" rowspan="2">
				<%=BbsRs("User")%>&nbsp;&nbsp;&nbsp;&nbsp;<br>
				<img src="images/noHeadPic.jpg" alt=" " width="91" height="84"><br>
				<img src="images/ip.gif"><%=BbsRs("IP")%><br>
				<%=BbsRs("AddDate")%>
			</td>
			<td class="hback" width="90%" height="5">
				&nbsp;
				<div align="right">
					第<%=int_Start%>楼
				</div>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<img src="Images/face<%=BbsRs("Face")%>.gif" width="22" height="22"><%=BbsRs("Body")%><br>
				<div align="right">
					<a href="EditBBS.asp?Act=Edit&BBSID=<%=BbsRs("ID")%>&ClassName=<%=ClassName%>">[编辑此贴子]</a>&nbsp;&nbsp;<a
						href="BBsDel.asp?Act=SinglDel&BBSID=<%=BbsRs("ID")%>&ClassID=<%=ClassID%>&ClassName=<%=ClassName%>&NoteID=<%=BbsRs("ParentID")%>"
						onclick="{if(confirm('确定要删除吗?')){return true;}return false;}">[删除此帖子]</a>&nbsp;&nbsp;</div>
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
		<td class="hback" colspan="2">
			<img src="images/postnew.gif" alt="发表贴子" width="85" height="26" style="cursor: hand"
				onmouseup="AddNote('<%=ClassID%>')">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/mreply.gif"
					alt="发表贴子" width="85" height="26" style="cursor: hand" onmouseup="showRep(REP)">
		</td>
		</tr>
		<tr id="REP" style="display:none;">
			<td class="hback" colspan="2" width="100%">
				<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
					<tr>
						<td class="hback" colspan="2" width="100%">
							快速回复帖子
						</td>
					</tr>
					<tr>
						<td class="hback" align="right" width="15%">
							用户名
						</td>
						<td class="hback">
							<input type="text" id="UserName" name="UserName" size="50" maxlength="80" value="<%=Temp_Admin_Name%>"
								readonly>
						</td>
					</tr>
					<tr>
						<td class="hback" align="right" height="25">
							表情
						</td>
						<td class="hback">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td>
										<input name="FaceNum" type="radio" value="1" checked>
										<img src="Images/face1.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="2">
										<img src="Images/face2.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="3">
										<img src="Images/face3.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="4">
										<img src="Images/face4.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="5">
										<img src="Images/face5.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="6">
										<img src="Images/face6.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="7">
										<img src="Images/face7.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="8">
										<img src="Images/face8.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="9">
										<img src="Images/face9.gif" width="22" height="22">
									</td>
								</tr>
								<tr>
									<td>
										<input type="radio" name="FaceNum" value="10">
										<img src="Images/face10.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="11">
										<img src="Images/face11.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="12">
										<img src="Images/face12.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="13">
										<img src="Images/face13.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="14">
										<img src="Images/face14.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="15">
										<img src="Images/face15.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="16">
										<img src="Images/face16.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="17">
										<img src="Images/face17.gif" width="22" height="22">
									</td>
									<td>
										<input type="radio" name="FaceNum" value="18">
										<img src="Images/face18.gif" width="22" height="22">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="hback" align="right">
							内容
						</td>
						<td class="hback" valign="top">
							<!--编辑器开始-->
							<iframe id='NewsContent' frameborder="0"
								scrolling="no" width='100%' height='280'></iframe>
							<input type="hidden" name="Content">
							<!--编辑器结束-->
						</td>
					</tr>
					<tr>
						<td class="hback">
							&nbsp;
							<input type="hidden" name="ClassID" value="<%=ClassID%>">
							<input type="hidden" name="NoteID" value="<%=NoteRs("ID")%>">
							<input type="hidden" name="Topic" value="<%=NoteRs("Topic")%>">
							<input type="hidden" name="ID" value="<%=ID%>">
						</td>
						<td class="hback">
							<input type="button" name="submit1" value="回复帖子" onclick="SubmitFun(this.form);">
							&nbsp;&nbsp;
							<input type="reset" name="reset" value=" 清  空 ">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		</form>
	</table>
	<%		
	end if
end if
Set NoteRs=nothing
if Request.QueryString("Act")="AddRe" then
	ID=NoSqlHack(Request.form("ID"))
	NoteID=NoSqlHack(Request.form("NoteID"))
	ClassID=NoSqlHack(Request.form("ClassID"))
	Topic=NoSqlHack(Request.form("Topic"))
	UserName=NoSqlHack(Request.form("UserName"))
	Face=NoSqlHack(Request.form("FaceNum"))
	Content=NoSqlHack(Request.form("Content"))
	ClassName=NoSqlHack(Request.form("ClassName"))
	if NoteID="" or ClassID="" or ID="" or Topic="" or UserName="" or Face="" or ClassName="" then
		Response.write ("<script>alert('参数出错啦!');history.back();</script>")
		response.end
	end if
	if Content="" then 
		Response.write ("<script>alert('留言内容不能为空!');history.back();</script>")
		response.end
	end if
	BbsRs.open "Select ID,ClassID,[User],ParentID,Topic,Body,AddDate,IsTop,State,Style,IsAdmin,Answer,Hit,LastUpdateDate,LastUpdateUser,Face,IP  From FS_WS_BBS Where 1=2",Conn,3,3
	BbsRs.Addnew
		BbsRs("ClassID")=NoSqlHack(ClassID)
		BbsRs("User")=NoSqlHack(UserName)
		BbsRs("ParentID")=NoSqlHack(ID)
		BbsRs("topic")=NoSqlHack(Topic)
		BbsRs("Body")=NoSqlHack(Content)
		BbsRs("AddDate")=now()
		BbsRs("LastUpdateDate")=now()
		BbsRs("LastUpdateUser")=NoSqlHack(UserName)
		BbsRs("Face")=NoSqlHack(Face)
		BbsRs("IP")=NoSqlHack(Request.ServerVariables("REMOTE_ADDR"))
	BbsRs.update
	Set BbsRs=nothing
		Response.Redirect("ShowNote.asp?NoteID="&NoteID&"&ClassName="&ClassName&"&ClassID="&ClassID&"")
end if

Set Conn=nothing
	%>

	<script type="text/jscript">
		function SubmitFun() {
			if (frames["NewsContent"].g_currmode != 'EDIT') { alert('其他模式下无法保存，请切换到设计模式'); return false; }
			document.getElementById('Content').value = frames["NewsContent"].GetNewsContentArray();
			document.getElementById('showBbs').submit();
		}
	</script>

</body>
</html>
