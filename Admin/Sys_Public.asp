<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Public") then Err_Show
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/common.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/coolWindowsCalendar.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.RefreshLen{
	height: 20px;
	width: 400px;
	border: 1px solid #104a7b;
	text-align: left;
	MARGIN-top:50px;
	margin-bottom: 5px;
}
</style>
<BODY oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="xingmu">
		<td class="xingmu">
			<p>��������</p>
		</td>
	</tr>
	<tr class="hback">
		<td><a href="Sys_Public.asp?Type=MF">վ����ҳ</a>��<a href="Sys_Public.asp?Type=NS">����</a>��
			<%if IsExist_SubSys("MS") then%>
			<a href="Sys_Public.asp?Type=MS">�̳�</a>��
			<%end if%>
			<a href="Sys_Public.asp?Type=DS">����</a>��<a href="SysRefreshset.asp">�Զ�ˢ�������ļ�</a></td>
	</tr>
</table>
<div id="RefreshMain">
	<%
	Dim str_type
	str_type = trim(Request.QueryString("Type"))
	select  case str_type
		case "MF"
			Call MF_Refresh()
		case "NS"
			Call NS_Refresh()
		case "MS"
			Call MS_Refresh()
		case "IB"
		
		case "DS"
			Call DS_Refresh()
		case "Log"
			call pub_log()
		case else
			Call MF_Refresh()
	end select
	sub pub_log()
			Dim Path,FileName,EditFile,FileContent,Result,strShowErr
			Result = Request.Form("Action")
			Path = "../FS_InterFace/Public_Log"
			FileName = "Refresh.ini"
			EditFile = Server.MapPath(Path) & "\" & FileName
			Dim FsoObj,FileObj,FileStreamObj
			Set FsoObj = Server.CreateObject(G_FS_FSO)
			Set FileObj = FsoObj.GetFile(EditFile)
			if Result = "" then
				Set FileStreamObj = FileObj.OpenAsTextStream(1)
				if Not FileStreamObj.AtEndOfStream then
					FileContent = FileStreamObj.ReadAll
				else
					FileContent = ""
				end if
			else
				Set FileStreamObj = FileObj.OpenAsTextStream(2)
				FileContent = Request.Form("ConstContent")
				FileStreamObj.Write FileContent
				if Err.Number <> 0 then
					strShowErr = "<li>����ʧ��</li><li>"& Err.Description &"</li>"
					Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				else
					strShowErr = "<li>ȫ�ֱ�������ɹ�</li>"
					Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
					Response.end
				end if
			end if
			Response.Write  "<table class=""table"" width=""98%"" align=""center""><tr class=""hback""><td class=""hback"" align=""center""><textarea name=""FileFresh"" style=""width:100%"" rows=""20"">"& FileContent &"</textarea></td></tr></table>"
%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr height="1" class="hback">
			<td height="25">��������
				<%
	if Request.QueryString("Type")="Log" then
			Response.Write("������־�鿴��<span class=""tx"">ע�⣺������������������ռ����û�а�װ����ִ����������޷�ʹ�ô˹��ܣ���Ҫ�ֹ�����</span>")
	elseif Request.QueryString("Type")="All" then
			Response.Write("��������")
	else
			response.Write Request.QueryString("Type")
	End if
	%>
			</td>
		</tr>
		<tr height="1" class="hback">
			<td height="25">
				<p>˵����<br>
					<strong>MF=Index</strong><br>
					վ����ҳ<br>
					����ϵͳ <br>
					<strong>NS=Index</strong><br>
					������ҳ<br>
					<strong>NS=Class(0)[1,2,3,6,7]</strong><br>
					ClassΪˢ��������Ŀ��(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���<br>
					<strong>NS=Class(1)[1,7]</strong><br>
					ClassΪˢ��������Ŀ��(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���<br>
					<strong>NS=News(0)[1,4,5,6,8]</strong><br>
					NewsΪˢ���������ҳ�棬(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���,(2)��ʾ��ʼʱ��ͽ���ʱ��<br>
					<strong>NS=News(1)[1,8]</strong><br>
					NewsΪˢ���������ҳ�棬(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���,(2)��ʾ��ʼʱ��ͽ���ʱ�� <br>
					<strong>NS=News(2)[2005-6-7,2006-6-8]</strong><br>
					NewsΪˢ���������ҳ�棬(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���,(2)��ʾ��ʼʱ��ͽ���ʱ�� <br>
					<strong>NS=Special(0)[1,5,6,7,8]</strong><br>
					SpecialΪˢ��ר�⣬(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ���<br>
					<strong>NS=Special(1)[1,8]</strong><br>
					SpecialΪˢ��ר�⣬(0)��ʾˢ��ָ��ID��(1)��ʾˢ�µ�ID��ʼ��ID�Ľ��� <br>
					<br>
					......<br>
					<br>
					��������
					<%
end sub
%>
			</td>
		</tr>
	</table>
	<%Sub MF_Refresh()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="3" class="xingmu">��������</td>
		</tr>
		<tr >
			<td width="13%" class="hback">
				<div align="right">��������</div>
			</td>
			<td width="87%" colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('MF','','index','');" name="Submit4" value="��ʼ����վ����ҳ">
			</td>
		</tr>
	</table>
	<%End sub%>
	<%
Sub NS_Refresh()
	Dim rs,str_ClassList
	str_ClassList =""
	set rs=Conn.execute("select ClassId,ClassName From FS_NS_NewsClass where ParentId = '0' and ReycleTF=0 and isUrl=0 order by OrderId desc,id desc")
	do while not rs.eof
		str_ClassList = str_ClassList & "<option value="""&rs("ClassId")&""">"&rs("ClassName")&"</option>"
		str_ClassList = str_ClassList & get_ChildClassList(rs("ClassId"),"��")
		rs.movenext
	loop
	rs.close:set rs=nothing
%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="3" class="xingmu">����������ҳ</td>
		</tr>
		<tr>
			<td width="13%" class="hback">
				<div align="right">����������ҳ</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('NS','','index','');" name="Submit" value="��ʼ����">
			</td>
		</tr>
		<tr>
			<td colspan="3" class="xingmu">��������</td>
		</tr>
		<tr>
			<td width="13%" class="hback">
				<div align="right">��������</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_News','nsallnews','');" name="Submit" value="��ʼ����">
			</td>
		</tr>
		<form name="Public_form_NS_ID_News" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">����ID����</div>
				</td>
				<td colspan="2" class="hback">
					<input name="startId" type="text" id="startId" value="1" size="10" maxlength="8">
					<input name="endId" type="text" id="endId" value="100" size="10" maxlength="10">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_News','nsidnews',this.form);" name="Submit5" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_NS_Last_News" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">��������</div>
				</td>
				<td colspan="2" class="hback">
					<input name="LastNews" type="text" id="LastNews" value="10" size="10" maxlength="5">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_News','nslastnews',this.form);" name="Submit6" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_NS_Date_News" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">�������ڷ���</div>
				</td>
				<td colspan="2" class="hback">
					<input name="startTime" type="text" id="startTime" value="<%=date()-1%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_NS_Date_News.startTime);document.Public_form_NS_Date_News.startTime.focus();" style="cursor:hand;">
					<input name="endTime" type="text" id="endTime" value="<%=date()%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_NS_Date_News.endTime);document.Public_form_NS_Date_News.endTime.focus();" style="cursor:hand;">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_News','nsdatenews',this.form);" name="Submit7" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_NS_Class_News" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">������Ŀ����</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="12" multiple id="ClassID" style="width:98%">
							<%=str_ClassList%>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_News','nsclassnews',this.form);" name="Submit8" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">������Ŀ</td>
			</tr>
		</form>
		<tr >
			<td class="hback">
				<div align="right">������Ŀ</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_NewsClass','nsallclass','');" name="Submit222" value="����������Ŀ">
			</td>
		</tr>
		<form name="Public_form_NS_Class" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ����Ŀ</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="10" multiple id="ClassID" style="width:98%">
							<%=str_ClassList%>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_NewsClass','nsclass',this.form);" name="Submit222" value="��ʼ����">
				</td>
			</tr>
		</form>
		<tr >
			<td colspan="3" class="xingmu">������ҳ</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">������ҳ</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_NewsClass','classpage','');" name="Submit222" value="�������е�ҳ">
			</td>
		</tr>
		<form name="Public_form_NS_Class_Page" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ��Ҫ�����ĵ�ҳ</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="10" multiple id="ClassID" style="width:98%">
							<%
							set rs = Conn.execute("select ID,ClassName From FS_NS_NewsClass where ParentId = '0' and ReycleTF=0 and isUrl=2 order by OrderId desc,id desc")
						 do while not rs.eof
							 response.Write"<option value="""&rs("Id")&""">"&rs("ClassName")&"</option>"
						 rs.movenext
						 loop
						 rs.close:set rs=nothing
							%>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_NewsClass','classpage',this.form);" name="Submit222" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">����ר��</td>
			</tr>
		</form>
		<form name="Public_form_NS_Special" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ��ר��</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="SpecialID" size="10" multiple id="SpecialID" style="width:98%">
							<%
			 set rs = Conn.execute("select SpecialID,SpecialCName From FS_NS_Special Where isLock=0 Order by SpecialID desc")
			 do while not rs.eof
				 response.Write"<option value="""&rs("SpecialID")&""">"&rs("SpecialCName")&"</option>"
			 rs.movenext
			 loop
			 rs.close:set rs=nothing
		  %>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('NS','FS_NS_Special','nsspecial',this.form);" name="Submit222" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
			</tr>
		</form>
	</table>
	<%
End sub
If str_type="NS" Then
	Function get_ChildClassList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassName from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "��"
		do while Not ChildTypeListRs.Eof
			get_ChildClassList = get_ChildClassList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_ChildClassList = get_ChildClassList & "��"&ChildTypeListRs("ClassName")&"</option>"
			get_ChildClassList = get_ChildClassList & get_ChildClassList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
end if
If str_type="MS" Then
	Function get_Child_MS_ClassList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassCName from FS_MS_ProductsClass where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "��"
		do while Not ChildTypeListRs.Eof
			get_Child_MS_ClassList = get_Child_MS_ClassList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_Child_MS_ClassList = get_Child_MS_ClassList & "��"&ChildTypeListRs("ClassCName")&"</option>"
			get_Child_MS_ClassList = get_Child_MS_ClassList & get_Child_MS_ClassList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
end if
If str_type="DS" Then
	Function get_Child_S_ClassList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassName from FS_DS_Class where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "��"
		do while Not ChildTypeListRs.Eof
			get_Child_S_ClassList = get_Child_S_ClassList &"<option value="""&ChildTypeListRs("ClassId")&""">"& TempStr
			get_Child_S_ClassList = get_Child_S_ClassList & "��"&ChildTypeListRs("ClassName")&"</option>"
			get_Child_S_ClassList = get_Child_S_ClassList & get_Child_S_ClassList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
	End Function
end if
Sub MS_Refresh()
	DIM str_ClassList,Rs
	str_ClassList =""
	set rs=Conn.execute("select ClassId,ClassCName From FS_MS_ProductsClass where ParentId = '0' and ReycleTF=0 order by OrderId desc,id desc")
	do while not rs.eof
		str_ClassList = str_ClassList & "<option value="""&rs("ClassId")&""">"&rs("ClassCName")&"</option>"
		str_ClassList = str_ClassList & get_Child_MS_ClassList(rs("ClassId"),"��")
		rs.movenext
	loop
%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="3" class="xingmu">�����̳���ҳ</td>
		</tr>
		<tr>
			<td width="13%" class="hback">
				<div align="right">�����̳���ҳ</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('MS','','index','');" name="Submit" value="��ʼ����">
			</td>
		</tr>
		<tr>
			<td colspan="3" class="xingmu">������Ʒ</td>
		</tr>
		<tr >
			<td width="13%" class="hback">
				<div align="right">��������</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Products','msallproduct','');" name="Submit" value="��ʼ����">
			</td>
		</tr>
		<form name="Public_form_MS_ID_Product" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">����ID����</div>
				</td>
				<td colspan="2" class="hback">
					<input name="startId" type="text" id="startId" value="1" size="10" maxlength="8">
					<input name="endId" type="text" id="endId" value="100" size="10" maxlength="10">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Products','msidproduct',this.form);" name="Submit2" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_MS_Last_Product" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">��������</div>
				</td>
				<td colspan="2" class="hback">
					<input name="LastNews" type="text" id="LastNews" value="10" size="10" maxlength="5">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Products','mslastproduct',this.form);" name="Submit10" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_MS_Date_Product" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">�������ڷ���</div>
				</td>
				<td colspan="2" class="hback">
					<input name="startTime" type="text" id="startTime" value="<%=date()-1%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_MS_Date_Product.startTime);document.Public_form_MS_Date_Product.startTime.focus();" style="cursor:hand;">
					<input name="endTime" type="text" id="endTime" value="<%=date()%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_MS_Date_Product.endTime);document.Public_form_MS_Date_Product.endTime.focus();" style="cursor:hand;">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Products','msdateproduct',this.form);" name="Submit11" value="��ʼ����">
				</td>
			</tr>
		</form>
		<form name="Public_form_MS_Class_Product" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">������Ŀ����</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="12" multiple id="ClassID" style="width:98%">
							<%=str_ClassList%>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Products','msclassproduct',this.form);" name="Submit12" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">������Ʒ��Ŀ</td>
			</tr>
		</form>
		<tr >
			<td class="hback">
				<div align="right">������Ʒ��Ŀ</div>
			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_ProductsClass','msallclass','');" name="Submit222" value="����������Ŀ">
			</td>
		</tr>
		<form name="Public_form_MS_Class" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ����Ŀ</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="10" multiple id="ClassID" style="width:98%">
							<%=str_ClassList%>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_ProductsClass','msclass',this.form);" name="Submit13" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">����ר��</td>
			</tr>
		</form>
		<form name="Public_form_MS_Special" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ��ר��</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="SpecialID" size="10" multiple id="SpecialID" style="width:98%">
							<%
			 set rs = Conn.execute("select SpecialID,SpecialCName From FS_MS_Special Where isLock=0 Order by SpecialID desc")
			 do while not rs.eof
				 response.Write"<option value="""&rs("SpecialID")&""">"&rs("SpecialCName")&"</option>"
			 rs.movenext
			 loop
			 rs.close:set rs=nothing
		  %>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('MS','FS_MS_Special','msspecial',this.form);" name="Submit14" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
			</tr>
		</form>
	</table>
	<%
End Sub
Sub DS_Refresh()
	Dim rs,str_ClassList
	str_ClassList =""
	set rs=Conn.execute("select ClassId,ClassName From FS_DS_Class where ParentId = '0' and ReycleTF=0 order by OrderId desc,id desc")
	do while not rs.eof
		str_ClassList = str_ClassList & "<option value="""&rs("ClassId")&""">"&rs("ClassName")&"</option>"
		str_ClassList = str_ClassList & get_Child_S_ClassList(rs("ClassId"),"��")
		rs.movenext
	loop
	rs.close:set rs=nothing
%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr>
			<td colspan="3" class="xingmu">����������ҳ</td>
		</tr>
		<tr >
			<td width="13%" class="hback">
				<div align="right">����������ҳ</div>			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('DS','','index','');" name="Submit9" value="��ʼ����">			</td>
		</tr>
		<tr>
			<td colspan="3" class="xingmu">��������</td>
		</tr>
		<tr >
			<td width="13%" class="hback">
				<div align="right">��������</div>			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_List','dsalldownload','');" name="Submit9" value="��ʼ����">			</td>
		</tr>
		<form name="Public_form_DS_ID_Download" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">����ID����</div>				</td>
				<td colspan="2" class="hback">
					<input name="startId" type="text" id="startId" value="1" size="10" maxlength="8">
					<input name="endId" type="text" id="endId" value="100" size="10" maxlength="10">
					<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_List','dsiddownload',this.form);" name="Submit5" value="��ʼ����">				</td>
			</tr>
		</form>
		<form name="Public_form_DS_Last_Download" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">��������</div>				</td>
				<td colspan="2" class="hback">
					<input name="LastNews" type="text" id="LastNews" value="10" size="10" maxlength="5">
					<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_List','dslastdownload',this.form);" name="Submit6" value="��ʼ����">				</td>
			</tr>
		</form>
		<form name="Public_form_DS_Date_Download" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">�������ڷ���</div>				</td>
				<td colspan="2" class="hback">
					<input name="startTime" type="text" id="startTime" value="<%=date()-1%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_DS_Date_Download.startTime);document.Public_form_DS_Date_Download.startTime.focus();" style="cursor:hand;">
					<input name="endTime" type="text" id="endTime" value="<%=date()%>" size="20" maxlength="30" readonly>
					<img src="../sys_images/calendar.gif" width="34" onClick="OpenWindowAndSetValue('CommPages/SelectDate.asp',280,150,window,document.Public_form_DS_Date_Download.endTime);document.Public_form_DS_Date_Download.endTime.focus();" style="cursor:hand;">
					<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_List','dsdatedownload',this.form);" name="Submit7" value="��ʼ����">				</td>
			</tr>
		</form>
		<form name="Public_form_DS_Class_Download" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">������Ŀ����</div>				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="ClassID" size="12" multiple id="ClassID" style="width:98%">
							<%=str_ClassList%>
						</select>
					</div>				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_List','dsclassdownload',this.form);" name="Submit8" value="��ʼ����">				</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">������Ŀ</td>
			</tr>
		</form>
		<tr >
			<td class="hback">
				<div align="right">������Ŀ</div>			</td>
			<td colspan="2" class="hback">
				<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_Class','dsallclass','');" name="Submit222" value="����������Ŀ">			</td>
		</tr>
		<form name="Public_form_DS_Class" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ����Ŀ</div>				</td>
				<td width="29%" class="hback">
				  <div align="center">
						<select name="select" size="10" multiple id="select" style="width:98%">
							<%=str_ClassList%>
						</select>
						<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_Class','dsclass',this.form);" name="Submit2222" value="��ʼ����">
					</div>				</td>
				<td width="58%" class="hback">&nbsp;</td>
			</tr>
			<tr >
				<td colspan="3" class="xingmu">����ר��</td>
			</tr>
		</form>
		<form name="Public_form_DS_Special" method="post" action="">
			<tr >
				<td class="hback">
					<div align="right">ѡ��ר��</div>
				</td>
				<td width="29%" class="hback">
					<div align="center">
						<select name="SpecialID" size="10" multiple id="SpecialID" style="width:98%">
							<%
			 set rs = Conn.execute("select SpecialID,SpecialCName From FS_DS_Special Where isLock=0 Order by SpecialID desc")
			 do while not rs.eof
				 response.Write"<option value="""&rs("SpecialID")&""">"&rs("SpecialCName")&"</option>"
			 rs.movenext
			 loop
			 rs.close:set rs=nothing
		  %>
						</select>
					</div>
				</td>
				<td width="58%" class="hback">
					<input type="button" onClick="Submit_Data_To_Refresh('DS','FS_DS_Special','dsspecial',this.form);" name="Submit14" value="��ʼ����">
				</td>
			</tr>
			<tr >
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
				<td class="hback">&nbsp;</td>
			</tr>
			
		</form>
	</table>
	<%
End sub
%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		<tr class="hback">
			<td><span class="tx">ע�⣺����������������û��������ķ���������ע��(��ûע��)"��������"�ؼ���������ʹ���Զ����񷢲����ܡ�<a href="../help?Lable=MF_PublicSite_Dll" target="_blank" style="cursor:help;"><img src="Images/_help.gif" width="50" height="17" border="0"></a></span></td>
		</tr>
	</table>
</div>
<div id="RefreshSchedule" style="display:none;" align="center"></div>
<!--<textarea id="TESTTEST" cols="130" rows="16"></textarea>-->
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
var G_REFRESH_NUM_TIME=<%= G_REFRESH_NUM_TIME %>;
countnum=1;
function opencat(cat)
{
  if(cat.style.display=="none") cat.style.display="";
  else cat.style.display="none"; 
}

function Submit_Data_To_Refresh(Sys,Table,Type,FormObj)
{
	var Action='',Str='',Obj=null;
	if (typeof(FormObj)=="object"){
		for(var i=0;i<FormObj.length;i++)
		{
			Obj = FormObj[i];
			if ((Obj.tagName=='INPUT')&&(Obj.type=='text')) Str=Obj.name+':'+Obj.value;
			if (Obj.tagName=='SELECT') {Str=Obj.name+':'+GetSelectID(Obj);}
			if (Str!='')
			{
				if (Action=='') Action=Str;
				else Action=Action+';'+Str;
			}
			Str='';
		}
	}else{
		Action="";
	}
	Action=Sys+'$'+Table+'$'+Type+'$'+Action+'$GO';
	Action="Action="+Action;
	$('RefreshMain').style.display="none";
	$('RefreshSchedule').style.display="";
	$('RefreshSchedule').innerHTML="<div class=\"RefreshLen\"><div class=\"xingmu\" id=\"RefreshLen\"></div></div>\
	<span id=\"result_str\"></span><br><br>";
	$("RefreshLen").style.width ="0%";
	$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">0%</span>";
	$('result_str').innerHTML="����׼��...&nbsp;&nbsp;";
	//alert(Action);
	Start_Refresh('PublicSite/Public_Refresh.asp',Action);
}

function Start_Refresh(url,Action){
	//alert(url+'**'+Action)
	var myAjax = new Ajax.Request(
		url,
		{method:'get',
		parameters:Action,
		onComplete:Refresh_Receive
		}
		);
}
function Refresh_Receive(OriginalRequest){
	var check,goback;
	var percent=0;
	//document.all.TESTTEST.value=OriginalRequest.responseText;
	goback="<a href=\"����\" onclick=\"$('RefreshMain').style.display='';$('RefreshSchedule').style.display='none';return false;\">����</a>";
	if (OriginalRequest.responseText.indexOf("$")>-1){
		check=OriginalRequest.responseText.split("$");
		switch (check[0]) {
			case "MF" :
				$("RefreshLen").style.width ="100%";
				$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">100%</span>";
				$('result_str').innerHTML="��ҳ��������&nbsp;&nbsp;"+check[3]+"&nbsp;&nbsp;<a href=\"http:\/\/<%= Request.Cookies("FoosunMFCookies")("FoosunMFDomain")&"\/"&Request.Cookies("FoosunMFCookies")("FoosunMFIndexFileName") %>\" target=\"_blank\">�����ҳ<\/a>&nbsp;&nbsp;"+goback;
				countnum=1;
				break;
			case "Next" :
				percent=(parseInt(check[2])/parseInt(check[1]))*100;
				percent=Math.round(percent);
				$("RefreshLen").style.width =percent+"%";
				$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">"+percent+"%</span>";
				$('result_str').innerHTML="�ܹ�Ҫ����" + check[1] + "������,�Ѿ�����" + check[2] + "������...";
				//Start_Refresh("PublicSite/Public_Refresh.asp","");
				
				if ((countnum % G_REFRESH_NUM_TIME)==0){
					window.setTimeout("Start_Refresh(\"PublicSite/Public_Refresh.asp\",\"\")",1000);
				}else{
					Start_Refresh("PublicSite/Public_Refresh.asp","");
				}
				countnum++;
				break;
			case "End" :
				$("RefreshLen").style.width ="100%";
				$("RefreshLen").innerHTML="&nbsp;<span class=\"xingmu\">100%</span>";
				$('result_str').innerHTML="�ܹ�Ҫ����" + check[1] + "������,�Ѿ�����" + (check[2]-1) + "������...";
				$('result_str').innerHTML=$('result_str').innerHTML+"<br />��������&nbsp;&nbsp;"+check[3]+"&nbsp;&nbsp;"+goback;
				countnum=1;
				break;
			case "No" :
				$('result_str').innerHTML="û��Ҫ����������&nbsp;&nbsp;"+goback;
				countnum=1;
				break;
			default :
				//alert(OriginalRequest.responseText);
				//$('result_str').innerHTML=OriginalRequest.responseText;
				//Start_Refresh("PublicSite/Public_Refresh.asp","");
				$('result_str').innerHTML="����ʧ�ܣ��������Ա��ϵ��&nbsp;&nbsp;"+goback+"<br>�����������£�ID��<span class=\"tx\">"+check[1]+"</span>��<span class=\"tx\">"+check[2]+"</span>";
				//Start_Refresh("PublicSite/Public_Refresh.asp","");
		} 
	}
	else{
		$('result_str').innerHTML="����ʧ�ܣ��������Ա��ϵ��&nbsp;&nbsp;"+goback+"<br>�����������£�"+OriginalRequest.responseText;
	}
}

function GetSelectID(Obj)
{
	var SelectObj=null,Str='';
	for(var i=0;i<Obj.options.length;i++)
	{
		SelectObj=Obj.options[i];
		if(SelectObj.selected)
		{
			if(Str=='') Str=SelectObj.value;
			else Str=Str+'*'+SelectObj.value;
		}
	}
	return Str;
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->