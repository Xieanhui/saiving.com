<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
if not MF_Check_Pop_TF("CS001") then Err_Show

Dim sRootDir,str_CurrPath
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/" & G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if


Dim RsEditObj,EditSql,SiteID,Auto_Collect_Time
Set RsEditObj = Server.CreateObject(G_FS_RS)
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & CintStr(SiteID)
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	Else
		If IsNull(RsEditObj("AutoCellectTime")) Or RsEditObj("AutoCellectTime") = "" Then
			Auto_Collect_Time = "no"
		Else
			Auto_Collect_Time = RsEditObj("AutoCellectTime")
		End If	
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Function SiteFolderList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = CollectConn.Execute("Select * from FS_SiteFolder where 1=1 order by ID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("Folder") = ClassListObj("ID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		SiteFolderList = SiteFolderList & "<option " & SelectStr & " value="&ClassListObj("ID")&"" & ">&nbsp;&nbsp;|--" & ClassListObj("SiteFolder") & "</option><br>"
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="Form" method="post" action="SiteTwoStep.asp">
<table width="98%" height="20" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback" >
            <td width="45" style="cursor:hand" align="center" alt="第二步" onClick="CheckData();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			<td width="35" style="cursor:hand" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result" value="Edit"></td>
        </tr>
  </table>  
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr class="hback" > 
      <td width="100" height="26"> 
      <div align="right">站点名称：</div></td>
      <td> 
        <input name="SiteName" style="width:100%;" type="text" id="SiteName" value="<%=RsEditObj("sitename")%>"> 
        <div align="right"> </div></td>
    </tr>
    <tr class="hback" > 
      <td height="26"> 
      <div align="right">采集对象页：</div></td>
      <td> 
        <input name="objURL" type="text" id="textarea" style="width:100%;" value="<%=RsEditObj("objURL")%>" size="50"></td>
    </tr>
	<tr class="hback">
			<td height="26"><div align="right">选择编码：</div></td>
			<td>
			<select name="WebCharset" style="width:100%;" id="WebCharset">
					<option value="GB2312" <% if RsEditObj("WebCharset") = "GB2312" Then Response.Write "selected" %>>GB2312</option>
					<option value="UTF-8" <% if RsEditObj("WebCharset") = "UTF-8" Then Response.Write "selected" %>>UTF-8</option>
				</select>
			</td>
		</tr>
    <tr class="hback" > 
      <td height="26"> 
      <div align="right">采集站点分类：</div></td>
      <td>
<select name="SiteFolder" style="width:100%;" id="SiteFolder">
          <option value="0">根栏目</option>
          <% = SiteFolderList %>
        </select></td>
    </tr>
    <tr class="hback" > 
      <td height="26"> 
      <div align="right">采集参数：</div></td>
      <td> 锁定 
        <input name="islock" type="checkbox" id="islock" value="1" <%if RsEditObj("islock")=true then response.Write("checked")%>>
        保存远程图片 
        <input type="checkbox" name="SaveRemotePic" value="1" <%if RsEditObj("SaveRemotePic")=true then response.Write("checked")%>>
        新闻是否已经审核 
        <input type="checkbox" name="Audit" value="1" <%if RsEditObj("Audit")=true then response.Write("checked")%>>
        是否倒序采集 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1" <%if RsEditObj("IsReverse")="1" then response.Write("checked")%>>
		<!-- 2007-02-25 Edit By Ken -->
		内容中包含图片时设置为图片新闻
		<input name="IsAutoPicNews" type="checkbox" id="IsAutoPicNews" value="1" <%if RsEditObj("IsAutoPicNews")="1" then response.Write("checked")%>>
		<!-- End -->
		</td>
    </tr>
    <tr class="hback" > 
      <td height="26"><div align="right">过滤选项：</div></td>
      <td>HTML 
        <input type="checkbox" name="TextTF" value="1" <% if RsEditObj("TextTF") = True then Response.Write("checked")%>>
        STYLE 
        <input type="checkbox" name="IsStyle" value="1" <% if RsEditObj("IsStyle") = True then Response.Write("checked")%>>
        DIV
        <input type="checkbox" name="IsDiv" value="1" <% if RsEditObj("IsDiv") = True then Response.Write("checked")%>>
        A
        <input type="checkbox" name="IsA" value="1" <% if RsEditObj("IsA") = True then Response.Write("checked")%>>
        CLASS
        <input type="checkbox" name="IsClass" value="1" <% if RsEditObj("IsClass") = True then Response.Write("checked")%>>
        FONT
        <input type="checkbox" name="IsFont" value="1" <% if RsEditObj("IsFont") = True then Response.Write("checked")%>>
        SPAN
        <input type="checkbox" name="IsSpan" value="1" <% if RsEditObj("IsSpan") = True then Response.Write("checked")%>>
        OBJECT
        <input type="checkbox" name="IsObject" value="1" <% if RsEditObj("IsObject") = True then Response.Write("checked")%>>
        IFRAME
        <input type="checkbox" name="IsIFrame" value="1" <% if RsEditObj("IsIFrame") = True then Response.Write("checked")%>>
        SCRIPT
        <input type="checkbox" name="IsScript" value="1" <% if RsEditObj("IsScript") = True then Response.Write("checked")%>> 
      </td>
    </tr>
	<tr class="hback">
		<td height="26"><div align="right">关键字过滤规则：</div></td>
		<td><input style="width:50%;" name="CS_SiteReKeyName" type="text" id="CS_SiteReKeyName" value="<% = CheckRulerIDStr(RsEditObj("RulerID"),"name") %>" readonly />
			<input style="width:50%;" name="CS_SiteReKeyID" type="hidden" id="CS_SiteReKeyID" value="<% = CheckRulerIDStr(RsEditObj("RulerID"),"ID") %>" />
			<% = GetAllRulerList() %>
		</td>
	</tr>
	<tr class="hback">
			<td height="26"><div align="right">入库栏目：</div></td>
			<td><select name="ToClass" style="width:100%;" id="ToClass">
				<%
					Dim obj_unite_rs,tmp_str_list
					Set obj_unite_rs = server.CreateObject(G_FS_RS)
					obj_unite_rs.Open "Select Orderid,id,ClassID,ClassName,ParentID from FS_NS_NewsClass where IsURL=0 And Parentid  = '0' and ReycleTF=0 Order by Orderid desc,ID desc",Conn,1,1
					  tmp_str_list  = ""
					  do while Not obj_unite_rs.eof 
						Dim Select_Str
						If RsEditObj("ToClassID") = obj_unite_rs("ClassID") Then
							Select_Str = "Selected"
						Else
							Select_Str = ""
						End If	
						tmp_str_list = tmp_str_list &"<option value="""& obj_unite_rs("ClassID") &""" " & Select_Str & ">+"& obj_unite_rs("ClassName") &"</option>"& Chr(13) & Chr(10)
						tmp_str_list = tmp_str_list &UniteChildNewsList(obj_unite_rs("ClassID"),"",RsEditObj("ToClassID"))
						 obj_unite_rs.movenext
					 Loop
					 obj_unite_rs.close
					 set obj_unite_rs = nothing
					 Response.Write tmp_str_list
				%>
				</select></td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">新闻模板：</div></td>
			<td>
			<input style="width:70%;" name="NewsTemp" type="text" id="NewsTemp" value="<% = RsEditObj("NewsTemplets") %>" readonly>
			<input name="Submit5" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir%>/<% = G_TEMPLETS_DIR %>',400,300,window,document.Form.NewsTemp);document.Form.NewsTemp.focus();">
			
				</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">定时采集设置：</div></td>
			<td>
			<select name="AutoCTF" id="AutoCTF" style="width:30%;" onChange="Javascript:Setautocellecttime(this.options[this.selectedIndex].value)">
				<option value="no" <%if Auto_Collect_Time = "no" Then Response.Write "selected" %>>不自动采集</option>
				<option value="day" <%if Left(Auto_Collect_Time,3) = "day" Then Response.Write "selected" %>>每天</option>
				<option value="week" <%if Left(Auto_Collect_Time,4) = "week" Then Response.Write "selected" %>>每周</option>
				<option value="month" <%if Left(Auto_Collect_Time,5) = "month" Then Response.Write "selected" %>>每月</option>
			</select>
			<%
				Dim Dis_W,W_Str
				if Left(Auto_Collect_Time,4) = "week" Then
					Dis_W = ""
					W_Str = Split(Split(Auto_Collect_Time,"$$$")(1),"|")(0)
				Else
					Dis_W = "none"
					W_Str = 0
				End If	
			%>
			<select name="TimeWeek" id="TimeWeek" style="width:30%; display:<% = Dis_W %>;">
				<option value="7" <% if Cint(W_Str) = 7 then Response.Write "selected" %>>周日</option>
				<option value="1" <% if Cint(W_Str) = 1 then Response.Write "selected" %>>周一</option>
				<option value="2" <% if Cint(W_Str) = 2 then Response.Write "selected" %>>周二</option>
				<option value="3" <% if Cint(W_Str) = 3 then Response.Write "selected" %>>周三</option>
				<option value="4" <% if Cint(W_Str) = 4 then Response.Write "selected" %>>周四</option>
				<option value="5" <% if Cint(W_Str) = 5 then Response.Write "selected" %>>周五</option>
				<option value="6" <% if Cint(W_Str) = 6 then Response.Write "selected" %>>周六</option>
			</select>
			<%
				Dim Dis_M,M_Str
				if Left(Auto_Collect_Time,5) = "month" Then
					Dis_M = ""
					M_Str = Split(Split(Auto_Collect_Time,"$$$")(1),"|")(0)
				Else
					Dis_M = "none"
					M_Str = 0
				End If	
			%>
			<select name="TimeMonth" id="TimeMonth" style="width:30%; display:<% = Dis_M %>;">
				<option value="1" <% if Cint(M_Str) = 1 then Response.Write "selected" %>>1</option>
				<option value="2" <% if Cint(M_Str) = 2 then Response.Write "selected" %>>2</option>
				<option value="3" <% if Cint(M_Str) = 3 then Response.Write "selected" %>>3</option>
				<option value="4" <% if Cint(M_Str) = 4 then Response.Write "selected" %>>4</option>
				<option value="5" <% if Cint(M_Str) = 5 then Response.Write "selected" %>>5</option>
				<option value="6" <% if Cint(M_Str) = 6 then Response.Write "selected" %>>6</option>
				<option value="7" <% if Cint(M_Str) = 7 then Response.Write "selected" %>>7</option>
				<option value="8" <% if Cint(M_Str) = 8 then Response.Write "selected" %>>8</option>
				<option value="9" <% if Cint(M_Str) = 9 then Response.Write "selected" %>>9</option>
				<option value="10" <% if Cint(M_Str) = 10 then Response.Write "selected" %>>10</option>
				<option value="11" <% if Cint(M_Str) = 11 then Response.Write "selected" %>>11</option>
				<option value="12" <% if Cint(M_Str) = 12 then Response.Write "selected" %>>12</option>
				<option value="13" <% if Cint(M_Str) = 13 then Response.Write "selected" %>>13</option>
				<option value="14" <% if Cint(M_Str) = 14 then Response.Write "selected" %>>14</option>
				<option value="15" <% if Cint(M_Str) = 15 then Response.Write "selected" %>>15</option>
				<option value="16" <% if Cint(M_Str) = 16 then Response.Write "selected" %>>16</option>
				<option value="17" <% if Cint(M_Str) = 17 then Response.Write "selected" %>>17</option>
				<option value="18" <% if Cint(M_Str) = 18 then Response.Write "selected" %>>18</option>
				<option value="19" <% if Cint(M_Str) = 19 then Response.Write "selected" %>>19</option>
				<option value="20" <% if Cint(M_Str) = 20 then Response.Write "selected" %>>20</option>
				<option value="21" <% if Cint(M_Str) = 21 then Response.Write "selected" %>>21</option>
				<option value="22" <% if Cint(M_Str) = 22 then Response.Write "selected" %>>22</option>
				<option value="23" <% if Cint(M_Str) = 23 then Response.Write "selected" %>>23</option>
				<option value="24" <% if Cint(M_Str) = 24 then Response.Write "selected" %>>24</option>
				<option value="25" <% if Cint(M_Str) = 25 then Response.Write "selected" %>>25</option>
				<option value="26" <% if Cint(M_Str) = 26 then Response.Write "selected" %>>26</option>
				<option value="27" <% if Cint(M_Str) = 27 then Response.Write "selected" %>>27</option>
				<option value="28" <% if Cint(M_Str) = 28 then Response.Write "selected" %>>28</option>
				<option value="29" <% if Cint(M_Str) = 29 then Response.Write "selected" %>>29</option>
				<option value="30" <% if Cint(M_Str) = 30 then Response.Write "selected" %>>30</option>
				<option value="31" <% if Cint(M_Str) = 31 then Response.Write "selected" %>>31</option>
			</select>
			<%
				Dim Dis_H,H_Str
				if Left(Auto_Collect_Time,2) = "no" Then
					Dis_H = "none"
					H_Str = ""
				Else
					Dis_H = ""
					If Left(Auto_Collect_Time,3) = "day" Then
						H_Str = Split(Auto_Collect_Time,"$$$")(1)
					Else
						H_Str = Split(Split(Auto_Collect_Time,"$$$")(1),"|")(1)
					End If	
				End If	
			%>
			<select name="TimeHour" id="TimeHour" style="width:30%; display:<% = Dis_H %>;">
				<option value="00:00" <% if H_Str = "00:00" Then Response.Write "selected" %>>00:00</option>
				<option value="00:30" <% if H_Str = "00:30" Then Response.Write "selected" %>>00:30</option>
				<option value="01:00" <% if H_Str = "01:00" Then Response.Write "selected" %>>01:00</option>
				<option value="01:30" <% if H_Str = "01:30" Then Response.Write "selected" %>>01:30</option>
				<option value="02:00" <% if H_Str = "02:00" Then Response.Write "selected" %>>02:00</option>
				<option value="02:30" <% if H_Str = "02:30" Then Response.Write "selected" %>>02:30</option>
				<option value="03:00" <% if H_Str = "03:00" Then Response.Write "selected" %>>03:00</option>
				<option value="03:30" <% if H_Str = "03:30" Then Response.Write "selected" %>>03:30</option>
				<option value="04:00" <% if H_Str = "04:00" Then Response.Write "selected" %>>04:00</option>
				<option value="04:30" <% if H_Str = "04:30" Then Response.Write "selected" %>>04:30</option>
				<option value="05:00" <% if H_Str = "05:00" Then Response.Write "selected" %>>05:00</option>
				<option value="05:30" <% if H_Str = "05:30" Then Response.Write "selected" %>>05:30</option>
				<option value="06:00" <% if H_Str = "06:00" Then Response.Write "selected" %>>06:00</option>
				<option value="06:30" <% if H_Str = "06:30" Then Response.Write "selected" %>>06:30</option>
				<option value="07:00" <% if H_Str = "07:00" Then Response.Write "selected" %>>07:00</option>
				<option value="07:30" <% if H_Str = "07:30" Then Response.Write "selected" %>>07:30</option>
				<option value="08:00" <% if H_Str = "08:00" Then Response.Write "selected" %>>08:00</option>
				<option value="08:30" <% if H_Str = "08:30" Then Response.Write "selected" %>>08:30</option>
				<option value="09:00" <% if H_Str = "09:00" Then Response.Write "selected" %>>09:00</option>
				<option value="09:30" <% if H_Str = "09:30" Then Response.Write "selected" %>>09:30</option>
				<option value="10:00" <% if H_Str = "10:00" Then Response.Write "selected" %>>10:00</option>
				<option value="10:30" <% if H_Str = "10:30" Then Response.Write "selected" %>>10:30</option>
				<option value="11:00" <% if H_Str = "11:00" Then Response.Write "selected" %>>11:00</option>
				<option value="11:30" <% if H_Str = "11:30" Then Response.Write "selected" %>>11:30</option>
				<option value="12:00" <% if H_Str = "12:00" Then Response.Write "selected" %>>12:00</option>
				<option value="12:30" <% if H_Str = "12:30" Then Response.Write "selected" %>>12:30</option>
				<option value="13:00" <% if H_Str = "13:00" Then Response.Write "selected" %>>13:00</option>
				<option value="13:30" <% if H_Str = "13:30" Then Response.Write "selected" %>>13:30</option>
				<option value="14:00" <% if H_Str = "14:00" Then Response.Write "selected" %>>14:00</option>
				<option value="14:30" <% if H_Str = "14:30" Then Response.Write "selected" %>>14:30</option>
				<option value="15:00" <% if H_Str = "15:00" Then Response.Write "selected" %>>15:00</option>
				<option value="15:30" <% if H_Str = "15:30" Then Response.Write "selected" %>>15:30</option>
				<option value="16:00" <% if H_Str = "16:00" Then Response.Write "selected" %>>16:00</option>
				<option value="16:30" <% if H_Str = "16:30" Then Response.Write "selected" %>>16:30</option>
				<option value="17:00" <% if H_Str = "17:00" Then Response.Write "selected" %>>17:00</option>
				<option value="17:30" <% if H_Str = "17:30" Then Response.Write "selected" %>>17:30</option>
				<option value="18:00" <% if H_Str = "18:00" Then Response.Write "selected" %>>18:00</option>
				<option value="18:30" <% if H_Str = "18:30" Then Response.Write "selected" %>>18:30</option>
				<option value="19:00" <% if H_Str = "19:00" Then Response.Write "selected" %>>19:00</option>
				<option value="19:30" <% if H_Str = "19:30" Then Response.Write "selected" %>>19:30</option>
				<option value="20:00" <% if H_Str = "20:00" Then Response.Write "selected" %>>20:00</option>
				<option value="20:30" <% if H_Str = "20:30" Then Response.Write "selected" %>>20:30</option>
				<option value="21:00" <% if H_Str = "21:00" Then Response.Write "selected" %>>21:00</option>
				<option value="21:30" <% if H_Str = "21:30" Then Response.Write "selected" %>>21:30</option>
				<option value="22:00" <% if H_Str = "22:00" Then Response.Write "selected" %>>22:00</option>
				<option value="22:30" <% if H_Str = "22:30" Then Response.Write "selected" %>>22:30</option>
				<option value="23:00" <% if H_Str = "23:00" Then Response.Write "selected" %>>23:00</option>
				<option value="23:30" <% if H_Str = "23:30" Then Response.Write "selected" %>>23:30</option>
			</select>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">单次采集数量：</div></td>
			<td><input style="width:100%;" name="NewsNum" type="text" id="NewsNum" value="<% = RsEditObj("CellectNewNum") %>" onKeyUp=if(isNaN(value)||event.keyCode==32)execCommand('undo');  onafterpaste=if(isNaN(value)||event.keyCode==32)execCommand('undo');>
			</td>
		</tr>
		<tr class="hback">
			<td height="26"><div align="right">图片保存路径：</div></td>
			<td><input name="PicSavePath" type="text" id="PicSavePath" style="width:40%" value="<% = RsEditObj("PicSavePath") %>" readonly>
			<INPUT type="button"  name="Submit4" value="选择路径" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%= str_CurrPath %>',300,250,window,document.Form.PicSavePath);document.Form.PicSavePath.focus();">
			保存远程图片自动水印
		<input type="checkbox" name="WaterPrint" value="1" <% IF RsEditObj("WaterPrintTF") = 1 Then Response.Write "checked" %>>
			</td>
		</tr>
  </table>
</form>
</body>
</html><%
Function UniteChildNewsList(TypeID,f_CompatStr,ClassID)  
	Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount,Select_Str
	Set f_ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassID,ParentID from FS_NS_NewsClass where IsURL=0 And ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by Orderid desc,id desc" )
	f_TempStr =f_CompatStr & "┄"
	do while Not f_ChildNewsRs.Eof
			if f_ChildNewsRs("ClassID") = ClassID Then
				Select_Str = "selected"
			Else
				Select_Str = ""
			End If	
			UniteChildNewsList = UniteChildNewsList & "<option value="""& f_ChildNewsRs("ClassID") &""" " & Select_Str & ">"
			UniteChildNewsList = UniteChildNewsList & "├" &  f_TempStr & f_ChildNewsRs("ClassName") 
			UniteChildNewsList = UniteChildNewsList & "</option>" & Chr(13) & Chr(10)
			UniteChildNewsList = UniteChildNewsList &UniteChildNewsList(f_ChildNewsRs("ClassID"),f_TempStr,ClassID)
		f_ChildNewsRs.MoveNext
	loop
	f_ChildNewsRs.Close
	Set f_ChildNewsRs = Nothing
End Function

Function CheckRulerIDStr(IDStr,Act)
	Dim Arr,i,Str_ID,CheckRs,Name_Str,Str_IDs
	Str_IDs = IDStr & ""
	Arr = Split(Str_IDs,",")
	For i = LBOund(Arr) To UBound(Arr)
		IF Arr(i) <> "" And Not IsNull(Arr(i)) Then
			Set CheckRs = CollectConn.ExeCute("Select ID,RuleName From FS_Rule Where ID = " & CintStr(Arr(i)) & "")
			IF CheckRs.Eof Then
				Str_ID = Str_ID & ""
				Name_Str = Name_Str & ""
			Else
				Str_ID = Str_ID & "," & Arr(i)
				Name_Str = Name_Str & "," & CheckRs(1)
			End If
			CheckRs.Close : Set CheckRs = Nothing	
		Else
			Str_ID = Str_ID & ""
			Name_Str = Name_Str & ""
		End If
	Next
	If Left(Str_ID,1) = "," Then
		Str_ID = Right(Str_ID,Len(Str_ID) - 1)
	End If
	If Left(Name_Str,1) = "," Then
		Name_Str = Right(Name_Str,Len(Name_Str) - 1)
	End If
	If Act = "ID" Then
		CheckRulerIDStr = Str_ID
	Else
		CheckRulerIDStr = Name_Str
	End If			
End Function


Function GetAllRulerList()
	Dim Rs
	GetAllRulerList = "<select name=""CS_SiteReKey"" id=""CS_SiteReKey"" style=""width:48%;"" onChange=""Javascript:SelectRuler(this.options[this.selectedIndex]);"">" & VbNewLine
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value="""">选择规则</option>" & VbNewLine
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value=""Clear"" style=""color:#FF0000;"">清除</option>" & VbNewLine
	Set Rs = CollectConn.ExeCute("Select ID,RuleName From FS_Rule Where ID > 0 Order By ID Desc")
	Do While Not Rs.Eof
		GetAllRulerList = GetAllRulerList & VbTab & VbTab & " <option value=""" & Rs(0) & """>" & Rs(1) & "</option>" & VbNewLine
	Rs.MoveNext
	Loop
	GetAllRulerList = GetAllRulerList & VbTab & VbTab & "</select>"	
	Rs.Close : Set Rs = Nothing
End Function


Set Conn = Nothing
Set CollectConn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function CheckData()
{
	if (document.Form.SiteName.value==''){alert('没有填写站点名称');document.Form.SiteName.focus();return;}
	if (document.Form.objURL.value==''){alert('没有填写采集对象页');document.Form.objURL.focus();return;}
	document.Form.submit();
}

function Setautocellecttime(str)
{
	if (str == "" || str == "no")
	{
		document.getElementById("TimeHour").style.display = "none";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "none";
	}
	if (str == "day")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "none";
	}
	else if (str == "week")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "";
		document.getElementById("TimeMonth").style.display = "none";
	}
	else if (str == "month")
	{
		document.getElementById("TimeHour").style.display = "";
		document.getElementById("TimeWeek").style.display = "none";
		document.getElementById("TimeMonth").style.display = "";
	}
}

function SelectRuler(obj)
{
	var Ruler_V = obj.value;
	var Ruler_Str = obj.innerText;
	var Name_Str = document.Form.CS_SiteReKeyName.value;
	var ID_Str = document.Form.CS_SiteReKeyID.value;
	if (Ruler_V == '')
	{
		return;
	}
	if (Ruler_V == 'Clear')
	{
		document.Form.CS_SiteReKeyName.value = '';
		document.Form.CS_SiteReKeyID.value = '';
	}
	else
	{
		if (Name_Str == '' || ID_Str == '')
		{
			document.Form.CS_SiteReKeyName.value = Ruler_Str;
			document.Form.CS_SiteReKeyID.value = Ruler_V;
		}
		else
		{
			if (ID_Str.indexOf(Ruler_V) != -1)
			{
				document.Form.CS_SiteReKeyName.value = Name_Str;
				document.Form.CS_SiteReKeyID.value = ID_Str;
			}
			else
			{
				document.Form.CS_SiteReKeyName.value = Name_Str + ',' + Ruler_Str;
				document.Form.CS_SiteReKeyID.value = ID_Str + ',' + Ruler_V;
			}
		}
	}
}
</script>





