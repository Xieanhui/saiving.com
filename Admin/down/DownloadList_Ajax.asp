<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim Conn,User_Conn,DS_Rs,DS_Sql ,DS_Rs1,DS_Sql1
MF_Default_Conn 
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("Down_List") then Err_Show
if request.QueryString("Act") = "DelAddr" then
	if request.QueryString("AddrID")<>"" then 
		Conn.execute("delete from FS_DS_Address where ID="&CintStr(request.QueryString("AddrID")))
		Call Edit_AddrList(NoSqlHack(request.QueryString("DownLoadID")))
	end if
elseif 	request.QueryString("Act") = "Check" then
	select case request.QueryString("stype")
	case "downname"
	DS_Sql = "select Count(*) from FS_DS_List where ClassID='"&NoSqlHack(request.QueryString("classid"))&"'  and Name='"&NoSqlHack(request.QueryString("name"))&"'"
	case "DownLoadID"
	DS_Sql = "select Count(*) from FS_DS_List where DownLoadID='"&NoSqlHack(request.QueryString("value"))&"'"
	case "addrname"
	DS_Sql = "select Count(*) from FS_DS_Address where AddressName='"&NoSqlHack(request.QueryString("value"))&"'"
	end select
	response.Write(Get_OtherTable_Value( DS_Sql ))
	response.End()
elseif request.QueryString("Act") = "GetExtName" then
	response.Write(getClass_FileExtName(NoSqlHack(request.QueryString("ClassID"))))
	response.End()
end if

Sub Edit_AddrList(DownID)
Dim rowii
rowii = 0
if DownID<>"" then
	DS_Sql1 = "select ID,AddressName,Url,Number from FS_DS_Address where DownLoadID = '"&NoSqlHack(DownID)&"' order by Number desc"
	set DS_Rs1 = Conn.execute(DS_Sql1)
	response.Write("<table border=""0"" width=""100%"" cellpadding=""3"" cellspacing=""1"" class=""table"">"&vbcrlf)
	do while not DS_Rs1.eof 
	rowii = rowii + 1%>
    <tr  class="hback"> 
      <td align="right">下载地址名称</td>
      <td>
		<input type="text" size="40" maxlength="50" name="AddressName" id="AddressName" value="<%=DS_Rs1("AddressName")%>">
	  </td>
    </tr>
    <tr>
      <td class="hback" align="right">下载地址</td>
      <td colspan="3" class="hback">
	  <input name="Url" type="text" id="Url" style="width:50%"  maxlength="100" value="<%=DS_Rs1("Url")%>"> 
      <input type="button" name="bnt_ChoosePic_rowBettween"  value="选择文件" onClick="SelectFile();">
	  <span id="Url_Alt"></span>
	  </td>       
    </tr>
    <tr  class="hback"> 
      <td align="right">下载地址排序</td>
      <td>
	  <input type="text" name="Number" id="Number" size="10" maxlength="1" value="<%=DS_Rs1("Number")%>">
	  <%if rowii>1 then%>
	  <input type="button" class="tx" value="删除这条下载" onClick="if(confirm('确定删除这条下载吗？')) {new Ajax.Updater('Ajax_AddrInfo','DownloadList_Ajax.asp?no-cache='+Math.random() , {method: 'get', parameters: 'Act=DelAddr&DownLoadID=<%=DownID%>&AddrID=<%=DS_Rs1("ID")%>' });disabled=true;}">
	  <%end if%>
      <span class=tx>默认排序请留空！</span></td>
    </tr>
<%	
		DS_Rs1.movenext
	loop
	response.Write("</table>")
	DS_Rs1.close
%>
	
<%end if

end Sub

Function getClass_FileExtName(ClassID)
	if ClassID<>"" then 
		set DS_Rs = Conn.execute("select FileExtName from FS_DS_Class where ClassID='"&NoSqlHack(ClassID)&"'")
		if not DS_Rs.eof then 
			getClass_FileExtName = DS_Rs("FileExtName")
		else
			getClass_FileExtName= "html"	
		end if
		DS_Rs.close
	end if
End Function

''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs,str_Chk_Info
	str_Chk_Info = ""
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			str_Chk_Info = str_Chk_Info & This_Fun_Rs(0)
			This_Fun_Rs.movenext
		loop
	else			
		if not This_Fun_Rs.eof then 
			str_Chk_Info = This_Fun_Rs(0)
		else
			str_Chk_Info = "OK"
		end if		
	end if	
	if Err.Number>0 then 
		Err.Clear
		Get_OtherTable_Value = "系统错误,"&Err.Description&",请联系管理员."
		exit function
	end if
	set This_Fun_Rs=nothing 
	if str_Chk_Info<>"" and cstr(str_Chk_Info)<>"0" then 
		str_Chk_Info = "<font color=red>重复:" & str_Chk_Info &"</font>"
	else
		str_Chk_Info = "<font color=green>OK</font>"
	end if		
	Get_OtherTable_Value = str_Chk_Info
End Function

User_Conn.close
Conn.close
%>





