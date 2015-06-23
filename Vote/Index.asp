<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
'''应用样式通过ID号：总的表格 id=Table_Vote ,提交 id=but_VoteSubmit 查看结果 id = but_ViewVote
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
response.Charset = "gb2312"
Dim Conn,User_Conn,VS_RS,VS_Sql,VS_RS1,VS_Sql1,VS_RS2,VS_Sql2,strShowErr ,IPInterView,IsSigned, TID,ShowType,ff_row_ii,MaxNum
Dim TmpStr,TmpStr1,Steps,DB_Steps,Cookie_Domain,OutHtmlID,PicW,Vs_Num
ff_row_ii = 1 : DB_Steps = 0

MF_Default_Conn

Cookie_Domain = Get_MF_Domain() 
if Cookie_Domain="" then  
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	
''前台页面,由JS调用得到 调用该文件必须给定一些参数.
TID = NoSqlHack(request.QueryString("TID"))
Steps = NoSqlHack(request.QueryString("Steps"))
OutHtmlID = NoSqlHack(request.QueryString("InfoID"))
PicW = request.QueryString("PicW")
if TID="" or not isnumeric(TID) then call HtmEnd("投票参数错误!")	
if Steps="" or isnull(Steps) or not isnumeric(Steps) then Steps = 1
if OutHtmlID = "" then OutHtmlID = "Vote_HTML_ID"
if PicW = "" or not isnumeric(PicW) then PicW = 100

''''''''''''''''''''' 
Call Show
''''''''''''''''''''

Sub Show()
	''选出最近的一步
	VS_Sql2 = "select top 1 * from FS_VS_Steps where TID = "&TID&" and Steps>="&Steps&" order by TID,Steps"
	Set VS_RS2 = CreateObject(G_FS_RS)
	VS_RS2.Open VS_Sql2,Conn,1,1
		if VS_RS2.eof then	
			VS_Sql = "select * from FS_VS_Theme where TID = "&TID
		else
			DB_Steps = Conn.execute("select Count(*) from FS_VS_Steps where TID = "&TID)(0)
			VS_Sql = "select * from FS_VS_Theme where TID = "&VS_RS2("QuoteID")
		end if
		If G_IS_SQL_DB = 1 Then	
			VS_Sql = VS_Sql&" AND (getdate() between StartDate and EndDate)"
		else
			VS_Sql = VS_Sql&" AND (now() between StartDate and EndDate)"
		end if
		'response.Write(VS_Sql)
		set VS_RS = Conn.execute(VS_Sql) 
		if Not VS_RS.eof then 
		if VS_RS("Type") = 2 then 
			MaxNum = VS_RS("MaxNum")
		else
			MaxNum = 1
		end if	
		response.Write("<table width=""100%"" id=""Table_Vote"" border=""0"" cellspacing=""3"" cellpadding=""0""> "&vbnewline)
		response.Write("<form name=voteForm id=voteForm method=""post"" onSubmit=""if(document.getElementById('ItemsInput')!=null){new Ajax.Updater('Ajax_TPInfo_"&VS_RS("TID")&"', 'http://"&Cookie_Domain&"/Vote/Vote_Ajax.asp?TID="&VS_RS("TID")&"&ItemsInput='+$('ItemsInput').value+'&MaxNum="&MaxNum&"&no-cache='+Math.random() , {method: 'post', parameters: Form.serialize(this) });}; else {new Ajax.Updater('Ajax_TPInfo_"&VS_RS("TID")&"', 'http://"&Cookie_Domain&"/Vote/Vote_Ajax.asp?TID="&VS_RS("TID")&"&MaxNum="&MaxNum&"&no-cache='+Math.random() , {method: 'post', parameters: Form.serialize(this) });};return false;""> "&vbnewline)
		response.Write("<tr> "&vbnewline)
		response.Write("<td width=""3%"" height=""18"">&nbsp;</td> "&vbnewline)
		response.Write("<td width=""3%""><img src=""Img/3.jpg"" width=""9"" height=""9""></td> "&vbnewline)
		response.Write("<td style=""cursor:hand"" onclick=""if(TR_"&VS_RS("TID")&".style.display=='none') TR_"&VS_RS("TID")&".style.display=''; else TR_"&VS_RS("TID")&".style.display='none';"">"&VS_RS("Theme")&"</td> "&vbnewline)
		response.Write("<td width=""3%"">&nbsp;</td> "&vbnewline)
		response.Write("</tr> "&vbnewline)
		response.Write("<tr id=""TR_"&VS_RS("TID")&""" style=""display:'none'""> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("<td><img src=""Img/5.jpg"" width=""6"" height=""6""></td> "&vbnewline)
		response.Write("<td style=""font-size:11px;color:#666666"">Begin:"&formatdatetime(VS_RS("StartDate"),0)&" End:"&formatdatetime(VS_RS("EndDate"),0)&"</td> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("</tr> "&vbnewline)
		response.Write("<tr> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("<td> "&vbnewline)
		if VS_RS2.eof then 
			VS_Sql1 = "select * from FS_VS_Items where TID = "&TID
		else	
			VS_Sql1 = "select * from FS_VS_Items where TID = "&VS_RS2("QuoteID")
		end if
		Set VS_RS1 = CreateObject(G_FS_RS)
		VS_RS1.Open VS_Sql1,Conn,1,1
		'---2007-01-15 Edit By Ken
		If Not VS_RS1.Eof Then
			Vs_Num = VS_RS1.recordcount
		Else
			Vs_Num = 0
		End If
		'------------		
		response.Write( "<!--++++++++++++++++循环主体开始++++++++++++++++-->"&vbnewline )
		do while Not VS_RS1.eof
			response.Write( radboxItemMode(VS_RS1("ItemMode")) &vbnewline )
			ff_row_ii = ff_row_ii + 1
			VS_RS1.movenext
		loop	
		response.Write( "<!--++++++++++++++++循环主体结束++++++++++++++++-->"&vbnewline )
		VS_RS1.close
		response.Write("</td> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("</tr> "&vbnewline)
		response.Write("<tr> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("<td>&nbsp;</td> "&vbnewline)
		response.Write("<td align=left> "&vbnewline)
		if VS_RS2.recordcount<=0 then 
			response.Write("<input type=""submit"" id=""but_VoteSubmit"" value="" 提 交 "" style=""width:80px;height:23px;""> "&vbnewline)
		else
			if cint(Steps)>1 then response.Write("<input type=""button"" value="" 上一步 "" style=""width:80px;height:23px; "" onClick=""new Ajax.Updater('"&OutHtmlID&"', 'http://"&Cookie_Domain&"/Vote/Index.asp?no-cache='+Math.random() , {method: 'get', parameters: 'TID="&TID&"&Steps="&Steps-1&"&InfoID="&OutHtmlID&"&PicW="&PicW&"' });""> "&vbnewline)
			response.Write("<input type=""submit""  id=""but_VoteSubmit"" value="" 提 交 "" style=""width:80px;height:23px; ""> "&vbnewline)				
			if cint(Steps)<DB_Steps then response.Write("<input type=""button"" value="" 下一步 "" style=""width:80px;height:23px; "" onClick=""new Ajax.Updater('"&OutHtmlID&"', 'http://"&Cookie_Domain&"/Vote/Index.asp?no-cache='+Math.random() , {method: 'get', parameters: 'TID="&TID&"&Steps="&Steps+1&"&InfoID="&OutHtmlID&"&PicW="&PicW&"' });""> "&vbnewline)
		end if
		response.Write(" <input type=""button"" id=""but_ViewVote"" value=""查看结果""  style=""width:80px;height:23px;"" onClick=""window.open('http://"&Cookie_Domain&"/Vote/View.asp?TID="&TID&"&Title="&server.URLEncode(Conn.execute("select Theme from FS_VS_Theme where TID = "&TID)(0))&"&Vs_Num=" & Vs_Num & "');"">")
		if DB_Steps>1 and VS_RS2.recordcount>0 then response.Write("&nbsp;"&Steps&"/"&DB_Steps&"")
			response.Write("</td> "&vbnewline)
			response.Write("<td>&nbsp;</td> "&vbnewline)
			response.Write("</tr> "&vbnewline)
			response.Write("<tr> "&vbnewline)
			response.Write("<td>&nbsp;</td> "&vbnewline)
			response.Write("<td>&nbsp;</td> "&vbnewline)
			response.Write("<td id=""Ajax_TPInfo_"&VS_RS("TID")&""">&nbsp;</td> "&vbnewline)
			response.Write("<td>&nbsp;</td> "&vbnewline)
			response.Write("</tr> "&vbnewline)
			response.Write("</form> "&vbnewline)
			response.Write("</table> "&vbnewline)
		else
			response.Write("没有记录. "&vbnewline)
		end if
		VS_RS.close	
	VS_RS2.close
End Sub

'选项模式:1:文字描述模式2:自主填写模式(文字后可以多个录入框)3:图片模式
''单选,多选
Function radboxItemMode(TType)
	''选项名称
	Dim ThisFun_Str,f_ItemValue_,rad_box
	TmpStr = "" : TmpStr1 = "" : ThisFun_Str = ""
	if VS_RS("Type") = 1 then 
		rad_box = "radio"
	elseif VS_RS("Type") = 2 then 
		rad_box = "checkbox"
	else
		exit function
	end if	
	f_ItemValue_ = VS_RS1("ItemValue")
	select case f_ItemValue_
	case "1-9"
		f_ItemValue_ = ff_row_ii &"."
	case "A-Z"
		f_ItemValue_ = chr(64+ff_row_ii) &"."
	case "a-z"
		f_ItemValue_ = chr(96+ff_row_ii) &"."
	case else
	end select
		
	if VS_RS1("DisColor")<>"" then 
		TmpStr = "<font color="""&VS_RS1("DisColor")&""">"&f_ItemValue_&" "&VS_RS1("ItemName")&"</font> "&vbNewLine
	else
		TmpStr = ""& f_ItemValue_&" "&VS_RS1("ItemName")&" "&vbNewLine
	end if		
	select case TType
	case 1
		if VS_RS("ItemMode") = 0 then 
			''横向
		else
			if ff_row_ii mod VS_RS("ItemMode") = 0 then 	
				TmpStr = TmpStr & "<br /> "&vbNewLine
			end if
		end if	
		if request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")<>"" and instr(cstr(request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")),","&cstr(VS_RS1("IID"))&",") then
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&""" checked> "&vbNewLine&TmpStr&vbNewLine
		else
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&"""> "&vbNewLine&TmpStr&vbNewLine
		end if
	case 2
		TmpStr = TmpStr&"<input type=text name=""ItemsInput"" id=""ItemsInput"" value="""&request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_ItemValue")&""" size=15 maxlength=25> "&vbNewLine		
		if VS_RS("ItemMode") = 0 then 
			''横向
		else
			if ff_row_ii mod VS_RS("ItemMode") = 0 then 	
				TmpStr = TmpStr & "<br /> "&vbNewLine
			end if
		end if	    
		if request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")<>"" and instr(cstr(request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")),","&cstr(VS_RS1("IID"))&",") then
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&""" checked> "&vbNewLine&TmpStr&vbNewLine
		else
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&"""> "&vbNewLine&TmpStr&vbNewLine
		end if	
	case 3
		if VS_RS1("PicSrc")<>"" then 
			TmpStr1 = "<img src="""&VS_RS1("PicSrc")&""" title=""点开查看大图:"&f_ItemValue_&" "&VS_RS1("ItemName")&""" onload=""if (this.offsetWidth>"&PicW&") this.width="&PicW&";"" style=""cursor:hand"" onclick=""window.open('"&VS_RS1("PicSrc")&"');"" />"	
		else
			TmpStr1 = "<img src=""http://"&Cookie_Domain&"/Vote/Img/NoPic.jpg"" title="""&f_ItemValue_&" "&VS_RS1("ItemName")&""" onload=""if (this.offsetWidth>"&PicW&") this.width="&PicW&";"" />"	
		end if
		if ff_row_ii = 1 then ThisFun_Str = ThisFun_Str & "<table border=0 cellpadding=""0"" cellspacing=""5""><tr> "&vbNewLine	
		ThisFun_Str = ThisFun_Str & "<td align=center>"&TmpStr1&"<br /> "&vbNewLine
		if request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")<>"" and instr(cstr(request.Cookies("FS_Cookies")(cstr(VS_RS1("TID"))&"_IID")),","&cstr(VS_RS1("IID"))&",") then
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&""" checked> "&vbNewLine&TmpStr&"</td> "&vbNewLine	
		else
			ThisFun_Str = ThisFun_Str & "<input name=""Items"" type="""&rad_box&""" value="""&VS_RS1("IID")&"""> "&vbNewLine&TmpStr&"</td> "&vbNewLine	
		end if
		if VS_RS("ItemMode") = 0 then 
			''横向
		else
			if ff_row_ii mod VS_RS("ItemMode") = 0 then 	
				ThisFun_Str = ThisFun_Str & "</tr><tr> "&vbNewLine
			else	
			end if
		end if
		if ff_row_ii = VS_RS1.recordcount then ThisFun_Str = ThisFun_Str & "</tr></table> "&vbNewLine
	end select
	radboxItemMode = ThisFun_Str
End Function
''====================

Function HtmEnd(Msg)
	ConnClose
	response.Write(""&Msg&" "&vbnewline)
	response.End()
End Function

Sub ConnClose()
	Set Conn = Nothing
End Sub
%>






