<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Charset="GB2312"
Dim title,condition,rs,i,pid
condition=request.QueryString("condition")
pid=CintStr(request.QueryString("pid"))
if condition= "jobname" then
	title="��ѡ����Ҫ������ְλ"
elseif condition= "workcity" then
	title="��ѡ���������Ĺ����ص�"
elseif condition= "workcity2" then
	title="��ѡ���������Ĺ����ص�"
else
	title="��ѡ��Ҫ������ʱ�䷶Χ"
end if
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
<tr>
<td class="xingmu" align="center" colspan="9"><%=title%></td>
<td width="10%"><a href="#" onclick="javascript:this.parentNode.parentNode.parentNode.parentNode.parentNode.style.display='none'" class="sd">[�ر�]</a></td>
</tr>
<%
	i=0
	if condition="jobname" then
		Set Rs=conn.execute("select distinct jobname from FS_AP_Job_Public")
		while not rs.eof
			if i Mod 10=0 then Response.Write("<tr>"&vbcrlf)
			i=i+1
			Response.Write("<td height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
			Response.Write("<a href='#' onClick=""chooseIt('JobName','"&Rs("jobname")&"')"">"&Rs("jobname")&"</a>")
			Response.Write("&nbsp;</td>")
			if i Mod 10=0 then Response.Write("</tr>"&vbcrlf)
			rs.movenext
		wend
		if i Mod 10<>0 then
			while  i Mod 10<>0
				Response.Write("<td class='hback'>&nbsp;</td>")
				i=i+1
			Wend
			Response.Write("</tr>")
		End if
	Elseif condition="workcity" then
		set Rs=Conn.execute("select pid,Province from FS_AP_Province" )
		while not rs.eof
			if i Mod 10=0 then Response.Write("<tr>"&vbcrlf)
			i=i+1
			Response.Write("<td height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
			Response.Write("<a href='#' onClick=""chooseIt('txt_WorkCity','"&Rs("Province")&"','"&Rs("pID")&"')"">"&Rs("Province")&"</a>")
			Response.Write("&nbsp;</td>")
			if i Mod 10=0 then Response.Write("</tr>"&vbcrlf)
			rs.movenext
		wend
		if i Mod 10<>0 then
			while  i Mod 10<>0
				Response.Write("<td class='hback'>&nbsp;</td>")
				i=i+1
			Wend
		Response.Write("</tr>")
		End if
	Elseif condition="workcity2" then
		if pid="" then pid=0
		set Rs=Conn.execute("select City from FS_AP_City where pid="&pid )
		while not rs.eof
			if i Mod 10=0 then Response.Write("<tr>"&vbcrlf)
			i=i+1
			Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
			Response.Write("<a href='#' onClick=""chooseIt('txt_WorkCity','"&Rs("City")&"')"">"&Rs("City")&"</a>")
			Response.Write("&nbsp;</td>")
			if i Mod 10=0 then Response.Write("</tr>"&vbcrlf)
			rs.movenext
		wend
		if i Mod 10<>0 then
			while  i Mod 10<>0
				Response.Write("<td class='hback'>&nbsp;</td>")
				i=i+1
			Wend
		Response.Write("</tr>")
		End if
	Elseif condition="publicdate" then
		Response.Write("<tr>"&vbcrlf)
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','��һ��','','1')"">��һ��</a>")
		Response.Write("&nbsp;</td>")
		
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','������','','2')"">������</a>")
		Response.Write("&nbsp;</td>")
		
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','������','','3')"">������</a>")
		Response.Write("&nbsp;</td>")		
		
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','��һ��','','7')"">��һ��</a>")
		Response.Write("&nbsp;</td>")
	
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','������','','14')"">������</a>")
		Response.Write("&nbsp;</td>")	
				
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','��һ��','','30')"">��һ��</a>")
		Response.Write("&nbsp;</td>")	
				
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','������','','42')"">������</a>")
		Response.Write("&nbsp;</td>")	
		
		Response.Write("<td width='10%' height='25' class='hback' onMouseOver=""this.className='hback_1'"" onMouseOut=""this.className='hback'"">&nbsp;")
		Response.Write("<a href='#' onClick=""chooseIt('PublicDate','������','','60')"">������</a>")
		Response.Write("&nbsp;</td>")	
	
		Response.Write("<td class='hback'>&nbsp;</td>")
		Response.Write("<td class='hback'>&nbsp;</td>")
		Response.Write("</tr>"&vbcrlf)
	End if
%>
</table>
<%
Conn.close
User_Conn.close
Set Conn=nothing
Set User_Conn=nothing
%>





