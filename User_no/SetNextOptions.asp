<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%Dim User_Conn,Conn,ReqSql,EquValue,sType,optionstr,SelectName
MF_Default_Conn
MF_User_Conn
response.Charset="gb2312"
SelectName = NoSqlHack(trim(request("SelectName")))
EquValue  = CintStr(request("EquValue"))
'fucxi 2008-7-21�޸� ��ҳֻ����AP_Person_Search.asp����ϵ���ɲ�����ReqSql�����������ж�SelectName����
Select Case SelectName
    Case "Job"
        ReqSql = "Select Job From FS_AP_Job where TID=" & EquValue      
    Case "City"
        ReqSql = "select City from FS_AP_City where PID=" & EquValue
    Case Else
        response.Write("ϵͳ��������ϵ����Ա��ֻ�ܽ����˲Ų�ѯ�Ĳ�������")
	    response.End()
End Select

'trim(request("ReqSql"))
sType = trim(request("sType"))
if not isnumeric(sType) then sType = 1
if SelectName = "" then SelectName = "NoName_Sys"
if instr(lcase(ReqSql),"select ")=0 then 
	response.Write("ϵͳ��������ϵ����Ա��")
	response.End()
end if
optionstr = Get_NextOptions(ReqSql,EquValue,sType)
if optionstr = "" then optionstr = "<option value="""">[��]</option>"
response.Write("<select name="""&SelectName&""" id="""&SelectName&""">"&vbNewLine)
response.Write("<option value="""">������</option>"&vbNewLine)
response.Write(optionstr)
response.Write("</select>"&vbNewLine)

Function Get_NextOptions(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql ����sql���,EquValue�����ݿ���ͬ��ֵ�����<option>�����selected,Get_Type=1Ϊ<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
if instr(This_Fun_Sql,"FS_ME_")>0 then 
	set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
else	
	set This_Fun_Rs = Conn.execute(This_Fun_Sql)
end if	
If Err.Number <> 0 then response.Redirect("error.asp?ErrCodes=<li>"&Err.description&"</li><li>��Ǹ,�����Sql���������.�����ֶβ�����.</li>")
do while not This_Fun_Rs.eof 	
	select case cstr(Get_Type)
	  case "1"
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
		exit do : Get_FildValue_List = "<option value="""">Get_Typeֵ�������</option>"&vbNewLine : exit Function
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_NextOptions = Get_Html
End Function



User_Conn.close
Conn.close
%>






