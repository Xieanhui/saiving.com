<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Public.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<%
Dim Conn,str_CurrPath,Types,NewsID,RsNewsObj,sRootDir
MF_Default_Conn
MF_Session_TF

If Request("NewsID")<>"" and Request("Types")<>"" then
   NewsID = Replace(Replace(Replace(Cstr(Request("NewsID")),")",""),"(",""),"'","")
   Types = Cstr(NoSqlHack(Request("Types")))
Else
	Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	Response.End
End if

Dim SP_List,SPRs,SP_IDStr
Set SPRs = Conn.ExeCute("Select SpecialEName,SpecialCName From FS_NS_Special Where isLock = 0 And SpecialID > 0 Order By SpecialID Desc")
IF SPRs.Eof Then
	SP_List = ""
Else
	Do While Not SPRs.Eof
		SP_List = SP_List & "<option value=""" & SPRs(0) & """>" & SPRs(1) & "</option>" & VbNewline
	SPRs.MoveNext
	Loop
End IF
SPRs.Close : Set SPRs = Nothing


NewsID = FormatIntArr(Replace(NewsID,"***",","))
IF Trim(Request.Form("action")) = "trues" Then
	SP_IDStr = NoSqlHack(Trim(Request.Form("SP_ID")))
	Conn.ExeCute("Update FS_NS_News Set SpecialEName = '" & NoSqlHack(SP_IDStr) & "' Where ID In(" & FormatIntArr(NewsID) & ")")
	Response.Write "<script>alert('设置成功');window.close();</script>"
	Response.End
End IF

%>
<html>
<head>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="../js/Public.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加新闻到专题</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" name="ToSpForm" id="ToSpForm" method="post" target="TempFream">
    <tr> 
      <td width="7%" height="5">&nbsp;</td>
      <td width="16%" height="5">&nbsp;</td>
      <td width="77%" height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>专题名称：</td>
      <td>
	  <input name="SP_Name" id="SP_Name" type="text" value="" style="width:80%;">
	  <input type="hidden" name="SP_ID" id="SP_ID" value="">
	  </td>
    </tr>
	<tr> 
      <td>&nbsp;</td>
      <td>选择专题：</td>
      <td>
	  <select name="SP_IDSE" id="SP_IDSE" style="width:50%;" onChange="Javascript:SelectSP(this.options[this.selectedIndex]);">
	  	<option value="" selected="selected">选择专题</option>
		<option value="Clear" style="color:#FF0000;">清除</option>
	  	<% = SP_List %>
	  </select>
	  </td>
    </tr>
    <tr> 
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit2" value=" 确 定 ">
          <input name="action" type="hidden" id="action" value="trues">
          <input type="button" name="Submit3" value=" 取 消 " onClick="window.close();">
        </div></td>
    </tr>
  </form>
</table>
<iframe name="TempFream" id="TempFream" src="" width="0" height="0" frameborder="0"></iframe>
</body>
</html>
<%
Conn.CLose : Set Conn = Nothing
%>
<script language="javascript" type="text/javascript">
<!--
function SelectSP(Obj)
{
	var Vstr = Obj.value;
	var Tstr = Obj.text;
	var Temp_Str = document.ToSpForm.SP_Name.value;
	var Temp_IDStr = document.ToSpForm.SP_ID.value;
	if (Vstr == '')
	{
		return;
	}
	if (Vstr == 'Clear')
	{
		document.ToSpForm.SP_Name.value = '';
		document.ToSpForm.SP_ID.value = '';
	}
	else
	{
		if (Temp_Str == '')
		{
			document.ToSpForm.SP_Name.value = Tstr;
			document.ToSpForm.SP_ID.value = Vstr;
		}
		else
		{
			if (Temp_Str.indexOf(Tstr) == -1)
			{
				document.ToSpForm.SP_Name.value = Temp_Str + ',' + Tstr;
				document.ToSpForm.SP_ID.value = Temp_IDStr + ',' + Vstr;
			}
			else
			{
				document.ToSpForm.SP_Name.value = Temp_Str;
				document.ToSpForm.SP_ID.value = Temp_IDStr;
			}	
		}
	}
}
-->
</script>