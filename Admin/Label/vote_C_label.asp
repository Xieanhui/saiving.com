<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	'session判断
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
<title>新闻标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <form  name="form1" method="post">
  <table width="98%" height="29" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="41%" class="xingmu"><strong>投票标签创建</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="关闭">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td width="19%" class="hback"><div align="right">选择栏目</div></td>
      <td width="81%" class="hback">
	  		<select name="TID" id="TID">
			<%
			Dim rs
			set rs = Conn.execute("select TID,Theme From FS_VS_Theme order by TID desc")
			do while not rs.eof
				if len(rs("Theme"))>20 then
					response.Write "<option value="""&rs("TID")&""">"&left(rs("Theme"),20)&"...</option>"
				else
					response.Write "<option value="""&rs("TID")&""">"&rs("Theme")&"</option>"
				end if
			rs.movenext
			loop
			rs.close:set rs = nothing
			%>
            </select>			</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">投票SPIN ID </div></td>
      <td class="hback"><label>
        <input name="spanid" type="text" id="spanid" value="Vote_HTML_ID">
      如果你要在一个页面调用多个投票，请把ID设置为不同</label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">图片宽度</div></td>
      <td class="hback"><label>
        <input name="PicWidth" type="text" id="PicWidth" value="100">
      如果投票中有图片，此项有效</label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="确定创建此标签">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" 取 消 "></td>
    </tr>
  </table>
	<label></label>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:VS=VSLIST┆';
		retV+='投票项$' + obj.TID.value+'┆';
		retV+='图片宽度$' + obj.PicWidth.value+'┆';
		retV+='SpanId$' + obj.spanid.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
 </form>
</body>
</html>







