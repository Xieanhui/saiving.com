<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp"-->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Response.Charset="gb2312"
Dim ExpressRs,PriceArray,i
Fs_User.Name=session("FS_UserName")
i=0
Set ExpressRs=Conn.execute("Select ComID,ComName,ComDescryption,ComAddress,ComWebSite,ComPrice from FS_MS_Company where IsLock=0 and ComClass=2")
%>
<table width="98%" border=0 align="center" cellPadding=5 cellSpacing=1 background="" class="table">
<tr><td class="xingmu" colspan="4">������˾ѡ��</td></tr>
<tr>
<td width="5%" class="hback_1"></td>
<td class="hback_1">��˾</td>
<td class="hback_1">��˾��ַ</td>
<td class="hback_1">�շѱ�׼(Ϊ�ο��۸񣬲���Ϊ������ã���Ʒ�����Ѱ����˷���)</td>
</tr>
<%
while not ExpressRs.eof
	Response.Write("<tr>"&vbcrlf)
	if i=0 then
		Response.Write("<td class=""hback"" align=""center""><input type=""radio"" name=""ExpressCompany"" value="""&ExpressRs("ComID")&""" checked=""checked""/></td>"&vbcrlf)
	Else
		Response.Write("<td class=""hback"" align=""center""><input type=""radio"" name=""ExpressCompany"" value="""&ExpressRs("ComID")&"""/></td>"&vbcrlf)
	End if
	i=i+1
	Response.Write("<td class=""hback""><a href=""http://"&ExpressRs("ComWebSite")&""" target=""_blank"">"&ExpressRs("ComName")&"</a></td>"&vbcrlf)
	Response.Write("<td class=""hback"">"&ExpressRs("ComAddress")&"</td>"&vbcrlf)
	Response.Write("<td class=""hback"">"&vbcrlf)
	Response.Write("<select name=""sel_express"">"&vbcrlf)
	if not ExpressRs("ComPrice")="" and not isNull(ExpressRs("ComPrice"))  then
		PriceArray=split(""&ExpressRs("ComPrice"),",")
		if isArray(PriceArray) then
			for i=0 to Ubound(PriceArray)
				Response.Write("<option>"&FormatCurrency(PriceArray(i))&"</option>"&vbcrlf)
			next
		End if
	End if
	Response.Write("<select>"&vbcrlf)
	Response.Write("</td>"&vbcrlf)
	Response.Write("</tr>"&vbcrlf)
	ExpressRs.movenext
Wend
%>
<tr>
<td colspan="4" class="hback">ȱ������
<input type="radio" name="LackDeal" value="1" />��Ǯ
<input type="radio" name="LackDeal" value="2" />�л���׷��
<input type="radio" name="LackDeal" value="3" checked="checked"/>��Ǯ�ݴ����̼ң������Ժ���������Ʒ
</td>
</tr>
<tr>
<td colspan="4" class="hback">����ʽ��
<input type="radio" name="M_Type" value="0" checked="checked"/>�ʼ�
<input type="radio" name="M_Type" value="1" />�ͻ�����
<input type="radio" name="M_Type" value="2" />����ȡ��
&nbsp;&nbsp;<button onClick="javascript:Element.toggle('UserInfo')">�޸��ջ���Ϣ</button>
</td>
</tr>
<tr>
<td align="center" class="hback" colspan="4">
	<table class="table" border="0" width="98%" id="UserInfo" style="display:none">
	<tr>
		<td class="hback" align="right" width="20%">�ջ�������:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.realname%>" width="50%" name="username" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">�ջ����Ա�:</td>
		<td class="hback" align="left">
		<input type="radio" value="0" name="sex" <%if Fs_User.sex=0 then Response.Write("checked")%>/>��
		<input type="radio" value="1" name="sex" <%if Fs_User.sex=1 then Response.Write("checked")%>/>Ů
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">ʡ:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Province%>" width="50%" name="M_Province" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">��:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.city%>" width="50%" name="M_City" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">��ַ:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Address%>" width="50%" name="M_Address" />
		</td>
	</tr>
		<td class="hback" align="right">��������:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.PostCode%>" width="50%" name="M_PostCode" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">�绰:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Tel%>" width="50%" name="M_Tel" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">�ƶ��绰:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Mobile%>" width="50%" name="Mobile" />
		</td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td colspan="4" class="hback">֧����ʽ��
<span style="display:"><input type="radio" name="M_PayStyle" value="0" checked="checked" />
����֧��</span><input type="radio" name="M_PayStyle" value="1" />��㣨���л�
<input type="radio" name="M_PayStyle" value="2" />�ʼ�
<input name="M_PayStyle" type="radio" value="3" />
��ҹ���(�ʻ�����)
<input type="radio" name="M_PayStyle" value="4" />�㿨
</td>
</tr>
<tr>
<td colspan="4" align="center" class="hback"><button onclick="makeOrder()">&nbsp;�ύ���� &gt;&gt; ��ɺ󵽶�������֧��</button></td>
</tr>
</table>
<%
ExpressRs.close
Set ExpressRs=nothing
Set Fs_User = Nothing
%>





