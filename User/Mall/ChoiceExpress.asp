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
<tr><td class="xingmu" colspan="4">物流公司选择</td></tr>
<tr>
<td width="5%" class="hback_1"></td>
<td class="hback_1">公司</td>
<td class="hback_1">公司地址</td>
<td class="hback_1">收费标准(为参考价格，不做为购买费用，产品费用已包含此费用)</td>
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
<td colspan="4" class="hback">缺货处理：
<input type="radio" name="LackDeal" value="1" />退钱
<input type="radio" name="LackDeal" value="2" />有货后追发
<input type="radio" name="LackDeal" value="3" checked="checked"/>将钱暂存于商家，用于以后购买其他物品
</td>
</tr>
<tr>
<td colspan="4" class="hback">购买方式：
<input type="radio" name="M_Type" value="0" checked="checked"/>邮寄
<input type="radio" name="M_Type" value="1" />送货上门
<input type="radio" name="M_Type" value="2" />上门取货
&nbsp;&nbsp;<button onClick="javascript:Element.toggle('UserInfo')">修改收货信息</button>
</td>
</tr>
<tr>
<td align="center" class="hback" colspan="4">
	<table class="table" border="0" width="98%" id="UserInfo" style="display:none">
	<tr>
		<td class="hback" align="right" width="20%">收货人姓名:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.realname%>" width="50%" name="username" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">收货人性别:</td>
		<td class="hback" align="left">
		<input type="radio" value="0" name="sex" <%if Fs_User.sex=0 then Response.Write("checked")%>/>男
		<input type="radio" value="1" name="sex" <%if Fs_User.sex=1 then Response.Write("checked")%>/>女
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">省:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Province%>" width="50%" name="M_Province" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">市:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.city%>" width="50%" name="M_City" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">地址:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Address%>" width="50%" name="M_Address" />
		</td>
	</tr>
		<td class="hback" align="right">邮政编码:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.PostCode%>" width="50%" name="M_PostCode" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">电话:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Tel%>" width="50%" name="M_Tel" />
		</td>
	</tr>
	<tr>
		<td class="hback" align="right">移动电话:</td>
		<td class="hback" align="left">
		<input type="input" value="<%=Fs_User.Mobile%>" width="50%" name="Mobile" />
		</td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td colspan="4" class="hback">支付方式：
<span style="display:"><input type="radio" name="M_PayStyle" value="0" checked="checked" />
在线支付</span><input type="radio" name="M_PayStyle" value="1" />电汇（银行汇款）
<input type="radio" name="M_PayStyle" value="2" />邮寄
<input name="M_PayStyle" type="radio" value="3" />
金币购买(帐户购买)
<input type="radio" name="M_PayStyle" value="4" />点卡
</td>
</tr>
<tr>
<td colspan="4" align="center" class="hback"><button onclick="makeOrder()">&nbsp;提交定单 &gt;&gt; 完成后到定单管理支付</button></td>
</tr>
</table>
<%
ExpressRs.close
Set ExpressRs=nothing
Set Fs_User = Nothing
%>





