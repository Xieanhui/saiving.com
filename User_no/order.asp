<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if Request.QueryString("Action") = "lock_order" then
	'进行判断
	dim rsOrderTF
	set rsOrderTF =User_Conn.execute("select M_state,IsSuccess From FS_ME_Order where OrderNumber='"& NoSqlHack(Request.QueryString("OrderNumber"))&"' and UserNumber='"& Fs_User.UserNumber &"'")
	if not rsOrderTF.eof then
		if rsOrderTF("M_state")=1 then
			strShowErr = "<li>已经处理的定单，您不能取消!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if 
		if rsOrderTF("IsSuccess")=1 then
			strShowErr = "<li>支付成功的定单，您不能取消!</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if 
	end if
	rsOrderTF.close:set rsOrderTF = nothing
	User_Conn.execute("Delete From FS_ME_Order  where OrderNumber='"& NoSqlHack(Request.QueryString("OrderNumber"))&"' and UserNumber='"& Fs_User.UserNumber &"'")
	User_Conn.execute("Delete From FS_ME_Order_detail  where OrderNumber='"& NoSqlHack(Request.QueryString("OrderNumber"))&"'")
	strShowErr = "<li>操作定单成功!</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%>-定单</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt;定单</td>
        </tr>
        <tr class="hback">
          <td class="hback"><a href="Order.asp">一般定单</a>┆<a href="Order_Pay.asp">在线支付定单</a></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="Order.asp">
          <tr class="hback"> 
            <td colspan="6" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="44%"> <strong> 
                    <%
				  dim strTmp,strLogType,strTmp1,OrderNumber
				  strLogType = NoSqlHack(Request.QueryString("LogTye"))
				  OrderNumber =  NoSqlHack(Request.QueryString("OrderNumber"))
				  if OrderNumber<>"" then OrderNumber = " and OrderNumber='"&OrderNumber&"' "
			     if Request.QueryString("LogTye")<>"" then
			  		strTmp =  " and LogType='"& strLogType &"'"
			     Else
			  		strTmp =  " "
			    End if
				Dim RsOrderObj,RsOrderSQL
				Dim strpage,strSQLs,StrOrders
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsOrderObj = Server.CreateObject(G_FS_RS)
				RsOrderSQL = "Select * From FS_ME_Order  where UserNumber='"& Fs_User.UserNumber &"' and OrderType<>3 "&OrderNumber&" order by  OrderID desc"
				RsOrderObj.Open RsOrderSQL,User_Conn,1,1
				Response.Write RsOrderObj.recordcount
				%>
                    </strong> 个定单</td>
                  <td width="56%"><div align="left"> </div></td>
                </tr>
              </table></td>
          </tr class="hback">
          <tr class="hback"> 
            <td width="20%" class="xingmu"><div align="left"><strong> 定单号(点定单查看详情)</strong></div></td>
            <td width="11%" class="xingmu"><div align="center">审核状态</div></td>
            <td width="21%" class="xingmu"><div align="center">成功日期</div></td>
            <td width="18%" class="xingmu"><strong>日期</strong></td>
            <td width="9%" class="xingmu"><div align="center"><strong>类型</strong></div></td>
            <td width="13%" class="xingmu"><div align="center"><strong>支付</strong></div></td>
          </tr>
          <%
		Dim select_count,select_pagecount,i
		if RsOrderObj.eof then
			   RsOrderObj.close
			   set RsOrderObj=nothing
			   set conn=nothing
			   set fs_user=nothing
			   Response.Write"<TR><TD colspan=""10""  class=""hback"">没有记录。</TD></TR>"
		else
				if Request("CountPage")="" or len(Request("CountPage"))<1 then
					RsOrderObj.pagesize = 20
				Else
					RsOrderObj.pagesize = CintStr(Request("CountPage"))
				End if
				RsOrderObj.absolutepage=CintStr(strpage)
				select_count=RsOrderObj.recordcount
				select_pagecount=RsOrderObj.pagecount
				for i=1 to RsOrderObj.pagesize
					if RsOrderObj.eof Then exit For 
		 %>
          <tr class="hback"> 
            <td class="hback"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Orderid<% = RsOrderObj("OrderID")%>);"  language=javascript><div align="left"> 
                <% = RsOrderObj("OrderNumber")%>
              </div></td>
            <td class="hback"> <div align="center"> 
                <%
					if RsOrderObj("isLock")=1 then
						Response.Write("<a href=""Order.asp?Action=lock_order&type=1&OrderNumber="& RsOrderObj("OrderNumber") &""" onClick=""{if(confirm('确定取消定单？')){return true;}return false;}"" title=""取消定单""><font color=red>审核中...</font></a>")
					Else
						Response.Write("已审核...")
					End if
					%>
              </div></td>
            <td class="hback"><div align="center"> 
                <% = RsOrderObj("M_PayDate")%>
              </div></td>
            <td class="hback"> 
              <% = RsOrderObj("AddTime")%>
            </td>
            <td class="hback"><div align="center"> 
                <%
			if RsOrderObj("OrderType")=0 then
				Response.Write("会员组")
			Elseif RsOrderObj("OrderType")=1 then
				Response.Write("商品")
			Elseif RsOrderObj("OrderType")=2 then
				Response.Write("点卡")
			Elseif RsOrderObj("OrderType")=3 then
				Response.Write("在线支付")
			Else
				Response.Write("其他")
			End if
			%>
              </div></td>
            <td class="hback"> <div align="center"> 
                <%
				if RsOrderObj("IsSuccess")=0 then
				%>
                <font color="#FF0000">未支付</font> 
                <%Else%>
                已支付 
                <%End if%>
              </div></td>
          </tr>
          <tr class="hback"  id="Orderid<% = RsOrderObj("OrderID")%>" style="display:none"> 
            <td height="106" colspan="6" class="hback"> <table width="100%" border="0" cellspacing="1" cellpadding="5" class="table">
                <tr class="hback"> 
                  <td width="11%" class="hback_1"><div align="center">产品</div></td>
                  <td colspan="3" class="hback"> <div align="left"> 
					<%
					Dim tmp_rs,tmp_SQL,tmp_i,sum_Moeny,p_rs
					Set tmp_rs = Server.CreateObject(G_FS_RS)
					tmp_SQL = "Select [DetailID],OrderNumber,ProductID,ProductNumber,M_state,Moneys From FS_ME_Order_detail  where OrderNumber='"& RsOrderObj("OrderNumber") &"' order by  DetailID desc"
					tmp_rs.Open tmp_SQL,User_Conn,1,3
					sum_Moeny = 0 
					for tmp_i = 1 to tmp_rs.recordcount
						if tmp_rs.eof then exit for
						if RsOrderObj("OrderType")=1 then
							set p_rs = Conn.execute("select ProductTitle From FS_MS_Products where id="&CintStr(tmp_rs("ProductID")))
							if not p_rs.eof then
								Response.Write "·<a href="& get_productsLink(tmp_rs("ProductID")) &">"&p_rs("ProductTitle") &"</a><br>"
							else
								Response.Write "·商品已经删除<br>"
							end if
						else
							Response.Write "<br />"
						end if
						sum_Moeny = sum_Moeny + tmp_rs("Moneys")
					tmp_rs.moveNext
					next
					%>
                    </div></td>
                  <td class="hback_1"><div align="center">总金额</div></td>
                  <td class="hback"> 
                    <% = formatCurrency(sum_Moeny) %>
                  </td>
                </tr>
                <tr class="hback"> 
                  <td class="hback_1"><div align="center">联系电话</div></td>
                  <td width="22%" class="hback"> 
                    <% = RsOrderObj("M_Tel")%>
                  </td>
                  <td width="9%" class="hback_1">移动电话</td>
                  <td width="26%" class="hback"> 
                    <% = RsOrderObj("M_Mobile")%>
                  </td>
                  <td width="9%" class="hback_1"><div align="center">性别</div></td>
                  <td width="23%" class="hback"> 
                    <%
				  if  RsOrderObj("M_Sex") = 0 then
				  		Response.Write("男")
					Else
				  		Response.Write("女")
					End if
				  %>
                  </td>
                </tr>
                <tr class="hback"> 
                  <td class="hback_1"><div align="center">购买方式</div></td>
                  <td class="hback"> 
                    <%
				  if  RsOrderObj("M_Type")=0 then
				  		Response.Write("邮寄")
				  Elseif RsOrderObj("M_Type") =1 then
				  		Response.Write("电汇（送货上门）")
				  Elseif RsOrderObj("M_Type") =1 then
				  		Response.Write("电汇（送货上门）")
				  Else
				  		Response.Write("上门取货")
				  End if
				  %>
                  </td>
                  <td class="hback_1"><div align="center">地址</div></td>
                  <td colspan="3" class="hback"> <div align="left"> 
                      <% = RsOrderObj("M_Province")%>
                      <% = RsOrderObj("M_City")%>
                      <% = RsOrderObj("M_Address")%>
                      　　邮编: 
                      <% = RsOrderObj("M_PostCode")%>
                    </div></td>
                </tr>
                <tr class="hback"> 
                  <td class="hback_1"><div align="center">收货人</div></td>
                  <td class="hback"> 
                    <% = RsOrderObj("M_UserName")%>
                  </td>
                  <td class="hback_1"><div align="center">物流公司</div></td>
                  <td class="hback"> 
                    <% = RsOrderObj("M_ExpressCompany")%>
                  </td>
                  <td class="hback">支付方式</td>
                  <td class="hback"> 
                  <%
				  if RsOrderObj("M_PayStyle") =0 then
				  		Response.Write("在线支付")	
				  Elseif RsOrderObj("M_PayStyle") =1 then
				  		Response.Write("电汇（银行汇款）")
				  Elseif RsOrderObj("M_PayStyle") =2 then
				  		Response.Write("邮寄")
				  Elseif RsOrderObj("M_PayStyle") =3 then
				  		Response.Write("帐户支付(金币)")
				  Else
				  		Response.Write("点卡")
				  End if
				  %>
                  </td>
                </tr>
                <tr class="hback"> 
                  <td class="hback_1"><div align="center">发货状态</div></td>
                  <td class="hback"> 
                    <%
					if RsOrderObj("M_state")=0 then
						Response.Write("未发货")
					elseif RsOrderObj("M_state")=1 then
						Response.Write("已发货")
						if RsOrderObj("OrderType")=1 then 
							''是商品则显示退货申请
							response.Write(" | <a href=""Mall/WithDraw_Apply.asp?Act=Add&OrderNumber="&RsOrderObj("OrderNumber")&""" title=""若你需要退货可点此进入退货申请或退货管理"">我要退货</a>")
						end if
					End if
					%>
                  </td>
                  <td class="hback_1"><div align="center">备注</div></td>
                  <td class="hback"> 
                    <% = RsOrderObj("Content")%>
                  </td>
                  <td class="hback_1"><div align="center">支付</div></td>
                  <td class="hback"> 
					<%
					if RsOrderObj("IsSuccess")=0 then
						if RsOrderObj("M_PayStyle") =0 then
						%>
						<a href="onlinepay.asp?OrderNumber=<% = RsOrderObj("OrderNumber")%>&Moneys=<%=sum_Moeny%>&OrderID=<%=RsOrderObj("OrderID")%>"><strong><font color="#FF0000">支付</font></strong></a> 
						<%
						Elseif RsOrderObj("M_PayStyle") =1 or RsOrderObj("M_PayStyle") =2 then
						%>
						<a href="PayCenter.asp?PayStyle=PostOrBank&OrderNumber=<% = RsOrderObj("OrderNumber")%>&Moneys=<%=sum_Moeny%>&OrderID=<%=RsOrderObj("OrderID")%>"><strong><font color="#FF0000">支付</font></strong></a> 
						<%
						Elseif RsOrderObj("M_PayStyle") =3 then
						%>
						<a href="PayCenter.asp?PayStyle=MySelfAcc&OrderNumber=<% = RsOrderObj("OrderNumber")%>&Moneys=<%=sum_Moeny%>&OrderID=<%=RsOrderObj("OrderID")%>"><strong><font color="#FF0000">支付</font></strong></a> 
						<%
						Else
						%>
						<a href="PayCenter.asp?PayStyle=Card&OrderNumber=<% = RsOrderObj("OrderNumber")%>&Moneys=<%=sum_Moeny%>&OrderID=<%=RsOrderObj("OrderID")%>"><strong><font color="#FF0000">支付</font></strong></a> 
						<%
						End if
					Else
					%>
					已支付 
					<%
					End if
					%>
                  </td>
                </tr>
              </table></td>
          </tr>
          <%
			  RsOrderObj.MoveNext
		  Next
		  %>
          <tr class="hback"> 
            <td colspan="6" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="80%"> <span class="top_navi"> 
                    <% 	Response.Write("每页:"& RsOrderObj.pagesize &"个,")
							Response.write"&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&LogType="&Request("LogTye")&">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&LogType="&Request("LogTye")&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&LogType="&Request("LogTye")&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&LogType="&Request("LogTye")&">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsOrderObj.close
								Set RsOrderObj=nothing
							End if
							%>
                    </SPAN></td>
                </tr>
              </table></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
'得到商品栏目地址
function get_productsLink(f_id)
		MFConfig_Cookies
		get_productsLink = ""
		dim rs,config_rs,config_mf_rs,class_rs
		dim SaveproductsPath,fileName,FileExtName,ClassId,IsDomain,LinkType,Mf_Domain,Url_Domain,ClassEName,c_Domain,c_SavePath
		set rs = Conn.execute("select ID,ClassId,SavePath,fileName,fileExtName from fS_MS_products where Id="&CintStr(f_id))
		SaveproductsPath = rs("SavePath")
		fileName = rs("fileName")
		fileExtName = rs("fileExtName")
		ClassId = rs("ClassId")
		set config_rs = Conn.execute("select top 1 IsDomain from fS_MS_SysPara")
		IsDomain = config_rs("IsDomain")
		LinkType = "1"
		config_rs.close:set config_rs=nothing
		Mf_Domain = Request.Cookies("foosunMfCookies")("foosunMfDomain")
		set class_rs = Conn.execute("select ClassEName,IsURL,URLAddress,[Domain],SavePath from fS_MS_productsClass where ClassId='"&NoSqlHack(ClassId)&"'")
		if not class_rs.eof then
			ClassEName = class_rs("ClassEName")
			c_Domain = class_rs("Domain")
			c_SavePath = class_rs("SavePath")
			class_rs.close:set class_rs=nothing
		else
			ClassEName = ""
			class_rs.close:set class_rs=nothing
		end if
		if not rs.eof then
			if trim(c_Domain)<>"" then
				Url_Domain = "http://"&c_Domain
			else
				Url_Domain = ""
			end if
			get_productsLink = Url_Domain & c_SavePath& "/" & ClassEName &SaveproductsPath &"/"&fileName&"."&fileExtName
		rs.close:set rs=nothing
	  else
			get_productsLink = ""
			rs.close:set rs=nothing
	  end if
	  get_productsLink = get_productsLink
End function
Set Fs_User = Nothing
set user_conn=nothing

%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





