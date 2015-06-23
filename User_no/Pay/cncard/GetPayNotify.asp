<!-- #include file="MD5.asp" -->
<%
'*******************************************
'文件名：GetPayNotify.asp
'主要功能：该示范程序主要完成接收云网支付网关支付通知信息，验证信息有效性，判断支付结果处理商务逻辑并给云网通知反馈
'版本：v1.7（Build2008-09-22）
'说明：
'	1.本页面请不要使用诸如response.redirect等页面转向的语句
'	2.请直接将通知反馈结果以XML的形式输出在本页，云网支付@网会解析您的输出结果
'	3.请务必在本页面进行处理商户自己的商务规则，进行发货等系列操作
'版权所有：北京云网无限网络技术有限公司
'技术支持联系方式：86-010-82888778-8054 转技术部
'*******************************************

'---获取云网支付网关向商户发送的支付通知信息(以下简称为通知信息)
	Const PayHandleURL,c_pass
	PayHandleURL	= "http://www.yoursitename.com/urlpath/GetPayHandle.asp"	'商户给用户显示处理结果页面的URL(对应范例文件：GetPayHandle.asp)	 	
	c_pass			= "Test"					'商户的支付密钥，登录商户管理后台(https://www.cncard.net/admin/)，在管理首页可找到该值

	Dim c_mid,c_order,c_orderamount,c_ymd,c_transnum,c_succmark,c_moneytype,c_cause,c_memo1,c_memo2,c_signstr
	c_mid			= request("c_mid")			'商户编号，在申请商户成功后即可获得，可以在申请商户成功的邮件中获取该编号
	c_order			= request("c_order")		'商户提供的订单号
	c_orderamount	= request("c_orderamount")	'商户提供的订单总金额，以元为单位，小数点后保留两位，如：13.05
	c_ymd			= request("c_ymd")			'商户传输过来的订单产生日期，格式为"yyyymmdd"，如20050102
	c_transnum		= request("c_transnum")		'云网支付网关提供的该笔订单的交易流水号，供日后查询、核对使用；
	c_succmark		= request("c_succmark")		'交易成功标志，Y-成功 N-失败			
	c_moneytype		= request("c_moneytype")	'支付币种，0为人民币
	c_cause			= request("c_cause")		'如果订单支付失败，则该值代表失败原因		
	c_memo1			= request("c_memo1")		'商户提供的需要在支付结果通知中转发的商户参数一
	c_memo2			= request("c_memo2")		'商户提供的需要在支付结果通知中转发的商户参数二
	c_signstr		= request("c_signstr")		'云网支付网关对已上信息进行MD5加密后的字符串

	'---校验信息完整性---
		IF c_mid="" or c_order="" or c_orderamount="" or c_ymd="" or c_moneytype="" or c_transnum="" or c_succmark="" or c_signstr="" THEN
			'支付信息有误		
			Call WriteNotice()			
		END IF

	'---将获得的通知信息拼成字符串，作为准备进行MD5加密的源串，需要注意的是，在拼串时，先后顺序不能改变
		
		
		srcStr = c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & c_pass

	'---对支付通知信息进行MD5加密
		r_signstr	= MD5(srcStr)

	'---校验商户网站对通知信息的MD5加密的结果和云网支付网关提供的MD5加密结果是否一致
		IF r_signstr<>c_signstr Then
			'签名验证失败			
			Call WriteNotice()			
		END IF

	'---校验商户编号
		Dim MerchantID	'商户自己的编号
		IF MerchantID<>c_mid Then			
			'提交的商户编号有误				
			Call WriteNotice()			
		END IF

	'---校验商户订单系统中是否有通知信息返回的订单信息
		Dim conn	'商户系统的数据链接
		sql="select top 1 数据列 from 商户的订单表 where 商户订单号="& c_order
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn
		IF rs.eof THEN
			'未找到该订单信息		
			Call WriteNotice()			
		END IF

	'---校验商户订单系统中记录的订单金额和云网支付网关通知信息中的金额是否一致
		Dim r_orderamount	'商户自己系统记录的订单金额
		r_orderamount=rs("订单金额")	'商户从自己订单系统获取该值
		IF ccur(r_orderamount)<>ccur(c_orderamount) THEN
			'支付金额有误
			Call WriteNotice()			
		END IF

	'---校验商户订单系统中记录的订单生成日期和云网支付网关通知信息中的订单生成日期是否一致
		Dim r_ymd	'商户自己系统记录的订单生成日期
		r_ymd=rs("订单生成日期")	'商户从自己订单系统获取该值
		IF r_ymd<>c_ymd THEN
			'订单时间有误			
			Call WriteNotice()			
		END IF

	'---校验商户系统中记录的需要在支付结果通知中转发的参数和云网支付网关通知信息中提供的参数是否一致
		Dim r_memo1	'商户自己系统记录的需要在支付结果通知中转发的参数一
		r_memo1 = rs("转发参数一")
		Dim r_memo2	'商户自己系统记录的需要在支付结果通知中转发的参二
		r_memo2 = rs("转发参数二")
		IF r_memo1<>c_memo1 or r_memo2<>c_memo2 THEN
			'参数提交有误
			Call WriteNotice()			
		END IF

	'---校验返回的支付结果的格式是否正确
		IF c_succmark<>"Y" and c_succmark<>"N" THEN
			'参数提交有误	
			Call WriteNotice()			
		END IF

	'---根据返回的支付结果，商户进行自己的发货等操作
		IF c_succmark="Y" Then			' c_succmark="Y"代表支付成功			
			'根据商户自己商务规则，进行发货等系列操作.如果此过程中出现错误需要停止代码的继续执行，则直接调用WriteNotice()这个过程，不要在页面输出错误信息。			
		Else
			'根据商户自己商务规则，进行支付失败时的系列操作，如果此过程中出现错误需要停止代码的继续执行，则直接调用WriteNotice()这个过程，不要在页面输出错误信息。			
		END If	
	'---通知云网已经收到通知，云网在收到您返回的信息后会自动解析您返回的信息并为用户显示出您设定的PayHandleURL的内容
		Call WriteNotice()

	'---输出支付结果通知反馈，告诉云网您已经收到了通知
		Sub WriteNotice()
			'<result>：值固定为1，表示商户已成功收到网关的支付成功的通知。
			'<reURL> ：商户显示给用户处理结果页面的URL(对应范例文件：GetPayHandle.asp)
			Response.Write("<result>1</result><reURL>" & PayHandleURL & "</reURL>")
			Response.End()
		End Sub	
%>