<!--#include file="Alipay_md5.asp" --><%
Class Alipay
	Public key
	Public partner
	public notify_url
	public return_url
	public show_url
	Public Function creatURL(subject,body,out_trade_no,price,quantity,seller_email,paymethod)
		Dim INTERFACE_URL
		Dim mystr,params,sign,itemURL,discount
		INTERFACE_URL = "https://www.alipay.com/cooperate/gateway.do?"
		mystr = Array("service=create_direct_pay_by_user","partner=" & partner,"subject=" & subject,"body=" & body,"out_trade_no=" & out_trade_no,"price=" & price,"discount=" & discount,"show_url=" & show_url,"quantity=" & quantity,"payment_type=1","seller_email=" & seller_email,"paymethod=" & paymethod,"notify_url=" & notify_url,"return_url=" & return_url)
		params = sort_blank(mystr)
		sign    = md5(params & key)
		itemURL = INTERFACE_URL & URLEncode(params) & "&sign=" & sign & "&sign_type=" & "MD5"
		'Response.Write(Server.HTMLEncode(URLEncode(params)))
		creatURL = itemURL
	End Function
	
	Public Function DelStr(Str)

		If IsNull(Str) Or IsEmpty(Str) Then
			Str = ""
		End If

		DelStr = Replace(Str,";","")
		DelStr = Replace(DelStr,"'","")
		DelStr = Replace(DelStr,"&","")
		DelStr = Replace(DelStr," ","")
		DelStr = Replace(DelStr,"　","")
		DelStr = Replace(DelStr,"%20","")
		DelStr = Replace(DelStr,"--","")
		DelStr = Replace(DelStr,"==","")
		DelStr = Replace(DelStr,"<","")
		DelStr = Replace(DelStr,">","")
		DelStr = Replace(DelStr,"%","")
	End Function
	
	Public Function Notify()
		Dim ResponseTxt,mysign
	
		ResponseTxt = CheckNotifyRequest(Request.QueryString)
		mysign = GetSign(Request.QueryString)

		'*************************交易状态返回处理*************************

		If mysign = Request.Form("sign") And ResponseTxt = "true" Then

			If Request.Form("trade_status") = "TRADE_FINISHED" Then
				
				Notify = "success"

			Else
				Notify = "fail"
			End If
		Else
			Notify= "fail"
		End If
	End Function
	
	Public Function Log(messages)
		dim fs,ts,currentPath
		on error resume next
		currentPath = Request.ServerVariables("URL")
		currentPath = Left(currentPath,InstrRev(currentPath,"/"))
		currentPath = currentPath&"Notify_DATA/"&replace(now(),":","")&".txt"
		currentPath = Server.MapPath(currentPath)
		set fs= createobject("scripting.filesystemobject")
		set ts=fs.createtextfile(currentPath,true)
		if err then
			Response.Write(currentPath)
		end if
		ts.writeline(messages)
		ts.close
		set ts=Nothing
		set fs=Nothing
	End Function
	
	Public Function CheckNotifyRequest(req)
		Dim alipayNotifyURL,Retrieval
		'**********************判断消息是不是支付宝发出********************
		alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
		alipayNotifyURL = alipayNotifyURL & "partner=" & partner & "&notify_id=" & req("notify_id")
		Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		Retrieval.setOption 2, 13056
		Retrieval.open "GET", alipayNotifyURL, False, "", ""
		Retrieval.send()
		CheckNotifyRequest   = Retrieval.ResponseText
		'Response.Write(Server.HTMLEncode(alipayNotifyURL)&vbcrlf&CheckNotifyRequest):Response.End()
		Set Retrieval = Nothing
		'*******************************************************************
	End Function
	
	Public Function NotifyPage()
		dim ResponseTxt,mysign
		ResponseTxt = CheckNotifyRequest(Request.QueryString)
		mysign = GetSign(Request.QueryString)
		'********************************************************
		'REsponse.Write(ResponseTxt):Response.End()
		If mysign = Request("sign") And ResponseTxt = "true"   Then
			NotifyPage="success"
		Else
			NotifyPage="fail"
		End If

	End Function
	
	Public Function GetSign(req)
		Dim varItem,mystr,Count,i,minmax,minmaxSlot,j,md5str
		'*******获取支付宝GET过来通知消息,判断消息是不是被修改过************
		For Each varItem in req
			mystr = varItem & "=" & req(varItem) & "^" & mystr
		Next

		If mystr <> "" Then
			mystr = Left(mystr,Len(mystr) - 1)
		End If
		mystr  = Split(mystr, "^")
		md5str = sort_blank(mystr)
		GetSign = md5(md5str & key)
		'Response.Write(Server.HTMLEncode(md5str)&"<br />"&GetSign):Response.End()
	End Function
	
	Private Function sort_blank(arr())
		Dim count,i,minmax,minmaxSlot,j,mark,temp,value,md5str
		count = Ubound(arr)
		'排序
		For i = count To 0 Step - 1
			minmax       = arr( 0 )
			minmaxSlot   = 0

			For j = 1 To i
				mark        = (arr( j ) > minmax)

				If mark Then
					minmax     = arr( j )
					minmaxSlot = j
				End If

			Next
			If minmaxSlot <> i Then
				temp = arr( minmaxSlot )
				arr( minmaxSlot ) = arr( i )
				arr( i ) = temp
			End If
		Next
		'删除空值
		md5str=""
		For j = 0 To count Step 1
			value = Split(arr( j ), "=")

			If  value(1) <> "" And value(0) <> "sign" And value(0) <> "sign_type"  Then

				If md5str="" Then
					md5str = arr( j )
				Else
					md5str = md5str & "&" & arr( j ) 
				End If

			End If

		Next
		sort_blank = md5str
	End Function
	
	Private Function URLEncode(url)
		Dim arr,c,arr1
		arr = Split(url,"&")
		for c=0 to ubound(arr)
			arr1=split(arr(c),"=")
			arr1(1) = Server.URLEncode(arr1(1))
			arr(c) = Join(arr1,"=")
		next
		URLEncode = Join(arr,"&")
	End Function

End Class

%>