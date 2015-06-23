<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
Dim Conn,User_Conn,VClass_Rs,VClass_Sql,sErrStr
Dim CheckStr,Need_Do_Pwd_Str
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("ME_Card") then Err_Show

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页

Function Get_Card(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "CardID,CardNumber,CardMoney,CardDateNumber,CardPoint,isBuy,CardOverDueTime,UserNumber,UserTime  ,AddTime,IsUse,CardPasswords"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_ME_Card"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(request(Str_Tmp))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "CardMoney","CardDateNumber","CardPoint","IsUse","isBuy"   ,"CardOverDueTime","UserTime","AddTime"
					''数字,日期
						regxp = "|<|>|=|<=|>=|<>|"
						if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
						elseif instr(Req_Str,"*")>0 then 
							if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
							if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
						else	
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
						end if		
					case else
					''字符
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	Str_Tmp = "":ii=0
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	On Error Resume Next
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	IF VClass_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
	VClass_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>VClass_Rs.PageCount Then cPageNo=VClass_Rs.PageCount 
	VClass_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""Card.asp?Act=Edit&CardID="&VClass_Rs("CardID")&""" class=""otherset"" title='点击修改'>"&VClass_Rs("CardNumber")&"</a></td>" & vbcrlf
		for ii=2 to 8
			select case ii
				case 2
				Str_Tmp = VClass_Rs(ii) & "元"
				case 5 
				if VClass_Rs(ii)=1 then 
					Str_Tmp="已售出"
				else
					Str_Tmp="未售出"
				end if		
				case else
				Str_Tmp = VClass_Rs(ii)
			end select		
				Get_Html = Get_Html & "<td align=""center"">"& Str_Tmp & "</td>" & vbcrlf
		next	
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" "&CheckStr&" name=""DelID"" id=""DelID"" value="""&VClass_Rs("CardID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		CheckStr = ""	
		VClass_Rs.MoveNext
 		if VClass_Rs.eof or VClass_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	VClass_Rs.close
	Get_Card = Get_Html
End Function

Sub Del()
	if not MF_Check_Pop_TF("ME019") then Err_Show '权限判断
	Dim Str_Tmp
	if request.QueryString("CardID")<>"" then 
		User_Conn.execute("Delete from FS_ME_Card where CardID = "&CintStr(request.QueryString("CardID")))
	else
		Str_Tmp = FormatIntArr(request.form("DelID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>")
		
		User_Conn.execute("Delete from FS_ME_Card where CardID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================
Function CheckCardCF(CardNumber)
''检查录入的卡是否重复,重复则返回卡号,不重复则返回""
	Dim CheckCardCF_Rs
	Set CheckCardCF_Rs = CreateObject(G_FS_RS)
	CheckCardCF_Rs.Open "select Count(*) from FS_ME_Card where CardNumber='"&NoSqlHack(CardNumber)&"'",User_Conn,1,1
	if  CheckCardCF_Rs(0)>0 then 
		CheckCardCF = CardNumber
	else 
		CheckCardCF = ""
	end if
	CheckCardCF_Rs.close	
End Function

Sub Save()
	Dim CardID,IsOk,Arr_Tmp,Str_Tmp,Arr_Tmp1,Str_Tmp1,Arr_Tmp2,Str_Tmp2,ErrInfo,tmpi
	Dim PutNum,CardAddStr,CardNum_Len,CardPwd_Len,Put_i,Put_Rs,Str_Tmp3,Put_Pwd_Type
    Dim Randchar,Randchararr,iR
	CardID = NoSqlHack(request.Form("CardID"))
	if not isnumeric(CardID) or CardID = "" then CardID = 0
	Str_Tmp = "CardNumber,CardPasswords,CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,isBuy,UserNumber,AddTime,UserTime"
	Arr_Tmp = split(Str_Tmp,",")
	for tmpi = 2 to 7 
		if trim(request.Form("frm_"&Arr_Tmp(tmpi))) = "" then sErrStr = Arr_Tmp(tmpi) : exit for
	next
	if sErrStr <>"" then response.Redirect("../error.asp?ErrCodes=<li>必要的参数 ["&sErrStr&"] 必须填写完整！！</li>")  : response.End()
	VClass_Sql = "select "&Str_Tmp&" from FS_ME_Card where CardID="&CardID
	Set VClass_Rs = CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,3,3
	if CardID >0 then 
	''修改
	if not MF_Check_Pop_TF("ME018") then Err_Show '权限判断
		for each Str_Tmp in Arr_Tmp
			if Str_Tmp = "CardPasswords" then 
				VClass_Rs(Str_Tmp) = Encrypt(NoSqlHack(request.Form("frm_"&Str_Tmp))) ''加密
			else
				if request.Form("frm_"&Str_Tmp)<>"" then 
					VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					VClass_Rs(Str_Tmp) = null
				end if	
			end if	
		next
		VClass_Rs.update
		VClass_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp?Act=Edit&CardID="&CardID )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	''新增
	  if not MF_Check_Pop_TF("ME017") then Err_Show '权限判断
	  if request.Form("ActPut")=1 then 
	  ''批量生成
	  	   Put_Pwd_Type = NoSqlHack(request.Form("Put_Pwd_Type"))
		   Put_Pwd_Type = replace(Put_Pwd_Type," ","")
		   ''卡的密码构成方式。数字，字母		   
		   PutNum=NoSqlHack(request.Form("Put_PutNum"))
		   if not isnumeric(PutNum) then PutNum = 1
		   CardAddStr=NoSqlHack(Trim(request.Form("Put_CardAddStr")))
		   CardNum_Len=NoSqlHack(request.Form("Put_CardNum_Len"))
		   if not isnumeric(CardNum_Len) then CardNum_Len = 12
		   CardPwd_Len=NoSqlHack(request.Form("Put_CardPwd_Len"))
		   if not isnumeric(CardPwd_Len) then CardPwd_Len = 8
		   ''如果前缀长度比冲值卡长度-2还大,则冲值卡长度=冲值卡长度+前缀长度＋２
		   if len(CardAddStr) > CardNum_Len-2 then CardNum_Len = len(CardAddStr) + CardNum_Len + 2
		 On Error Resume Next
		   if User_Conn.execute("select Count(*) from FS_ME_CardPut")(0)=0 then 
		   	User_Conn.execute("insert into FS_ME_CardPut (PutNum,CardAddStr,CardNum_Len,CardPwd_Len) VALUES ("&PutNum&",'"&CardAddStr&"',"&CardNum_Len&","&CardPwd_Len&")")
		   else
		   	User_Conn.execute("update FS_ME_CardPut set PutNum="&PutNum&",CardAddStr='"&CardAddStr&"',CardNum_Len="&CardNum_Len&",CardPwd_Len="&CardPwd_Len&"")
		   end if
		   If Err.Number<>0 Then Err.Clear : response.Redirect("../Error.asp?ErrCodes=<li>批量生成卡基础设置存盘失败.请检查拼写是否正确.</li>") : response.End()
		   Str_Tmp = "CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,AddTime,isBuy"
		   Arr_Tmp = split(Str_Tmp,",")	 
		 for Put_i = 1 to PutNum
		   Str_Tmp1 = "" :   Str_Tmp2 = "" :   Str_Tmp3 = ""	
		   ''生成卡号，密码并入库
		   ''''''''''''''添加到FS_ME_CardPut表
		   Str_Tmp2=Str_Tmp2&GetRamCode(CardPwd_Len)
		   'response.Write("密码"&Put_i&"次："&Str_Tmp2&"<br>")
		   ''''''''''''''''''''''''''''''''''
		   ''卡号加密，取消06-07-05
		   ''第一次
		   Str_Tmp1 = Str_Tmp1 &GetRamCode( CardNum_Len - len(CardAddStr) )
		   '''''''''''''''''''''''生成后再检查数据库,若存在则再生成一次
		   do while CheckCardCF(CardAddStr & Str_Tmp1)<>""		   
			   Str_Tmp1 = GetRamCode( CardNum_Len - len(CardAddStr) )			 
		   loop
		   'response.Write("卡号"&Put_i&"次："&Str_Tmp1&"<br>")
		   '''''''''''''''''''''''
		   On Error Resume Next
		   VClass_Rs.AddNew		   
		   ''''''''''''''添加到FS_ME_Card表
		   Str_Tmp3 = ""
		   VClass_Rs("CardNumber") = CardAddStr & Str_Tmp1
		   VClass_Rs("CardPasswords") = Encrypt(Str_Tmp2)
		   for each Str_Tmp3 in Arr_Tmp
				if request.Form("frm_"&Str_Tmp3)<>"" then 
					VClass_Rs(Str_Tmp3) = NoSqlHack(request.Form("frm_"&Str_Tmp3))
				else
					VClass_Rs(Str_Tmp3) = null
				end if	
				'response.Write(Str_Tmp3&":"&NoSqlHack(request.Form("frm_"&Str_Tmp3))&"<br>")				
		   next
		   VClass_Rs.update
		   'If Err.Number<>0 Then Err.Clear : response.Redirect("../Error.asp?ErrCodes=<li>卡数据存盘失败.请检查拼写是否正确.</li>") : response.End()
		   'response.Write("Str_Tmp1:"&Str_Tmp1&"  "&"Str_Tmp2:"&Str_Tmp2&"  "&"Str_Tmp3:"&Str_Tmp3&"<br>")
	     next
 	     VClass_Rs.close 
		 'response.End()
		 response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "User/Card.asp" )&"&ErrCodes=<li>恭喜，批量生成成功。</li>")		   
	  else
	  '''''''''''''''''''''  		
		if request.Form("AddMode")=1 then 
		  ''单张
			if CheckCardCF(NoSqlHack(request.Form("frm_CardNumber")))<>"" then 
				response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>卡号:"&request.Form("frm_CardNumber")&"已存在。</li>")
				response.End()
			end if
		  VClass_Rs.AddNew
		  for each Str_Tmp in Arr_Tmp
			if Str_Tmp = "CardPasswords" then 
				VClass_Rs(Str_Tmp) = Encrypt(NoSqlHack(request.Form("frm_"&Str_Tmp))) ''加密
			else
				if request.Form("frm_"&Str_Tmp)<>"" then 
					VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
				else
					VClass_Rs(Str_Tmp) = null
				end if	
			end if	
			'response.Write(NoSqlHack(request.Form("frm_"&Str_Tmp))&":"&Str_Tmp&"<br>")
    	  next
		  VClass_Rs.update
		  Put_i=1
		else
		  ''批量添加多张
		  Str_Tmp1 = Trim(request.form("More_Mode_Area"))
		  if Str_Tmp1="" or instr(Str_Tmp1,"|")=0 or len(Str_Tmp1)<3 then 
		  	response.Redirect("../Error.asp?ErrCodes=批量添加时，卡号和密码不能为空或格式不正确或长度不够。")	: response.End()	
		  else
		   ErrInfo = "" : Put_i=0 : iR=0  : Str_Tmp2 = ""
		   Str_Tmp = "CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,AddTime,isBuy"
		   Arr_Tmp = split(Str_Tmp,",")
		   if instr(Str_Tmp1,"&nbsp;") then Str_Tmp1 = left(Str_Tmp1,len(Str_Tmp1) - len("&nbsp;"))
		   Str_Tmp1 = replace(replace(Str_Tmp1,"&nbsp;",vbcr)," ","")
		   Arr_Tmp1 = split(Str_Tmp1,vbcr) ''每一行数据
		   ErrInfo = ErrInfo & "<li>你提交了: "&ubound(Arr_Tmp1)+1&" 张卡.</li>" 
		   '''''''''''''''''''''判断后再造字符串
		   for iR = lbound(Arr_Tmp1) to ubound(Arr_Tmp1)
		  	  Arr_Tmp2=split(Arr_Tmp1(iR),"|") ''拆分卡号和密码
			  if ubound(Arr_Tmp2)<>1 then 
				ErrInfo = ErrInfo & "<li>批量添加中："&asc(Arr_Tmp1(iR))&"未入库，原因：分割符“|”数量不正确。</li>" 
			  else
				if CheckCardCF( Trim(Arr_Tmp2(0)) ) <> "" then 
					ErrInfo = ErrInfo & "<li>卡号: "&Trim(Arr_Tmp2(0))&" 在库中已存在.未入库。</li>"
				else
					Str_Tmp2 = Str_Tmp2 & Trim(Arr_Tmp2(0)) & "|" & Arr_Tmp2(1) & vbcr
				end if
			  end if
		   next
		   
 		   Str_Tmp1 = Str_Tmp2  	      
		   
		  if Str_Tmp2 <> "" then 
		  
		   Arr_Tmp1 = split(Str_Tmp1,vbcr) ''判断后的每行数据
		   for iR = lbound(Arr_Tmp1) to ubound(Arr_Tmp1) - 1
			   	Arr_Tmp2=split(Arr_Tmp1(iR),"|") ''拆分卡号和密码
				''+=========================
				VClass_Rs.Addnew
				VClass_Rs("CardNumber") = Trim(Arr_Tmp2(0)) : VClass_Rs("CardPasswords") = Encrypt( Arr_Tmp2(1) )
		 		for each Str_Tmp in Arr_Tmp
					if request.Form("frm_"&Str_Tmp)<>"" then 
						VClass_Rs(Str_Tmp) = NoSqlHack(request.Form("frm_"&Str_Tmp))
					else
						VClass_Rs(Str_Tmp) = null
					end if	
				next
				VClass_Rs.update
				''+=========================
				Put_i = Put_i + 1
		    next
			Str_Tmp2 = replace(Str_Tmp2,vbcr,"<br />&nbsp;&nbsp;&nbsp;")
			ErrInfo = ErrInfo & "<li>共入库 "&Put_i&" 张卡.</li>" 
		   end if
		  end if
		end if
 	    VClass_Rs.close
		'response.End()
		if Put_i>0 then 
			response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode("User/Card.asp?Act=Add" )&"&ErrCodes="&server.URLEncode("<li>恭喜，新增成功。</li>详细信息:"&ErrInfo))
		else
			response.Redirect("../Error.asp?ErrCodes="&server.URLEncode(ErrInfo))
		end if		
	  end if
	'''''''''''''''''''''  
	end if
End Sub
''=========================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//点击标题排序
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;
function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&filterorderby=")==-1&&Old_Sql.indexOf("?filterorderby=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&filterorderby=" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?filterorderby=" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);
				break;	
			}
		}
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName):New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
-->
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return true;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback"> 
    <td class="xingmu" >冲值卡管理</td>
  </tr>
  <tr  class="hback"> 
    <td><a href="Card.asp?Act=View">管理首页</a> | <a href="Card.asp?Act=Add">新建</a> 
      | <a href="Card.asp?Act=Put">批量生成</a> | <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("IsUse=0"))%>">未使用</a> 
      | <a href="Card.asp?Act=View&Add_Sql=<%=Encrypt("IsUse=1")%>">已使用</a> 
	  | 
	  <% If G_IS_SQL_User_DB=1 then%>
	  <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("datediff(s,CardOverdueTime,'"&DateValue(Now())&"')>0"))%>">已过期</a> 
	  <%Else%>
	  <a href="Card.asp?Act=View&Add_Sql=<%=server.URLEncode(Encrypt("datediff('s',CardOverdueTime,'"&DateValue(Now())&"')>0"))%>">已过期</a>
	  <%End if%>
      | <a href="Card.asp?Act=Search">查询</a></td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "","View","SearchGo"
		View
	case "Add","Edit","Put"
		Add_Edit
	case "Save"
		Save
	case "Del"
		Del
	case "Search"
		Search
end select

'******************************************************************
Sub View()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" action="?Act=Del">
    <tr  class="hback"> 
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('CardNumber')" class="sd"><b>〖卡号〗</b></a> <span id="Show_Oder_CardNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardMoney')" class="sd"><b>面值</b></a> <span id="Show_Oder_CardMoney"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('CardDateNumber')" class="sd"><b>天数</b></a> <span id="Show_Oder_CardDateNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardPoint')" class="sd"><b>点数</b></a> <span id="Show_Oder_CardPoint"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('isBuy')" class="sd"><b>状态</b></a> <span id="Show_Oder_isBuy"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('CardOverDueTime')" class="sd"><b>过期时间</b></a> <span id="Show_Oder_CardOverDueTime"></span></td>
	  <td align="center" class="xingmu"><a href="javascript:OrderByName('UserNumber')" class="sd"><b>使用者</b></a> <span id="Show_Oder_UserNumber"></span></td>
      <td align="center" class="xingmu"><a href="javascript:OrderByName('UserTime')" class="sd"><b>冲值时间</b></a> <span id="Show_Oder_UserTime"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_Card( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
  </form>
</table>
<%End Sub

Sub Add_Edit()
Dim CardID,Bol_IsEdit,isuse,UserTime
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then 
	CardID = request.QueryString("CardID")
	if CardID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的CardID没有提供</li>") : response.End()
	VClass_Sql = "select CardID,CardNumber,CardPasswords,CardMoney,CardDateNumber,CardPoint,CardOverDueTime,IsUse,UserNumber,UserTime,AddTime,isBuy from FS_ME_Card where CardID="&CardID
	Set VClass_Rs	= CreateObject(G_FS_RS)
	VClass_Rs.Open VClass_Sql,User_Conn,1,1
	if VClass_Rs.eof then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>没有相关的内容,或该内容已不存在.</li>") : response.End()
	Bol_IsEdit = True
	isuse = VClass_Rs("IsUse")
	UserTime = VClass_Rs("UserTime")
	if UserTime<>"" then if isdate(UserTime) then UserTime = formatdatetime(UserTime,2)
else
	isuse = 0	
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_Save" id="form_Save" onSubmit="return chkinput(this);" method="post" action="?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then response.Write("修改点卡信息<input type=""hidden"" name=""CardID"" value="""&VClass_Rs(0)&""">") else response.Write("新增点卡") end if%></td>
	</tr>
<!---------->
	<tr  class="hback" <%if Bol_IsEdit = True or request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>> 
      <td width="20%" align="right">添加方式</td>
      <td> 
        <input name="AddMode" type="radio" onClick="CNum.style.display='';CPwd.style.display='';More_Mode.style.display='none';frm_CardNumber.require='true';frm_CardPasswords.require='true';More_Mode_Area.require='false';" value="1" checked>
          单张冲值卡		  
        <input type="radio" name="AddMode" value="0" onClick="CNum.style.display='none';CPwd.style.display='none';More_Mode.style.display='';frm_CardNumber.require='false';frm_CardPasswords.require='false';More_Mode_Area.require='true';"> 
          批量添加冲值卡
      </td>
    </tr>
<!--多张-->
	<tr  class="hback" id="More_Mode" style="display:none;"> 
      <td width="20%" align="right">格式文本</td>
      <td> 
	  	<textarea name="More_Mode_Area" cols="45" rows="10" dataType="Require" min="5" max="10000" msg="长度必须在5-10000之间"></textarea>
        <br>
        分隔符为<span class="tx">|</span>，请按照每行一张卡，回车换行，每张卡格式为“<span class="tx">卡号|密码</span>”
        <br>
		<span class="tx">自动生成:
        <input type="text" name="PutCardNum" size="10" value="10" maxlength="8" title="只能数字" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
		<input type="button" value="Go" onClick="try{GetCodes(PutCardNum.value)} catch(e){return};"></span>
		</td>
    </tr>
<!--单张-->
	<tr  class="hback" id="CNum" <%if request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>> 
      <td width="20%" align="right">卡号</td>
      <td> 
        <input type="text" name="frm_CardNumber" size="40" value="<%if Bol_IsEdit then response.Write( VClass_Rs(1) ) else if request.QueryString("Act")<>"Search" then response.Write(GetRamCode(14)) end if end if%>"<%if request.QueryString("Act")="Put" then response.Write(" require=""false"" ") else response.Write(" require=""true"" ") end if%> dataType="LimitB" min="2" max="30" msg="长度必须在2-30之间">
        <span class="tx">不填则自动生成</span>
	  </td>
    </tr>
    <tr  class="hback" id="CPwd" <%if request.QueryString("Act")="Put" then response.Write(" style=""display='none'"" ") end if%>>
      <td align="right">卡号密码</td>
      <td>
	  <input type="text" name="frm_CardPasswords" size="40" value="<%if Bol_IsEdit then response.Write(Decrypt( VClass_Rs(2) )) else if request.QueryString("Act")<>"Search" then response.Write(GetRamCode(6)) end if end if%>"<%if request.QueryString("Act")="Put" then response.Write(" require=""false"" ") else response.Write(" require=""true"" ") end if%>dataType="LimitB" min="2" max="30" msg="长度必须在2-30之间">
      </td>
    </tr>
<!---->	
<!----批量生成时显示------>
<%if request.QueryString("Act")="Put" then
Dim Put_Rs,Put_Sql,Bol_Put_IsEdit
Bol_Put_IsEdit = false
Put_Sql = "select top 1 PutNum,CardAddStr,CardNum_Len,CardPwd_Len from FS_ME_CardPut order by PutID desc"
''取最后一次的设置
set Put_Rs=User_Conn.execute(Put_Sql)
if not Put_Rs.eof then Bol_Put_IsEdit = true
%>
	<tr  class="hback"> 
      <td width="20%" align="right">生成数量</td>
      <td> 
	    <input type="hidden" name="ActPut" value="1">
		<input type="text" name="Put_PutNum" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(0)) else response.Write("1") end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">冲值卡号前缀</td>
      <td> 
		<input type="text" name="Put_CardAddStr" onKeyUp="value=value.replace(/[^a-zA-Z0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9]/g,''))" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(1)) else response.Write("FS2006") end if%>" datatype="LimitB" min="-1" max="20" msg="前缀不能超过20个字符">
        默认为FS2006可为空</td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">冲值卡号长度</td>
      <td> 
		<input type="text" name="Put_CardNum_Len" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(2)) else response.Write("12") end if%>" dataType="Compare" to="1" msg="必须在区间[前缀长+2,30]以内" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">冲值卡密码长度</td>
      <td> 
		<input type="text" name="Put_CardPwd_Len" size="40" value="<%if Bol_Put_IsEdit then response.Write(Put_Rs(3)) else response.Write("8") end if%>" dataType="Compare" msg="必须>=0" to="0" operator="GreaterThanEqual">
	  </td>
    </tr>
	<tr  class="hback"> 
      <td width="20%" align="right">卡密码构成方式</td>
      <td> 
        <input name="Put_Pwd_Type" type="checkbox" value="1" checked<%'if Bol_Put_IsEdit then if instr(Put_Rs(4),"1")>0 then response.Write(" checked") end if else response.Write(" checked") end if%>>
        数字
        <input type="checkbox" name="Put_Pwd_Type" value="2" checked<%'if Bol_Put_IsEdit then if instr(Put_Rs(4),"2")>0 then response.Write(" checked") end if end if%>>
        字母
	  </td>
    </tr>
<%
Put_Rs.close
set Put_Rs=nothing
end if%>	
<!----批量生成时显示结束------>
    <tr  class="hback"> 
      <td align="right">点卡面值</td>
      <td>
	  <input type="text" name="frm_CardMoney" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(3)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
        元 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡天数</td>
      <td>
	  <input type="text" name="frm_CardDateNumber" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(4)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡点数</td>
      <td>
	  <input type="text" name="frm_CardPoint" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(5)) end if%>" onKeyUp="value=value.replace(/[^0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9]/g,''))">
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right"><span class="tx"><strong>点卡过期时间</strong></span></td>
      <td>
	  <input type="text" class="tx" name="frm_CardOverDueTime" size="27" value="<%if Bol_IsEdit then response.Write(VClass_Rs(6)) else response.Write( dateadd("d",30,date()) ) end if%>" readonly> <input name="SelectDate" type="button" id="SelectDate" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_CardOverDueTime);">	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡是否已使用</td>
      <td>
	  <select name="frm_IsUse" id="frm_IsUse" onChange="if(this.options[this.selectedIndex].value='0'){frm_UserNumber.value='';frm_UserTime.value=''}">
		<%=PrintOption(isuse,"1:已使用,0:未使用")%>
      </select>
        如果只是新增卡,请保持未使用</td>
    </tr>

    <tr  class="hback"> 
      <td align="right">点卡拥有者编号</td>
      <td>
	  <input type="text" name="frm_UserNumber" size="40" value="<%if Bol_IsEdit then response.Write(VClass_Rs(8)) end if%>">
        如果只是新增卡,请留空</td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">冲值日期</td>
      <td>
	  <input type="text" name="frm_UserTime" size="40" value="<%=UserTime%>">
        如果只是新增卡,请留空</td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡添加日期</td>
      <td>
	  <input type="text" name="frm_AddTime" size="27" value="<%if Bol_IsEdit then response.Write(VClass_Rs(10)) else response.Write(date()) end if%>" DataType="Require" msg="添加日期必须填写" readonly> <input name="SelectDate5" type="button" id="SelectDate5" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);"></td>
	  
    </tr>
    <tr  class="hback"> 
      <td align="right">是否销售</td>
      <td>
        <input type="radio" name="frm_isBuy" value="1" <%if Bol_IsEdit then if VClass_Rs(11)=1 then response.Write(" checked ") end if end if%>>
          已销售		  
        <input type="radio" name="frm_isBuy" value="0" <%if Bol_IsEdit then if VClass_Rs(11)=0 then response.Write(" checked ") end if else response.Write(" checked ") end if%>>
          未销售
	  </td>
	  
    </tr>
    <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 保存 " /> <!--<%IF request.QueryString("Act")="Put" then%> onClick="Put_CardNum_Len.to = (Put_CardAddStr.value.length+2).toString();Put_CardNum_Len.msg='长度必须大于等于'+(Put_CardAddStr.value.length+2).toString()" <%end if%>-->
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="if(confirm('确定删除该项目吗？')) location='<%=server.URLEncode("Card.asp?Act=Del&CardID="&CardID)%>'">
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%End Sub

Sub Search()
%>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" method="post" action="?Act=SearchGo">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >查询点卡</td>
    </tr>
    <tr  class="hback"> 
      <td width="20%" align="right">卡号</td>
      <td> <input type="text" name="frm_CardNumber" size="40" value="">
        模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">卡号密码</td>
      <td> <input type="password" name="frm_CardPasswords" size="40" value="">
        准确查询 </td>
    <tr  class="hback"> 
      <td align="right">点卡面值</td>
      <td> <select name="CardMoneyDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardMoney" size="30" value="">
      数字,可在开头加上简单比较符号,*号表示模糊查询</td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡天数</td>
      <td> <select name="CardDateNumberDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardDateNumber" size="30" value="">
      数字,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡点数</td>
      <td> <select name="CardPointDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_CardPoint" size="30" value="">
      数字,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡过期时间</td>
      <td>
        <!-- dataType="Date" msg="日期格式错误" require="false" -->
         <select name="CardOverDueTimeDate" style="width:55">
           <option value="" selected="selected"></option>
           <option value="*">*</option>
           <option value="&gt;">&gt;</option>
           <option value="&lt;">&lt;</option>
           <option value="=">=</option>
           <option value="&gt;=">&gt;=</option>
           <option value="&lt;=">&lt;=</option>
           <option value="&lt;&gt;">&lt;&gt;</option>
         </select>
         <input type="text" name="frm_CardOverDueTime" size="17" value="" readonly>
        <input name="SelectDate4" type="button" id="SelectDate4" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_CardOverDueTime);">
        日期,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡是否已使用</td>
      <td> <input type="radio" name="frm_IsUse"  value="1">
        已使用 
        <input type="radio" name="frm_IsUse"  value="0">
        未使用 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡拥有者编号</td>
      <td> <input type="text" name="frm_UserNumber" size="40" value="">
        模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">冲值日期</td>
      <td>  <select name="UserTimeDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_UserTime" size="17" value="" readonly=>
        <input name="SelectDate2" type="button" id="SelectDate2" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_UserTime);">
        日期,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">点卡添加日期</td>
      <td> <select name="AddTimeDate" style="width:55">
        <option value="" selected="selected"></option>
        <option value="*">*</option>
        <option value="&gt;">&gt;</option>
        <option value="&lt;">&lt;</option>
        <option value="=">=</option>
        <option value="&gt;=">&gt;=</option>
        <option value="&lt;=">&lt;=</option>
        <option value="&lt;&gt;">&lt;&gt;</option>
      </select>
         <input type="text" name="frm_AddTime" size="17" value="" readonly>
        <input name="SelectDate3" type="button" id="SelectDate3" value="选择时间" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.frm_AddTime);">
        日期,可在开头加上简单比较符号,*号表示模糊查询 </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">是否销售</td>
      <td> <input type="radio" name="frm_isBuy" value="1">
        已销售 
        <input type="radio" name="frm_isBuy" value="0">
        未销售 </td>
    </tr>
    <tr  class="hback"> 
      <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit" name="submit" value=" 执行查询 " /> 
              &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" 重置 " /> 
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>

<%End Sub%>
</body>
<%
Set VClass_Rs=nothing
User_Conn.close
Set User_Conn=nothing
%>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
//打开后根据规则显示箭头
var Req_FildName;
if (Old_Sql.indexOf("filterorderby=")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);break;}	
		}
	}	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + "csed")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
///////////////////////////////////////////////////////// 
function GetCodes(number)
{
	(isNaN(number))?10:number;
	new Ajax.Updater('More_Mode_Area', 'Card_ajax.asp?no-cache='+Math.random() , {method: 'get',parameters:'number='+number});	
}

function chkinput(form)
{
	var value_=form.frm_CardMoney.value+form.frm_CardDateNumber.value+form.frm_CardPoint.value;
	if (value_=='')
		{alert('点卡面值,点卡天数,点卡点数必须至少填写一项并且必须是数字。');form.frm_CardMoney.focus();return false;}
	else
		return Validator.Validate(form,3); 
}
-->
</script>

</html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. --> 






