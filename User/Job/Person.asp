<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 
Dim Ap_Rs,Ap_Sql,sysaprs,int_PerCount
Dim UserNumber,CorpID,UserName,CorpName,searchID,LeftCount,InitCount,IsCorporation

UserNumber = request.QueryString("UID")
searchID = UserNumber
CorpID = Session("FS_UserNumber")
CorpName = Session("FS_UserName")


if not UserNumber<>"" then 
	UserNumber = Session("FS_UserNumber")
	UserName  = Get_OtherTable_Value("select RealName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
	if UserName = "" then UserName = Get_OtherTable_Value("select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
end if	

IsCorporation = Get_OtherTable_Value("select IsCorporation from FS_ME_Users where UserNumber='"&NoSqlHack(CorpID)&"'")

set Ap_Rs = Conn.execute("select * from FS_AP_Resume_BaseInfo where UserNumber='"&NoSqlHack(UserNumber)&"'")

if IsCorporation = "1" then 
	if searchID="" then response.Redirect("../lib/error.asp?ErrCodes=<li>必要的参数必须提供.</li>") : response.End()
	if Ap_Rs.eof then response.Redirect("../lib/error.asp?ErrCodes=<li>该用户尚未添加求职简历.</li>") : response.End()
	rem 扣除积分的处理
	if Get_OtherTable_Value("select Count(*) from FS_AP_UserList where UserNumber='"&NoSqlHack(UserNumber)&"' and GroupLevel=3")=0 then 
		''VIP用户不需要扣除
		''得到消费标准
		int_PerCount = 0 :InitCount = 0
		set sysaprs = Conn.execute("select top 1 PerCount,InitCount from FS_AP_SysPara")
		if not sysaprs.eof then 
			if not isnull(sysaprs(0)) then int_PerCount = clng(sysaprs(0))
			if not isnull(sysaprs(1)) then InitCount = clng(sysaprs(1))
		end if
		''得到查看的求职者的姓名
		UserName = Get_OtherTable_Value("select RealName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
		if UserName="" then UserName = Get_OtherTable_Value("select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
		''判断公司角色
		if session("AP_EndDate")="" then 
			session("AP_EndDate") = Get_OtherTable_Value("select EndDate from FS_AP_UserList where UserNumber='"&NoSqlHack(UserNumber)&"' and GroupLevel=2")        
		end if
		if isdate(session("AP_EndDate")) then 
			''公司属于包月用户
			if datediff("day",session("AP_EndDate"),date())>0 then 
				response.Redirect("../lib/error.asp?ErrCodes=<li>您的点卡已到期，请冲值。</li>&ErrorUrl=../card.asp") : response.End()
			end if
		else
			''得到公司现有的点数
			LeftCount = Get_OtherTable_Value("select LeftCount from FS_AP_UserList where UserNumber='"&NoSqlHack(CorpID)&"' and GroupLevel=1")	
			if LeftCount="" then LeftCount=0
			LeftCount = InitCount + clng(LeftCount)
			
			if  clng(LeftCount) < int_PerCount then 
				response.Redirect("../lib/error.asp?ErrCodes=<li>您的点数不够用请进行冲值.要求每次消费点数（"&int_PerCount&"）您现有点数（"&clng(LeftCount)&"）</li>&ErrorUrl=../card.asp") : response.End()
			end if
		
			''如果设置了要扣点，则扣除
			if int_PerCount<>0 then 
				''插入消费记录
				Conn.execute("insert into FS_AP_Consume (UserNumber,ConsumeDate,SearchUId,SearchUName,LeftCount,LeftDateNumber) values ('"&NoSqlHack(CorpID)&"','"&now()&"','"&NoSqlHack(UserNumber)&"','"&NoSqlHack(UserName)&"',"&clng(LeftCount) - int_PerCount&",0)")
				''更新招聘公司的点卡值
				Conn.execute("update FS_AP_UserList set LeftCount=LeftCount-"&int_PerCount&" where UserNumber='"&NoSqlHack(CorpID)&"'")
			end if
		end if	
	end if	
''-----------------------------------	
else
''-----------------------------------	
	if Ap_Rs.eof then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>您尚未添加求职简历.</li>&ErrorUrl=../job/job_resume.asp") : response.End()
	end if	
end if

Conn.execute("update FS_AP_Resume_BaseInfo set click=click+1")


''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then
		if instr(This_Fun_Sql," in ")>0 then 
			do while not This_Fun_Rs.eof
				Get_OtherTable_Value  = Get_OtherTable_Value &""& This_Fun_Rs(0)	
				This_Fun_Rs.movenext		
			loop
		else
			Get_OtherTable_Value = This_Fun_Rs(0) & ""
		end if
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=UserName%> 的简历</title>
<STYLE>
TABLE {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '宋体'
}
TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '宋体'
}
P {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '宋体'
}
.ResTitleFot {
	FONT-SIZE: 14px; FONT-FAMILY: '宋体'
}
.ResTitleTbBd {
	BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid
}
.ResTitleTbBd1 {
	BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-BOTTOM: #ffffff 1px solid
}
.STYLE1 {color: #FF0000}
</STYLE>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1%"><img src="Images/m_1.gif" width="12" height="13"></td>
    <td width="99%" align="right" background="Images/m_2.gif"><img src="Images/spacer.gif" width="224" height="1"><img src="Images/m_3.gif" width="12" height="13"></td>
  </tr>
</table>

<table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="42" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="42" align="center" background="Images/m_bgtop.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="42" align="center"><img src="Images/m_m.gif" width="135" height="27"></td>
        </tr>
      </table></td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>
<!--基本信息--->

<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Images/m_xx.gif"><img src="Images/spacer.gif" width="1" height="1"></td>
        </tr>
        <tr>
          <td bgcolor="EEFAFE"><table cellspacing=1 cellpadding=1 width=98% align=center border=0 >
              <!--DWLayoutTable-->
              <tbody>
                <tr height=20> 
                  <td width="87" height="16">姓　　名： </td>
                  <td width="196" height="16"><%=Ap_Rs("Uname")%></td>
                  <td width="88" height="16">性　别： </td>
                  <td width="194" height="16"><%=Replacestr(Ap_Rs("Sex"),"0:男,1:女")%></td>
                  <td width="117" rowspan=10 align="center" valign="middle"> 
                    <table width="100%" height="0%" border="0" cellpadding="0" cellspacing="0">
                      <tr><td align="center" valign="top">
					  	<%
							if Ap_Rs("PictureExt") <>"" And Not IsNull(Ap_Rs("PictureExt")) then 
								Dim photo
								photo = Trim(Ap_Rs("PictureExt"))
								If Right(LCase(photo),4) = ".jpg" Or Right(LCase(photo),4) = ".bmp" Or Right(LCase(photo),4) = ".gif" Or Right(LCase(photo),4) = ".png" Then
									response.Write("<img border=0 src="""&photo&""" id=""Resume_Photo"" style=""cursor:pointer;"" onclick=""window.open('"&photo&"');"">")
								else
									response.Write("<b>照片显示区<br>请上传你的照片<br>并且选择一张置顶！</b>")		
								end if
							else
								response.Write("<b>照片显示区<br>请上传你的照片<br>并且选择一张置顶！</b>")	
							end if	
						%>					  	
					  </td>
                      </tr>
                    </table>
					<script language="javascript" type="text/javascript">
					<!--
						function getThePhotoToSmall(PicID)
						{
							var Obj = document.getElementById(PicID);
							var _Width = Obj.width;
							if (_Width > 150)
							{
								Obj.width = 150;
							}
						}
						window.setTimeout("getThePhotoToSmall('Resume_Photo')",50);
					//-->
					</script> 
            </td>
                </tr>
                <tr height=20> 
                  <td height="16">出生日期：</td>
                  <td height="16"><%=Ap_Rs("Birthday")%></td>
                  <td height="16">居住地：</td>
                  <td height="16"><%=Ap_Rs("Province")&"、"&Ap_Rs("City")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">工作年限：</td>
                  <td height="16"><%=Replacestr(Ap_Rs("WorkAge"),"1:在读学生,2:应届毕业生,3:一年以上,4:二年以上,5:三年以上,6:五年以上,7:八年以上,8:十年以上")%></td>
                  <td height="16">身　高： </td>
                  <td height="16"><%=Ap_Rs("shengao")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">最高学历：</td>
                  <td colspan=3 height="16"><%=Ap_Rs("xueli")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">地　　址： </td>
                  <td colspan=2 height="16"><%=Ap_Rs("address")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">电子邮件：</td>
                  <td colspan=3 height="16"><%=Ap_Rs("Email")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">联系电话：</td>
                  <td colspan=3 height="16"><%=Ap_Rs("HomeTel")%></td>
                </tr>
				 
                <tr height=20> 
                  <td height="16">公司电话：</td>
                  <td colspan=3 height="16"><%=Ap_Rs("CompanyTel")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">移动电话：</td>
                  <td colspan=3 height="16"><%=Ap_Rs("Mobile")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">多久到岗：</td>
                  <td colspan=3><%=Ap_Rs("howday")%></td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
        <tr>
          <td background="Images/m_xx.gif"><img src="Images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table> </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>
<!--over-->
<%
Ap_Rs.close : set Ap_Rs = nothing
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_Language where UserNumber='"&NoSqlHack(UserNumber)&"'")
%>

<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><table cellspacing=0 cellpadding=1 width=98% align=center border=0>
      <tbody>
        <tr  >
          <td background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>语 言 能 力</strong></font> </td>
        </tr>
		<%do while not Ap_Rs.eof%>
        <tr>
          <td class=ResTbLfPd id=Cur_Val width="589"><%=Ap_Rs("Language")%></td>
          <td class=ResTbLfPd id=Cur_Val width="589"><%=Ap_Rs("Degree")%></td>
        </tr>
		<%
		Ap_Rs.movenext
		loop		
		%>
        <tr>
          <td class=ResTbLfPd 
      id=Cur_Val width="589"></td>
          <td class=ResTbLfPd 
      id=Cur_Val width="589" colspan="2"></td>
        </tr>
      </tbody>
    </table></td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>

<%
Ap_Rs.close : set Ap_Rs = nothing
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_Intention where UserNumber='"&NoSqlHack(UserNumber)&"'")
Dim WorkType,Salary,SelfAppraise
if not Ap_Rs.eof then 
	WorkType = replacestr(Ap_Rs("WorkType"),"1:全职,2:兼职,3:实习,4:全职/兼职")
	Salary = replacestr(Ap_Rs("Salary"),":面议,0:面议,1:1500元以下,2:1500-1999元,3:2000-2999元,4:3000-4499元,5:4500-5999元,6:6000-7999元,7:8000-9999元,8:10000-14999元,9:15000-19999元,10:20000-29999元,11:30000-49999元,12:50000元以上")
	SelfAppraise = Ap_Rs("SelfAppraise")
end if 
Ap_Rs.close
%>

<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR  > 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>求 职 意 向</strong></font> </TD>
          </TR>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>工作性质：</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=WorkType%></TD>
          </TR>
		  
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>工作地点：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Get_OtherTable_Value("select Province+' '+City+' ' from FS_AP_Resume_WorkCity where UserNumber in ('"&FormatStrArr(UserNumber)&"')")%></TD>
          </tr>
		
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>希望行业：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Get_OtherTable_Value("select Trade+'/'+Job+' | ' from FS_AP_Resume_Position where UserNumber in ('"&FormatStrArr(UserNumber)&"')")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>期望工资：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Salary%></TD>
          </tr>
        </TBODY>
      </TABLE>
	  
	  
	  </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>
<!--项 目 经 验--->

<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>项 目 经 验</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_ProjectExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>项目名称：</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Project")%></TD>
          </TR>
		  
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>项目时间：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&""&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </tr>
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>软件环境：</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("SoftSettings")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>硬件环境：</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("HardSettings")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>开发工具：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Tools")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>项目描述：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("ProjectDescript")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>责任描述：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Duty")%></TD>
          </tr>
		  
<%Ap_Rs.movenext
loop
Ap_Rs.close%>		  
        </TBODY>
      </TABLE>
	  </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>

<!------>

<!--工作经验-->	  
<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>工 作 经 验</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_WorkExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val><span class="STYLE1">工作时间：</span>&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&""&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>公司名称：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("CompanyName")%></TD>
          </tr>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>公司性质：&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("CompanyKind")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>行    业：&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Trade")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>职    位：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Job")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>部    门：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Department")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>工作描述：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Description")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>证明人：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Certifier")%></TD>
          </tr>
          <tr> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>联系证明人：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("CertifierTel")%></TD>
          </tr>
		  
<%Ap_Rs.movenext
loop
Ap_Rs.close%>		  
        </TBODY>
      </TABLE>
	  
	  
<!---->	  
    </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>


<!--教 育 培 训-->	  
<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>教 育 培 训</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_EducateExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val><span class="STYLE1">培训时间：</span>&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&"到"&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </TR>
	
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>培训机构：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("SchoolName")%></TD>
          </tr>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>学习内容：&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Specialty")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>学历证书：&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Diploma")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>相关描述：&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Description")%></TD>
          </tr>
	  
<%Ap_Rs.movenext
loop
Ap_Rs.close%>		  
        </TBODY>
      </TABLE>
	  
	  
<!---->	  
    </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>

<TABLE cellSpacing=0 cellPadding=0 width="70%" align=center border=0>
  <TBODY>
  <TR>
    <TD width="2%"><IMG height=39 src="Images/m_4.gif" width=12></TD>
    <TD align=middle width="97%" 
    background=Images/m_5.gif>相信我，我会做得更好！</TD>
    <TD align=right width="1%" background=Images/m_5.gif><IMG 
      height=39 src="Images/m_6.gif" 
width=12></TD></TR></TBODY></TABLE>

<%
Set Fs_User = Nothing
set AP_Rs = Nothing
%>

</body></html>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





