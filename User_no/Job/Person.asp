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
	if searchID="" then response.Redirect("../lib/error.asp?ErrCodes=<li>��Ҫ�Ĳ��������ṩ.</li>") : response.End()
	if Ap_Rs.eof then response.Redirect("../lib/error.asp?ErrCodes=<li>���û���δ�����ְ����.</li>") : response.End()
	rem �۳����ֵĴ���
	if Get_OtherTable_Value("select Count(*) from FS_AP_UserList where UserNumber='"&NoSqlHack(UserNumber)&"' and GroupLevel=3")=0 then 
		''VIP�û�����Ҫ�۳�
		''�õ����ѱ�׼
		int_PerCount = 0 :InitCount = 0
		set sysaprs = Conn.execute("select top 1 PerCount,InitCount from FS_AP_SysPara")
		if not sysaprs.eof then 
			if not isnull(sysaprs(0)) then int_PerCount = clng(sysaprs(0))
			if not isnull(sysaprs(1)) then InitCount = clng(sysaprs(1))
		end if
		''�õ��鿴����ְ�ߵ�����
		UserName = Get_OtherTable_Value("select RealName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
		if UserName="" then UserName = Get_OtherTable_Value("select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(UserNumber)&"'")
		''�жϹ�˾��ɫ
		if session("AP_EndDate")="" then 
			session("AP_EndDate") = Get_OtherTable_Value("select EndDate from FS_AP_UserList where UserNumber='"&NoSqlHack(UserNumber)&"' and GroupLevel=2")        
		end if
		if isdate(session("AP_EndDate")) then 
			''��˾���ڰ����û�
			if datediff("day",session("AP_EndDate"),date())>0 then 
				response.Redirect("../lib/error.asp?ErrCodes=<li>���ĵ㿨�ѵ��ڣ����ֵ��</li>&ErrorUrl=../card.asp") : response.End()
			end if
		else
			''�õ���˾���еĵ���
			LeftCount = Get_OtherTable_Value("select LeftCount from FS_AP_UserList where UserNumber='"&NoSqlHack(CorpID)&"' and GroupLevel=1")	
			if LeftCount="" then LeftCount=0
			LeftCount = InitCount + clng(LeftCount)
			
			if  clng(LeftCount) < int_PerCount then 
				response.Redirect("../lib/error.asp?ErrCodes=<li>���ĵ�������������г�ֵ.Ҫ��ÿ�����ѵ�����"&int_PerCount&"�������е�����"&clng(LeftCount)&"��</li>&ErrorUrl=../card.asp") : response.End()
			end if
		
			''���������Ҫ�۵㣬��۳�
			if int_PerCount<>0 then 
				''�������Ѽ�¼
				Conn.execute("insert into FS_AP_Consume (UserNumber,ConsumeDate,SearchUId,SearchUName,LeftCount,LeftDateNumber) values ('"&NoSqlHack(CorpID)&"','"&now()&"','"&NoSqlHack(UserNumber)&"','"&NoSqlHack(UserName)&"',"&clng(LeftCount) - int_PerCount&",0)")
				''������Ƹ��˾�ĵ㿨ֵ
				Conn.execute("update FS_AP_UserList set LeftCount=LeftCount-"&int_PerCount&" where UserNumber='"&NoSqlHack(CorpID)&"'")
			end if
		end if	
	end if	
''-----------------------------------	
else
''-----------------------------------	
	if Ap_Rs.eof then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>����δ�����ְ����.</li>&ErrorUrl=../job/job_resume.asp") : response.End()
	end if	
end if

Conn.execute("update FS_AP_Resume_BaseInfo set click=click+1")


''�õ���ر��ֵ��
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
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=UserName%> �ļ���</title>
<STYLE>
TABLE {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '����'
}
TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '����'
}
P {
	FONT-SIZE: 12px; LINE-HEIGHT: 120%; FONT-FAMILY: '����'
}
.ResTitleFot {
	FONT-SIZE: 14px; FONT-FAMILY: '����'
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
<!--������Ϣ--->

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
                  <td width="87" height="16">�ա������� </td>
                  <td width="196" height="16"><%=Ap_Rs("Uname")%></td>
                  <td width="88" height="16">�ԡ��� </td>
                  <td width="194" height="16"><%=Replacestr(Ap_Rs("Sex"),"0:��,1:Ů")%></td>
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
									response.Write("<b>��Ƭ��ʾ��<br>���ϴ������Ƭ<br>����ѡ��һ���ö���</b>")		
								end if
							else
								response.Write("<b>��Ƭ��ʾ��<br>���ϴ������Ƭ<br>����ѡ��һ���ö���</b>")	
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
                  <td height="16">�������ڣ�</td>
                  <td height="16"><%=Ap_Rs("Birthday")%></td>
                  <td height="16">��ס�أ�</td>
                  <td height="16"><%=Ap_Rs("Province")&"��"&Ap_Rs("City")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">�������ޣ�</td>
                  <td height="16"><%=Replacestr(Ap_Rs("WorkAge"),"1:�ڶ�ѧ��,2:Ӧ���ҵ��,3:һ������,4:��������,5:��������,6:��������,7:��������,8:ʮ������")%></td>
                  <td height="16">���ߣ� </td>
                  <td height="16"><%=Ap_Rs("shengao")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">���ѧ����</td>
                  <td colspan=3 height="16"><%=Ap_Rs("xueli")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">�ء���ַ�� </td>
                  <td colspan=2 height="16"><%=Ap_Rs("address")%></td>
                </tr>
                <tr height=20> 
                  <td height="16">�����ʼ���</td>
                  <td colspan=3 height="16"><%=Ap_Rs("Email")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">��ϵ�绰��</td>
                  <td colspan=3 height="16"><%=Ap_Rs("HomeTel")%></td>
                </tr>
				 
                <tr height=20> 
                  <td height="16">��˾�绰��</td>
                  <td colspan=3 height="16"><%=Ap_Rs("CompanyTel")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">�ƶ��绰��</td>
                  <td colspan=3 height="16"><%=Ap_Rs("Mobile")%></td>
                </tr>
				
                <tr height=20> 
                  <td height="16">��õ��ڣ�</td>
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
      id=Cur_Val><font color="05A3D8"><strong>�� �� �� ��</strong></font> </td>
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
	WorkType = replacestr(Ap_Rs("WorkType"),"1:ȫְ,2:��ְ,3:ʵϰ,4:ȫְ/��ְ")
	Salary = replacestr(Ap_Rs("Salary"),":����,0:����,1:1500Ԫ����,2:1500-1999Ԫ,3:2000-2999Ԫ,4:3000-4499Ԫ,5:4500-5999Ԫ,6:6000-7999Ԫ,7:8000-9999Ԫ,8:10000-14999Ԫ,9:15000-19999Ԫ,10:20000-29999Ԫ,11:30000-49999Ԫ,12:50000Ԫ����")
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
      id=Cur_Val><font color="05A3D8"><strong>�� ְ �� ��</strong></font> </TD>
          </TR>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>�������ʣ�</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=WorkType%></TD>
          </TR>
		  
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>�����ص㣺&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Get_OtherTable_Value("select Province+' '+City+' ' from FS_AP_Resume_WorkCity where UserNumber in ('"&FormatStrArr(UserNumber)&"')")%></TD>
          </tr>
		
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>ϣ����ҵ��&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Get_OtherTable_Value("select Trade+'/'+Job+' | ' from FS_AP_Resume_Position where UserNumber in ('"&FormatStrArr(UserNumber)&"')")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>�������ʣ�&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Salary%></TD>
          </tr>
        </TBODY>
      </TABLE>
	  
	  
	  </td>
    <td width="3" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
  </tr>
</table>
<!--�� Ŀ �� ��--->

<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>�� Ŀ �� ��</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_ProjectExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>��Ŀ���ƣ�</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Project")%></TD>
          </TR>
		  
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��Ŀʱ�䣺&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&""&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </tr>
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>���������</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("SoftSettings")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>Ӳ��������</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("HardSettings")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>�������ߣ�&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Tools")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��Ŀ������&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("ProjectDescript")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>����������&nbsp;</TD>
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

<!--��������-->	  
<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>�� �� �� ��</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_WorkExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val><span class="STYLE1">����ʱ�䣺</span>&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&""&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��˾���ƣ�&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("CompanyName")%></TD>
          </tr>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��˾���ʣ�&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("CompanyKind")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��    ҵ��&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Trade")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>ְ    λ��&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Job")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��    �ţ�&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Department")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>����������&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Description")%></TD>
          </tr>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>֤���ˣ�&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val colspan="2"><%=Ap_Rs("Certifier")%></TD>
          </tr>
          <tr> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val>��ϵ֤���ˣ�&nbsp;</TD>
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


<!--�� �� �� ѵ-->	  
<table width="70%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="3" height="16" background="Images/m_bg.gif"><img src="Images/m_bg.gif" width="3" height="2"></td>
    <td height="16" bgcolor="f4f4f4"><TABLE cellSpacing=0 cellPadding=1 width=98% align=center border=0>
        <TBODY>
          <TR> 
            <TD background="Images/m_d.gif" height="19" colspan="3" class=ResTbLfPd 
      id=Cur_Val><font color="05A3D8"><strong>�� �� �� ѵ</strong></font> </TD>
          </TR>
<%
set Ap_Rs = Conn.execute("select * from FS_AP_Resume_EducateExp where UserNumber='"&NoSqlHack(UserNumber)&"'")
do while not Ap_Rs.eof
%>
          <TR> 
            <TD width="100" class=ResTbLfPd 
      id=Cur_Val><span class="STYLE1">��ѵʱ�䣺</span>&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=formatdatetime(Ap_Rs("BeginDate"),1)&"��"&formatdatetime(Ap_Rs("EndDate"),1)%></TD>
          </TR>
	
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>��ѵ������&nbsp;</TD>
            <TD class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("SchoolName")%></TD>
          </tr>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>ѧϰ���ݣ�&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Specialty")%></TD>
          </TR>
		  
          <TR> 
            <TD class=ResTbLfPd 
      id=Cur_Val>ѧ��֤�飺&nbsp;</TD>
            <TD colspan="2" class=ResTbLfPd 
      id=Cur_Val><%=Ap_Rs("Diploma")%></TD>
          </TR>
          <tr> 
            <TD class=ResTbLfPd 
      id=Cur_Val>���������&nbsp;</TD>
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
    background=Images/m_5.gif>�����ң��һ����ø��ã�</TD>
    <TD align=right width="1%" background=Images/m_5.gif><IMG 
      height=39 src="Images/m_6.gif" 
width=12></TD></TR></TBODY></TABLE>

<%
Set Fs_User = Nothing
set AP_Rs = Nothing
%>

</body></html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





