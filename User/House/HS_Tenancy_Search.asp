<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="default.asp">����</a>����������</td>
          </tr>
        </table>

        <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form name="form3" method="post" action="HS_Tenancy_Search_Result.asp">
            <tr  class="hback"> 
              <td colspan="3" align="left" class="xingmu" >��ѯ��Դ ������Ϣ</td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">ID</td>
              <td> <input type="text" name="TID" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">��;</td>
              <td> <select name="UseFor">
                  <option value="">����</option>
                  <option value="1">ס��</option>
                  <option value="2">д�ּ�</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">����</td>
              <td> <select name="Class">
                  <option value="">����</option>
                  <option value="1">����</option>
                  <option value="2">����</option>
                  <option value="3">����</option>
                  <option value="4">��</option>
                  <option value="5">����</option>
                  <option value="6">ת��</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">��Դ��ַ</td>
              <td> <input type="text" name="Position" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">����</td>
              <td> <input type="text" name="CityArea" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">�۸�</td>
              <td> <input type="text" name="Price" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">����</td>
              <td> <input type="text" name="Hou���??? seStyle" size="40" value=""> <span class="tx">����,�洢��ʽ:l,m,nl:��m:��n:��</span></td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">���</td>
              <td> <input type="text" name="Area" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">¥��</td>
              <td> <input type="text" name="Floor" size="40" value=""> <span class="tx">¥��,�洢��ʽ:m,nm:�ܲ�n:�ڼ���</span></td>
            </tr>
            <tr  class="hback"> 
              <td align="right">������ʩ</td>
              <td> <input type="text" name="equip" size="40" value=""> <span class="tx">�����ʽ:��l,m,n,x,y,zl:ͨˮm:��n:��x:�绰y:����z:��ʾ���1��ʾ��,0��ʾ��</span></td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">װ�����</td>
              <td> <select name="Decoration">
                  <option value="">����</option>
                  <option value="1">��װ��</option>
                  <option value="2">�е�װ��</option>
                  <option value="3">�ߵ�װ��</option>
                </select> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">��ϵ��</td>
              <td> <input type="text" name="LinkMan" size="40" value=""> </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">��Ч��</td>
              <td> <select name="Period">
                  <option value="">����</option>
                  <option value="һ��">һ��</option>
                  <option value="����">����</option>
                  <option value="����">����</option>
                  <option value="һ��">һ��</option>
                  <option value="����">����</option>
                </select> <span class="tx">��Ч��:һ��,����,����,һ��,����),����ֻ��������</span> 
              </td>
            </tr>
            <tr  class="hback"> 
              <td width="20%" align="right">��������</td>
              <td> <input type="text" name="PubDate" size="40" value=""> <span class="tx">�����á�>2006-08������ʽ</span> 
              </td>
            </tr>
            <tr  class="hback"> 
              <td align="right">�Ƿ�ͨ�����</td>
              <td> <input type="radio" name="Audited"  value="1">
                ����� 
                <input type="radio" name="Audited"  value="0">
                δ���</td>
            </tr>
            <tr  class="hback"> 
              <td colspan="4"> <table border="0" width="100%" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td align="center"> <input type="submit" name="submit" value=" ִ�в�ѯ " /> 
                      &nbsp; <input type="reset" name="ReSet" id="ReSet" value=" ���� " /> 
                    </td>
                  </tr>
                </table></td>
            </tr>
          </form>
        </table>


       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Fs_User = Nothing
%>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





