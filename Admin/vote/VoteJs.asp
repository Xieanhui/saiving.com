<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<%session.CodePage="936"%>
<%
response.Charset="gb2312"
Dim Steps,TID,Cookie_Domain,OutHtmlID
''ǰ̨ҳ��,��JS���õõ� ���ø��ļ��������һЩ����.
TID = request.QueryString("TID")
OutHtmlID = request.QueryString("InfoID")
if OutHtmlID = "" then OutHtmlID = "Vote_HTML_ID"
if TID = "" or not isnumeric(TID) then response.Write("document.writeln('�ڲ�����:����ʱ,TID�����ṩ.\n');"&vbNewLine)

Cookie_Domain = request.Cookies("FoosunMFCookies")("FoosunMFDomain")
if Cookie_Domain="" then 
	Cookie_Domain = "localhost"
else
	if left(lcase(Cookie_Domain),len("http://"))="http://" then Cookie_Domain = mid(Cookie_Domain,len("http://")+1)
	if right(Cookie_Domain,1)="/" then Cookie_Domain = mid(Cookie_Domain,1,len(Cookie_Domain) - 1)
end if	
response.Write("<!--"&vbNewLine)
response.Write("//��һ�ε����� "&vbNewLine)
response.Write("function f_FirstGetVote_awen() {new Ajax.Updater('"&NoSqlHack(OutHtmlID)&"', 'http://"&Cookie_Domain&"/Vote/Index.asp?no-cache='+Math.random() , {method: 'get', parameters: 'TID="&NoSqlHack(TID)&"&InfoID="&NoSqlHack(OutHtmlID)&"' })}; "&vbNewLine)
response.Write("setTimeout('f_FirstGetVote_awen()',200);"&vbNewLine)
response.Write("-->"&vbNewLine)
%>





