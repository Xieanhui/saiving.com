<script language="javascript">
function openScript(url, width, height){
var Win=window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no' );
}
</script>
<%
MF_Default_Conn
Dim TellRs,TelTopic,sqlTess,AddDate,DisID
Set TellRs= server.CreateObject (G_FS_RS)
sqlTess="select Top 1 ID,Topic,Content,Person,IsUse,PV,AddUser,AddDate From FS_WS_NewsTell Order By ID Desc"
TellRs.open sqlTess,Conn,1,1
if not TellRs.eof then
	DisID = TellRs("ID")
	TelTopic=trim(TellRs("Topic"))
	AddDate=Trim(TellRs("AddDate"))
end if
set TellRs=nothing
%>
<table cellspacing=1 cellpadding=3 align=center border=0 width=98%><tr><td align=center width=100% valign=middle colspan=2><strong>公告信息:</strong>
<%
if TelTopic="" or IsNull(TelTopic) then
%>
	<B>当前没有公告</B>(<%=now()%>)
<%
else
%>
	<a href="javascript:openScript('announcements.asp?action=showone&boardid=0&DisID=<% = DisID %>',500,300)"><B><%=TelTopic%></B></a>(<%=AddDate%>)
<%
end if
%>
</td></tr></table>






