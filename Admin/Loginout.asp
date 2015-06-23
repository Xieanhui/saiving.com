<%
Option Explicit
	Session("Admin_Name") = ""
	Session("Admin_Pass_Word") = ""
	Session("Admin_Parent_Admin") = ""
	Session("Admin_Is_Super") = ""
	Session("Admin_Pop_List") = ""
	Session("Admin_Add_Admin") = ""
	Session("Admin_Style_Num") = ""
	Session("Admin_FilesTF") = ""
	
	Response.Cookies("FoosunAdminCookie")("Admin_Name") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Pass_Word") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Parent_Admin") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Is_Super") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Pop_List") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Add_Admin") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_Style_Num") = ""
	Response.Cookies("FoosunAdminCookie")("Admin_FilesTF") = ""
	'session.Abandon
	Response.Cookies("FoosunSUBCookie")=""
	Response.Cookies("FoosunSUBCookie")=Empty
	Response.Cookies("FoosunMFCookies")=""
	Response.Cookies("FoosunMFCookies")=Empty
	Response.Cookies("FoosunNSCookies")=""
	Response.Cookies("FoosunNSCookies")=Empty
	Response.Cookies("FoosunDSCookies")=""
	Response.Cookies("FoosunDSCookies")=Empty
Response.Redirect "Login.asp"
%>





