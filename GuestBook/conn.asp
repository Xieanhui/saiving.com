<%
Dim f_ConnStr
	If G_IS_SQL_DB = 1 Then
		f_ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;"& G_DATABASE_CONN_STR &";"
	Else
		f_ConnStr = "DBQ=" + Server.MapPath(Add_Root_Dir(G_DATABASE_CONN_STR)) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	End If
	On Error Resume Next
	Set Conn = Server.CreateObject(G_FS_CONN)
	Conn.Open f_ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "<font size=""2"">[���ݿ����Ӵ���]<br>����ϵͳ��������>>վ�㳣������,����/FS_Inc/Const.asp�ļ�!</font>"
		Response.End
	End If
%>





