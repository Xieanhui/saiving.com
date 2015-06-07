<%
'获得SQL函数和ACCESS函数
Sub GetFunctionstr()	
	If G_IS_SQL_DB=0 Then
		CharIndexStr="Mid"
	Else
		CharIndexStr="Substring"
	End If
End Sub
'生成xml文件
Function Makexml(f_parentID)
	dim f_obj_xml_rs,xml_list,savepath_xml
	xml_list = "<?xml version=""1.0"" encoding=""gb2312""?>" & chr(13) & chr(10)
	xml_list =  xml_list & "<rss version=""2.0"">"& chr(13) & chr(10)
	xml_list =  xml_list & "<classlist>"& chr(13) & chr(10)
	Set f_obj_xml_rs = Server.CreateObject(G_FS_RS)
	f_obj_xml_rs.open "select id,ClassName,ClassEName,classid,Addtime,isShow,[Domain],Parentid,ReycleTF,OrderID From FS_NS_NewsClass where Parentid='"& f_parentID &"' and ReycleTF=0 order by OrderID desc,id desc",Conn,0,1
	do while not f_obj_xml_rs.eof
		xml_list =  xml_list & "<item>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<classcname>"& f_obj_xml_rs("ClassName")&"</classcname>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<link>"& f_obj_xml_rs("ClassEName")&"</link>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<classename>"& f_obj_xml_rs("ClassEName")&"</classename>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<classid>"& f_obj_xml_rs("classid")&"</classid>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<addtime>"& f_obj_xml_rs("Addtime")&"</addtime>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<isshow>"& f_obj_xml_rs("isShow")&"</isshow>"& chr(13) & chr(10)
		xml_list =  xml_list & "		<domain>"& len(f_obj_xml_rs("Domain"))&"</domain>"& chr(13) & chr(10)
		xml_list =  xml_list & "</item>"& chr(13) & chr(10)
		f_obj_xml_rs.movenext
	loop
	xml_list =  xml_list & "</classlist>"& chr(13) & chr(10)
	xml_list =  xml_list & "</rss>"& chr(13) & chr(10)
	savepath_xml = Replace("\"&G_VIRTUAL_ROOT_DIR&"\FS_InterFace\xml\","\\","\")
	Call SaveFile(xml_list, f_parentID ,"xml",savepath_xml,"NS")
End Function
'获得xml栏目子类
Function News_ChildNewsRss(TypeID,f_CompatStr)  
	Dim f_ChildNewsRs_1,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
	Set f_ChildNewsRs_1 = Conn.Execute("Select id,orderid,ClassName,ClassID,ParentID from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by Orderid desc,id desc" )
	f_TempStr =f_CompatStr & "┄"
	do while Not f_ChildNewsRs_1.Eof
			News_ChildNewsRss = News_ChildNewsRss & "├" & f_TempStr &  f_ChildNewsRs_1("ClassName") 
			News_ChildNewsRss = News_ChildNewsRss & "RssFeed:xml/NS/"& f_ChildNewsRs_1("ClassID") &".xml"
			News_ChildNewsRss = News_ChildNewsRss & "</option>" & Chr(13) & Chr(10)
			News_ChildNewsRss = News_ChildNewsRss &News_ChildNewsRss(f_ChildNewsRs_1("ClassID"),f_TempStr)
		f_ChildNewsRs_1.MoveNext
	loop
	f_ChildNewsRs_1.Close
	Set f_ChildNewsRs_1 = Nothing
End Function
%>





