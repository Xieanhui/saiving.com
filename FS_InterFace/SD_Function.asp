<%
'返回权限列表
Function SD_Pops_List()
	SD_Pop_List = ""
End Function
'生成 
Function SD_Refresh(Lable_ID)
	SD_Refresh = ""
End Function
'全部标签列表
Function SD_Lables_List()
	SD_Lables_List = ""
End Function
'获得每一个标签的参数
Function SD_Lables_Para(LableID)
	SD_Lables_Para = ""
End Function
'会员添加修改一条供求信息，参数为数据库里面的字段排列数组
'数组的第一为为ID，如果ID为空，则是添加，如果ID不为空，则是修改 
Function SD_Add_OR_Edit_News(f_News_Cont_Array)
	SD_Lables_Para = True
End Function
%>





