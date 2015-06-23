<%
function table1(total,table_x,table_y,thickness,table_width,all_width,all_height,table_type,Vs_Num)
'-----生成投票结果柱形图函数
	Dim line_color,left_width,length,total_no,temp1,temp6,i,temp2,temp3,temp4,line0,table_space
	Dim temp_space
	'-------------------------------------
	'---2007-01-15 Edit By Ken    将原固定数组改为动态二维数组,解决超出数组范围则下标越界问题
	Dim ColorStr,CArrayNum,k,j,StrColor,ArrayNum,h,l,t
	ColorStr = "#d1ffd1,#ffbbbb,#ffe3bb,#cff4f3,#d9d9e5,#ffc7ab,#ecffb7"
	StrColor = "#00ff00,#ff0000,#ff9900,#33cccc,#666699,#993300,#99cc00"
	CArrayNum = Split(ColorStr,",")
	ArrayNum = Split(StrColor,",")
	Dim tb_color()
	If Vs_Num <= 7 Then
		For k = 1 To 7
			ReDim Preserve tb_color(7,2)
			tb_color(k,1) = CArrayNum(k-1)
			tb_color(k,2) = ArrayNum(k-1)
		Next
	Else
		j = Cint(Vs_Num\7)
		Vs_Num = Cint(j+1)*7
		For l = 0 To j
			ReDim Preserve tb_color(Vs_Num,2)
			h = l*7
			For k = 1 To 7
				t = Cint(h+k)
				If t >= Vs_Num Then Exit For
				tb_color(t,1) = CArrayNum(k-1)
				tb_color(t,2) = ArrayNum(k-1)
			Next
		Next		
	End If

	line_color="#69f"
	left_width=70
	length=thickness/2
	total_no=ubound(total,1)
	
	temp1=0
	temp6=0
	for i=1 to total_no
		if temp1<total(i,1) then temp1=total(i,1)
		if temp6>total(i,1) then temp6=total(i,1)
	next
	'if temp6=1 then
	'else'出现值小于0的情况
	temp1=int(temp1)
	if temp6>=0 then
		temp6=int(temp6)
		if temp6>10 then
			temp2=mid(cstr(temp6),2,1)
			if temp2>4 then 
				temp3=(int(temp6/(10^(len(cstr(temp6))-1)))-1)*10^(len(cstr(temp6))-1)
			else
				temp3=(int(temp6/(10^(len(cstr(temp6))-1)))-0.5)*10^(len(cstr(temp6))-1)
			end if
			temp6=temp3
		else
			temp6=0
		end if
	'	if temp6-10<0 then temp6=0 else temp6=temp6-10
	else
		temp6=int(0-temp6)
		if temp6>10 then
			temp2=mid(cstr(temp6),2,1)
			if temp2>4 then 
				temp3=(int(temp6/(10^(len(cstr(temp6))-1)))+1)*10^(len(cstr(temp6))-1)
			else
				temp3=(int(temp6/(10^(len(cstr(temp6))-1)))+0.5)*10^(len(cstr(temp6))-1)
			end if
			temp6=0-temp3
		else
			temp6=-10
		end if
	end if
	if temp1>9 then
		temp2=mid(cstr(temp1),2,1)
		if temp2>4 then 
			temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
		else
			temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
		end if
	else
		if temp1>4 then temp3=10 else temp3=5
	end if
	temp4=temp3
	response.write "<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:-1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='-45' focus='100%' type='gradient'/></v:rect>"
	select case table_type
	case "A"
		for i=0 to all_height step all_height/5
			if i<>all_height then response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>"
			if i<>all_height then response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height-length-i&"px' to='"&table_x+left_width+length&"px,"&table_y+all_height-i&"px' strokecolor='"&line_color&"'/>"
			response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>"
			response.write ""
			response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:1'>"
			response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
			temp4=temp4-(temp3-temp6)/5
		next
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y&"px' to='"&table_x+left_width+length&"px,"&table_y+all_height-length&"px' strokecolor='"&line_color&"'/>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width-50&"px;top:"&table_y-20&"px;width:100px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>纵坐标</td></tr></table></v:textbox></v:shape>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width&"px;top:"&table_y+all_height-9&"px;width:100px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>横坐标</td></tr></table></v:textbox></v:shape>"
	
		line0=(temp3/(temp3-temp6))*all_height
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:2' from='"&table_x+left_width&"px,"&table_y+line0&"px' to='"&table_x+all_width+left_width&"px,"&table_y+line0&"px'></v:line>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+line0-length&"px' to='"&table_x+left_width+length&"px,"&table_y+line0&"px' strokecolor='"&line_color&"'/>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+line0-length&"px' to='"&table_x+all_width+left_width&"px,"&table_y+line0-length&"px' strokecolor='"&line_color&"'/>"
	
		table_space=(all_width-table_width*total_no)/total_no
	for i=1 to total_no
		temp_space=table_x+left_width+table_space/2+table_space*(i-1)+table_width*(i-1)
		response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;"
		response.write "left:"&temp_space&"px;top:"
		if total(i,1)>0 then
			response.write table_y+line0-(total(i,1)/(temp3-temp6))*all_height
		else
			response.write table_y+line0
		end if
		response.write "px;width:"&table_width&"px;height:"&all_height*(abs(total(i,1))/(temp3-temp6))
		response.write "px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
		response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' type='gradient'/>"
		response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
		response.write "</v:rect>"
	
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&temp_space&"px;top:"
		if total(i,1)>0 then
			response.write table_y+line0-(total(i,1)/(temp3-temp6))*all_height-table_width
		else
			response.write table_y+line0-table_width
		end if
		response.write "px;width:"&table_space+15&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"
	
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&temp_space-table_space/2&"px;top:"&table_y+all_height+1&"px;width:"&table_space+table_width&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,2)&"</td></tr></table></v:textbox></v:shape>"
	next
	
	Case "B"
		table_space=(all_height-table_width*total_no)/total_no
		temp4=temp6
	
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+left_width&"px,"&table_y+all_height+15&"px'/>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width-left_width&"px;top:"&table_y+all_height&"px;width:"&left_width&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
		temp4=temp6+(temp3-temp6)/5
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+all_width&"px,"&table_y+all_height-length&"px' strokecolor='"&line_color&"'/>"
	for i=0 to all_width-1 step all_width/5
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+i&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+length+i&"px,"&table_y+all_height&"px' strokecolor='"&line_color&"'/>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length+i&"px,"&table_y+all_height-length&"px' to='"&table_x+left_width+length+i&"px,"&table_y&"px' strokecolor='"&line_color&"'/>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+i+all_width/5&"px,"&table_y+all_height&"px' to='"&table_x+left_width+i+all_width/5&"px,"&table_y+all_height+15&"px'/>"
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+i+all_width/5-left_width&"px;top:"&table_y+all_height&"px;width:"&left_width&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
		temp4=temp4+(temp3-temp6)/5
	next
	
	for i=1 to total_no
		temp_space=table_space/2+table_space*(i-1)+table_width*(i-1)
		if total(i,1)>=0 then
			response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px;top:"&table_y+temp_space&"px;width:"&all_width*temp3/(temp3-temp6)*(total(i,1)/temp3)&"px;height:"&table_width&"px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
			response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' angle='-90' focus='100%' type='gradient'/>"
			response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
			response.write "</v:rect>"
			response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)+all_width*temp3/(temp3-temp6)*(total(i,1)/temp3)+thickness/2&"px;top:"&table_y+temp_space&"px;width:"&table_space+15&"px;height:18px;z-index:1'>"
			response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"
		else
			response.write "<v:rect id='_x0000_s1025' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)-all_width*temp3/(temp3-temp6)*abs(total(i,1)/temp3)&"px;top:"&table_y+temp_space&"px;width:"&all_width*temp3/(temp3-temp6)*abs(total(i,1)/temp3)&"px;height:"&table_width&"px;z-index:1' fillcolor='"&tb_color(i,2)&"'>"
			response.write "<v:fill color2='"&tb_color(i,1)&"' rotate='t' angle='-90' focus='100%' type='gradient'/>"
			response.write "<o:extrusion v:ext='view' backdepth='"&thickness&"pt' color='"&tb_color(i,2)&"' on='t'/>"
			response.write "</v:rect>"
			response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)+thickness/2&"px;top:"&table_y+temp_space-table_space/4&"px;width:"&table_space+30&"px;height:18px;z-index:1'>"
			response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center'>"&total(i,1)&"</td></tr></table></v:textbox></v:shape>"
		end if
	
		response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+temp_space&"px;width:"&left_width&"px;height:18px;z-index:1'>"
		response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&total(i,2)&"</td></tr></table></v:textbox></v:shape>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px,"&table_y&"px' to='"&table_x+left_width+all_width*abs(temp6)/(temp3-temp6)&"px,"&table_y+all_height&"px'></v:line>"
	next
	
	case else
	end select
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>"
		response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>"
	'end if

end function
%>




