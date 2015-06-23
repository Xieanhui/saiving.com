//�滻
/*
����ֵ:�滻���ֵ
����˵��:str_raw,ԭֵ;strFind Ҫ�滻��ֵ; strReplace�滻��
*/
function f_replace(str_raw,strFind,strReplace)
{
    var tmpval=str_raw.toString();
    return tmpval.split(strFind).join(strReplace);
} 

//����ʱ��,��������
/*����ֵ:true��,false��
����˵��:valobj,����Ŀؼ�����
*/
function f_validatetime(valobj)
{
	timeStr=valobj.value.replace("��",":")
	if (timeStr=="") return false;
	if (timeStr.length==4)
	timeStr=timeStr.substr(0,2)+":"+timeStr.substr(2,2)

	valobj.value=timeStr
	var timePat = /^(\d{1,2}):(\d{1,2})$/;

	var matchArray = timeStr.match(timePat);
	if (matchArray == null) 
	{
		alert("�����ʱ����������ո�ʽ:Сʱ:����!");
		valobj.value="08:00"
		return false;
	}
	hour = matchArray[1];
	minute = matchArray[2];

	if (hour < 0  || hour > 23) 
	{
		alert("Сʱ��������00--23֮��!");
		hour=8;
	}

	if (minute < 0 || minute > 59) 
	{
		alert ("������������00--59֮��!");
		minute=0;
	}
	valobj.value=('00'+hour).substr(('00'+hour).length-2,2)+":"+('00'+minute).substr(('00'+minute).length-2,2)
	return true;
}

//��������,��������
/*
����ֵ:true��,false��
����˵��:valobj,����Ŀؼ�����
*/
function f_validatedate(valobj) 
{
	dar=f_replace(valobj.value,".","-")
	if (dar=="") return false;
	if(dar.split("-")[0].length==2)
	{
		var Current_Date = new Date();
		var Current_year = Current_Date.getYear();
		dar=Current_year.toString().substr(0,2)+dar
	}
	var datePat = /^(\d{4})(-)(\d{1,2})(-)(\d{1,2})$/;

	var matchArray = dar.match(datePat); // is the format ok?
	if (matchArray == null) 
	{
		alert("�����������������ո�ʽ:��������-����-����!");
		valobj.value=""
		return false;
	}
	month = matchArray[3]; // parse date into variables
	day = matchArray[5];
	year = matchArray[1];
	if (month < 1 || month > 12) 
	{ // check month range
		alert("�·ݳ���!");
		month=1;
	}
	if (day < 1 || day > 31) 
	{
		alert("���ڳ���!");
		day=1;
	}
	if ((month==4 || month==6 || month==9 || month==11) && day==31) 
	{
		alert(month+"��û��31��!");
		day=30;
	}
	if (month == 2) { // check for february 29th
	var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
	if (day>29) 
	{
		alert("2�²��ܳ���29��!");
		day=28;
	}
	if (day==29 && !isleap) 
	{
		alert(year + "�겻�����꣬2��û��29��!");
		day=28;
	}
	}
	valobj.value=year+"-"+('0'+month).substr(('0'+month).length-2,2)+"-"+('0'+day).substr(('0'+day).length-2,2)
	return true;
}
