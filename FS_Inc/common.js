//替换
/*
返回值:替换后的值
参数说明:str_raw,原值;strFind 要替换的值; strReplace替换成
*/
function f_replace(str_raw,strFind,strReplace)
{
    var tmpval=str_raw.toString();
    return tmpval.split(strFind).join(strReplace);
} 

//检验时间,网上下载
/*返回值:true真,false假
参数说明:valobj,检验的控件对象
*/
function f_validatetime(valobj)
{
	timeStr=valobj.value.replace("：",":")
	if (timeStr=="") return false;
	if (timeStr.length==4)
	timeStr=timeStr.substr(0,2)+":"+timeStr.substr(2,2)

	valobj.value=timeStr
	var timePat = /^(\d{1,2}):(\d{1,2})$/;

	var matchArray = timeStr.match(timePat);
	if (matchArray == null) 
	{
		alert("输入的时间有误，请参照格式:小时:分钟!");
		valobj.value="08:00"
		return false;
	}
	hour = matchArray[1];
	minute = matchArray[2];

	if (hour < 0  || hour > 23) 
	{
		alert("小时数必须在00--23之间!");
		hour=8;
	}

	if (minute < 0 || minute > 59) 
	{
		alert ("分钟数必须在00--59之间!");
		minute=0;
	}
	valobj.value=('00'+hour).substr(('00'+hour).length-2,2)+":"+('00'+minute).substr(('00'+minute).length-2,2)
	return true;
}

//检验日期,网上下载
/*
返回值:true真,false假
参数说明:valobj,检验的控件对象
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
		alert("输入的日期有误，请参照格式:年年年年-月月-日日!");
		valobj.value=""
		return false;
	}
	month = matchArray[3]; // parse date into variables
	day = matchArray[5];
	year = matchArray[1];
	if (month < 1 || month > 12) 
	{ // check month range
		alert("月份超界!");
		month=1;
	}
	if (day < 1 || day > 31) 
	{
		alert("日期超界!");
		day=1;
	}
	if ((month==4 || month==6 || month==9 || month==11) && day==31) 
	{
		alert(month+"月没有31日!");
		day=30;
	}
	if (month == 2) { // check for february 29th
	var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
	if (day>29) 
	{
		alert("2月不能超过29日!");
		day=28;
	}
	if (day==29 && !isleap) 
	{
		alert(year + "年不是闰年，2月没有29日!");
		day=28;
	}
	}
	valobj.value=year+"-"+('0'+month).substr(('0'+month).length-2,2)+"-"+('0'+day).substr(('0'+day).length-2,2)
	return true;
}
