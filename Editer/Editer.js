var g_editerpath='';
var g_contentid='';
var g_picpath='';
var g_Buttons=new CLS_Buttons();
var g_editingcontrol=null;
var g_popupmenu = window.createPopup();
var g_currmode='EDIT';
var g_editertype=0;
var g_ParentDispalyNone=0;
var g_tablebordershown=1;
function LoadEditFile(f_picpath,f_editerpath,f_contentid,f_editertype){
	g_editerpath=f_editerpath;
	g_contentid=f_contentid;
	g_picpath=f_picpath;
	g_editertype=f_editertype;
	EditArea.focus();
	InitialButtons();
	var f_ParentDispalyNone=0,f_Obj=window.parent.document.getElementById("ParentDispalyNone");
	if((f_Obj)&&(f_Obj.value.toLowerCase()=='none'))f_ParentDispalyNone=1;
	if(f_ParentDispalyNone==0)LayoutAndSetContent();
	return true;
}
function LayoutAndSetContent(){
		LayoutButtons();
		SetNewsContentArray();
		ShowTableBorders();
		SetBodyStyle();
		InitialPopupMenu();
		window.onresize=LayoutButtons;
		document.oncontextmenu=function(){return true;}
		window.onerror=function(msg, url, line){return true;}
}
function InitialButtons(){
	g_Buttons.Add(new CLS_Button('undo.gif','撤消','Format(\'undo\')'));
	g_Buttons.Add(new CLS_Button('redo.gif','恢复','Format(\'redo\')'));
	g_Buttons.Add(new CLS_Button('find.gif','查找 / 替换','SearchStr()'));
	g_Buttons.Add(new CLS_Button('calculator.gif','计算器','Calculator()'));
	g_Buttons.Add(new CLS_Button('date.gif','插入当前日期','InsertDate()'));
	g_Buttons.Add(new CLS_Button('time.gif','插入当前时间','InsertTime()'));
	g_Buttons.Add(new CLS_Button('geshi.gif','删除所有HTML标识','DelAllHtmlTag()'));
	g_Buttons.Add(new CLS_Button('clear.gif','删除文字格式','Format(\'removeformat\')'));
	g_Buttons.Add(new CLS_Button('url.gif','插入超级连接','Format(\'CreateLink\')'));
	g_Buttons.Add(new CLS_Button('nourl.gif','取消超级链接','Format(\'unLink\')'));
	g_Buttons.Add(new CLS_Button('htm.gif','插入网页','InsertPage()'));
	g_Buttons.Add(new CLS_Button('fieldset.gif','插入栏目框','InsertFrame()'));
	g_Buttons.Add(new CLS_Button('Marquee.gif','插入滚动文本','InsertMarquee()'));
	g_Buttons.Add(new CLS_Button('PicAlign.gif','图文并排','PicAndTextArrange()'));
	g_Buttons.Add(new CLS_Button('PasteText.gif','文本粘贴','PureTextPaste()'));
	g_Buttons.Add(new CLS_Button('PasteWord.gif','清理 Word 垃圾代码','PasteWord()'));
	g_Buttons.Add(new CLS_Button('bold.gif','加粗','Format(\'bold\')'));
	g_Buttons.Add(new CLS_Button('italic.gif','斜体','Format(\'italic\')'));
	g_Buttons.Add(new CLS_Button('underline.gif','下划线','Format(\'underline\')'));
	g_Buttons.Add(new CLS_Button('TextColor.gif','文字颜色','TextColor()'));
	g_Buttons.Add(new CLS_Button('fgbgcolor.gif','文字背景色','TextBGColor()'));
	g_Buttons.Add(new CLS_Button('sline.gif','插入特殊水平线','SpecialHR()'));
	g_Buttons.Add(new CLS_Button('line.gif','插入普通水平线','InsertHR()'));
	g_Buttons.Add(new CLS_Button('chars.gif','插入换行符号','InsertBR()'));
	g_Buttons.Add(new CLS_Button('outdent.gif','减少缩进量','Format(\'outdent\')'));
	g_Buttons.Add(new CLS_Button('indent.gif','增加缩进量','Format(\'indent\')'));
	g_Buttons.Add(new CLS_Button('num.gif','编号','Format(\'insertorderedlist\')'));
	g_Buttons.Add(new CLS_Button('list.gif','项目符号','Format(\'insertunorderedlist\')'));
	g_Buttons.Add(new CLS_Button('Aleft.gif','左对齐','Format(\'justifyleft\')'));
	g_Buttons.Add(new CLS_Button('Acenter.gif','居中','Format(\'justifycenter\')'));
	g_Buttons.Add(new CLS_Button('Aright.gif','右对齐','Format(\'justifyright\')'));
	g_Buttons.Add(new CLS_Button('Inserttable.gif','插入表格','InsertTable()'));
	if(g_editertype==1){	
		g_Buttons.Add(new CLS_Button('Load.gif','插入附件，支持格式为：zip、rar等','InsertLoad()'));
		g_Buttons.Add(new CLS_Button('img.gif','插入图片，支持格式为：jpg、gif、bmp、png等','InsertPicture()'));
		g_Buttons.Add(new CLS_Button('flash.gif','插入flash多媒体文件','InsertFlash()'));
		g_Buttons.Add(new CLS_Button('wmv.gif','插入视频文件，支持格式为：avi、wmv、asf、mpg','InsertVideo()'));
		g_Buttons.Add(new CLS_Button('rm.gif','插入RealPlay文件，支持格式为：rm、ra、ram','InsertRM()'));
	}
	Options=new Array(new Option('字体',''),new Option('宋体','宋体')
	,new Option('黑体','黑体'),new Option('楷体','楷体_GB2312'),new Option('仿宋','仿宋_GB2312')
	,new Option('隶书','隶书'),new Option('幼圆','幼圆'),new Option('Arial','Arial')
	,new Option('Arial Blac','Arial Blac'),new Option('Arial Narrow','Arial Narrow'),new Option('Brush Script MT','Brush Script MT')
	,new Option('Century Gothic','Century Gothic'),new Option('Comic Sans MS','Comic Sans MS'),new Option('Courier','Courier')
	,new Option('Courier New','Courier New'),new Option('MS Sans Serif','MS Sans Serif'),new Option('Script','Script')
	,new Option('System','System'),new Option('Times New Roman','Times New Roman'),new Option('Verdana','Verdana')
	,new Option('Wide Latin','Wide Latin'),new Option('Wingdings','Wingdings'));
	g_Buttons.Add(new CLS_Select(Options,'Format(\'fontname\',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();',133));
	Options=new Array(new Option('字号',''),new Option('一号','7')
		,new Option('二号','6'),new Option('三号','5'),new Option('四号','4')
		,new Option('五号','3'),new Option('六号','2'),new Option('七号','1'));
	g_Buttons.Add(new CLS_Select(Options,'Format(\'fontsize\',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();',52));
	Options=new Array(new Option('段落样式',''),new Option('普通','&lt;P&gt;')
		,new Option('标题一','&lt;H1&gt;'),new Option('标题二','&lt;H2&gt;'),new Option('标题三','&lt;H3&gt;')
		,new Option('标题四','&lt;H4&gt;'),new Option('标题五','&lt;H5&gt;'),new Option('标题六','&lt;H6&gt;')
		,new Option('段落','&lt;p&gt;'),new Option('&lt;dd&gt;','定义'),new Option('术语定义','&lt;dt&gt;')
		,new Option('目录列表','&lt;dir&gt;'),new Option('菜单列表','&lt;menu&gt;'),new Option('已编排格式','&lt;PRE&gt;'))
	g_Buttons.Add(new CLS_Select(Options,'Format(\'FormatBlock\',this[this.selectedIndex].value);this.selectedIndex=0;EditArea.focus();',92));
}
function InitialPopupMenu(){
	EditArea.document.body.contentEditable="true";
	EditArea.document.onmouseup=new Function("return SearchObject(EditArea.event);");
	EditArea.document.oncontextmenu=new Function("return ShowMouseRightMenu(EditArea.event);");	
}
function CLS_Button(_Pic,_Title,_Fun,_Width){
	this.Pic=_Pic;
	this.Title=_Title;
	this.Fun=_Fun;
	if(_Width)this.Width=_Width;else this.Width=30;
	this.HTML=function(){
		return '<div class=\"Toolbar_row_Button\"><a class=\"Btn\" href=\"javascript:void(0);\"><img class=\"Btn\" border=\"0\" src="'+g_picpath+this.Pic+'" title="'+this.Title+'" onClick="'+this.Fun+'"></a></div>'
	}
}
function Option(_text,_value){
	this.text=_text;
	this.value=_value;
}
function CLS_Select(_Options,_Fun,_Width){
	this.Options=_Options;
	this.Fun=_Fun;
	if(_Width)this.Width=_Width;else this.Width=30;
	this.HTML=function(){
		var _HTML='';
		for(var i=0;i<this.Options.length;i++){
			_HTML+='<option value=\"'+this.Options[i].value+'\">'+this.Options[i].text+'</option>';
		}
		return '<div class="Toolbar_row_Button"><select class=\"ToolSelectStyle\" onChange=\"'+_Fun+'\">'+_HTML+'</SELECT></div>';
	}
}
function CLS_Buttons(){
	this.Buttons=new Array();
	this.Add=function(_Button){
		this.Buttons.push(_Button);
	}
	this.Items=function(_i){
		return this.Buttons[_i];
	}
	this.length=function(){
		return this.Buttons.length;
	}
	this.RowHTML=function(_B,_E){
		var _HTML='';
		for(var i=_B;i<=_E;i++)
			_HTML+=this.Buttons[i].HTML();
		return '<div class=\"Toolbar_row\">'+_HTML+'</div>';
	}
}
function LayoutButtons(){
	var BodyWidth=parseInt(document.body.clientWidth);
	var BodyHeight=parseInt(document.body.clientHeight);
	var _RowIsEnd=false,_ButtonsWidth=0,_ButtonsHeight=0,_BIndex=_Index=0,_EIndex=0,_HTML='',_Button,_ButtonsLength=g_Buttons.length();
	while(_Index<_ButtonsLength){
		_Button=g_Buttons.Items(_Index);
		_ButtonsWidth+=_Button.Width;
		if(_ButtonsWidth>BodyWidth){
			_EIndex=_Index-1;
			_HTML+=g_Buttons.RowHTML(_BIndex,_EIndex);
			_RowIsEnd=true;
			_ButtonsHeight+=30;
			_ButtonsWidth=0;
			_BIndex=_Index;
		}else{
			_Index++;
			_RowIsEnd=false;
		}
	}
	if(!_RowIsEnd){
		_HTML+=g_Buttons.RowHTML(_BIndex,_ButtonsLength-1);
		_ButtonsHeight+=30;
	}
	Toolbar.innerHTML=_HTML;
	Toolbar.style.height=_ButtonsHeight;
	var EditAreaHeight=BodyHeight-parseInt(document.all.Toolbar.style.height)-63;
	document.all.EditArea.height=EditAreaHeight;
}
function SetNewsContentArray(){
	var IsAlert=false;
	if(g_contentid!=''){
		var ContentObj=window.parent.document.getElementById(g_contentid);
		if(ContentObj){
			EditArea.document.designMode="On";
			EditArea.document.open();
			EditArea.document.write("<head></head><body MONOSPACE>"+unescape(ContentObj.value)+"</body>");
			EditArea.document.body.contentEditable="true";
			EditArea.document.execCommand("2D-Position",true,true);
			EditArea.document.execCommand("MultipleSelection", true, true);
			EditArea.document.execCommand("LiveResize", true, true);
			EditArea.document.close();
			//unescapeALink();
			ShowTableBorders();
		}else{IsAlert=true;}
	}else{IsAlert=true;}
	if(IsAlert){alert('内容容器不存在，请和系统管理员联系！');}
}
function GetNewsContentArray(){
	return EditArea.document.body.innerHTML;
}
function SetBodyStyle(){
	EditArea.document.body.runtimeStyle.fontSize='9pt';
}
function YCancelEvent() {
	event.returnValue=false;
	event.cancelBubble=true;
	return false;
}
function setMode(NewMode){  
	if (NewMode!=g_currmode){   
		if (NewMode=='TEXT'){
			if (!confirm("警告！切换到纯文本模式会丢失您所有的HTML格式，您确认切换吗？")) return false;
		}
		var sBody='';
		switch(g_currmode){
			case "CODE":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerText;
				break;
			case "TEXT":
				sBody=EditArea.document.body.innerText;
				sBody=HTMLEncode(sBody);
				break;
			case "EDIT":
			case "VIEW":
				if (NewMode=="TEXT") sBody=EditArea.document.body.innerText;
				else sBody=EditArea.document.body.innerHTML;
				break;
		}
		//sBody=sBody.replace(/　/ig,'&nbsp;');
		//alert(sBody);
		document.all["Editer_CODE"].className='Toolbar_row_Button2 ModeBarBtnOff';
		document.all["Editer_EDIT"].className='Toolbar_row_Button2 ModeBarBtnOff';
		document.all["Editer_TEXT"].className='Toolbar_row_Button2 ModeBarBtnOff';
		document.all["Editer_VIEW"].className='Toolbar_row_Button2 ModeBarBtnOff';
		document.all["Editer_"+NewMode].className='Toolbar_row_Button2 ModeBarBtnOn';
		switch (NewMode){
			case "CODE":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head></head><body MONOSPACE>");
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "EDIT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head></head><body MONOSPACE>"+sBody+"</body>");
				EditArea.document.body.contentEditable="true";
				EditArea.document.execCommand("2D-Position",true,true);
				EditArea.document.execCommand("MultipleSelection", true, true);
				EditArea.document.execCommand("LiveResize", true, true);
				EditArea.document.close();
				ShowTableBorders();
				DisabledAllBtn(false);
				break;
			case "TEXT":
				EditArea.document.designMode="On";
				EditArea.document.open();
				EditArea.document.write("<head></head><body MONOSPACE>");
				EditArea.document.body.innerText=sBody;
				EditArea.document.body.contentEditable="true";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "VIEW":
				EditArea.document.designMode="off";
				EditArea.document.open();
				EditArea.document.write("<head></head><body MONOSPACE>"+sBody);
				EditArea.document.body.contentEditable="false";
				EditArea.document.close();
				DisabledAllBtn(true);
				break;
		}
		g_currmode=NewMode;
		if (NewMode!='EDIT') EmptyShowObject(true);
		else {EmptyShowObject(false);InitialPopupMenu();}
		SetBodyStyle();
	}
	EditArea.focus();
}
function EmptyShowObject(Flag){
	document.all.ShowObject.innerHTML='&nbsp;';
	document.all.ShowObject.disabled=Flag;
}
function HTMLEncode(text){
	text = text.replace(/&/g, "&amp;") ;
	text = text.replace(/"/g, "&quot;") ;
	text = text.replace(/</g, "&lt;") ;
	text = text.replace(/>/g, "&gt;") ;
	text = text.replace(/'/g, "&#146;") ;
	text = text.replace(/\ /g,"&nbsp;");
	text = text.replace(/\n/g,"<br>");
	text = text.replace(/\t/g,"&nbsp;&nbsp;&nbsp;&nbsp;");
	return text;
}
function ShowMouseRightMenu(event){
	var width=86;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=g_popupmenu.document;
	var ObjPopBody=g_popupmenu.document.body;
	var MenuStr='';
	MenuStr+=FormatMenuRow("selectall", "全选","SelectAll.gif");
	MenuStr+=FormatMenuRow("cut", "剪切","Cut.gif");
	MenuStr+=FormatMenuRow("copy", "复制","Copy.gif");
	MenuStr+=FormatMenuRow("paste", "粘贴","Paste.gif");
	MenuStr+=FormatMenuRow("delete", "删除","Del.gif");
	height+=100;
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=86><tr><td width=86 class=RightBg><TABLE border=0 cellpadding=0 cellspacing=0>"+MenuStr
	MenuStr=MenuStr+"<\/TABLE><\/td><\/tr><\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href=\""+g_editerpath+"MenuCSS.css\" type=\"text/css\" rel=\"stylesheet\"></head><body ondrag=\"return false;\" scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=5;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	g_popupmenu.show(lefter, topper, width, height, EditArea.document.body);
	return false;
}
function GetMenuRowStr(DisabledStr, MenuOperation, MenuImage, MenuDescripion){
	var MenuRowStr='';
	MenuRowStr="<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=81><tr "+DisabledStr+"><td valign=middle height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
	if (DisabledStr==''){
		MenuRowStr += " onclick=\"parent."+MenuOperation+";parent.g_popupmenu.hide();\"";
	}
	MenuRowStr+=">"
	if (MenuImage!=""){
		MenuRowStr+="&nbsp;<img border=0 src='"+g_picpath+MenuImage+"' width=20 height=20 align=absmiddle "+DisabledStr+">&nbsp;";
	}
	else{
		MenuRowStr+="&nbsp;";
	}
	MenuRowStr+=MenuDescripion+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
	return MenuRowStr;
}
function FormatMenuRow(MenuStr,MenuDescription,MenuImage){
	var DisabledStr='';
	var ShowMenuImage='';
	if (!EditArea.document.queryCommandEnabled(MenuStr)){
		DisabledStr="disabled";
	}
	var MenuOperation="Format('"+MenuStr+"')";
	if (MenuImage){
		ShowMenuImage=MenuImage;
	}
	return GetMenuRowStr(DisabledStr,MenuOperation,ShowMenuImage,MenuDescription)
}
function SearchObject(){
	UpdateToolbar();
}
function MouseRightMenuItem(CommandString, CommandId){
	this.CommandString = CommandString;
	this.CommandId = CommandId;
}
function GetEditAreaSelectionType(){
	return EditArea.document.selection.type;
}
function ExeEditAttribute(){
	OpenWindow(g_editerpath+'AttributeWindow.htm',360,120,window)
	EditArea.focus();
}
function AddLink(){
	var sText = EditArea.document.selection.createRange();
	if (!sText==""){
		var temp=EditArea.document.execCommand("CreateLink");
		if (sText.parentElement().tagName == "A"){
			sText.parentElement().innerText=sText.parentElement().href;
			EditArea.document.execCommand("ForeColor","false","#FF0033");
		}
	}else{
		alert("Please select some blue text!");
	}
}
function unescapeALink(){
	var Reg=new RegExp('http:\/\/'+location.hostname+':'+location.port+'/.*/|http:\/\/'+location.hostname+'/.*/|http:\/\/'+location.hostname);
	var Links=EditArea.document.getElementsByTagName('A');
	for(var i=0;i<Links.length;i++){
		Links[i].href=unescape(Links[i].href);
		var LinkStr=Links[i].href?Links[i].href:'';
		if(LinkStr.match(Reg))LinkStr=LinkStr.replace(Reg,'');
		Links[i].href=LinkStr;
	}
}
function InsertHTMLStr(Str){
	EditArea.focus();
	if (EditArea.document.selection.type.toLowerCase() != "none"){
		EditArea.document.selection.clear() ;
	}
	EditArea.document.selection.createRange().pasteHTML(Str); 
	//if(g_currmode=='EDIT')EditArea.document.selection.createRange().pasteHTML(Str); 
	//else{
		//EditArea.document.selection.createRange().pasteHTML(escape(Str));
	//}
	EditArea.focus();
	ShowTableBorders();
	//unescapeALink();
}
function InsertPicture(){
	var ReturnValue=OpenWindow('Picture.asp',420,180,window);
	if (ReturnValue!=''){
		var TempArray=ReturnValue.split("$$$");
		InsertHTMLStr(TempArray[0]);
	}
	EditArea.focus();
}
function QueryCommand(CommandID){
	var State=EditArea.QueryStatus(CommandID)
	if (State==3) return true;
	else return false;
}
function Format(Operation,Val){
	EditArea.focus();
	if (Val=="RemoveFormat"){
		Operation=Val;
		Val=null;
	}
	if (Val==null) EditArea.document.execCommand(Operation);
	else EditArea.document.execCommand(Operation,"",Val);
	EditArea.focus();
}
function TextBGColor(){
	EditArea.focus();
	var EditRange = EditArea.document.selection.createRange();
	var RangeType = EditArea.document.selection.type;
	if (RangeType!="Text"){
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow(g_editerpath+'SelectColor.htm',230,190,window);
	if (ReturnValue!=null){
		EditRange.pasteHTML("<span style='background-color:"+ReturnValue+"'>"+EditRange.text+"</span> ");
		EditRange.select();
	}
	EditArea.focus();
}
function Print(CommandID){
	EditArea.focus();
	if (EditArea.QueryStatus(CommandID)!=3) EditArea.ExecCommand(CommandID,0);
	EditArea.focus();
}
function InsertTable(){
	var ReturnValue=OpenWindow(g_editerpath+'InsertTable.htm',290,110,window);
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertPage(){
	var ReturnValue=OpenWindow(g_editerpath+'Page.htm',320,110,window);
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertMarquee(){
	EditArea.focus();
	var ReturnValue=OpenWindow(g_editerpath+'Marquee.htm',260,50,window); 
	InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function Calculator(){
	EditArea.focus();
	var ReturnValue=OpenWindow(g_editerpath+'Calculator.htm',160,180,window);
	if (ReturnValue!=null){
		var TempArray,ParameterA,ParameterB;
		TempArray=ReturnValue.split("*")
		ParameterA=TempArray[0];
		InsertHTMLStr(ParameterA);
	}
	EditArea.focus();
}
function InsertDate(){
	EditArea.focus();
	var NowDate = new Date();
	var FormateDate=NowDate.getYear()+"年"+(NowDate.getMonth() + 1)+"月"+NowDate.getDate() +"日";
	InsertHTMLStr(FormateDate);
	EditArea.focus();
}
function InsertTime(){
	EditArea.focus();
	var NowDate=new Date();
	var FormatTime=NowDate.getHours() +":"+NowDate.getMinutes()+":"+NowDate.getSeconds();
	InsertHTMLStr(FormatTime);
	EditArea.focus();
}
function InsertFrame(){
	EditArea.focus();
	var ReturnVlaue =OpenWindow(g_editerpath+'Frame.htm',280,118,window);
	if (ReturnVlaue != null){
		InsertHTMLStr(ReturnVlaue);
	}
	EditArea.focus();
}
function InsertBR(Index){
	EditArea.focus();
	InsertHTMLStr('<br>');
	EditArea.focus();
}
function DelAllHtmlTag(){
	var TempStr;
	TempStr=EditArea.document.body.innerHTML;
	var re=/<\/*[^<>]*>/ig
	TempStr=TempStr.replace(re,'');
	EditArea.document.body.innerHTML=TempStr;
	EditArea.focus();
}
function InsertFlash(){
  var ReturnValue = OpenWindow('Flash.asp',380,100,window); 
  if (ReturnValue!=''){
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function InsertVideo(){
  var ReturnValue=OpenWindow('Video.asp',400,100,window);
  if (ReturnValue!=''){
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function InsertRM(){
  var ReturnValue=OpenWindow('RM.asp',400,100,window);  
  if (ReturnValue!=''){
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function InsertLoad(){
  var ReturnValue=OpenWindow('Load.asp',400,50,window);  
  if (ReturnValue!=''){
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  EditArea.focus();
}
function SpecialHR(){
	EditArea.focus();
	var ReturnValue = OpenWindow(g_editerpath+'SpecialHR.htm',320,120,window); 
	if (ReturnValue!= null) InsertHTMLStr(ReturnValue);
	EditArea.focus();
}
function InsertHR(){
	EditArea.focus();
	InsertHTMLStr('<hr>');
	EditArea.focus();
}
function ShowTableBorders(){
	AllTables=EditArea.document.body.getElementsByTagName("TABLE");
	for(var i=0;i<AllTables.length;i++){
		if ((AllTables[i].border==null)||(AllTables[i].border=='0')){
			AllTables[i].runtimeStyle.borderTop=AllTables[i].runtimeStyle.borderLeft="1px dotted #709FCB";
			AllRows = AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++){
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++){
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom="1px dotted #709FCB";
				}
			}
		}
		else
		{
			AllTables[i].runtimeStyle.borderTop='';
			AllRows=AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++){
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++){
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom='';
				}
			}
		}
	}
  g_tablebordershown=g_tablebordershown?0:1;
}
function ImageSelected(){
	EditArea.focus();
	if (EditArea.document.selection.type=="Control"){
		var oControlRange=EditArea.document.selection.createRange();
		if (oControlRange(0).tagName.toUpperCase()=="IMG"){
			selectedImage=EditArea.document.selection.createRange()(0);
			return true;
		}	
	}
}
function TextColor(){	
	EditArea.focus();
	var EditRange = EditArea.document.selection.createRange();
	var RangeType = EditArea.document.selection.type;
	if (RangeType!="Text"){
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow(g_editerpath+'SelectColor.htm',230,190,window);
	if (ReturnValue!=null){
		EditRange.pasteHTML("<font color='"+ReturnValue+"'>"+EditRange.text+"</font>");
		EditRange.select();
	}
	EditArea.focus();
}
function PicAndTextArrange(){
	if(ImageSelected())	{
		sPrePos=selectedImage.style.position;
		var ReturnValue=OpenWindow(g_editerpath+'SelectPicStyle.htm',380,130,window);
		if(ReturnValue)	{
			for(key in ReturnValue)
			if(key=='style') for(sub_key in ReturnValue.style) selectedImage.style[sub_key]=ReturnValue.style[sub_key];
			else selectedImage[key]=ReturnValue[key];
			if(!ReturnValue.align) selectedImage.removeAttribute('align');
			if(sPrePos.match(/^absolute$/i) && !selectedImage.style.position.match(/^absolute$/i)){
				sFired = selectedImage.parentElement;
				while(!sFired.tagName.match(/^table$|^body$/i))
				sFired = sFired.parentElement;
				if(sFired.tagName.match(/^table$/i) && sFired.style.position.match(/absolute/i));
				sFired.outerHTML=selectedImage.outerHTML;
			}else{
				if(!sPrePos.match(/^absolute$/i) && selectedImage.style.position.match(/^absolute$/i)) selectedImage.outerHTML='<table style="position: absolute;"><tr><td>' + selectedImage.outerHTML + '</td></tr></table>';
			}
		}
	}
	else alert('请选择图片');
}
function GetAllAncestors(){
	var p = GetParentElement();
	var a = [];
	while (p && (p.nodeType==1)&&(p.tagName.toLowerCase()!='body')){
		a.push(p);
		p=p.parentNode;
	}
	a.push(EditArea.document.body);
	return a;
}
function GetParentElement(){
	var sel=GetSelection();
	var range=CreateRange(sel);
	switch (sel.type){
		case "Text":
		case "None":
			return range.parentElement();
		case "Control":
			return range.item(0);
		default:
			return EditArea.document.body;
	}
}
function GetSelection(){
	return EditArea.document.selection;
}
function CreateRange(sel){
	return sel.createRange();
}
function UpdateToolbar(){
	var ancestors=null;
	ancestors=GetAllAncestors();
	document.all.ShowObject.innerHTML='&nbsp;';
	for (var i=ancestors.length;--i>=0;){
		var el = ancestors[i];
		if (!el) continue;
		var a=document.createElement("span");
		a.href="#";
		a.el=el;
		a.editor=this;
		if (i==0){
			a.className='AncestorsMouseUp';
			g_editingcontrol=a.el;
		}
		else a.className='AncestorsStyle';
		a.onmouseover=function(){
			if (this.className=='AncestorsMouseUp') this.className='AncestorsMouseUpOver';
			else if (this.className=='AncestorsStyle') this.className='AncestorsMouseOver';
		};
		a.onmouseout=function(){
			if (this.className=='AncestorsMouseUpOver') this.className='AncestorsMouseUp';
			else if (this.className=='AncestorsMouseOver') this.className='AncestorsStyle';
		};
		a.onmousedown=function(){this.className='AncestorsMouseDown';};
		a.onmouseup=function(){this.className='AncestorsMouseUpOver';};
		a.ondragstart=YCancelEvent;
		a.onselectstart=YCancelEvent;
		a.onselect=YCancelEvent;
		a.onclick=function(){
			this.blur();
			SelectNodeContents(this);
			return false;
		};
		var txt='<'+el.tagName.toLowerCase();
		a.title=el.style.cssText;
		if (el.id) txt += "#" + el.id;
		if (el.className) txt += "." + el.className;
		txt=txt+'>';
		a.appendChild(document.createTextNode(txt));
		document.all.ShowObject.appendChild(a);
	}
}
function SelectNodeContents(Obj,pos){
	for (var i=0;i<document.all.ShowObject.children.length;i++){
		if (document.all.ShowObject.children(i).className=='AncestorsMouseUp') document.all.ShowObject.children(i).className='AncestorsStyle';
	}
	var node=Obj.el;
	g_editingcontrol=node;
	EditArea.focus();
	var collapsed=(typeof pos!='undefined');
	var range = EditArea.document.body.createTextRange();
	range.moveToElementText(node);
	(collapsed) && range.collapse(pos);
	range.select();
}
function DeleteHTMLTag(){
	var AvailableDeleteTagName='p,a,div,span';
	if (g_editingcontrol!=null){
		var DeleteTagName=g_editingcontrol.tagName.toLowerCase();
		if (AvailableDeleteTagName.indexOf(DeleteTagName)!=-1){
			g_editingcontrol.parentElement.innerHTML=g_editingcontrol.innerHTML;
		}
	}
	UpdateToolbar();
	ShowTableBorders();
}
function InsertHref(Operation){
	EditArea.focus();
	EditArea.document.execCommand(Operation,true);
	EditArea.focus();
}
function SearchStr(){
	var Temp=window.showModalDialog(g_editerpath+"Search.htm", window, "dialogWidth:320px; dialogHeight:170px; help: no; scroll: no; status: no");
	EditArea.focus();
}
function DisabledAllBtn(Flag){
	var AllBtnArray=document.body.getElementsByTagName('IMG'),CurrObj=null;
	for (var i=0;i<AllBtnArray.length;i++){
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='Btn') CurrObj.disabled=Flag;
	}
	var AllBtnArray=document.body.getElementsByTagName('A'),CurrObj=null;
	for (var i=0;i<AllBtnArray.length;i++){
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='Btn') CurrObj.disabled=Flag;
	}
	AllBtnArray=document.body.getElementsByTagName('SELECT');
	for (var i=0;i<AllBtnArray.length;i++){
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='ToolSelectStyle') CurrObj.disabled=Flag;
	}
}

function PureTextPaste(){
	EditArea.focus();
	var sText = HTMLEncode(clipboardData.getData("Text")) ;
	InsertHTMLStr(sText);
	EditArea.focus();
} 
function PasteWord(){
	EditArea.focus();
	var html = EditArea.document.body.innerHTML
	EditArea.document.body.innerHTML='';
	html = html.replace(/<o:p>\s*<\/o:p>/g, "") ;
	html = html.replace(/<o:p>.*?<\/o:p>/g, "&nbsp;") ;
	html = html.replace( /\s*mso-[^:]+:[^;"]+;?/gi, "" ) ;
	html = html.replace( /\s*MARGIN: 0cm 0cm 0pt\s*;/gi, "" ) ;
	html = html.replace( /\s*MARGIN: 0cm 0cm 0pt\s*"/gi, "\"" ) ;
	html = html.replace( /\s*TEXT-INDENT: 0cm\s*;/gi, "" ) ;
	html = html.replace( /\s*TEXT-INDENT: 0cm\s*"/gi, "\"" ) ;
	html = html.replace( /\s*TEXT-ALIGN: [^\s;]+;?"/gi, "\"" ) ;
	html = html.replace( /\s*PAGE-BREAK-BEFORE: [^\s;]+;?"/gi, "\"" ) ;
	html = html.replace( /\s*FONT-VARIANT: [^\s;]+;?"/gi, "\"" ) ;
	html = html.replace( /\s*tab-stops:[^;"]*;?/gi, "" ) ;
	html = html.replace( /\s*tab-stops:[^"]*/gi, "" ) ;
	html = html.replace( /\s*face="[^"]*"/gi, "" ) ;
	html = html.replace( /\s*face=[^ >]*/gi, "" ) ;
	html = html.replace( /\s*FONT-FAMILY:[^;"]*;?/gi, "" ) ;
	html = html.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, "<$1$3") ;
	html =  html.replace( /\s*style="\s*"/gi, '' ) ;
	html = html.replace( /<SPAN\s*[^>]*>\s*&nbsp;\s*<\/SPAN>/gi, '&nbsp;' ) ;
	html = html.replace( /<SPAN\s*[^>]*><\/SPAN>/gi, '' ) ;
	html = html.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, "<$1$3") ;
	html = html.replace( /<SPAN\s*>(.*?)<\/SPAN>/gi, '$1' ) ;
	html = html.replace( /<FONT\s*>(.*?)<\/FONT>/gi, '$1' ) ;
	html = html.replace(/<\\?\?xml[^>]*>/gi, "") ;
	html = html.replace(/<\/?\w+:[^>]*>/gi, "") ;
	html = html.replace( /<H\d>\s*<\/H\d>/gi, '' ) ;
	html = html.replace( /<H1([^>]*)>/gi, '<div$1><b><font size="6">' ) ;
	html = html.replace( /<H2([^>]*)>/gi, '<div$1><b><font size="5">' ) ;
	html = html.replace( /<H3([^>]*)>/gi, '<div$1><b><font size="4">' ) ;
	html = html.replace( /<H4([^>]*)>/gi, '<div$1><b><font size="3">' ) ;
	html = html.replace( /<H5([^>]*)>/gi, '<div$1><b><font size="2">' ) ;
	html = html.replace( /<H6([^>]*)>/gi, '<div$1><b><font size="1">' ) ;
	html = html.replace( /<\/H\d>/gi, '</font></b></div>' ) ;
	html = html.replace( /<(U|I|STRIKE)>&nbsp;<\/\1>/g, '&nbsp;' ) ;
	html = html.replace( /<([^\s>]+)[^>]*>\s*<\/\1>/g, '' ) ;
	html = html.replace( /<([^\s>]+)[^>]*>\s*<\/\1>/g, '' ) ;
	html = html.replace( /<([^\s>]+)[^>]*>\s*<\/\1>/g, '' ) ;
	var re = new RegExp("(<P)([^>]*>.*?)(<\/P>)","gi") ;
	html = html.replace( re, "<div$2</div>" ) ;
	InsertHTMLStr(html);
	EditArea.focus();
}
function URLPara(){
	var URL=window.location.href;
	var URLArray,Paras,ParasArray,StrReturn='';
	if(arguments[arguments.length-1]=="#")URLArray=URL.split("#");
	else URLArray=URL.split("?");
	if (URLArray.length==1)Paras='';
	else Paras=URLArray[1]; 
	if(Paras!=''){
		ParasArray=Paras.split("&");
		var ParasLength=ParasArray.length;
		var CheckStr=arguments[0].toLowerCase()+"=";
		for(i=0;i<ParasLength;i++){
			if(ParasArray[i].toLowerCase().indexOf(CheckStr)==0){
				StrReturn=ParasArray[i].toLowerCase().replace(CheckStr,"");
				break;
			}
		}
	}
	return StrReturn;
}