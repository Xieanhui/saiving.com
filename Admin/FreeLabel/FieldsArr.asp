<%
'=======================================
'定义数据表中的字段名和对应的中文名数组
'查询字段
'排序字段
'查询条件字段
'=======================================
Dim GetStrFun
If G_IS_SQL_DB = 1 Then
	GetStrFun = "SubString"
Else
	GetStrFun = "Mid"
End IF

'=======================================
' 新闻表
'=======================================
Dim NSAllFCNArr,NSAllFENArr,NSAllFTypeArr,NSConArr,NSOrderArr
Dim NewAllENameFields,NewAllENameFieldsType
'---所有新闻表字段中文名
NSAllFCNArr = Array("编号","新闻编号","权重","所属栏目编号","所属专题","新闻标题","副标题","内容导读","标题显示评论","标题颜色","标题是否粗体","标题是否斜体","标题新闻","标题新闻地址","新闻内容","是否图片新闻","大图地址","小图地址","模板地址","浏览权限","来源","编辑","关键字","作者","点击数","保存路径","生成文件名","文件扩展名","图片头条","是否锁定","是否删除","添加时间","是否草稿","画中画广告","广告宽度","广告高度","广告连接","广告地址","推荐","滚动","允许评论","投稿","远程存图","文字头条","热点","精彩","不规则","公告","幻灯")
'---所有新闻表字段字段名
NewAllENameFields = "ID||NewsID||PopId||ClassID||SpecialEName||NewsTitle||CurtTitle||NewsNaviContent||isShowReview||TitleColor||titleBorder||TitleItalic||IsURL||URLAddress||Content||isPicNews||NewsPicFile||NewsSmallPicFile||Templet||isPop||Source||Editor||Keywords||Author||Hits||SaveNewsPath||FileName||FileExtName||TodayNewsPic||isLock||isRecyle||addtime||isdraft||IsAdPic||" & GetStrFun & "(AdPicWH,1,1)||" & GetStrFun & "(AdPicWH,3,1)||AdPicLink||AdPicAdress||" & GetStrFun & "(NewsProperty,1,1)||" & GetStrFun & "(NewsProperty,3,1)||" & GetStrFun & "(NewsProperty,5,1)||" & GetStrFun & "(NewsProperty,7,1)||" & GetStrFun & "(NewsProperty,9,1)||" & GetStrFun & "(NewsProperty,11,1)||" & GetStrFun & "(NewsProperty,13,1)||" & GetStrFun & "(NewsProperty,15,1)||" & GetStrFun & "(NewsProperty,17,1)||" & GetStrFun & "(NewsProperty,19,1)||" & GetStrFun & "(NewsProperty,21,1)"
NSAllFENArr = Split(NewAllENameFields,"||")
'---所有新闻表字段类型
NewAllENameFieldsType = "数字,文本型ID,数字,文本型ID,文本型ID,文本,文本,备注,判断型数字,文本,判断型数字,判断型数字,判断型数字,文本,备注,判断型数字,文本,文本,文本,判断型数字,文本,文本,文本,文本,数字,文本,文本,文本,判断型数字,判断型数字,判断型数字,日期时间型,判断型数字,判断型数字,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本"
NSAllFTypeArr = Split(NewAllENameFieldsType,",")
'---可以用来查询的新闻表字段
NSConArr = Array("编号","新闻编号","新闻标题","副标题","内容导读","标题颜色","标题新闻地址","新闻内容","大图地址","小图地址","模板地址","来源","编辑","关键字","作者","点击数","添加时间","广告连接","广告地址","所属专题")
'---可以用来排序的新闻表字段
NSOrderArr = Array("编号","权重","点击数","添加时间")

'=======================================
' 新闻栏目表
'=======================================
Dim NS_CAllFCNArr,NS_CAllENArr,NS_CAllTypeArr,NS_CConArr,NS_COrderArr
Dim NCAllENameFields,NCAllENameFieldsType
'---新闻栏目表所有字段中文名
NS_CAllFCNArr = Array("编号","栏目编号","栏目权重","栏目中文名","栏目英文名","是否外部栏目","外部栏目地址","父栏目编号","栏目模板","栏目新闻模板","二级域名","栏目管理员ID","浏览权限","文件扩展名","添加时间","允许投稿","归档时间","栏目导航显示","刷新信息条数","导航说明","导航图片","自定义字段分类ID","新闻默认审核","添加新闻模式","保存路径","保存格式","删除已审投稿","栏目关键字","栏目描述","是否删除","画中画广告","广告宽度","广告高度","广告连接","广告地址")
'---新闻栏目表所有字段名
NCAllENameFields = "ID||ClassID||OrderID||ClassName||ClassEName||IsURL||UrlAddress||ParentID||Templet||NewsTemplet||Domain||ClassAdmin||isPop||FileExtName||Addtime||isConstr||Oldtime||isShow||RefreshNumber||ClassNaviContent||ClassNaviPic||DefineID||NewsCheck||AddNewsType||SavePath||FileSaveType||isConstrDel||ClassKeywords||Classdescription||ReycleTF||IsAdPic||" & GetStrFun & "(AdPicWH,1,1)||" & GetStrFun & "(AdPicWH,3,1)||AdPicLink||AdPicAdress"
NS_CAllENArr = Split(NCAllENameFields,"||")
'---新闻栏目表所有字段类型
NCAllENameFieldsType = "数字,文本型ID,数字,文本,文本,判断型数字,文本,文本型ID,文本,文本,文本,文本型ID,判断型数字,文本,日期时间型,判断型数字,数字,判断型数字,数字,备注,文本,数字,判断型数字,判断型数字,文本,判断型数字,判断型数字,文本,文本,判断型数字,判断型数字,文本,文本,文本,文本"
NS_CAllTypeArr = Split(NCAllENameFieldsType,",")
'---查询的新闻栏目表字段
NS_CConArr = Array("编号","栏目编号","栏目中文名","栏目英文名","外部栏目地址","父栏目编号","栏目模板","栏目新闻模板","二级域名","添加时间","导航说明","导航图片","保存路径","栏目关键字","栏目描述","广告宽度","广告高度","广告连接","广告地址")
'---排序的新闻表字段
NS_COrderArr = Array("编号","栏目权重","添加时间")


'=======================================
'下载表
'=======================================
Dim DSAllFCNArr,DSAllFENArr,DSAllFTypeArr,DSConArr,DSOrderArr
Dim DownAllENameFields,DownAllENameFieldsType 
'---下载表中所有字段中文名
DSAllFCNArr = Array("编号","下载编号","所属栏目编号","简介","授权","添加时间","星级","是否审核","浏览权限","下载次数","修改时间","提供者EMAIL","文件扩展名","文件名","软件大小","语言版本","下载名称","下载模板","解压密码","图片地址","下载性质","开发商","提供者地址","推荐属性","允许评论","评论需要审核","系统平台","下载类型","软件版本","过期天数","消费点数","保存路径","点击次数","所属专题")
'---下载表中所有字段英文名
DownAllENameFields = "ID||DownLoadID||ClassID||Description||Accredit||AddTime||Appraise||AuditTF||BrowPop||ClickNum||EditTime||EMail||FileExtName||FileName||FileSize||Language||Name||NewsTemplet||PassWord||Pic||Property||Provider||ProviderUrl||RecTF||ReviewTF||ShowReviewTF||SystemType||Types||Version||OverDue||ConsumeNum||SavePath||Hits||SpeicalID"
DSAllFENArr = Split(DownAllENameFields,"||")
'---下载表所有字段类型
DownAllENameFieldsType = "数字,文本型ID,文本型ID,备注,判断型数字,日期时间型,判断型数字,判断型数字,文本,数字,日期时间型,文本,文本,文本,文本,文本,文本,文本,文本,文本,判断型数字,文本,文本,判断型数字,判断型数字,判断型数字,文本,判断型数字,文本,数字,数字,文本,数字,数字,文本"
DSAllFTypeArr = Split(DownAllENameFieldsType,",")
'---查询的下载表字段
DSConArr = Array("编号","下载编号","简介","授权","添加时间","星级","下载次数","修改时间","提供者EMAIL","软件大小","语言版本","下载名称","解压密码","图片地址","下载性质","开发商","提供者地址","系统平台","下载类型","软件版本","过期天数","消费点数","点击次数","推荐属性","所属专题")
'---排序的下载表字段
DSOrderArr = Array("编号","添加时间","下载次数","修改时间","消费点数","点击次数")

'=======================================
'下载栏目表
'=======================================
Dim DCAllFCNArr,DCAllFENArr,DCAllFTypeArr,DCConArr,DCOrderArr
Dim D_CAllENameFields,D_CAllENameFieldsType
'---下载栏目表所有字段中文名
DCAllFCNArr = Array("编号","栏目编号","栏目权重","栏目中文名","栏目英文名","是否外部栏目","外部栏目地址","父栏目编号","栏目模板","下载模板","二级域名","栏目管理员编号","浏览权限","文件扩展名","添加时间","允许投稿","导航显示","刷新数量","导读内容","导读图片","自定义字段编号","默认审核","保存路径","首页保存类型","投稿允许删除","栏目关键字","栏目描述","是否删除")
'---下载栏目表所有字段英文名
D_CAllENameFields = "ID||ClassID||OrderID||ClassName||ClassEName||IsURL||UrlAddress||ParentID||Templet||NewsTemplet||Domain||ClassAdmin||isPop||FileExtName||Addtime||isConstr||isShow||RefreshNumber||ClassNaviContent||ClassNaviPic||DefineID||NewsCheck||SavePath||FileSaveType||isConstrDel||ClassKeywords||Classdescription||ReycleTF"
DCAllFENArr = Split(D_CAllENameFields,"||")
'---下载栏目表所有字段类型
D_CAllENameFieldsType = "数字,文本型ID,数字,文本,文本,判断型数字,文本,文本型ID,文本,文本,文本,文本型ID,判断型数字,文本,日期时间型,判断型数字,判断型数字,数字,备注,文本,数字,判断型数字,文本,判断型数字,判断型数字,文本,文本,判断型数字"
DCAllFTypeArr = Split(D_CAllENameFieldsType,",")
'---下载栏目表查询字段
DCConArr = Array("编号","栏目编号","栏目中文名","栏目英文名","外部栏目地址","栏目模板","下载模板","二级域名","添加时间","导读内容","导读图片","栏目关键字","栏目描述")
'---下载栏目表排序字段
DCOrderArr = Array("编号","栏目权重","添加时间")


'=======================================
'商品表
'=======================================
Dim MSAllFCNArr,MSAllFENArr,MSAllFTypeArr,MSConArr,MSOrderArr
Dim MallAllENameFields,MallAllENameFieldsType
'---商品表所有字段中文名
MSAllFCNArr = Array("编号","商品名称","条形码","商品序列号","所属栏目编号","关键字","所属专题ID","库存","库存警告","原价","现价","允许批发","商品描述","保修条款","商品模板","制造商","产地","是否有发票","点击数","出厂日期","保存路径","文件名","文件扩展名","小图地址","大图地址","推荐","热点","特价","锁定","促销","幻灯","滚动","销售类型","添加日期","打折率","打折开始日期","打折结束日期","是否删除","添加者姓名","售出数量","商品权重","显示评论")
'---商品表所有字段英文名
MallAllENameFields = "ID||ProductTitle||Barcode||Serialnumber||ClassID||Keyword||SpecialID||Stockpile||StockpileWarn||OldPrice||NewPrice||" & GetStrFun & "(IsWholesale,0,1)||ProductContent||RepairContent||TempletFile||MakeFactory||ProductsAddress||IsInvoice||Click||MakeTime||SavePath||FileName||FileExtName||smallPic||BigPic||" & GetStrFun & "(StyleFlagBit,1,1)||" & GetStrFun & "(StyleFlagBit,3,1)||" & GetStrFun & "(StyleFlagBit,5,1)||" & GetStrFun & "(StyleFlagBit,7,1)||" & GetStrFun & "(StyleFlagBit,9,1)||" & GetStrFun & "(StyleFlagBit,11,1)||" & GetStrFun & "(StyleFlagBit,13,1)||SaleStyle||AddTime||Discount||DiscountStartDate||DiscountEndDate||ReycleTF||AddMember||saleNumber||popid||isShowReview"
MSAllFENArr = Split(MallAllENameFields,"||")
'---商品表所有字段类型
MallAllENameFieldsType = "数字,文本,文本,文本,文本型ID,文本,数字ID,文本,数字,货币,货币,判断型数字,备注,备注,文本,文本,文本,判断型数字,数字,日期时间型,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,文本,判断型数字,日期时间型,数字,日期时间型,日期时间型,判断型数字,文本,数字,判断型数字,判断型数字"
MSAllFTypeArr = Split(MallAllENameFieldsType,",")
'---商品表查询字段
MSConArr = Array("编号","商品名称","条形码","商品序列号","关键字","库存","库存警告","原价","现价","允许批发","商品描述","保修条款","制造商","产地","是否有发票","点击数","出厂日期","小图地址","大图地址","销售类型","添加日期","打折率","打折开始日期","打折结束日期","添加者姓名","售出数量","所属专题ID")
'---商品表排序字段
MSOrderArr = Array("编号","原价","现价","点击数","出厂日期","添加日期","打折率","打折开始日期","打折结束日期","售出数量","商品权重")


'=======================================
'商品栏目表
'=======================================
Dim MCAllFCNArr,MCAllFENArr,MCAllFTypeArr,MCConArr,MCOrderArr
Dim M_CAllENameFields,M_CAllENameFieldsType
'---商品栏目表所有字段中文名
MCAllFCNArr = Array("编号","栏目编号","父栏目编号","栏目权重","栏目英文名","栏目中文名","栏目模板","商品模板","是否外部栏目","外部栏目地址","添加日期","二级域名","浏览权限","是否继承父栏目设置","栏目管理员","导航显示","栏目导读内容","栏目导读图片","自定义字段ID","栏目关键字","栏目描述","是否删除","文件扩展名","保存路径","文件保存类型")
'---商品栏目表所有字段英文名
M_CAllENameFields = "ID||ClassID||ParentID||OrderID||ClassEName||ClassCName||ClassTemplet||ProductsTemplet||IsUrl||UrlAddress||Addtime||Domain||IsLimited||IsInherit||ClassAdmin||NaviShow||NaviContent||NaviPic||DefineID||Keywords||Description||ReycleTF||FileExtName||SavePath||FileSaveType"
MCAllFENArr = Split(M_CAllENameFields,"||")
'---商品栏目表所有字段类型
M_CAllENameFieldsType = "数字,文本型ID,文本型ID,数字,文本,文本,文本,文本,判断型数字,文本,日期时间型,文本,判断型数字,判断型数字,判断型数字,判断型数字,备注,文本,数字ID,备注,备注,判断型数字,文本,文本,判断型数字"
MCAllFTypeArr = Split(M_CAllENameFieldsType,",")
'---商品栏目表查询字段
MCConArr = Array("编号","栏目编号","栏目英文名","栏目中文名","栏目模板","商品模板","外部栏目地址","添加日期","二级域名","栏目导读内容","栏目导读图片","栏目关键字","栏目描述")
'---商品栏目表排序字段
MCOrderArr = Array("编号","栏目权重","添加日期")

'=======================================
' 返回字段名在数组中的位置
'=======================================
Function GetInnerFieldsNum(FieldCName,FieldCNameArray)
	Dim i,FiledName
	FiledName = FieldCName & ""
	For i = 0 to UBound(FieldCNameArray)
		if FieldCNameArray(i) & "" = FiledName Then
			GetInnerFieldsNum = i
			Exit For
			Exit Function
		End if
	Next
	GetInnerFieldsNum = GetInnerFieldsNum
End Function
%>







