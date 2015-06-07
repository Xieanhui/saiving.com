define(['eventUtil', 'ajax', 'utilities','cookie', 'modal'],function(eventUtil, ajax, utilities, cookie, modal){

    var arrAgency = [ [ "\u5317\u4eac", "beijing" ], [ "\u6cb3\u5317", "beijing" ], [ "\u5929\u6d25", "beijing" ], [ "\u5185\u8499\u53e4", "ruimong" ], [ "\u65b0\u7586", "xinjiang" ], [ "\u7518\u8083", "ganshu" ], [ "\u9655\u897f", "xianxi" ], [ "\u9752\u6d77", "qinghai" ], [ "\u8fbd\u5b81", "dongshanxin" ], [ "\u5409\u6797", "dongshanxin" ], [ "\u9ed1\u9f99\u6c5f", "dongshanxin" ], [ "\u5b81\u590f", "ganshu" ], [ "\u6cb3\u5357", "henan" ], [ "\u5c71\u897f", "shanxi" ], [ "\u5c71\u4e1c", "shandong" ], [ "\u56db\u5ddd", "sichung" ], [ "\u91cd\u5e86", "sichung" ], [ "\u6c5f\u82cf", "jiangshu" ], [ "\u6e56\u5317", "hubei" ], [ "\u5b89\u5fbd", "anhui" ], [ "\u4e0a\u6d77", "shanghai" ], [ "\u6d59\u6c5f", "jiejiang" ], [ "\u6e56\u5357", "hunan" ], [ "\u798f\u5efa", "fujian" ], [ "\u6c5f\u897f", "jiangxi" ], [ "\u5e7f\u897f", "guangxi" ], [ "\u5e7f\u4e1c", "guangdong" ], [ "\u8d35\u5dde", "guizhou" ], [ "\u4e91\u5357", "yunnan" ], [ "\u6d77\u5357", "guangdong" ], [ "\u897f\u85cf", "qinghai" ], [ "\u53f0\u6e7e", "beijing" ], [ "\u9999\u6e2f", "beijing" ], [ "\u6fb3\u95e8", "beijing" ] ];

    function showAgency(uAgency, oAgency) {
        switch (uAgency) {
            case "\u5317\u4eac"://北京
            case "\u5317\u4EAC\u5E02"://北京市
                oAgency.innerHTML = "\u5317\u4eac\u603b\u90e8";
                break;

            case "\u6cb3\u5317"://河北
            case "\u6cb3\u5317\u7701"://河北省
            case "\u5929\u6d25"://天津
            case "\u5929\u6d25\u5e02"://天津市
            case "\u5185\u8499\u53e4"://内蒙古
            case "\u5185\u8499\u53e4\u7701"://内蒙古省
            case "\u65b0\u7586"://新疆
            case "\u65b0\u7586\u81ea\u6cbb\u533a"://新疆自治区
            case "\u7518\u8083"://甘肃
            case "\u7518\u8083\u7701"://甘肃省
            case "\u9655\u897f"://陕西
            case "\u9655\u897f\u7701"://陕西省
            case "\u9752\u6d77": //青海
            case "\u9752\u6d77\u7701": //青海省
            case "\u8fbd\u5b81": //辽宁
            case "\u8fbd\u5b81\u7701": //辽宁省
            case "\u5409\u6797": //吉林
            case "\u5409\u6797\u7701": //吉林省
            case "\u9ed1\u9f99\u6c5f"://黑龙江
            case "\u9ed1\u9f99\u6c5f\u7701": //黑龙江省
            case "\u5b81\u590f": //宁夏
            case "\u5b81\u590f\u81ea\u6cbb\u533a": //宁夏自治区
            case "\u6cb3\u5357"://河南
            case "\u6cb3\u5357\u7701": //河南省
            case "\u5c71\u897f": //山西
            case "\u5c71\u897f\u7701": //山西省
            case "\u56db\u5ddd"://四川
            case "\u56db\u5ddd\u7701"://四川省
            case "\u91cd\u5e86": //重庆
            case "\u91cd\u5e86\u5e02": //重庆市
            case "\u6c5f\u82cf"://江苏
            case "\u6c5f\u82cf\u7701"://江苏省
            case "\u6e56\u5317"://湖北
            case "\u6e56\u5317\u7701"://湖北省
            case "\u5b89\u5fbd"://安徽
            case "\u5b89\u5fbd\u7701"://安徽省
            case "\u4e0a\u6d77"://上海
            case "\u4e0a\u6d77\u5E02"://上海市
            case "\u6d59\u6c5f"://浙江
            case "\u6d59\u6c5f\u7701"://浙江省
            case "\u6e56\u5357"://湖南
            case "\u6e56\u5357\u7701"://湖南省
            case "\u798f\u5efa"://福建
            case "\u798f\u5efa\u7701"://福建省
            case "\u6c5f\u897f"://江西
            case "\u6c5f\u897f\u7701"://江西省
            case "\u5e7f\u897f"://广西
            case "\u5e7f\u897f\u7701"://广西省
            case "\u5e7f\u4e1c"://广东
            case "\u5e7f\u4e1c\u7701"://广东省
            case "\u8d35\u5dde"://贵州
            case "\u8d35\u5dde\u7701"://贵州省
            case "\u4e91\u5357"://云南
            case "\u4e91\u5357\u7701"://云南省
            case "\u6d77\u5357"://海南
            case "\u6d77\u5357\u7701"://海南省
            case "\u897f\u85cf"://西藏
            case "\u897f\u85cf\u81ea\u6cbb\u533a": //西藏自治区
                oAgency.innerHTML = uAgency + "\u529e\u4e8b\u5904";
                break;

            case "\u53f0\u6e7e": //台湾
            case "\u53f0\u6e7e\u7701": //台湾省
            case "\u9999\u6e2f": //香港
            case "\u9999\u6e2f\u7279\u522b\u884c\u653f\u533a": //香港特别行政区
            case "\u6fb3\u95e8": //澳门
            case "\u6fb3\u95e8\u7279\u522b\u884c\u653f\u533a": //澳门特别行政区
                oAgency.innerHTML = "\u6e2f\u6fb3\u53f0\u53ca\u6d77\u5916\u529e\u4e8b\u5904";
                break;

            default:
                oAgency.innerHTML = "\u5317\u4eac\u603b\u90e8";
                break;
        }
    }

    function setAnchorLink(uAgency, arrAgency, arrContact) {//设置链接锚和添加事件

        var i, j;
        for (i = 0; i < arrAgency.length; i++) {
            if (arrAgency[i][0] == uAgency) {
                for (j = 0; j < arrContact.length; j++) {
                    arrContact[j].setAttribute("href", "http://www.saiving.com/about/contact/index.html#" + arrAgency[i][1]);
                    eventUtil.addHandler(arrContact[j], "click", function (event) {
                        eventUtil.preventDefault();
                        var sUrl = eventUtil.target.href;
                        var hash = sUrl.split("#")[1] == "" ? "beijing" : sUrl.split("#")[1];
                        ajax.loadXMLDoc("http://www.saiving.com/contactHtml/" + hash + ".txt", modal.oModalContent);
                        modal.displayModalWindow('block');
                    });
                }
                break;
            }
        }
    }

    function AgencyAndAnchorLink() {
        var oAgency = document.getElementById("agency");
        var arrContact = document.getElementsByName("contact");
        var cookieAgency = cookie.getCookie("agency");
        if (cookieAgency) {
            showAgency(cookieAgency, oAgency);
            setAnchorLink(cookieAgency, arrAgency, arrContact);
        } else {
            try {
                utilities.proxyResponse("http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=js", function(){
                    if (remote_ip_info["province"] == null || remote_ip_info["province"] == "\u4fdd\u7559\u5730\u5740" || remote_ip_info["province"] == "") {
                        showAgency("\u5317\u4eac", oAgency);
                        setAnchorLink("\u5317\u4eac", arrAgency, arrContact);
                        cookie.setCookie("agency", "\u5317\u4eac", 10);
                    } else {
                        showAgency(remote_ip_info["province"], oAgency);
                        setAnchorLink(remote_ip_info["province"], arrAgency, arrContact);
                        cookie.setCookie("agency", remote_ip_info["province"], 10);
                    }
                });

            } catch (err) {
                    utilities.proxyResponse("http://pv.sohu.com/cityjson", function(){
                        showAgency(returnCitySN["cname"], oAgency);
                        setAnchorLink(returnCitySN["cname"], arrAgency, arrContact);
                        cookie.setCookie("agency",returnCitySN["cname"], 10);
                    });
            }
        }
    }

    return {
        AgencyAndAnchorLink: AgencyAndAnchorLink
    };

});
