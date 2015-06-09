define(['eventUtil', 'ajax', 'modal'], function(eventUtil, ajax, modal) {

    var arrAgency = [
        ["\u5317\u4eac", "beijing"], //北京
        ["\u5317\u4EAC\u5E02", "beijing"],
        ["\u6cb3\u5317", "beijing"], //河北
        ["\u6cb3\u5317\u7701", "beijing"],
        ["\u5929\u6d25", "beijing"], //天津
        ["\u5929\u6d25\u5e02", "beijing"],
        ["\u5185\u8499\u53e4", "ruimong"], //内蒙古
        ["\u5185\u8499\u53e4\u81ea\u6cbb\u533a", "ruimong"],
        ["\u65b0\u7586", "xinjiang"], //新疆
        ["\u65b0\u7586\u81ea\u6cbb\u533a", 'xinjiang'],
        ["\u7518\u8083", "ganshu"], //甘肃
        ["\u7518\u8083\u7701", "ganshu"],
        ["\u9655\u897f", "xianxi"], //陕西
        ["\u9655\u897f\u7701", "xianxi"],
        ["\u9752\u6d77", "qinghai"], //青海
        ["\u9752\u6d77\u7701", "qinghai"],
        ["\u8fbd\u5b81", "dongshanxin"], //辽宁
        ["\u8fbd\u5b81\u7701", "dongshanxin"],
        ["\u5409\u6797", "dongshanxin"], //吉林
        ["\u5409\u6797\u7701", "dongshanxin"],
        ["\u9ed1\u9f99\u6c5f", "dongshanxin"], //黑龙江
        ["\u9ed1\u9f99\u6c5f\u7701", "dongshanxin"],
        ["\u5b81\u590f", "ganshu"], //宁夏
        ["\u5b81\u590f\u81ea\u6cbb\u533a", "ganshu"],
        ["\u6cb3\u5357", "henan"], //河南
        ["\u6cb3\u5357\u7701", "henan"],
        ["\u5c71\u4e1c", "shandong"], //山东
        ["\u5c71\u4e1c\u7701", "shandong"],
        ["\u5c71\u897f", "shanxi"], //山西
        ["\u5c71\u897f\u7701", "shanxi"],
        ["\u56db\u5ddd", "sichung"], //四川
        ["\u56db\u5ddd\u7701", "sichung"],
        ["\u91cd\u5e86", "sichung"], //重庆
        ["\u91cd\u5e86\u5e02", "sichung"],
        ["\u6c5f\u82cf", "jiangshu"], //江苏
        ["\u6c5f\u82cf\u7701", "jiangshu"],
        ["\u6e56\u5317", "hubei"], //湖北
        ["\u6e56\u5317\u7701", "hubei"]
        ["\u5b89\u5fbd", "anhui"], //安徽
        ["\u5b89\u5fbd\u7701", "anhui"],
        ["\u4e0a\u6d77", "shanghai"], //上海
        ["\u4e0a\u6d77\u5E02", "shanghai"]
        ["\u6d59\u6c5f", "jiejiang"], //浙江
        ["\u6d59\u6c5f\u7701", "jiejiang"]
        ["\u6e56\u5357", "hunan"], //湖南
        ["\u6e56\u5357\u7701", "hunan"],
        ["\u798f\u5efa", "fujian"], //福建
        ["\u798f\u5efa\u7701", "fujian"],
        ["\u6c5f\u897f", "jiangxi"], //江西
        ["\u6c5f\u897f\u7701", "jiangxi"],
        ["\u5e7f\u897f", "guangxi"], //广西
        ["\u5e7f\u897f\u7701", "guangxi"],
        ["\u5e7f\u4e1c", "guangdong"], //广东
        ["\u5e7f\u4e1c\u7701", "guangdong"],
        ["\u8d35\u5dde", "guizhou"], //贵州
        ["\u8d35\u5dde\u7701", "guizhou"],
        ["\u4e91\u5357", "yunnan"], //云南
        ["\u4e91\u5357\u7701", "yunnan"],
        ["\u6d77\u5357", "guangdong"], //海南
        ["\u6d77\u5357\u7701", "guangdong"],
        ["\u897f\u85cf", "qinghai"], //西藏
        ["\u897f\u85cf\u81ea\u6cbb\u533a", "qinghai"],
        ["\u53f0\u6e7e", "beijing"], //台湾
        ["\u53f0\u6e7e\u7701", "beijing"],
        ["\u9999\u6e2f", "beijing"], //香港
        ["\u9999\u6e2f\u7279\u522b\u884c\u653f\u533a", "beijing"],
        ["\u6fb3\u95e8", "beijing"], //澳门
        ["\u6fb3\u95e8\u7279\u522b\u884c\u653f\u533a", "beijing"]
    ];

    function getHash(agencyStr) {
        for (var i = 0; i < arrAgency.length; i++) {
            if (arrAgency[i][0] == agencyStr) {
                return arrAgency[i][1];
            }
        }
    }

    function setAnchor(agencyStr) {
        var hash = getHash(agencyStr) || "beijing";
        var arrContact = document.getElementsByName("contact");
        var i;

        for (i = 0; i < arrContact.length; i++) {
            arrContact[i].setAttribute("href", "http://www.saiving.com/about/contact/index.html#" + hash);
        }

        eventUtil.addHandler(document, "click", function(event) {
            var ev = event || window.event;
            var tg = ev.target || ev.srcElement;
            var url = tg.href || "";
            url = url.toString();
            if (url && url.substring(0, url.indexOf("#")) == "http://www.saiving.com/about/contact/index.html") {
                eventUtil.preventDefault(ev);
                ajax.loadXMLDoc("http://www.saiving.com/contactHtml/" + hash + ".txt", modal.oModalContent);
                modal.displayModalWindow('block');
            }
        });
    }

    return {
        setAnchor: setAnchor
    }
});
