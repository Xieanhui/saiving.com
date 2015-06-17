define(function() {

    function g(id) {
        return document.getElementById(id);
    }

    function getElementsByClassName(oElm, strTagName, strClassName) {
        var arrElements = (strTagName == "*" && oElm.all) ? oElm.all : oElm.getElementsByTagName(strTagName);
        var arrReturnElements = new Array();
        strClassName = strClassName.replace(/\-/g, "\\-");
        var oRegExp = new RegExp("(^|\\s)" + strClassName + "(\\s|$)");
        var oElement;
        for (var i = 0; i < arrElements.length; i++) {
            oElement = arrElements[i];
            if (oRegExp.test(oElement.className)) {
                arrReturnElements.push(oElement);
            }
        }
        return (arrReturnElements);
    }

    var get = {
        byId: function(id) {
            return document.getElementById(id)
        },
        byClass: function(sClass, oParent) {
            if (oParent.getElementsByClass) {
                return (oParent || document).getElementsByClass(sClass)
            } else {
                var aClass = [];
                var reClass = new RegExp("(^| )" + sClass + "( |$)");
                var aElem = this.byTagName("*", oParent);
                for (var i = 0; i < aElem.length; i++) reClass.test(aElem[i].className) && aClass.push(aElem[i]);
                return aClass
            }
        },
        byTagName: function(elem, obj) {
            return (obj || document).getElementsByTagName(elem)
        }
    };

    function displayElts(arrElements, status) { //操作元素:显示或隐藏

        var _len = arrElements.length;
        if (_len) {
            for (var i = 0; i < arrElements.length; i++) {
                arrElements[i].style.display = status;
            }
        } else {
            arrElements.style.display = status;
        }

    }

    function proxyResponse(url, fn) { //按需载入js等类型的文档

        var head = document.getElementsByTagName("head")[0];
        var js = document.createElement("script");
        js.src = url;
        js.defer = true;
        js.onload = js.onreadystatechange = function() {
            if (!this.readyState || this.readyState == "loaded" || this.readyState == "complete") {
                fn();
                //head.removeChild(js);
                //JS加载完毕了. 类似于ajax请求完成.
                //执行是否登陆成功的判断
            }
        };
        head.appendChild(js);
    }

    return {
        g: g,
        getElementsByClassName: getElementsByClassName,
        displayElts: displayElts,
        proxyResponse: proxyResponse
    };

});
