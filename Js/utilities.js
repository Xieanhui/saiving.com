define(function(){

    function displayElts(arrElements, status) {//操作元素:显示或隐藏
        try {
            if (arrElements instanceof Array) {
                if (arrElements.length > 0) {
                    for (var i = 0; i < arrElements.length; i++) {
                        arrElements[i].style.display = status;
                    }
                }
            }
            else if (arrElements != null) {
                arrElements.style.display = status;
            }
        } catch (error) {
            console("argument must Input a document element:" + error.name + ":" + error.message);
        }
    }

    function proxyResponse(url,fn){//按需载入js等类型的文档

        var head = document.getElementsByTagName("head")[0];
        var js = document.createElement("script");
        js.src = url;
        js.onload = js.onreadystatechange = function()
        {
            if (!this.readyState || this.readyState == "loaded" || this.readyState == "complete")
            {
                fn();
                head.removeChild(js);
                //JS加载完毕了. 类似于ajax请求完成.
                //执行是否登陆成功的判断
            }
        };
        head.appendChild(js);
    }

    return {

        displayElts : displayElts,
        proxyResponse : proxyResponse

    };

});
