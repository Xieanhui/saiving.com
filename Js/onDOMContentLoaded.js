function onDOMContentLoaded(onready, config) {
    //浏览器检测相关对象，在此为节省代码未实现，实际使用时需要实现。   
    //var Browser = {};   
    //设置是否在FF下使用DOMContentLoaded（在FF2下的特定场景有Bug）   
    this.conf = {
        enableMozDOMReady: true
    };
    if (config) {
        for (var p in config)
            this.conf[p] = config[p];
    }

    var isReady = false;

    function doReady() {
            if (isReady) return;
            //确保onready只执行一次   
            isReady = true;
            onready();
        }
        /*IE*/
    if (Browser.ie) {
        (function() {
            if (isReady) return;
            try {
                document.documentElement.doScroll("left");
            } catch (error) {
                setTimeout(arguments.callee, 0);
                return;
            }
            doReady();
        })();
        window.attachEvent('onload', doReady);
    }
    /*Webkit*/
    else if (Browser.webkit && Browser.version < 525) {
        (function() {
            if (isReady) return;
            if (/loaded|complete/.test(document.readyState))
                doReady();
            else
                setTimeout(arguments.callee, 0);
        })();
        window.addEventListener('load', doReady, false);
    }
    /*FF Opera 高版webkit 其他*/
    else {
        if (!Brower.gecko || Browser.version != 2 || this.conf.enableMozDOMReady)
            document.addEventListener("DOMContentLoaded", function() {
                document.removeEventListener("DOMContentLoaded", arguments.callee, false);
                doReady();
            }, false);
        window.addEventListener('load', doReady, false);
    }

}
