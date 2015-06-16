define(['eventUtil', 'Browser'], function(eventUtil, Browser) {

    function preventDefault(ev) {
        if (ev && ev.preventDefault) {
            ev.preventDefault();
        } else {
            ev.returnValue = false;
        }
    }

    function stopPropagation(ev) {

        if (ev && ev.stopPropagation) {
            ev.stopPropagation();
        } else {
            ev.cancelBubble = true;
        }
    }

    function addHandler(el, type, fn) { //给对象添加事件
        // if (element.addEventListener) {
        //     element.addEventListener(type, handler, false);
        // } else if (element.attachEvent) {
        //     element.attachEvent("on" + type, handler);
        // } else {
        //     element['on' + type] = handler;
        // }
        if (document.addEventListener) {
            var _len = el.length;
            if (_len) {
                for (var i = 0; i < _len; i++) {
                    arguments.callee(el[i], type, fn);
                }
            } else {
                el.addEventListener(type, fn, false);
            }
        } else if (document.attachEvent) {

            var _len = el.length;
            if (_len) {
                for (var i = 0; i < _len; i++) {
                    arguments.callee(el[i], type, fn);
                }
            } else {
                el.attachEvent('on' + type, function() {
                    return fn.call(el, window.event);
                });
            }
        }
    }

    function removeHandler(element, type, handler) { //移除对象事件
        // if (element.removeEventListener) {
        //     element.removeEventListener(type, handler, false);
        // } else if (element.detachEvent) {
        //     element.detachEvent("on" + type, handler);
        // }

        if (document.removeEventListener) {
            var _len = el.length;
            if (_len) {
                for (var i = 0; i < _len; i++) {
                    arguments.callee(el[i], type, fn);
                }
            } else {
                el.removeEventListener(type, fn, false);
            }
        } else if (document.detachEvent) {
            var _len = el.length;
            if (_len) {
                for (var i = 0; i < _len; i++) {
                    arguments.callee(el[i], type, fn);
                }
            } else {
                el.detachEvent('on' + type, function() {
                    return fn.call(el, window.event);
                });
            }
        }

    }

    function addLoadEvent(func) { //添加文档ONLOAD事件处理函数
        var oldLoad = window.onload;
        if (typeof window.onload != "function") {
            window.onload = func;
        } else {
            window.onload = function() {
                oldLoad();
                func();
            }
        }
    }

    // var timer;

    // function fireContentLoadedEvent() {
    //     if (document.loaded) return;
    //     if (timer) window.clearInterval(timer);
    //     document.fire("dom:loaded");
    //     document.loaded = true;
    // }

    // if (document.addEventListener) {
    //     if (Browser.webkit) {
    //         timer = window.setInterval(function() {
    //             if (/loaded|complete/.test(document.readyState))
    //                 fireContentLoadedEvent();
    //         }, 0);

    //         eventUtil.addHandler(window, "load", fireContentLoadedEvent);

    //     } else {
    //         document.addEventListener("DOMContentLoaded",
    //             fireContentLoadedEvent, false);
    //     }

    // } else {
    //     document.write("<" + "script id=__onDOMContentLoaded defer src=//:><\/script>");
    //     g("__onDOMContentLoaded").onreadystatechange = function() {
    //         if (this.readyState == "complete") {
    //             this.onreadystatechange = null;
    //             fireContentLoadedEvent();
    //         }
    //     };
    // }

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
            if (!Browser.gecko || Browser.version != 2 || this.conf.enableMozDOMReady)
                document.addEventListener("DOMContentLoaded", function() {
                    document.removeEventListener("DOMContentLoaded", arguments.callee, false);
                    doReady();
                }, false);
            window.addEventListener('load', doReady, false);
        }
    }

    function addScrollEvent(func) { //添加滚动条事件处理函数
        var oldScroll = window.onscroll;
        if (typeof window.onscroll != "function") {
            window.onscroll = func;
        } else {
            window.onscroll = function() {
                oldScroll();
                func();
            }
        }
    }

    function addResizeEvent(func) { //添加滚动条事件处理函数
        var oldResize = window.onresize;
        if (typeof window.onresize != "function") {
            window.onresize = func;
        } else {
            window.onresize = function() {
                oldResize();
                func();
            }
        }
    }

    return {
        preventDefault: preventDefault,
        stopPropagation: stopPropagation,
        addHandler: addHandler,
        removeHandler: removeHandler,
        addLoadEvent: addLoadEvent,
        onDOMContentLoaded: onDOMContentLoaded,
        addScrollEvent: addScrollEvent,
        addResizeEvent: addResizeEvent
    };

});
