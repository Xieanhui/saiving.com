define(function() {

<<<<<<< HEAD
    function preventDefault(event) {
        var ev = event || window.event;
=======
    function preventDefault(ev) {
>>>>>>> dev
        if (ev && ev.preventDefault) {
            ev.preventDefault();
        } else {
            ev.returnValue = false;
        }
    }

<<<<<<< HEAD
    function stopPropagation(event) {
        var ev = event || window.event;
=======
    function stopPropagation(ev) {
>>>>>>> dev
        if (ev && ev.stopPropagation) {
            ev.stopPropagation();
        } else {
            ev.cancelBubble = true;
        }
    }

    function addHandler(args, sEvent, handler) { //给对象添加事件
        if (args.addEventListener) {
            args.addEventListener(sEvent, handler, false);
        } else if (args.attachEvent) {
            args.attachEvent("on" + sEvent, handler);
        } else {
            args['on' + sEvent] = handler;
        }
    }

    function removeHandler(args, sEvent, handler) { //移除对象事件
        if (args.removeEventListener) {
            args.removeEventListener(sEvent, handler, false);
        } else if (args.detachEvent) {
            args.detachEvent("on" + sEvent, handler);
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
        addScrollEvent: addScrollEvent,
        addResizeEvent: addResizeEvent
    };

});
