define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    function askOnline() {
        var oprodBorad = utilities.g('prodBorad');

        var sUrl = 'http://qiao.baidu.com/v3/?module=default&controller=im&action=index&ucid=7225288&type=n&siteid=4628760',
            sName = '\u5728\u7ebf\u54a8\u8be2',
            iWidth = 800,
            iHeight = 550,
            iTop = (window.screen.height - iHeight) / 2,
            iLeft = (window.screen.width - iWidth) / 2,
            sParam = 'height=' + iHeight + ',width=' + iWidth + ',top=' + iTop + ',left=' + iLeft + ',toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no';

        eventUtil.addHandler(oprodBorad, 'click', function(event) {
            var ev = event || window.event;
            var target = ev.target || ev.srcElement;
            var regEx = /(^|\s)btn(\s|$)/g;
            if (target.nodeType == 1) {
                if (regEx.test(target.className)) {
                    eventUtil.preventDefault(event);
                    window.open(sUrl, sName, sParam);
                }
            }

        });
    }

    return {
        askOnline: askOnline
    }

});
