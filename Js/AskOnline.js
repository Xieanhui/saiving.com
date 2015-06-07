define(['eventUtil'], function(eventUtil){

    function askOnline(){

        var oAskOnline = document.getElementsByName("askOnline") || null;
        if (oAskOnline != null) {
            var sUrl = 'http://qiao.baidu.com/v3/?module=default&controller=im&action=index&ucid=7225288&type=n&siteid=4628760';
            var sName = '\u5728\u7ebf\u54a8\u8be2';
            var iWidth = 800;
            var iHeight = 550;
            var iTop = (window.screen.height - iHeight) / 2;
            var iLeft = (window.screen.width - iWidth) / 2;
            var sParam = 'height=' + iHeight + ',width=' + iWidth + ',top=' + iTop + ',left=' + iLeft + ',toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no';
            for (var i = 0; i < oAskOnline.length; i++){
                eventUtil.addHandler(oAskOnline.item(i),"click", function(event){
                    eventUtil.preventDefault();
                    window.open(sUrl, sName, sParam);
                });
            }
        }
    }

    return {
        askOnline:askOnline
    };
});



