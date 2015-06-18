define(['utilities', 'eventUtil'], function(utilities, eventUtil) {

    var olive = utilities.g("live"),
        step = 0,
        timer = null,
        current = 0,
        oLvCt = utilities.g("liveCtrl");
    var oLiveCtrl = (oLvCt != null) ? oLvCt.getElementsByTagName('li') : null;

    if (olive != null) {
        oLiveCtrl[0].style.backgroundColor = "#ffffff";

        for (var i = 0; i < oLiveCtrl.length; i++) {
            oLiveCtrl[i].index = i;
        }

        eventUtil.addHandler(oLvCt, 'click', function(event) {
            var ev = event || window.event;
            var target = ev.target || ev.srcElement;
            if (timer) {
                clearInterval(timer);
                clearTimeout(timer);
            }
            for (var j = 0; j < oLiveCtrl.length; j++) {
                oLiveCtrl[j].style.backgroundColor = "#555555";
            }
            if (!isNaN(target.index)) {
                step = -(target.index * 2000);
            }
            olive.style.left = step + "px";
            oLiveCtrl[target.index].style.backgroundColor = "#ffffff";
            timer = setTimeout(slide, 3000);
        });
    }


    function slide() {
        if (olive != null) {
            clearTimeout(timer);
            timer = setInterval(function() {
                if (olive.offsetLeft <= -4000) step = 0;
                if (step % 2000 == 0) {
                    current = Math.abs(Math.floor(step / 2000));
                    if (current == oLiveCtrl.length) current = 0;
                    for (var j = 0; j < oLiveCtrl.length; j++) {
                        oLiveCtrl[j].style.backgroundColor = "#555555";
                    }
                    oLiveCtrl[current].style.backgroundColor = "#ffffff";
                    if (timer) clearInterval(timer);
                    timer = setTimeout(slide, 3000);
                }
                step -= 100;
                olive.style.left = step + "px";

            }, 10);
        }
    }

    return {
        slide: slide
    }
});
