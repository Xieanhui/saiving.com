define(['eventUtil'], function(eventUtil) {
    var olivePrev = document.getElementById("livePrev");
    var oliveNext = document.getElementById("liveNext");
    var olive = document.getElementById("live");
    var step = 10;

    eventUtil.addHandler(olivePrev, "click", function() {


    });


    var timer = setInterval(function() {
    	if(olive.offsetLeft == -4000) clearInterval(timer);
    	step -= 10;
    	alert(step);
        olive.style.left = step + "px";

    }, 100);

})
