define(['eventUtil'],function (eventUtil) {

    var oBackToTop = document.getElementById("backToTop");
    var scrollTop, clientHeight;
    var step = 0;
    var timer = null;
    var isTop = true;

   function stopScroll(){
       eventUtil.addScrollEvent(function(){
           clientHeight = document.documentElement.clientHeight || document.body.clientHeight;
           scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
           (scrollTop >= clientHeight)? (oBackToTop.style.display = "block"):(oBackToTop.style.display = "none");

           if(!isTop){
               clearInterval(timer);
           }
           isTop = false;
       });
    }

    function backToTop(){
        eventUtil.addHandler(oBackToTop, "click", function () {
            clearInterval(timer);
            timer = setInterval(function(){
                scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
                step = Math.ceil(scrollTop / 5);
                document.body.scrollTop -= step;
                document.documentElement.scrollTop -=  step;
                isTop = true;
                if(scrollTop <= 0){
                    clearInterval(timer);
                    document.body.scrollTop = 0;
                    document.documentElement.scrollTop =  0;
                }
            }, 2);
        });
    }

    return {
        stopScroll:stopScroll,
        backToTop:backToTop
    };

});