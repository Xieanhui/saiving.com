define(['eventUtil','utilities'],function(eventUtil,utilities){

    var oModalWindow = document.getElementById("modalWindow"),
        oModalContent = document.getElementById("modalContent"),
        oModalCloser = document.getElementById("modalCloser");

    eventUtil.addHandler(oModalCloser, "click", function(event){
        utilities.displayElts(oModalWindow, "none");//隐藏模态框
    });

    function displayModalWindow(status) {//显示模态框 or 隐藏模态框
        utilities.displayElts(oModalWindow, status);
    }

    return {
        oModalContent : oModalContent,
        displayModalWindow: displayModalWindow
    };
});