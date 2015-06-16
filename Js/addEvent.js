var addEvent = (function(){
        if( document.addEventListener ){
            return function(el, type, fn){
                var _len = el.length;
                if( _len ){
                    for(var i=0; i<_len; i++){
                        addEvent(el[i], type, fn);
                    }   
                }else{
                    el.addEventListener(type, fn, false);
                }   
            };  
        }else if( document.attachEvent ){
            return function(el, type, fn){
                var _len = el.length;
                if( _len ){
                    for(var i=0; i<_len; i++){
                        addEvent(el[i], type, fn);
                    }   
                }else{
                    el.attachEvent('on' + type, function(){
                        return fn.call(el, w.event);
                    }); 
                }   
            };  
        }   
    })();