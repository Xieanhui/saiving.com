window.ajax = (function() {

		var func = function (url, suc, fls) {
		
			var httpRequest;
			
			try	{
			
				httpRequest = new XMLHttpRequest();
				
			} catch(ex) {
			
				httpRequest = new ActiveXObject('microsoft.xmlhttp');
				
			}
			
			if (httpRequest !== null) {
			
				httpRequest.onreadystatechange = suc(httpRequest);
				httpRequest.open('POST', url, true);
				httpRequest.send(null);
				
			} else {
			
				fls();
				
			}
			
		};
		
		return func;
		
	}());