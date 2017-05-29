app.controller('googleSheetController', function($scope, $http, $window, $location) {
	$scope.emailId='';
	$scope.emailDiv=true;

	$scope.responseStatus = '';
	$scope.spreadSheetId='';
	$scope.spreadSheetDiv=false;
	$scope.code='';
	$scope.sheetGgrid=false;
	$scope.sheets=null;
	$scope.findCustomer = function() {
		$http.get("api/findCustomer/"+$scope.email)
	    .then(function(response) {
	    	$scope.emailDiv=false;
	    	$scope.responseStatus=response.data.status;
	        if (response.data.status == 'NEED_AUTHENTICATION') {
	        	$window.location.href=response.data.redirectUril;
	        	//$location.path(response.data.redirectUril);
	        	//$scope.$apply();
	        } 
	    });
	}
	
	$scope.generateToken = function() {
		var dataObj = $scope.code;
		$http.post("api/generateToken", dataObj)
	    .then(function(response) {
	    	$scope.responseStatus=response.data.status;
	    	$scope.email=response.data.customerName;
	    });
	}
	
	$scope.getSpreadSheets = function() {
		$scope.sheets=null;
		$http.get("api/getSpreadSheets/"+$scope.email+"/"+$scope.spreadSheetId)
	    .then(function(response) {
	    	$scope.responseStatus=response.data.status;
	        if (response.data.sheets != null) {
	        	$scope.sheets=response.data.sheets;
	        	$scope.sheetGgrid=true;
	        } else if ( $scope.responseStatus == 'Session Expired' || $scope.responseStatus == 'Autherization Problem') {
	        	$scope.sheetGgrid=false;
	        	$scope.emailDiv=true;
	        }
	    });
	}
	$scope.onload = function() {
		var codeRes = $scope.getParam('code');
		if (codeRes != undefined && codeRes!='') {
			$scope.code=codeRes;
			$scope.generateToken();
			$scope.spreadSheetDiv=true;
			$scope.emailDiv=false;

		} else {
			
			$scope.emailDiv=true;
		}
	}
	
	$scope.getParam = function(param) {
        var sPageURL =  window.location.search.substring(1);
            var sURLVariables = sPageURL.split(/[&||?]/);
            var res='';

        for (var i = 0; i < sURLVariables.length; i += 1) {
            var paramName = sURLVariables[i],
                sParameterName = (paramName || '').split('=');

            if (sParameterName[0] === param) {
                res = sParameterName[1];
            }
        }

        return res;
}
	 
});