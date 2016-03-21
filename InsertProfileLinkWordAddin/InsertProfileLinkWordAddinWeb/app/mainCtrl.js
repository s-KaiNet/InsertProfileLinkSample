(function () {
	'use strict';

	mainCtrl.$inject = ['$scope', 'adalAuthenticationService', 'o365Service'];

	function mainCtrl($scope, adal, o365) {
		var that = this;
		this.adal = adal;
		this.o365 = o365;
		this.$scope = $scope;
		$scope.log = '';

		if (!adal.userInfo.isAuthenticated) {
			$scope.$on('adal:loginSuccess', function () {
				that.onInit();
			});
		} else {
			that.onInit();
		}

		$scope.search = function() {
			$scope.log += $scope.searchText;
		}
	}

	mainCtrl.prototype.onInit = function () {
		var that = this;
		this.o365.getMyProperties()
		.then(function (data) {
			that.$scope.username = data.DisplayName;
			that.$scope.profileImageUrl = data.PictureUrl;
		}, function(err) {
				that.$scope.log += JSON.stringify(err);
			});

	}

	angular.module('profile.link.app')
		.controller('MainCtrl', mainCtrl);
})();