(function () {
	'use strict';

	mainCtrl.$inject = ['$scope', 'adalAuthenticationService', 'o365Service'];

	function mainCtrl($scope, adal, o365) {
		var that = this;
		this.adal = adal;
		this.o365 = o365;
		this.$scope = $scope;
		$scope.log = '';
		$scope.searchResults = [];

		if (!adal.userInfo.isAuthenticated) {
			$scope.$on('adal:loginSuccess', function () {
				that.onInit();
			});
		} else {
			that.onInit();
		}

		$scope.search = function () {

			that.o365.search($scope.searchText)
			.then(function (data) {
				if (data.query && data.query.PrimaryQueryResult && data.query.PrimaryQueryResult.RelevantResults) {
					var results = data.query.PrimaryQueryResult.RelevantResults;

					if (results.RowCount < 1) {
						return;
					}

					$scope.searchResults = [];

					for (var i = 0; i < results.Table.Rows.results.length; i++) {
						var result = results.Table.Rows.results[i];
						var name = that.getPropertyValue('PreferredName', result.Cells.results);
						var path = that.getPropertyValue('Path', result.Cells.results);

						$scope.searchResults.push({name: name, path: path});
					}
				}
			}, function (err) {
				that.$scope.log += JSON.stringify(err);
			});
		}

		$scope.insertProfileUrl = function(user) {
			Word.run(function(ctx) {
				ctx.document.body.insertHtml('<a href="' + user.path + '">' + user.name + '</a>', 'end');

				return ctx.sync()
					.then(function() {
						
					})
					.catch(function(err) {
						that.$scope.log += JSON.stringify(err);
					});
			});
		}
	}

	mainCtrl.prototype.onInit = function () {
		var that = this;

		this.o365.getMyProperties()
		.then(function (data) {
			that.$scope.username = data.DisplayName;
		}, function (err) {
			that.$scope.log += JSON.stringify(err);
		});

		this.o365.getProfileImageUrl()
		.then(function (data) {
			that.$scope.profileImageUrl = data;
		}, function (err) {
			that.$scope.log += JSON.stringify(err);
		});
	}

	mainCtrl.prototype.getPropertyValue = function (propertyName, cells) {
		for (var i = 0; i < cells.length; i++) {
			if (cells[i].Key === propertyName) {
				return cells[i].Value;
			}
		}

		return null;
	}

	angular.module('profile.link.app')
		.controller('MainCtrl', mainCtrl);
})();