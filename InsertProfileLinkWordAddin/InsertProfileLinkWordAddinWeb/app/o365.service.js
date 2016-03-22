(function () {
	'use strict';

	o365Service.$inject = ['adalAuthenticationService', '$http', '$q'];

	function o365Service(adal, $http, $q) {
		this.$q = $q;
		this.$http = $http;
		this.adal = adal;
	}

	o365Service.prototype.getProfileImageUrl = function () {

		var deferred = this.$q.defer();
		this.$http.get('https://graph.microsoft.com/v1.0/me/photo/$value', {
			responseType: 'arraybuffer'
		})
		.then(function (result) {
			var data = new Uint8Array(result.data);
			var raw = String.fromCharCode.apply(null, data);
			var base64 = btoa(raw);

			deferred.resolve('data:image//jpeg;base64,' + base64);
		},
		function (err) {
			if (err.status === 404) {
				deferred.resolve(null);
			} else {
				deferred.reject(err);
			}
		});

		return deferred.promise;
	}

	o365Service.prototype.getDiscoveryData = function () {
		var deferred = this.$q.defer();
		var that = this;

		this.adal.acquireToken('https://api.office.com/discovery/')
			.then(function (token) {
				return that.$http.post('api/discovery/services', JSON.stringify({ token: token }), {
					headers: {
						'Content-Type': 'application/json'
					}
				});
			})
		.then(function (result) {
			var data = {};
			data.spResourceId = result.data.rootSite.serviceResourceId;
			data.spResourceUrl = result.data.rootSite.serviceEndpointUri;

			deferred.resolve(data);
		},
		function (err) {
			deferred.reject(err);
		});

		return deferred.promise;
	}

	o365Service.prototype.search = function (terms) {
		var deferred = this.$q.defer();
		var that = this;
		this.getDiscoveryData()
			.then(function (result) {
				that.adal.config.endpoints[result.spResourceUrl] = result.spResourceId;

				return that.$http.get(result.spResourceUrl + '/search/query?querytext=\'' + terms +
					'\'&sourceid=\'B09A7990-05EA-4AF9-81EF-EDFAB16C4E31\'&selectproperties=\'Path,PreferredName\'', {
					headers: { 'accept': 'application/json;odata=verbose' }
				});
			}, function (err) {
				deferred.reject(err);
			})
		.then(function (response) {
			deferred.resolve(response.data.d);
		},
		function (err) {
			deferred.reject(err);
		});

		return deferred.promise;
	}

	o365Service.prototype.getMyProperties = function () {
		var deferred = this.$q.defer();
		var that = this;
		this.getDiscoveryData()
			.then(function (result) {
				that.adal.config.endpoints[result.spResourceUrl] = result.spResourceId;

				return that.$http.get(result.spResourceUrl + '/sp.userprofiles.peoplemanager/getmyproperties', {
					headers: { 'accept': 'application/json;odata=verbose' }
				});
			}, function (err) {
				deferred.reject(err);
			})
		.then(function (response) {
			deferred.resolve(response.data.d);
		},
		function (err) {
			deferred.reject(err);
		});

		return deferred.promise;
	}

	angular.module('profile.link.app')
		.service('o365Service', o365Service);
})();