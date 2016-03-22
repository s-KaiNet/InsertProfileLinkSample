(function () {
	'use strict';

	angular.module('profile.link.app', ['ngRoute', 'AdalAngular', 'officeuifabric.core', 'officeuifabric.components', 'angular-loading-bar'])
	.config(config);

	config.$inject = [
		'$routeProvider',
		'$httpProvider',
		'adalAuthenticationServiceProvider',
		'cfpLoadingBarProvider'];

	function config($routeProvider, $httpProvider, adalProvider, cfpLoadingBarProvider) {
		$routeProvider
			.when('/', {
				controller: 'MainCtrl',
				templateUrl: 'main.html',
				requireADLogin: true
			})
			.otherwise({
				redirectTo: '/'
			});

		adalProvider.init(
			{
				instance: 'https://login.microsoftonline.com/',
				clientId: '58f46aab-fc72-42f7-8d76-7eb7dadc61f0',
				extraQueryParameter: 'nux=1',
				endpoints: {
					'https://api.office.com/discovery/v1.0/me/': 'https://api.office.com/discovery/',
					'https://graph.microsoft.com': 'https://graph.microsoft.com'
				}
			},
			$httpProvider
		);

		cfpLoadingBarProvider.includeSpinner = false;
	}
})();