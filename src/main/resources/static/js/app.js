var app = angular.module('myApp', ['ngRoute','ngResource']);
app.config(function($routeProvider){
    $routeProvider
        .when('/login',{
            templateUrl: '/login.html',
            controller: 'housieController'
        })
        .when('/register',{
            templateUrl: '/register.html',
            controller: 'housieController'
        })
        .otherwise(
            { redirectTo: '/'}
        );
});