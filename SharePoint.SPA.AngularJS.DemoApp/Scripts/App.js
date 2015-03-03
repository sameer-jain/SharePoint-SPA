
    'use strict';

    // create app
    var app = angular.module('employeeApp', ['ngRoute']);
    // 'ngRoute' for app route (url path) support

    // startup code
    app.run(['$route',  function ($route) {

    }]);

    //Routing
    app.config(['$routeProvider',
      function ($routeProvider) {
          $routeProvider
              .when('/', {
                  templateUrl: '../Views/List.html'
              })
              .when('/List', {
                  templateUrl: '../Views/List.html'
              })
              .when('/Add', {
                  templateUrl: '../Views/Add.html'
              })
              .when('/Edit/:Id', {
                  templateUrl: '../Views/Edit.html'
              }).otherwise({ redirectTo: '/' });
      }
    ]);


    //Add form controllers
    app.controller('AddFormCtrl', 
    function ($scope, $location, sharePointListServices) {
        $scope.addEmpClick = function (event) {
            if ($('form')[0].checkValidity() === false) {
                event.stopPropagation();
                return;
            }

            $scope.isDisabled = true;
            var emp = new employeeItem();
            emp.Title = $scope.EmployeeItem.Title;
            emp.LastName = $scope.EmployeeItem.LastName;
            emp.MobileNumber = $scope.EmployeeItem.MobileNumber;
            emp.JobTitle = $scope.EmployeeItem.JobTitle;
            emp.Email = $scope.EmployeeItem.Email;

            sharePointListServices.addListItem("EmployeeList", emp).then(function (data) {
                alert("Employee Info submitted successfully.");
                $location.path('/List');
            }).then(function () {
                $scope.EmployeeItem = null;
            });
            event.preventDefault();
        };

        $scope.cancelClick = function (event) {
            $location.path('/List');
            event.preventDefault();
        };
       

    });


    //edit form controllers
    app.controller('EditFormCtrl', 
    function ($scope, $location, $route, $routeParams, sharePointListServices) {
        var itemId = $routeParams.Id || 0;
        var query = "?$select=ID,Title,LastName,MobileNumber,JobTitle,Email";
        sharePointListServices.getListItemById("EmployeeList", itemId, query).then(function (data) {
            $scope.EmployeeItem = data;
        });

        $scope.updateClick = function (event) {
            $scope.isDisabled = true;
            var employeeItem = $scope.EmployeeItem;
            sharePointListServices.updateListItems("EmployeeList", employeeItem).then(function (data) {
                alert("Employee data Updated successfully.");
                $location.path('/List');
            });
            event.preventDefault();
        };       

        $scope.cancelClick = function (event) {
            $location.path('/List');
            event.preventDefault();
        };


    });

    //List form controllers
    app.controller('ListFormCtrl', 
    function ($scope, $location , sharePointListServices) {
        $scope.isItemExists = true;
        var query = "?$select=ID,Title,LastName,MobileNumber,JobTitle,Email";
        sharePointListServices.getListItems("EmployeeList", query).then(function (result) {
            $scope.EmployeeItems = result;
           
            if (angular.equals([], $scope.EmployeeItems)) {
                $scope.isItemExists = false;
            }
            else {
                $scope.isItemExists = true;
            }

        });

        $scope.deleteItem = function (item, event) {

            var wantToDelete = confirm('Are you sure you want to delete?');
            if (wantToDelete == true) {
                sharePointListServices.deleteListItem('EmployeeList', item.ID).then(function (data) {
                    alert("Item Deleted successfully.");                   
                }).then(function () {
                    var idx = $scope.EmployeeItems.indexOf(item);
                    if (idx > -1) {
                        $scope.EmployeeItems.splice(idx, 1);
                    }

                    if (angular.equals([], $scope.EmployeeItems)) {
                        $scope.isItemExists = false;
                    }
                    else {
                        $scope.isItemExists = true;
                    }
                });
            }
            event.preventDefault();
        };
    });



    app.service('sharePointListServices', ['$http', '$q', function ($http, $q) {

        //Get the ListItem
        this.getListItemById = function (listTitle, itemId, query) {
            var dfd = $q.defer();
            setHttpHeaders($http, actionType.Get);
            var restUrl = _spPageContextInfo.webServerRelativeUrl+ "/_api/web/lists/getbytitle('" + listTitle + "')/items(" + itemId + ")" + query;
            $http.get(restUrl).success(function (data) {
                dfd.resolve(data.d);
            }).error(function (data) {
                dfd.reject("error in getting items");
            });
            return dfd.promise;
        }

        //Get the ListItem
        this.getListItems = function (listTitle, query) {
            var dfd = $q.defer();
            setHttpHeaders($http, actionType.Get);
            var restUrl = _spPageContextInfo.webServerRelativeUrl+ "/_api/web/lists/getbytitle('" + listTitle + "')/items" + query;
            $http.get(restUrl).success(function (data) {
                dfd.resolve(data.d.results);
            }).error(function (data) {
                dfd.reject("error in getting items");
            });
            return dfd.promise;
        }

        //Create a ListItem
        this.addListItem = function (listTitle, listItem) {
            var dfd = $q.defer();
            setHttpHeaders($http, actionType.Add);
            var restUrl =_spPageContextInfo.webServerRelativeUrl+ "/_api/web/lists/getbytitle('" + listTitle + "')/items";
            $http.post(restUrl, listItem).success(function (data) {
                //resolve the new data
                dfd.resolve(data.d);
            }).error(function (data) {
                dfd.reject("failed to get items");
            });
            return dfd.promise;
        }

        //Update a ListItem
        this.updateListItems = function (listTitle, listItem) {
            var dfd = $q.defer();
            setHttpHeaders($http, actionType.Update);
            var restUrl =_spPageContextInfo.webServerRelativeUrl+  "/_api/web/lists/getbytitle('" + listTitle + "')/items(" + listItem.ID + ")";
            $http.post(restUrl, listItem).success(function (data) {
                //resolve something
                dfd.resolve(true);
            }).error(function (data) {
                dfd.reject("error updating item");
            });
            return dfd.promise;
        }

        //Delete a list item
        this.deleteListItem = function (listTitle, itemId) {
            var dfd = $q.defer();
            setHttpHeaders($http, actionType.Delete);
            var restUrl = _spPageContextInfo.webServerRelativeUrl+ "/_api/web/lists/getbytitle('" + listTitle + "')/items(" + itemId + ")";
            $http.post(restUrl)
                .success(function (data) {
                    //resolve something
                    dfd.resolve(true);
                }).error(function (data) {
                    dfd.reject("error deleting item");
                });
            return dfd.promise;
        }

    }]);

    var actionType = {
        Get: 0,
        Add: 1,
        Update: 2,
        Delete: 4
    };

    function setHttpHeaders($http, action) {
        //Define the http headers actionType
        $http.defaults.headers.common.Accept = "application/json;odata=verbose";
        $http.defaults.headers.post['Content-Type'] = 'application/json;odata=verbose';
        $http.defaults.headers.post['X-Requested-With'] = 'XMLHttpRequest';
        $http.defaults.headers.post['X-RequestDigest'] = angular.element(document.querySelector('#__REQUESTDIGEST')).val();
        $http.defaults.headers.post['If-Match'] = "*";
        var httpMethod;
        switch (action) {
            case actionType.Update:
                httpMethod = "MERGE";
                break;
            case actionType.Delete:
                httpMethod = "DELETE";
                break;
            default:
                httpMethod = "";
                break;
        }
        $http.defaults.headers.post['X-HTTP-Method'] = httpMethod;
    }

    var employeeItem = function () {
        this.ID = undefined;
        this.Title = undefined;
        this.LastName = undefined;
        this.MobileNumber = undefined;
        this.JobTitle = undefined;
        this.Email = undefined;
        this.__metadata = {
            type: 'SP.Data.EmployeeListListItem'
        };
    };