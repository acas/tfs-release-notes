var releaseNotes = angular.module('releaseNotes', ['ui.bootstrap'])

releaseNotes.controller('releaseNotes-controller', ['$scope', '$http', '$q', '$timeout', '$modal', function ($scope, $http, $q, $timeout, $modal) {
    $scope.releaseNotes = new function () {
        var utilities = {
            loadGenerators: function() {
                api.generators = [
                    { name: "excel" },
                    { name: "html"  }
                ]
            },
            loadConfigurations: function () {
                var deferred = $q.defer();
                $http({ method: 'GET', url: 'api/Configuration/Load' })
                .success(function (data) {
                    api.configurations = data;
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            saveConfiguration: function (fields) {
                var deferred = $q.defer();
                $http({ method: 'POST', url: 'api/Configuration/Save', data: fields })
                .success(function (data) {
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            deleteConfiguration: function (fields) {
                var deferred = $q.defer();
                $http({ method: 'POST', url: 'api/Configuration/Delete', data: fields })
                .success(function (data) {
                    utilities.loadConfigurations()
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            load: function() {
                utilities.loadGenerators()
                utilities.loadConfigurations();
            },
            initialize: function () {
                utilities.load()
            }
        }

        var api = {
            configurations: [],
            selectedConfiguration: null,
            trashyConfiguration: null,
            newConfiguration: null,
            generators: [],
            fields: {
                teamProjectPath: null,
                projectName: null,
                projectSubpath: null,
                iteration: null,
                database: null,
                databaseServer: null,
                webServer: null,
                webLocation: null,
                generator: null,
            },
            saveConfiguration: function () {
                if (api.selectedConfiguration) {
                    api.fields['configurationName'] = api.selectedConfiguration.configurationName
                }
                if (api.newConfiguration) {
                    api.fields['configurationName'] = api.newConfiguration
                }
                utilities.saveConfiguration(api.fields).then(function () {
                    utilities.loadConfigurations().then(function () {
                        var selectedConfiguration = _.find(api.configurations, function (configuration) {
                            return api.fields.configurationName === configuration.configurationName
                        })
                        if (selectedConfiguration) api.selectedConfiguration = api.trashyConfiguration = selectedConfiguration
                        api.newConfiguration = null
                    })
                })
                //delete api.fields['configurationName']
            },
            deleteConfiguration: function () {
                if (api.trashyConfiguration) {
                    utilities.deleteConfiguration(api.fields)
                }
            },
            configurationChanged: function() {
                if (api.selectedConfiguration) {
                    api.fields = api.selectedConfiguration
                }
            },
            generate: function () {
                var modalInstance = $modal.open({
                    templateUrl: 'Home/Waiting',
                    controller: 'releaseNotes-progressModal-controller',
                    size: 'md'
                })
                $http({ method: 'POST', url: 'api/ReleaseNotes', data: api.fields, responseType: 'arraybuffer' })
                .success(function (response, status, xhr) {
                    modalInstance.close()
                    // thanks StackOverflow for this snippet of code that works in all browsers
                    // modified slightly from its original format for Angular

                    // check for a filename
                    var filename = "";
                    var disposition = xhr('Content-Disposition');
                    if (disposition && disposition.indexOf('attachment') !== -1) {
                        var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                        var matches = filenameRegex.exec(disposition);
                        if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
                    }

                    var type = xhr('Content-Type');
                    var blob = new Blob([response], { type: type });

                    if (typeof window.navigator.msSaveBlob !== 'undefined') {
                        window.navigator.msSaveBlob(blob, filename);
                    } else {
                        var URL = window.URL || window.webkitURL;
                        var downloadUrl = URL.createObjectURL(blob);

                        if (filename) {
                            var a = document.createElement("a");
                            if (typeof a.download === 'undefined') {
                                window.location = downloadUrl;
                            } else {
                                a.href = downloadUrl;
                                a.download = filename;
                                document.body.appendChild(a);
                                a.click();
                            }
                        } else {
                            window.location = downloadUrl;
                        }

                        $timeout(function () { URL.revokeObjectURL(downloadUrl); }, 100);
                    }
                })
                .error(function (error) {
                    console.log(error)
                    modalInstance.close()
                }) 
                // window.location.href = 'api/ReleaseNotes?data=' + encodeURIComponent(angular.toJson(api.fields))
            }
        }
        utilities.initialize()
        return api;
    }
}])

releaseNotes.controller('releaseNotes-progressModal-controller', ['$scope', '$modalInstance', function ($scope, $modalInstance) {
    $scope.progressModal = new function () {
        var api = {
            close: function () {
                $modalInstance.close()
            }
        }
        return api
    }
}])