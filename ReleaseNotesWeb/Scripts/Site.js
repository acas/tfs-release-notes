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
            load: function() {
                utilities.loadGenerators()
            },
            initialize: function () {
                utilities.load()
            }
        }

        var api = {
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
                        // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                        window.navigator.msSaveBlob(blob, filename);
                    } else {
                        var URL = window.URL || window.webkitURL;
                        var downloadUrl = URL.createObjectURL(blob);

                        if (filename) {
                            // use HTML5 a[download] attribute to specify filename
                            var a = document.createElement("a");
                            // safari doesn't support this yet
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

                        $timeout(function () { URL.revokeObjectURL(downloadUrl); }, 100); // cleanup
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