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
            loadPresets: function () {
                var deferred = $q.defer();
                $http({ method: 'GET', url: 'api/Presets' })
                .success(function (data) {
                    api.presets = data;
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            savePreset: function (fields) {
                var deferred = $q.defer();
                $http({ method: 'PUT', url: 'api/Presets', data: fields })
                .success(function (data) {
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            deletePreset: function (fields) {
                var deferred = $q.defer();
                $http({ method: 'DELETE', url: 'api/Presets', data: fields })
                .success(function (data) {
                    utilities.loadPresets()
                    deferred.resolve()
                })
                .error(function () { deferred.reject() })
                return deferred.promise
            },
            load: function() {
                utilities.loadGenerators()
                utilities.loadPresets();
            },
            initialize: function () {
                utilities.load()
            }
        }

        var api = {
            showPresetSection: true,
            showRequiredSection: false,
            showOptionalSection: true,
            presets: [],
            selectedPreset: null,
            newPreset: null,
            generators: [],
            selectedGenerator: null,
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
            savePreset: function () {
                if (api.selectedPreset) {
                    api.fields['presetName'] = api.selectedPreset.presetName
                }
                if (api.newPreset) {
                    api.fields['presetName'] = api.newPreset
                }
                utilities.savePreset(api.fields).then(function () {
                    utilities.loadPresets().then(function () {
                        var selectedPreset = _.find(api.presets, function (preset) {
                            return api.fields.presetName === preset.presetName
                        })
                        if (selectedPreset) api.selectedPreset = api.trashyPreset = selectedPreset
                        api.newPreset = null
                    })
                })
                //delete api.fields['presetName']
            },
            deletePreset: function () {
                if (api.selectedPreset) {
                    utilities.deletePreset(api.fields)
                }
            },
            presetChanged: function() {
                if (api.selectedPreset) {
                    api.fields = api.selectedPreset
                    api.selectedGenerator = _.find(api.generators, function (generator) {
                        return generator.name === api.fields.generator
                    })
                }
            },
            setGenerator: function() {
                api.fields.generator = api.selectedGenerator.name;
            },
            generate: function () {
                var modalInstance = $modal.open({
                    templateUrl: 'Home/Waiting',
                    controller: 'releaseNotes-progressModal-controller',
                    size: 'md',
					backdrop: 'static'
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
                    if (error) {
                        console.log(error)
                    }
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