(function () {
    'use strict';

    function main($rootScope, $scope, $http, $window, Upload, $timeout, $mdSidenav) {
        $scope.jid = randomString();
        $scope.templateFile = {};
        $scope.datasourceFile = {};
        $scope.datasourceTableIndex = 0;
        $scope.datasourceName = "";
        $scope.stage = 1;
        $scope.REQUESTED_FILEFORMAT = REQUESTED_FILEFORMAT;

        if (REQUESTED_EXTENSION === null) {
            $scope.REQUESTED_EXTENSION = null;
        } else {
            $scope.REQUESTED_EXTENSION = REQUESTED_EXTENSION.toUpperCase();
        }

        $scope.uploadTemplateFile = function () {
            $scope.uploadFile(
                $scope.templateFile.file,
                function (response) {
                    $scope.stage = 2;
                },
                function (response) {
                    alertError(response);
                },
                function (e) {
                    $scope.templateFile.loadedSize = e.loaded;
                    $scope.templateFile.totalSize = e.total;
                    $scope.templateFile.progress = Math.min(100, parseInt(100.0 * e.loaded / e.total));
                }
            );
        };

        $scope.uploadDatasourceFile = function () {
            $scope.uploadFile(
                $scope.datasourceFile.file,
                function () {
                    $scope.stage = 3;
                },
                function (response) {
                    alertError(response);
                },
                function (e) {
                    $scope.datasourceFile.loadedSize = e.loaded;
                    $scope.datasourceFile.totalSize = e.total;
                    $scope.datasourceFile.progress = Math.min(100, parseInt(100.0 * e.loaded / e.total));
                }
            );
        };

        $scope.assembleDocument = function () {
            $http({
                method: "POST",
                url: GROUPDOCS_ASSEMBLY_API_BASE + "Assemble?" + $.param({
                    jid: $scope.jid,
                    templateFilename: $scope.templateFile.file.name,
                    datasourceFilename: $scope.datasourceFile.file.name,
                    assembledFilename: "assembled-" + $scope.templateFile.file.name,
                    datasourceName: $scope.datasourceName,
                    datasourceTableIndex: $scope.datasourceTableIndex,
                }),
                responseType: "application/json",
            }).then(
                function (response) {
                    window.location.href = GROUPDOCS_ASSEMBLY_API_BASE + "Download?" + $.param({
                        jid: $scope.jid,
                        filename: response.data.filename,
                    });
                },
                function (response) {
                    alertError(response)
                }
            );
        };

        $scope.uploadFile = function (file, success, error, progress) {
            Upload.upload({
                url: GROUPDOCS_ASSEMBLY_API_BASE + "Upload?" + $.param({
                    jid: $scope.jid,
                }),
                data: {
                    file: file,
                }
            }).then(success, error, progress);
        };

    }

    function randomString() {
        return Math.random().toString(36).substring(2)
            + Math.random().toString(36).substring(2)
            + Math.random().toString(36).substring(2);
    }

    function alertError(response) {
        console.error("Error occurred while processing request.", response.data);
        alert(response.data.Message
            + " "
            + response.data.ExceptionMessage
            + "\n\n"
            + "See console for details."
        );
    }

    angular.module('GroupDocsAssemblyApp').controller('Main', main);
})();
