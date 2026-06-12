(function () {
  'use strict';

  angular.module('ExcelApp', [])
    .controller('MainController', ['$scope', '$http', '$timeout', function ($scope, $http, $timeout) {
      $scope.dbStatus = 'unknown'; // 'unknown', 'online', 'offline'
      $scope.tableStatus = 'unknown'; // 'unknown', 'found', 'missing'
      $scope.recordCount = 0;
      
      $scope.fileSelected = null;
      $scope.fileName = '';
      $scope.isProcessing = false;
      
      $scope.errorMsg = '';
      $scope.successMsg = '';
      $scope.selectedCols = {};

      // Fetch database and table configuration status
      $scope.checkDbStatus = function () {
        $scope.errorMsg = '';
        $http.get('default.aspx?action=status')
          .then(function (response) {
            var data = response.data;
            if (data.connected) {
              $scope.dbStatus = 'online';
              $scope.mappableColumns = data.mappableColumns || [];
              angular.forEach($scope.mappableColumns, function (col, index) {
                $scope.selectedCols[col.dbColumn] = (index === 0);
              });
              if (data.tableExists) {
                $scope.tableStatus = 'found';
                $scope.recordCount = data.recordCount;
              } else {
                $scope.tableStatus = 'missing';
                $scope.recordCount = 0;
              }
            } else {
              $scope.dbStatus = 'offline';
              $scope.tableStatus = 'unknown';
              $scope.errorMsg = 'Failed to connect to Oracle database: ' + data.error;
            }
          }, function (error) {
            $scope.dbStatus = 'offline';
            $scope.tableStatus = 'unknown';
            $scope.errorMsg = 'Error communicating with server handler.';
            console.error(error);
          });
      };


      // File Selection logic
      $scope.onFileSelect = function (element) {
        $scope.$apply(function () {
          if (element.files.length > 0) {
            $scope.fileSelected = element.files[0];
            $scope.fileName = $scope.fileSelected.name;
            $scope.errorMsg = '';
            $scope.successMsg = '';
          }
        });
      };

      // Drag and drop event binders
      $timeout(function() {
        var dropZone = document.getElementById('drop-zone');
        if (dropZone) {
          dropZone.addEventListener('dragover', function(e) {
            e.preventDefault();
            dropZone.classList.add('dragover');
          });

          dropZone.addEventListener('dragleave', function(e) {
            e.preventDefault();
            dropZone.classList.remove('dragover');
          });

          dropZone.addEventListener('drop', function(e) {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
              $scope.$apply(function() {
                $scope.fileSelected = e.dataTransfer.files[0];
                $scope.fileName = $scope.fileSelected.name;
                $scope.errorMsg = '';
                $scope.successMsg = '';
              });
            }
          });
        }
      });

      // Clear current file selection
      $scope.clearFile = function () {
        $scope.fileSelected = null;
        $scope.fileName = '';
        $scope.successMsg = '';
        $scope.errorMsg = '';
        var fileInput = document.getElementById('fileInput');
        if (fileInput) {
          fileInput.value = '';
        }
      };

      // Upload and Process Excel
      $scope.processExcelFile = function () {
        if (!$scope.fileSelected) {
          $scope.errorMsg = 'Please select an Excel file first.';
          return;
        }

        var cols = [];
        angular.forEach($scope.selectedCols, function(value, key) {
          if (value) {
            cols.push(key);
          }
        });

        if (cols.length === 0) {
          $scope.errorMsg = 'Please select at least one database column to pull.';
          return;
        }

        $scope.isProcessing = true;
        $scope.errorMsg = '';
        $scope.successMsg = '';

        var fd = new FormData();
        fd.append('excelFile', $scope.fileSelected);
        fd.append('columns', cols.join(','));

        $http.post('default.aspx?action=process', fd, {
          transformRequest: angular.identity,
          headers: { 'Content-Type': undefined },
          responseType: 'blob' // Essential to read binary Excel file from response stream
        })
        .then(function (response) {
          $scope.isProcessing = false;
          
          // Check if response is JSON (error) or binary (success)
          var contentType = response.headers('Content-Type');
          if (contentType && contentType.indexOf('application/json') !== -1) {
            // Read blob as text to parse error message
            var reader = new FileReader();
            reader.onload = function() {
              $scope.$apply(function() {
                try {
                  var errObj = JSON.parse(reader.result);
                  $scope.errorMsg = 'Processing failed: ' + errObj.error;
                } catch(e) {
                  $scope.errorMsg = 'An error occurred during sheet processing.';
                }
              });
            };
            reader.readAsText(response.data);
            return;
          }

          // Trigger download in browser
          var blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          var downloadUrl = URL.createObjectURL(blob);
          var a = document.createElement('a');
          
          // Create name with _Processed suffix
          var nameParts = $scope.fileName.split('.');
          var ext = nameParts.pop();
          var baseName = nameParts.join('.');
          
          a.href = downloadUrl;
          a.download = baseName + '_Processed.' + ext;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(downloadUrl);

          $scope.successMsg = 'Excel file processed successfully and download started!';
        }, function (error) {
          $scope.isProcessing = false;
          $scope.errorMsg = 'Failed to process the Excel file. Verify database connection and schema.';
          console.error(error);
        });
      };

      // Initialize status checks on load
      $scope.checkDbStatus();
    }]);
})();
