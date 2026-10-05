(function () {
  'use strict';

  angular.module('ExcelApp', [])
    .controller('MainController', ['$scope', '$http', '$timeout', function ($scope, $http, $timeout) {
      $scope.dbStatus = 'unknown'; // 'unknown', 'online', 'offline'
      $scope.tableStatus = 'unknown'; // 'unknown', 'found', 'missing'
      $scope.tableName = 'V_EMP_DETAILS';
      $scope.recordCount = 0;
      
      $scope.fileSelected = null;
      $scope.fileName = '';
      $scope.isProcessing = false;
      
      $scope.errorMsg = '';
      $scope.successMsg = '';

      // Excel inspection results
      $scope.sheetNames = [];
      $scope.selectedSheets = {};
      $scope.excelColumns = [];
      
      // Column Mapping State
      $scope.autoDetected = false;
      $scope.detectedColumnName = '';
      $scope.detectedKeyType = '';
      
      $scope.selectedInputCol = '';
      $scope.selectedKeyType = 'PIS'; // 'PCNO', 'PIS', 'ACCNO', 'GPFPRAN'

      // Supported Key Types for Dropdown
      $scope.availableKeyTypes = [
        { id: 'PCNO', name: 'PC Number (PCNO)', desc: 'Matches against Employee PC Number' },
        { id: 'PIS', name: 'PIS Number (PIS)', desc: 'Matches against PIS Number' },
        { id: 'ACCNO', name: 'Account Number (ACCNO)', desc: 'Matches against Account Number (GPF / PRAN)' },
        { id: 'GPFPRAN', name: 'GPF / PRAN (GPFPRAN)', desc: 'Matches against GPF / PRAN Number' }
      ];

      // Available Output Columns to append to Excel
      $scope.targetColumns = [
        { id: 'PCNO', name: 'PCNO', label: 'PC Number (PCNO)', desc: 'Employee PC Number' },
        { id: 'PIS', name: 'PIS', label: 'PIS Number (PIS)', desc: 'Personnel Information System ID' },
        { id: 'ACCNO', name: 'ACCNO', label: 'Account Number (ACCNO)', desc: 'Account / Pension ID' },
        { id: 'GPFPRAN', name: 'GPFPRAN', label: 'GPF / PRAN (GPFPRAN)', desc: 'GPF / PRAN Account ID' }
      ];

      // Selected Output Columns checklist
      $scope.outputCols = {
        'PCNO': true,
        'PIS': false,
        'ACCNO': true,
        'GPFPRAN': false
      };

      // Check database connection and table/view status
      $scope.checkDbStatus = function () {
        $scope.errorMsg = '';
        $http.get('default.aspx?action=status')
          .then(function (response) {
            var data = response.data;
            if (data.connected) {
              $scope.dbStatus = 'online';
              $scope.tableName = data.tableName || 'V_EMP_DETAILS';
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
              $scope.errorMsg = 'Unable to connect to master records service: ' + (data.error || 'Service unavailable');
            }
          }, function (error) {
            $scope.dbStatus = 'offline';
            $scope.tableStatus = 'unknown';
            $scope.errorMsg = 'Error communicating with records service.';
            console.error(error);
          });
      };

      // When an Excel file is selected, inspect headers and sheets
      $scope.setFile = function (file) {
        $scope.fileSelected = file;
        $scope.fileName = file.name;
        $scope.errorMsg = '';
        $scope.successMsg = '';
        $scope.sheetNames = [];
        $scope.selectedSheets = {};
        $scope.excelColumns = [];
        $scope.autoDetected = false;
        $scope.detectedColumnName = '';
        $scope.detectedKeyType = '';

        $scope.inspectUploadedFile();
      };

      // Inspect file headers and worksheets via backend
      $scope.inspectUploadedFile = function () {
        if (!$scope.fileSelected) return;
        $scope.isProcessing = true;

        var fd = new FormData();
        fd.append('excelFile', $scope.fileSelected);

        $http.post('default.aspx?action=inspect', fd, {
          transformRequest: angular.identity,
          headers: { 'Content-Type': undefined }
        })
        .then(function (response) {
          $scope.isProcessing = false;
          if (response.data.success) {
            $scope.sheetNames = response.data.sheets || [];
            angular.forEach($scope.sheetNames, function (sheet) {
              $scope.selectedSheets[sheet] = true;
            });

            $scope.excelColumns = response.data.columns || [];

            // Auto-detection logic
            if (response.data.detectedKeyType) {
              $scope.autoDetected = true;
              $scope.detectedColumnName = response.data.detectedColumn;
              $scope.detectedKeyType = response.data.detectedKeyType;
              $scope.selectedInputCol = response.data.detectedColumn;
              $scope.selectedKeyType = response.data.detectedKeyType;
            } else {
              $scope.autoDetected = false;
              $scope.detectedColumnName = '';
              $scope.detectedKeyType = '';
              if ($scope.excelColumns.length > 0) {
                $scope.selectedInputCol = $scope.excelColumns[0].name;
              }
              $scope.selectedKeyType = 'PIS';
            }

            // Set smart default output columns
            $scope.applyDefaultOutputCols();
          } else {
            $scope.errorMsg = 'Failed to inspect file: ' + (response.data.error || 'Unknown error');
            $scope.clearFileDirect();
          }
        }, function (error) {
          $scope.isProcessing = false;
          $scope.errorMsg = 'Error communicating with server while reading Excel headers.';
          $scope.clearFileDirect();
        });
      };

      // Smart defaults for output columns based on input key type
      $scope.applyDefaultOutputCols = function () {
        var key = $scope.selectedKeyType;
        if (key === 'GPFPRAN' || key === 'ACCNO') {
          $scope.outputCols = { 'PCNO': true, 'PIS': true, 'ACCNO': false, 'GPFPRAN': false };
        } else if (key === 'PIS') {
          $scope.outputCols = { 'PCNO': true, 'PIS': false, 'ACCNO': true, 'GPFPRAN': false };
        } else if (key === 'PCNO') {
          $scope.outputCols = { 'PCNO': false, 'PIS': true, 'ACCNO': true, 'GPFPRAN': false };
        }
      };

      // Handler when user manually changes input key type
      $scope.onKeyTypeChange = function () {
        $scope.applyDefaultOutputCols();
      };

      // File Selection logic
      $scope.onFileSelect = function (element) {
        $scope.$apply(function () {
          if (element.files.length > 0) {
            $scope.setFile(element.files[0]);
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
                $scope.setFile(e.dataTransfer.files[0]);
              });
            }
          });
        }
      });

      // Clear current file selection
      $scope.clearFileDirect = function () {
        $scope.fileSelected = null;
        $scope.fileName = '';
        $scope.sheetNames = [];
        $scope.selectedSheets = {};
        $scope.excelColumns = [];
        $scope.autoDetected = false;
        $scope.detectedColumnName = '';
        $scope.detectedKeyType = '';
        $scope.selectedInputCol = '';
        var fileInput = document.getElementById('fileInput');
        if (fileInput) {
          fileInput.value = '';
        }
      };

      $scope.clearFile = function () {
        $scope.clearFileDirect();
        $scope.successMsg = '';
        $scope.errorMsg = '';
      };

      // Get count of selected output columns
      $scope.getSelectedOutputCount = function () {
        var count = 0;
        angular.forEach($scope.outputCols, function (val) {
          if (val) count++;
        });
        return count;
      };

      // Process Excel file and download result
      $scope.processExcelFile = function () {
        if (!$scope.fileSelected) {
          $scope.errorMsg = 'Please select an Excel or CSV file first.';
          return;
        }

        if (!$scope.selectedInputCol) {
          $scope.errorMsg = 'Please select which column in your Excel contains the identifier.';
          return;
        }

        if (!$scope.selectedKeyType) {
          $scope.errorMsg = 'Please specify what type of identifier the column represents.';
          return;
        }

        var outputList = [];
        angular.forEach($scope.outputCols, function (val, key) {
          if (val) outputList.push(key);
        });

        if (outputList.length === 0) {
          $scope.errorMsg = 'Please select at least one output column to append to the Excel file.';
          return;
        }

        var selectedSheetsList = [];
        angular.forEach($scope.selectedSheets, function (val, key) {
          if (val) selectedSheetsList.push(key);
        });

        if ($scope.sheetNames.length > 0 && selectedSheetsList.length === 0) {
          $scope.errorMsg = 'Please select at least one worksheet to process.';
          return;
        }

        $scope.isProcessing = true;
        $scope.errorMsg = '';
        $scope.successMsg = '';

        var fd = new FormData();
        fd.append('excelFile', $scope.fileSelected);
        fd.append('inputCol', $scope.selectedInputCol);
        fd.append('keyType', $scope.selectedKeyType);
        fd.append('outputCols', outputList.join(','));
        fd.append('sheets', selectedSheetsList.join(','));

        $http.post('default.aspx?action=process', fd, {
          transformRequest: angular.identity,
          headers: { 'Content-Type': undefined },
          responseType: 'blob'
        })
        .then(function (response) {
          $scope.isProcessing = false;

          // Check if response is JSON error
          var contentType = response.headers('Content-Type');
          if (contentType && contentType.indexOf('application/json') !== -1) {
            var reader = new FileReader();
            reader.onload = function () {
              $scope.$apply(function () {
                try {
                  var errObj = JSON.parse(reader.result);
                  $scope.errorMsg = 'Processing failed: ' + errObj.error;
                } catch (e) {
                  $scope.errorMsg = 'An error occurred during spreadsheet processing.';
                }
              });
            };
            reader.readAsText(response.data);
            return;
          }

          // Trigger download in browser
          var blobType = $scope.fileName.endsWith('.csv') ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
          var blob = new Blob([response.data], { type: blobType });
          var downloadUrl = URL.createObjectURL(blob);
          var a = document.createElement('a');

          var nameParts = $scope.fileName.split('.');
          var ext = nameParts.pop();
          var baseName = nameParts.join('.');

          a.href = downloadUrl;
          a.download = baseName + '_Processed.' + ext;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(downloadUrl);

          $scope.successMsg = 'Spreadsheet processed successfully and downloaded! (Matched using ' + $scope.selectedKeyType + ')';
        }, function (error) {
          $scope.isProcessing = false;
          $scope.errorMsg = 'Failed to process spreadsheet. Please verify service connection and file format.';
          console.error(error);
        });
      };

      // Initialize status on load
      $scope.checkDbStatus();
    }]);
})();
