<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="ExcelProcessor._default" %>
<!DOCTYPE html>
<html lang="en" ng-app="ExcelApp">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
  <meta name="description" content="Offline Excel & Oracle Data Processor" />
  <title>Excel &amp; Oracle Data Processor</title>
  
  <!-- Local Offline Stylesheets -->
  <link rel="stylesheet" href="lib/bootstrap.min.css" />
  <link rel="stylesheet" href="style.css" />
</head>
<body ng-controller="MainController">

  <div class="app-container">
    
    <!-- Title Section -->
    <div class="text-center mb-5">
      <h1 class="app-title">Excel &amp; Oracle Data Processor</h1>
      <p class="subtitle">Process sheets and inject account numbers offline using PIS matching</p>
    </div>

    <!-- Alerts -->
    <div class="row">
      <div class="col-12">
        <div class="alert alert-danger" ng-if="errorMsg" role="alert" style="border-radius: 8px; background: rgba(255, 8, 68, 0.15); border-color: rgba(255, 8, 68, 0.3); color: #ffccd5;">
          <strong>Error:</strong> {{ errorMsg }}
        </div>
        <div class="alert alert-success" ng-if="successMsg" role="alert" style="border-radius: 8px; background: rgba(0, 255, 135, 0.15); border-color: rgba(0, 255, 135, 0.3); color: #d2ffd6;">
          <strong>Success:</strong> {{ successMsg }}
        </div>
      </div>
    </div>

    <div class="row">
      
      <!-- Database & Configuration Panel -->
      <div class="col-lg-5 col-md-12">
        <div class="glass-panel">
          <div class="card-title-container">
            <h4 class="mb-0 text-white font-weight-bold">Oracle Database</h4>
            <button class="btn btn-sm btn-glass-secondary py-1 px-2" ng-click="checkDbStatus()" ng-disabled="isProcessing" title="Refresh Status">
              Refresh
            </button>
          </div>
          
          <hr style="border-color: rgba(255, 255, 255, 0.1);" />

          <!-- Connection Status Info -->
          <div class="mb-4">
            <div class="d-flex justify-content-between align-items-center mb-2">
              <span class="text-muted">Connection String</span>
              <span class="badge badge-pis">Web.config</span>
            </div>
            <div class="d-flex justify-content-between align-items-center mb-3">
              <span class="text-muted">DB Status</span>
              <span class="status-pill status-online" ng-if="dbStatus === 'online'">
                &#9679; Online
              </span>
              <span class="status-pill status-offline" ng-if="dbStatus === 'offline'">
                &#9679; Disconnected
              </span>
              <span class="status-pill status-unknown" ng-if="dbStatus === 'unknown'">
                &#9679; Unknown
              </span>
            </div>

            <div class="d-flex justify-content-between align-items-center mb-3">
              <span class="text-muted">Employee Table</span>
              <span class="status-pill status-online" ng-if="tableStatus === 'found'">
                &#9679; Active ({{ recordCount }} rows)
              </span>
              <span class="status-pill status-offline" ng-if="tableStatus === 'missing'">
                &#9679; Missing Table
              </span>
              <span class="status-pill status-unknown" ng-if="tableStatus === 'unknown'">
                &#9679; Unknown
              </span>
            </div>
          </div>


        </div>
      </div>

      <!-- File Processing Panel -->
      <div class="col-lg-7 col-md-12">
        <div class="glass-panel">
          <h4 class="mb-3 text-white font-weight-bold">Upload &amp; Process Sheet</h4>
          <p class="text-muted small mb-4">
            Select an Excel file (<code>.xlsx</code> or <code>.xls</code>). The server will check each row's <strong>PIS NO</strong> against the database, append an <strong>Account Number</strong> column, and download a processed copy.
          </p>

          <!-- Upload Drop Zone -->
          <div id="drop-zone" class="drop-zone mb-4" ng-show="!fileSelected">
            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="mb-3" viewBox="0 0 24 24" style="stroke: url(#cyanBlueGradient); filter: drop-shadow(0 0 8px rgba(0, 242, 254, 0.3));">
              <defs>
                <linearGradient id="cyanBlueGradient" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stop-color="#00f2fe" />
                  <stop offset="100%" stop-color="#4facfe" />
                </linearGradient>
              </defs>
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
            <h5 class="text-white">Drag &amp; drop Excel file here</h5>
            <p class="text-muted small">or click to browse local files</p>
            <input type="file" id="fileInput" accept=".xlsx, .xls" onchange="angular.element(this).scope().onFileSelect(this)" />
          </div>

          <!-- Selected File Panel -->
          <div class="p-3 mb-4 rounded d-flex justify-content-between align-items-center" ng-show="fileSelected" style="background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1);">
            <div class="d-flex align-items-center">
              <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#00f2fe" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 12px; filter: drop-shadow(0 0 4px rgba(0, 242, 254, 0.3));">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
              <div>
                <div class="text-white font-weight-bold text-truncate" style="max-width: 280px;">{{ fileName }}</div>
                <div class="text-muted small">{{ (fileSelected.size / 1024) | number:1 }} KB</div>
              </div>
            </div>
            <button type="button" class="btn btn-sm btn-glass-danger font-weight-bold py-1 px-3" ng-click="clearFile()" ng-disabled="isProcessing">
              Clear
            </button>
          </div>

          <!-- Spinner and Loading -->
          <div class="spinner-container" ng-if="isProcessing">
            <div class="glow-spinner"></div>
            <p class="text-muted mt-3 mb-0 small">Processing spreadsheet &amp; matching database records...</p>
          </div>

          <!-- Actions -->
          <div class="mt-4" ng-show="fileSelected && !isProcessing">
            <!-- Dynamic Column Selector Checklist -->
            <div class="mb-4 p-3 rounded" style="background: rgba(255, 255, 255, 0.02); border: 1px solid rgba(255, 255, 255, 0.05);" ng-show="mappableColumns.length > 0">
              <label class="text-muted small d-block mb-3 font-weight-bold text-uppercase" style="letter-spacing: 0.5px;">Columns to Pull from Database:</label>
              <div class="d-flex flex-wrap" style="gap: 20px;">
                <label ng-repeat="col in mappableColumns" class="d-flex align-items-center text-white small mb-0" style="cursor: pointer; user-select: none;">
                  <input type="checkbox" ng-model="selectedCols[col.dbColumn]" style="width: 16px; height: 16px; margin-right: 8px; cursor: pointer;" />
                  {{ col.displayName }}
                </label>
              </div>
            </div>

            <button class="btn btn-glass w-100 font-weight-bold text-uppercase py-3" ng-click="processExcelFile()" ng-disabled="isProcessing || dbStatus !== 'online'">
              Process File &amp; Download Copy
            </button>
            <p class="text-danger small mt-2 text-center" ng-if="dbStatus !== 'online'">
              * Please connect to Oracle database to enable processing.
            </p>
          </div>

          <!-- Schema Guideline Mockup -->
          <div class="mt-4">
            <label class="text-muted small">Expected Upload File Column Mapping</label>
            <div class="preview-table-container">
              <table class="preview-table">
                <thead>
                  <tr>
                    <th>Serial Number</th>
                    <th>PIS NO</th>
                    <th>Name</th>
                    <th>Rank</th>
                    <th>Other Columns...</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>1</td>
                    <td>1001 <span class="badge-pis">Matches PIS</span></td>
                    <td>John Doe</td>
                    <td>Manager</td>
                    <td>...</td>
                  </tr>
                  <tr>
                    <td>2</td>
                    <td>1002 <span class="badge-pis">Matches PIS</span></td>
                    <td>Jane Smith</td>
                    <td>Officer</td>
                    <td>...</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

    </div>

    <!-- Footer -->
    <div class="text-center text-muted small mt-5">
      <p>Designed and Configured for Visual Studio 2015 &amp; Oracle 11g | Working Offline</p>
    </div>

  </div>

  <!-- Local Offline AngularJS Loader -->
  <script src="lib/angular.min.js"></script>
  <!-- Main AngularJS App Script -->
  <script src="app.js"></script>

</body>
</html>
