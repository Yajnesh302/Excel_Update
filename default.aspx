<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="ExcelProcessor._default" %>
<!DOCTYPE html>
<html lang="en" ng-app="ExcelApp">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
  <meta name="description" content="Offline Employee Data Sheet Matcher and Record Enricher" />
  <title>Employee Data Processor</title>
  
  <!-- Local Offline Stylesheets -->
  <link rel="stylesheet" href="lib/bootstrap.min.css" />
  <link rel="stylesheet" href="style.css" />
</head>
<body ng-controller="MainController">

  <div class="app-container">
    
    <!-- Top Navigation Header -->
    <header class="top-navbar">
      <div class="brand-wrapper">
        <div class="brand-icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
        </div>
        <div>
          <h1 class="app-title">Employee Data Processor</h1>
          <p class="subtitle">Auto-detects employee identifiers, matches master records, and enriches spreadsheets offline</p>
        </div>
      </div>

      <div class="d-flex align-items-center" style="gap: 12px;">
        <button class="btn btn-sm btn-glass-secondary py-2 px-3" ng-click="checkDbStatus()" ng-disabled="isProcessing" title="Refresh System Status">
          Refresh Status
        </button>
      </div>
    </header>

    <!-- Top Quick Metric Stats Grid -->
    <div class="quick-stats-grid">
      <div class="stat-card">
        <div class="stat-icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
            <ellipse cx="12" cy="5" rx="9" ry="3"></ellipse>
            <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path>
            <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path>
          </svg>
        </div>
        <div class="stat-content">
          <span class="stat-label">Master Service</span>
          <span class="stat-value" ng-if="dbStatus === 'online'" style="color: #059669;">Connected (Online)</span>
          <span class="stat-value" ng-if="dbStatus !== 'online'" style="color: #e11d48;">Disconnected</span>
        </div>
      </div>

      <div class="stat-card">
        <div class="stat-icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
            <circle cx="9" cy="7" r="4"></circle>
            <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
            <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
          </svg>
        </div>
        <div class="stat-content">
          <span class="stat-label">Employee Directory</span>
          <span class="stat-value">{{ recordCount }} Active Records</span>
        </div>
      </div>

      <div class="stat-card">
        <div class="stat-icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
            <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
            <line x1="8" y1="21" x2="16" y2="21"></line>
            <line x1="12" y1="17" x2="12" y2="21"></line>
          </svg>
        </div>
        <div class="stat-content">
          <span class="stat-label">Supported Identifiers</span>
          <span class="stat-value">PIS, PCNO, GPF, Bank A/C</span>
        </div>
      </div>
    </div>

    <!-- Main Content Row -->
    <div class="row">
      
      <!-- Left Column: System Overview & Guidelines -->
      <div class="col-xl-4 col-lg-4 col-md-12">
        
        <!-- How It Works Panel -->
        <div class="glass-panel">
          <h5 class="panel-heading mb-3 d-flex align-items-center">
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 8px;">
              <circle cx="12" cy="12" r="10"></circle>
              <line x1="12" y1="16" x2="12" y2="12"></line>
              <line x1="12" y1="8" x2="12.01" y2="8"></line>
            </svg>
            System Overview &amp; Features
          </h5>
          <ul class="feature-list mb-0 pl-3">
            <li><strong>Smart Auto-Detection:</strong> Uploaded spreadsheets are automatically scanned to detect <code>GPFPRAN</code>, <code>PIS</code>, <code>PCNO</code>, or Bank Account (<code>ACCNO</code>) headers.</li>
            <li><strong>Custom Mapping:</strong> If your sheet uses unique column headers, simply choose the matching identifier column from the dropdown.</li>
            <li><strong>Customizable Columns:</strong> Select any combination of employee fields to append directly to your downloaded spreadsheet.</li>
            <li><strong>Supported Formats:</strong> Works seamlessly with <strong>.xlsx</strong>, <strong>.xls</strong>, and <strong>.csv</strong> files.</li>
          </ul>
        </div>

      </div>

      <!-- Right Column: Interactive Processing Studio -->
      <div class="col-xl-8 col-lg-8 col-md-12">
        <div class="glass-panel">
          
          <div class="d-flex justify-content-between align-items-center mb-3">
            <div>
              <h4 class="mb-1 panel-heading">Spreadsheet Processing Studio</h4>
              <p class="text-muted small mb-0">Upload any spreadsheet to automatically match records and download an enriched copy.</p>
            </div>
          </div>

          <!-- Alert Notifications -->
          <div class="custom-alert-danger" ng-if="errorMsg" role="alert">
            <strong>Notice:</strong> {{ errorMsg }}
          </div>
          <div class="custom-alert-success" ng-if="successMsg" role="alert">
            <strong>Success:</strong> {{ successMsg }}
          </div>

          <!-- Upload Drop Zone (Visible when no file is selected) -->
          <div id="drop-zone" class="drop-zone mb-4" ng-show="!fileSelected">
            <svg xmlns="http://www.w3.org/2000/svg" width="56" height="56" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" class="mb-3" viewBox="0 0 24 24" style="stroke: url(#blueIndigoGradient);">
              <defs>
                <linearGradient id="blueIndigoGradient" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stop-color="#2563eb" />
                  <stop offset="100%" stop-color="#3b82f6" />
                </linearGradient>
              </defs>
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
            <h5 class="drop-zone-title mb-1">Drag &amp; drop your Excel or CSV file here</h5>
            <p class="text-muted small mb-0">or click to browse local files (.xlsx, .xls, .csv)</p>
            <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" onchange="angular.element(this).scope().onFileSelect(this)" />
          </div>

          <!-- Selected File Header (Visible when file is selected) -->
          <div class="selected-file-card mb-4" ng-show="fileSelected">
            <div class="d-flex align-items-center">
              <div class="stat-icon" style="margin-right: 14px; width: 44px; height: 44px;">
                <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                  <line x1="16" y1="13" x2="8" y2="13"></line>
                  <line x1="16" y1="17" x2="8" y2="17"></line>
                  <polyline points="10 9 9 9 8 9"></polyline>
                </svg>
              </div>
              <div>
                <div class="selected-filename text-truncate" style="max-width: 480px;">{{ fileName }}</div>
                <div class="text-muted small">{{ (fileSelected.size / 1024) | number:1 }} KB &bull; {{ sheetNames.length || 1 }} Sheet(s) detected</div>
              </div>
            </div>
            <button type="button" class="btn btn-sm btn-glass-danger font-weight-bold py-2 px-3" ng-click="clearFile()" ng-disabled="isProcessing">
              Remove File
            </button>
          </div>

          <!-- Spinner and Loading -->
          <div class="spinner-container" ng-if="isProcessing">
            <div class="glow-spinner"></div>
            <p class="text-muted mt-3 mb-0 small">Processing spreadsheet &amp; matching records...</p>
          </div>

          <!-- Interactive Mapping & Configuration Section -->
          <div ng-show="fileSelected && !isProcessing">
            
            <!-- Flow Banner -->
            <div class="flow-banner">
              <div class="flow-step">
                <span>Input:</span>
                <strong>{{ selectedInputCol || 'None' }}</strong>
                <span class="badge badge-pis">{{ selectedKeyType }}</span>
              </div>
              <div class="flow-arrow">&rarr;</div>
              <div class="flow-step">
                <span>Master Records:</span>
                <strong style="color: #059669;">Matched</strong>
              </div>
              <div class="flow-arrow">&rarr;</div>
              <div class="flow-step">
                <span>Appending:</span>
                <strong style="color: #2563eb;">{{ getSelectedOutputCount() }} Column(s)</strong>
              </div>
            </div>

            <!-- Card 1: Column Detection & Mapping -->
            <div class="mapping-card">
              <div class="d-flex justify-content-between align-items-center mb-3">
                <label class="section-label small font-weight-bold text-uppercase mb-0" style="letter-spacing: 0.5px;">1. Column Identification &amp; Mapping</label>
                
                <span class="badge-detected" ng-if="autoDetected">
                  &#10003; Auto-detected: {{ detectedColumnName }} &rarr; {{ detectedKeyType }}
                </span>
                <span class="badge-manual" ng-if="!autoDetected">
                  &#9888; Manual Mapping
                </span>
              </div>

              <div class="row">
                <div class="col-md-6 mb-3">
                  <label class="text-muted small d-block mb-1">Column in Your Sheet:</label>
                  <select class="form-select-glass" ng-model="selectedInputCol">
                    <option ng-repeat="col in excelColumns" value="{{ col.name }}">
                      {{ col.name }} (Column {{ col.index }})
                    </option>
                  </select>
                </div>

                <div class="col-md-6 mb-3">
                  <label class="text-muted small d-block mb-1">Identifier Type:</label>
                  <select class="form-select-glass" ng-model="selectedKeyType" ng-change="onKeyTypeChange()">
                    <option ng-repeat="kt in availableKeyTypes" value="{{ kt.id }}">
                      {{ kt.name }}
                    </option>
                  </select>
                </div>
              </div>
            </div>

            <!-- Card 2: Output Columns Checklist -->
            <div class="mapping-card">
              <div class="d-flex justify-content-between align-items-center mb-3">
                <label class="section-label small font-weight-bold text-uppercase mb-0" style="letter-spacing: 0.5px;">2. Output Columns to Append to File</label>
                <span class="text-muted small">{{ getSelectedOutputCount() }} of 4 selected</span>
              </div>

              <div class="row">
                <div class="col-xl-3 col-md-6 mb-3" ng-repeat="tc in targetColumns">
                  <label class="output-checkbox-card" ng-class="{'active': outputCols[tc.id]}">
                    <input type="checkbox" ng-model="outputCols[tc.id]" />
                    <div>
                      <div class="font-weight-bold small checkbox-title">{{ tc.name }}</div>
                      <div class="text-muted" style="font-size: 11px;">{{ tc.desc }}</div>
                    </div>
                  </label>
                </div>
              </div>
            </div>

            <!-- Card 3: Multi-Sheet Selection (if applicable) -->
            <div class="sheet-selection-card mb-4" ng-show="sheetNames.length > 1">
              <label class="text-muted small d-block mb-3 font-weight-bold text-uppercase" style="letter-spacing: 0.5px;">Worksheets to Process:</label>
              <div class="d-flex flex-wrap" style="gap: 20px;">
                <label ng-repeat="sheet in sheetNames" class="sheet-label d-flex align-items-center small mb-0">
                  <input type="checkbox" ng-model="selectedSheets[sheet]" style="width: 16px; height: 16px; margin-right: 8px; cursor: pointer; accent-color: var(--primary-color);" />
                  {{ sheet }}
                </label>
              </div>
            </div>

            <!-- Action Button -->
            <button class="btn btn-glass w-100 font-weight-bold text-uppercase py-3" ng-click="processExcelFile()" ng-disabled="isProcessing || dbStatus !== 'online'">
              Process File &amp; Download Enriched Copy
            </button>
            
            <p class="text-danger small mt-2 text-center" ng-if="dbStatus !== 'online'">
              * Master records service must be connected to process files.
            </p>
          </div>

        </div>
      </div>

    </div>

    <!-- Footer -->
    <div class="text-center text-muted small mt-4 mb-2">
    </div>

  </div>

  <!-- Local Offline AngularJS Loader -->
  <script src="lib/angular.min.js"></script>
  <!-- Main AngularJS App Script -->
  <script src="app.js"></script>

</body>
</html>
