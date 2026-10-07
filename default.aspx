<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="ExcelProcessor._default" %>
<!DOCTYPE html>
<html lang="en" ng-app="ExcelApp">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
  <meta name="description" content="Offline Employee Data Sheet Matcher and Record Enricher" />
  <title>Employee Record Mapper</title>
  
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
          <h1 class="app-title">Employee Record Mapper</h1>
          <p class="subtitle">Auto-detects employee identifiers, matches master records, and enriches spreadsheets offline</p>
        </div>
      </div>

      <div class="d-flex align-items-center" style="gap: 12px;">
        <% if (Session["UserPCNO"] != null || User.Identity.IsAuthenticated) { %>
        <div class="d-flex align-items-center" style="gap: 8px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 9999px; padding: 4px 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04);">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
            <circle cx="12" cy="7" r="4"></circle>
          </svg>
          <span style="font-size: 0.84rem; font-weight: 600; color: #1e293b;">
            <%= Session["UserName"] != null ? Session["UserName"] : (Session["UserPCNO"] ?? User.Identity.Name) %>
          </span>
          <span style="font-size: 0.75rem; color: #2563eb; background: #eff6ff; border: 1px solid #bfdbfe; padding: 2px 8px; border-radius: 9999px; font-weight: 600;">
            <%= Session["UserDivName"] != null ? Session["UserDivName"] : (Session["UserPCNO"] ?? User.Identity.Name) %>
          </span>
        </div>
        <a href="default.aspx?action=logout" class="btn btn-sm btn-outline-danger py-2 px-3" style="font-weight: 600; border-radius: 8px; text-decoration: none;" title="Sign out of system">
          Sign Out
        </a>
        <% } %>
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
          <div class="custom-alert-warning" ng-if="duplicateCount > 0" role="alert">
            <div class="d-flex align-items-start">
              <span style="font-size: 1.3rem; margin-right: 12px; line-height: 1.2;">⚠️</span>
              <div>
                <strong>Duplicate Entries Detected:</strong>
                <div>{{ duplicateCount }} row(s) contain repeated identifier values and have been highlighted with a soft yellow fill in your downloaded Excel sheet.</div>
                <div class="mt-2 small" style="opacity: 0.95;">
                  &bull; <strong>If duplicate is not required:</strong> Simply select and delete the row in Excel.<br />
                  &bull; <strong>If duplicate is required:</strong> Select the highlighted row/cells in Excel and click <em>Fill Color &rarr; No Fill</em> to clear the highlight easily.
                </div>
              </div>
            </div>
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
                <div class="text-muted small">{{ (fileSelected.size / 1024) | number:1 }} KB &bull; {{ sheetNames.length || 1 }} Sheet(s) &bull; Active: <strong>{{ selectedSheet }}</strong></div>
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
              <div class="flow-step" ng-if="sheetNames.length > 1">
                <span>Active Sheet:</span>
                <strong style="color: #4338ca;">{{ selectedSheet }}</strong>
              </div>
              <div class="flow-arrow" ng-if="sheetNames.length > 1">&rarr;</div>
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

            <!-- Sheet Selection Card (when file has multiple sheets) -->
            <div class="mapping-card mb-3" ng-show="sheetNames.length > 1" style="background: #f8faff; border-color: #cbd5e1;">
              <div class="d-flex justify-content-between align-items-center mb-2">
                <label class="section-label small font-weight-bold text-uppercase mb-0" style="letter-spacing: 0.5px; color: #1e40af;">
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24" style="margin-right: 6px; vertical-align: -2px;">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                    <line x1="3" y1="9" x2="21" y2="9"></line>
                    <line x1="9" y1="21" x2="9" y2="9"></line>
                  </svg>
                  Active Worksheet to Process
                </label>
                <span style="font-size: 0.76rem; background: #e0e7ff; color: #3730a3; padding: 3px 10px; border-radius: 9999px; font-weight: 600;">
                  {{ sheetNames.length }} Sheets Detected
                </span>
              </div>
              <div class="row align-items-center">
                <div class="col-md-7">
                  <label class="text-muted small d-block mb-1 font-weight-bold">Select Sheet:</label>
                  <select class="form-select-glass" ng-model="selectedSheet" ng-change="onSheetChange()" style="font-weight: 600; font-size: 0.95rem; border-color: #93c5fd;">
                    <option ng-repeat="s in sheetNames" value="{{ s }}">
                      📄 {{ s }}
                    </option>
                  </select>
                </div>
                <div class="col-md-5 mt-2 mt-md-0">
                  <div class="small text-muted" style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; line-height: 1.4;">
                    <span style="color: #059669; font-weight: 600;">✓ Safe Isolation:</span> Only the selected sheet will be enriched. All other sheets remain untouched.
                  </div>
                </div>
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

            <!-- Card 3: Duplicate Detection & Highlighting -->
            <div class="mapping-card">
              <div class="d-flex justify-content-between align-items-center mb-2">
                <label class="section-label small font-weight-bold text-uppercase mb-0" style="letter-spacing: 0.5px;">3. Duplicate Identifier Handling</label>
                <span class="badge" style="background: #fef3c7; color: #92400e; border: 1px solid #fde68a; font-size: 11px; padding: 4px 8px; border-radius: 6px;">
                  ⚠️ Duplicate Detection
                </span>
              </div>
              <label class="output-checkbox-card" ng-class="{'active': highlightDuplicates}" style="cursor: pointer; display: flex; align-items: flex-start; gap: 12px; margin-bottom: 0;">
                <input type="checkbox" ng-model="highlightDuplicates" style="margin-top: 3px; cursor: pointer;" />
                <div>
                  <div class="font-weight-bold small checkbox-title">Highlight Duplicate Rows in Excel (Soft Yellow Fill)</div>
                  <div class="text-muted" style="font-size: 11.5px; line-height: 1.4; margin-top: 3px;">
                    When enabled, any rows containing duplicate values in the identifier column will be highlighted across the whole row. If you don't need a duplicate, you can delete it in Excel; if you need to keep it, simply click <strong>Fill Color &rarr; No Fill</strong>.
                  </div>
                </div>
              </label>
            </div>

            <!-- Action Button -->
            <button class="btn btn-glass w-100 font-weight-bold text-uppercase py-3" ng-click="processExcelFile()" ng-disabled="isProcessing || dbStatus !== 'online'">
              Process Sheet &amp; Download Enriched Copy
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
