define([
    'qlik',
    'jquery',
    'require',
    'text!./qxport.css'
], function(qlik, $, requirejs, cssContent) {
    'use strict';

    $('<style>').html(cssContent).appendTo('head');
    var qxportSheetContextCache = {};
    var excelJsLoadPromise = null;

    function capitalizeWords(str) {
        return String(str)
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/[_\-]/g, ' ')
            .replace(/\b\w/g, function(c) {
                return c.toUpperCase();
            });
    }

    return {
        definition: {
            type: 'items',
            component: 'accordion',
            items: {
                settings: {
                    uses: 'settings',
                    items: {
                        exportSettings: {
                            type: 'items',
                            label: 'Export Settings',
                            items: {
                                buttonText: {
                                    type: 'string',
                                    label: 'Button Text',
                                    ref: 'buttonText',
                                    defaultValue: 'Export'
                                },
                                fileName: {
                                    type: 'string',
                                    label: 'File Name',
                                    ref: 'fileName',
                                    defaultValue: 'QlikSense_Export'
                                },
                                maxRows: {
                                    type: 'number',
                                    label: 'Max Rows per Sheet',
                                    ref: 'maxRows',
                                    defaultValue: 100000,
                                    min: 100,
                                    max: 1000000
                                }
                            }
                        }
                    }
                },
                about: {
                    label: 'About',
                    type: 'items',
                    items: {
                        name: {
                            label: 'QXport',
                            type: 'string',
                            component: 'text'
                        },
                        version: {
                            label: 'Version: 1.1.0',
                            type: 'string',
                            component: 'text'
                        },
                        author: {
                            label: 'Author: Eli Gohar',
                            type: 'string',
                            component: 'text'
                        }
                    }
                }
            }
        },

        initialProperties: {
            showTitles: false,
            showDetails: true,
            qHyperCubeDef: {
                qDimensions: [],
                qMeasures: [],
                qInitialDataFetch: [{
                    qTop: 0,
                    qLeft: 0,
                    qHeight: 100,
                    qWidth: 10
                }]
            }
        },

        paint: function($element, layout) {
            var app = qlik.currApp();
            var buttonText = layout.buttonText || 'Export';
            var objectId = layout && layout.qInfo && layout.qInfo.qId ? layout.qInfo.qId : 'qxport-default';
            var excludedVisualizations = [
                'qxport',
                'action-button',
                'navigation-button',
                'filterpane',
                'textbox',
                'button',
                'qlik-date-picker',
                'tcmenu',
                'variable-input',
                'sheet-title'
            ];

            var uiState = {
                activeTab: 'sheet',
                appLoaded: false,
                appSheets: [],
                searchTerm: ''
            };

            $element.empty();

            var $container = $(`
                <div class="export-excel-container qxport-mode-sheet">
                    <div class="qxport-tabs">
                        <button type="button" class="qxport-tab-btn qxport-tab-sheet active" data-tab="sheet">Current Sheet</button>
                        <button type="button" class="qxport-tab-btn qxport-tab-app" data-tab="app">App Export</button>
                    </div>

                    <div class="export-list-title"></div>

                    <div class="qxport-app-toolbar" style="display:none;">
                        <input type="text" class="qxport-search-input" placeholder="Search sheets or charts..." />
                        <div class="qxport-toolbar-actions">
                            <button type="button" class="qxport-secondary-btn qxport-select-visible-btn">Select All Visible</button>
                            <button type="button" class="qxport-secondary-btn qxport-clear-visible-btn">Clear All</button>
                            <button type="button" class="qxport-secondary-btn qxport-expand-all-btn">Expand All</button>
                            <button type="button" class="qxport-secondary-btn qxport-collapse-all-btn">Collapse All</button>
                        </div>
                    </div>

                    <div class="chart-list-container">
                        <ul class="export-object-list"><li>Loading charts...</li></ul>
                    </div>

                    <div class="qxport-selection-summary" style="display:none;"></div>

                    <div class="qxport-button-row">
                        <button class="export-excel-btn qxport-main-export-btn" type="button">
                            <span class="export-icon">📊</span>
                            ${escapeHtml(buttonText)}
                        </button>
                    </div>

                    <div class="export-status" style="display: none;"></div>
                    <div class="export-progress" style="display: none;">
                        <div class="progress-bar">
                            <div class="progress-fill"></div>
                        </div>
                        <div class="progress-text">0%</div>
                    </div>
                </div>
            `);

            $element.append($container);

            var $root = $element.find('.export-excel-container');
            var $title = $element.find('.export-list-title');
            var $objectList = $element.find('.export-object-list');
            var $toolbar = $element.find('.qxport-app-toolbar');
            var $searchInput = $element.find('.qxport-search-input');
            var $selectionSummary = $element.find('.qxport-selection-summary');
            var $mainExportBtn = $element.find('.qxport-main-export-btn');
            var $tabButtons = $element.find('.qxport-tab-btn');

            var currentSheet = qlik.navigation.getCurrentSheetId();
            var sheetId = currentSheet && currentSheet.sheetId;

            if (sheetId) {
                qxportSheetContextCache[objectId] = sheetId;
            }

            var effectiveSheetId = sheetId || qxportSheetContextCache[objectId];
            var hasCurrentSheet = !!effectiveSheetId;

            function setRenderMode(mode) {
                $root.removeClass('qxport-mode-sheet qxport-mode-app');
                if (mode === 'app') {
                    $root.addClass('qxport-mode-app');
                } else {
                    $root.addClass('qxport-mode-sheet');
                }
            }

            function setActiveTab(tabName) {
                uiState.activeTab = tabName;
                $tabButtons.removeClass('active');
                $element.find('.qxport-tab-btn[data-tab="' + tabName + '"]').addClass('active');
            }

            function updateMainExportButton() {
                var $icon = $mainExportBtn.find('.export-icon');
                $mainExportBtn.contents().filter(function() {
                    return this.nodeType === 3;
                }).remove();

                if (uiState.activeTab === 'app') {
                    $icon.text('🗂️');
                } else {
                    $icon.text('📊');
                }

                $mainExportBtn.append(' ' + buttonText);

                $mainExportBtn.off('click').on('click', function() {
                    if (uiState.activeTab === 'app') {
                        exportSelectedFromAppMode(app, layout, $element, uiState);
                    } else {
                        if (!hasCurrentSheet) {
                            var $status = $element.find('.export-status');
                            $status.show().text('❗ Current sheet is not available in this view. Please use App Export.').css('color', '#dc3545');
                            return;
                        }
                        exportFromCurrentSheet(app, layout, $element);
                    }
                });
            }

            function updateSelectionSummary() {
                if (uiState.activeTab !== 'app') {
                    $selectionSummary.hide();
                    return;
                }

                var selectedCheckboxes = $element.find('.qxport-app-chart-checkbox:checked');
                var sheetSet = {};

                selectedCheckboxes.each(function() {
                    sheetSet[$(this).data('sheetId')] = true;
                });

                var chartCount = selectedCheckboxes.length;
                var sheetCount = Object.keys(sheetSet).length;

                $selectionSummary
                    .text('Selected: ' + chartCount + ' chart(s) from ' + sheetCount + ' sheet(s)')
                    .show();
            }

            function applyAppFilter() {
                var term = String($searchInput.val() || '').trim().toLowerCase();
                uiState.searchTerm = term;

                var $groups = $element.find('.qxport-sheet-group');

                $groups.each(function() {
                    var $group = $(this);
                    var sheetTitle = String($group.data('sheetTitle') || '').toLowerCase();
                    var groupMatches = !term || sheetTitle.indexOf(term) !== -1;
                    var visibleCount = 0;

                    $group.find('.qxport-chart-item').each(function() {
                        var $item = $(this);
                        var chartTitle = String($item.data('chartTitle') || '').toLowerCase();
                        var visible = !term || groupMatches || chartTitle.indexOf(term) !== -1;
                        $item.toggle(visible);
                        if (visible) visibleCount++;
                    });

                    var showGroup = groupMatches || visibleCount > 0;
                    $group.toggle(showGroup);
                });

                updateSelectionSummary();
            }

            function bindAppModeHandlers() {
                $searchInput.off('input').on('input', function() {
                    applyAppFilter();
                });

                $element.find('.qxport-select-visible-btn').off('click').on('click', function() {
                    $element.find('.qxport-app-chart-checkbox:visible').prop('checked', true);
                    updateSelectionSummary();
                });

                $element.find('.qxport-clear-visible-btn').off('click').on('click', function() {
                    $element.find('.qxport-app-chart-checkbox:visible').prop('checked', false);
                    updateSelectionSummary();
                });

                $element.find('.qxport-expand-all-btn').off('click').on('click', function() {
                    $element.find('.qxport-sheet-group:visible').each(function() {
                        var $group = $(this);
                        $group.find('.qxport-sheet-body').removeClass('collapsed');
                        $group.find('.qxport-sheet-toggle').text('▼');
                    });
                });

                $element.find('.qxport-collapse-all-btn').off('click').on('click', function() {
                    $element.find('.qxport-sheet-group:visible').each(function() {
                        var $group = $(this);
                        $group.find('.qxport-sheet-body').addClass('collapsed');
                        $group.find('.qxport-sheet-toggle').text('▶');
                    });
                });

                $element.find('.qxport-sheet-toggle').off('click').on('click', function() {
                    var targetSheetId = $(this).data('sheetId');
                    var $body = $element.find('.qxport-sheet-body[data-sheet-id="' + targetSheetId + '"]');
                    var isOpen = !$body.hasClass('collapsed');
                    $body.toggleClass('collapsed', isOpen);
                    $(this).text(isOpen ? '▶' : '▼');
                });

                $element.find('.qxport-sheet-select-all').off('click').on('click', function() {
                    var targetSheetId = $(this).data('sheetId');
                    $element.find('.qxport-app-chart-checkbox[data-sheet-id="' + targetSheetId + '"]:visible').prop('checked', true);
                    updateSelectionSummary();
                });

                $element.find('.qxport-sheet-clear').off('click').on('click', function() {
                    var targetSheetId = $(this).data('sheetId');
                    $element.find('.qxport-app-chart-checkbox[data-sheet-id="' + targetSheetId + '"]:visible').prop('checked', false);
                    updateSelectionSummary();
                });

                $element.find('.qxport-app-chart-checkbox').off('change').on('change', function() {
                    updateSelectionSummary();
                });
            }

            async function renderCurrentSheetTab() {
                setActiveTab('sheet');
                setRenderMode('sheet');
                updateMainExportButton();

                $title.text('Select charts to export from this sheet');
                $toolbar.hide();
                $selectionSummary.hide();

                if (!hasCurrentSheet) {
                    $objectList.html('<li>Could not identify current sheet in this view.</li>');
                    return;
                }

                $objectList.html('<li>Loading charts...</li>');

                try {
                    var sheet = await app.getObject(effectiveSheetId);
                    var sheetLayout = await sheet.getLayout();
                    var objects = await getAllExportableObjects(sheetLayout, app, excludedVisualizations);

                    $objectList.empty();

                    if (!objects.length) {
                        $objectList.append('<li>No charts found on this sheet.</li>');
                        return;
                    }

                    var promises = objects.map(async function(obj) {
                        try {
                            var model = await app.getObject(obj.qInfo.qId);
                            var objLayout = await model.getLayout();
                            var title = (objLayout.title || objLayout.qMeta && objLayout.qMeta.title || '').trim();
                            var type = objLayout.visualization || 'Chart';
                            var displayName = title || capitalizeWords(type);

                            return '<li><label><input type="checkbox" class="chart-checkbox" value="' + obj.qInfo.qId + '" checked /> 📊 ' + escapeHtml(displayName) + '</label></li>';
                        } catch (err) {
                            console.warn('Could not load title for object ' + obj.qInfo.qId, err);
                            return '<li><label><input type="checkbox" class="chart-checkbox" value="' + obj.qInfo.qId + '" checked /> 📊 Chart</label></li>';
                        }
                    });

                    var items = await Promise.all(promises);
                    $objectList.html(items.filter(Boolean).join(''));
                } catch (err) {
                    console.error('Failed to load sheet objects:', err);
                    $objectList.html('<li>Failed to load charts.</li>');
                }
            }

            async function renderAppTab() {
                setActiveTab('app');
                setRenderMode('app');
                updateMainExportButton();

                $title.text('Select charts to export from the app');
                $toolbar.show();

                $objectList.html('<li>Loading app sheets and charts...</li>');

                try {
                    if (!uiState.appLoaded) {
                        uiState.appSheets = await loadAppSelectionData(app, excludedVisualizations, $element);
                        uiState.appLoaded = true;
                    }

                    if (!uiState.appSheets.length) {
                        $objectList.html('<li>No exportable charts found in the app.</li>');
                        updateSelectionSummary();
                        return;
                    }

                    var html = uiState.appSheets.map(function(sheetGroup) {
                        var chartCountText = sheetGroup.charts.length + ' chart' + (sheetGroup.charts.length > 1 ? 's' : '');
                        var chartsHtml = sheetGroup.charts.map(function(chart) {
                            return `
                                <li class="qxport-chart-item" data-chart-title="${escapeHtmlAttr(chart.chartTitle)}">
                                    <label>
                                        <input
                                            type="checkbox"
                                            class="qxport-app-chart-checkbox"
                                            data-sheet-id="${escapeHtmlAttr(sheetGroup.sheetId)}"
                                            data-obj-id="${escapeHtmlAttr(chart.objId)}"
                                        />
                                        📊 ${escapeHtml(chart.chartTitle)}
                                    </label>
                                </li>
                            `;
                        }).join('');

                        return `
                            <li class="qxport-sheet-group" data-sheet-title="${escapeHtmlAttr(sheetGroup.sheetTitle)}">
                                <div class="qxport-sheet-header">
                                    <button type="button" class="qxport-sheet-toggle" data-sheet-id="${escapeHtmlAttr(sheetGroup.sheetId)}">▶</button>
                                    <div class="qxport-sheet-title-wrap">
                                        <span class="qxport-sheet-title">${escapeHtml(sheetGroup.sheetTitle)} (${chartCountText})</span>
                                    </div>
                                    <div class="qxport-sheet-actions">
                                        <button type="button" class="qxport-link-btn qxport-sheet-select-all" data-sheet-id="${escapeHtmlAttr(sheetGroup.sheetId)}">Select all</button>
                                        <button type="button" class="qxport-link-btn qxport-sheet-clear" data-sheet-id="${escapeHtmlAttr(sheetGroup.sheetId)}">Clear</button>
                                    </div>
                                </div>
                                <ul class="qxport-sheet-body collapsed" data-sheet-id="${escapeHtmlAttr(sheetGroup.sheetId)}">
                                    ${chartsHtml}
                                </ul>
                            </li>
                        `;
                    }).join('');

                    $objectList.html(html);

                    $searchInput.val(uiState.searchTerm || '');
                    bindAppModeHandlers();
                    applyAppFilter();
                    updateSelectionSummary();
                } catch (err) {
                    console.error('Failed to load app mode:', err);
                    $objectList.html('<li>Failed to load app charts.</li>');
                    updateSelectionSummary();
                }
            }

            async function loadAppSelectionData(appInstance, excludedVis) {
                var $status = $element.find('.export-status');
                var $progress = $element.find('.export-progress');
                var $progressFill = $element.find('.progress-fill');
                var $progressText = $element.find('.progress-text');

                $status.show().text('Scanning app sheets...').css('color', '#007acc');
                $progress.show();
                $progressFill.css('width', '0%');
                $progressText.text('0%');

                try {
                    var sheetInfos = await getAppSheets(appInstance);
                    var groups = [];

                    for (var s = 0; s < sheetInfos.length; s++) {
                        var sheetInfo = sheetInfos[s];
                        var currentSheetId = sheetInfo.qInfo && sheetInfo.qInfo.qId;

                        if (!currentSheetId) {
                            continue;
                        }

                        $status.text('Scanning sheet ' + (s + 1) + ' of ' + sheetInfos.length + '...');
                        var scanProgress = Math.round(((s + 1) / Math.max(sheetInfos.length, 1)) * 100);
                        $progressFill.css('width', scanProgress + '%');
                        $progressText.text(scanProgress + '%');

                        try {
                            var sheetModel = await appInstance.getObject(currentSheetId);
                            var sheetLayout = await sheetModel.getLayout();
                            var sheetTitle = getSheetTitle(sheetLayout, sheetInfo, s + 1);
                            var sheetObjects = await getAllExportableObjects(sheetLayout, appInstance, excludedVis);
                            var chartEntries = [];

                            for (var i = 0; i < sheetObjects.length; i++) {
                                try {
                                    var objModel = await appInstance.getObject(sheetObjects[i].qInfo.qId);
                                    var objLayout = await objModel.getLayout();
                                    var chartTitle =
                                        objLayout.title ||
                                        objLayout.qMeta && objLayout.qMeta.title ||
                                        capitalizeWords(objLayout.visualization || 'Chart');

                                    chartEntries.push({
                                        objId: sheetObjects[i].qInfo.qId,
                                        chartTitle: chartTitle
                                    });
                                } catch (objErr) {
                                    console.warn('Failed loading app chart title for ' + sheetObjects[i].qInfo.qId, objErr);
                                }
                            }

                            if (chartEntries.length) {
                                groups.push({
                                    sheetId: currentSheetId,
                                    sheetTitle: sheetTitle,
                                    charts: chartEntries
                                });
                            }
                        } catch (sheetErr) {
                            console.warn('Failed scanning sheet ' + currentSheetId, sheetErr);
                        }
                    }

                    return groups;
                } finally {
                    $progress.hide();
                    $status.hide();
                }
            }

            $tabButtons.off('click').on('click', async function() {
                var tab = $(this).data('tab');
                if (tab === 'app') {
                    await renderAppTab();
                } else {
                    await renderCurrentSheetTab();
                }
            });

            renderCurrentSheetTab();

            return qlik.Promise.resolve();
        }
    };

    async function getAllExportableObjects(sheetLayout, app, excludedVisualizations) {
        var allObjects = [];

        if (!sheetLayout || !sheetLayout.qChildList || !sheetLayout.qChildList.qItems) {
            return allObjects;
        }

        for (var i = 0; i < sheetLayout.qChildList.qItems.length; i++) {
            var obj = sheetLayout.qChildList.qItems[i];

            if (obj.qExtendsId === 'excel-export') {
                continue;
            }

            try {
                var model = await app.getObject(obj.qInfo.qId);
                var layout = await model.getLayout();

                if (excludedVisualizations.indexOf(layout.visualization) !== -1) {
                    continue;
                }

                if (layout.visualization === 'container' && typeof model.getChildInfos === 'function') {
                    var children = await model.getChildInfos();

                    for (var j = 0; j < children.length; j++) {
                        try {
                            var childModel = await app.getObject(children[j].qId);
                            var childLayout = await childModel.getLayout();

                            if (excludedVisualizations.indexOf(childLayout.visualization) !== -1) {
                                continue;
                            }

                            allObjects.push({
                                qInfo: children[j]
                            });
                        } catch (childErr) {
                            console.warn('Failed loading child object ' + children[j].qId, childErr);
                        }
                    }
                } else {
                    allObjects.push(obj);
                }
            } catch (err) {
                console.warn('Failed loading object ' + obj.qInfo.qId, err);
            }
        }

        return dedupeObjectsById(allObjects);
    }

    async function exportFromCurrentSheet(app, layout, $element) {
        var $status = $element.find('.export-status');
        var $button = $element.find('.qxport-main-export-btn');
        var $progress = $element.find('.export-progress');
        var $progressFill = $element.find('.progress-fill');
        var $progressText = $element.find('.progress-text');

        var selectedIds = $element.find('.chart-checkbox:checked').map(function() {
            return $(this).val();
        }).get();

        if (!selectedIds.length) {
            $status.show().text('❗ No charts selected.').css('color', '#dc3545');
            return;
        }

        try {
            $button.prop('disabled', true);
            $status.show().text('Preparing export...').css('color', '#007acc');
            $progress.show();
            $progressFill.css('width', '0%');
            $progressText.text('0%');

            await ensureExcelJsLoaded();

            var exportData = [];
            var maxRows = layout.maxRows || 1000;
            var failedObjects = [];

            for (var i = 0; i < selectedIds.length; i++) {
                var objId = selectedIds[i];
                var progress = Math.round(((i + 1) / selectedIds.length) * 100);

                $status.text('Processing object ' + (i + 1) + ' of ' + selectedIds.length + '...');
                $progressFill.css('width', progress + '%');
                $progressText.text(progress + '%');

                try {
                    var objModel = await app.getObject(objId);
                    var objLayout = await objModel.getLayout();
                    var data = await extractTableData(app, objModel, objLayout, maxRows);

                    var sheetName = sanitizeSheetName(
                        objLayout.title ||
                        objLayout.qMeta && objLayout.qMeta.title ||
                        objLayout.visualization ||
                        ('Object_' + (i + 1))
                    );

                    if (!data || !data.length) {
                        data = [['No data']];
                    }

                    exportData.push({
                        name: sheetName,
                        data: data
                    });
                } catch (err) {
                    console.warn('Failed to extract data from ' + objId + ':', err);
                    failedObjects.push(objId + ' - ' + (err && err.message ? err.message : 'Unknown error'));
                }

                await delay(30);
            }

            if (!exportData.length) {
                throw new Error('No data could be exported from the selected objects');
            }

            var fileName = sanitizeFileName(layout.fileName || 'QlikSense_Export');
            await exportSheetsToXlsx(exportData, fileName + '.xlsx');

            if (failedObjects.length) {
                $status.text('✅ Exported ' + exportData.length + ' sheet(s). Skipped ' + failedObjects.length + ' object(s).').css('color', '#d97706');
            } else {
                $status.text('✅ Exported ' + exportData.length + ' sheet(s) to XLSX').css('color', '#28a745');
            }

            $progress.delay(1500).fadeOut();
            $status.delay(4000).fadeOut();
        } catch (err) {
            console.error('Export failed:', err);
            $status.text('❌ Export failed: ' + err.message).css('color', '#dc3545');
            $progress.hide();
        } finally {
            $button.prop('disabled', false);
        }
    }

    async function exportSelectedFromAppMode(app, layout, $element, uiState) {
        var $status = $element.find('.export-status');
        var $button = $element.find('.qxport-main-export-btn');
        var $progress = $element.find('.export-progress');
        var $progressFill = $element.find('.progress-fill');
        var $progressText = $element.find('.progress-text');

        var selectedItems = $element.find('.qxport-app-chart-checkbox:checked').map(function() {
            return {
                sheetId: $(this).data('sheetId'),
                objId: $(this).data('objId')
            };
        }).get();

        if (!selectedItems.length) {
            $status.show().text('❗ No charts selected in app export tab.').css('color', '#dc3545');
            return;
        }

        try {
            $button.prop('disabled', true);
            $status.show().text('Preparing app export...').css('color', '#007acc');
            $progress.show();
            $progressFill.css('width', '0%');
            $progressText.text('0%');

            await ensureExcelJsLoaded();

            var exportData = [];
            var failedObjects = [];
            var maxRows = layout.maxRows || 1000;

            for (var i = 0; i < selectedItems.length; i++) {
                var item = selectedItems[i];
                var progress = Math.round(((i + 1) / selectedItems.length) * 100);

                $status.text('Exporting chart ' + (i + 1) + ' of ' + selectedItems.length + '...');
                $progressFill.css('width', progress + '%');
                $progressText.text(progress + '%');

                try {
                    var objModel = await app.getObject(item.objId);
                    var objLayout = await objModel.getLayout();
                    var data = await extractTableData(app, objModel, objLayout, maxRows);
                    var sheetGroup = findSheetGroupById(uiState.appSheets, item.sheetId);
                    var sheetTitle = sheetGroup ? sheetGroup.sheetTitle : 'Sheet';
                    var chartTitle =
                        objLayout.title ||
                        objLayout.qMeta && objLayout.qMeta.title ||
                        objLayout.visualization ||
                        ('Object_' + (i + 1));

                    var excelSheetName = sanitizeSheetName(sheetTitle + ' - ' + chartTitle);

                    if (!data || !data.length) {
                        data = [['No data']];
                    }

                    exportData.push({
                        name: excelSheetName,
                        data: data
                    });
                } catch (err) {
                    console.warn('Failed to extract selected app object ' + item.objId + ':', err);
                    failedObjects.push(item.objId + ' - ' + (err && err.message ? err.message : 'Unknown error'));
                }

                await delay(30);
            }

            if (!exportData.length) {
                throw new Error('No data could be exported from the selected app charts');
            }

            var fileName = sanitizeFileName((layout.fileName || 'QlikSense_Export') + '_App');
            await exportSheetsToXlsx(exportData, fileName + '.xlsx');

            if (failedObjects.length) {
                $status.text('✅ Exported ' + exportData.length + ' sheet(s) from app. Skipped ' + failedObjects.length + ' object(s).').css('color', '#d97706');
            } else {
                $status.text('✅ Exported ' + exportData.length + ' sheet(s) from app').css('color', '#28a745');
            }

            $progress.delay(1500).fadeOut();
            $status.delay(4000).fadeOut();
        } catch (err) {
            console.error('App export failed:', err);
            $status.text('❌ Export App failed: ' + err.message).css('color', '#dc3545');
            $progress.hide();
        } finally {
            $button.prop('disabled', false);
        }
    }

    function getAppSheets(app) {
        return new Promise(function(resolve, reject) {
            try {
                app.getList('sheet', function(reply) {
                    var items = reply &&
                        reply.qAppObjectList &&
                        reply.qAppObjectList.qItems
                        ? reply.qAppObjectList.qItems
                        : [];

                    resolve(items);
                });
            } catch (err) {
                reject(err);
            }
        });
    }

    async function extractTableData(app, objModel, layout, maxRows) {
        var hc = layout.qHyperCube;

        if (!hc) {
            return [['No data']];
        }

        if (layout.visualization === 'treemap') {
            return await extractTreemapDataViaTempCube(app, objModel, maxRows);
        }

        return await extractStandardHyperCubeData(objModel, hc, maxRows);
    }

    async function extractStandardHyperCubeData(objModel, hc, maxRows) {
        var data = [];
        var headers = [];

        (hc.qDimensionInfo || []).forEach(function(dim) {
            headers.push(dim.qFallbackTitle || 'Dim');
        });

        (hc.qMeasureInfo || []).forEach(function(meas) {
            headers.push(meas.qFallbackTitle || 'Meas');
        });

        if (!headers.length) {
            return [['No data']];
        }

        data.push(headers);

        var totalRows = Math.min(hc.qSize && hc.qSize.qcy || 0, maxRows);
        var pageSize = 1000;
        var qcx = hc.qSize && hc.qSize.qcx || headers.length;

        for (var i = 0; i < totalRows; i += pageSize) {
            var pages = await objModel.getHyperCubeData('/qHyperCubeDef', [{
                qTop: i,
                qLeft: 0,
                qWidth: qcx,
                qHeight: Math.min(pageSize, totalRows - i)
            }]);

            (pages[0] && pages[0].qMatrix || []).forEach(function(row) {
                var values = row.map(function(cell) {
                    return getCellValue(cell);
                });

                while (values.length < headers.length) {
                    values.push('');
                }

                data.push(values);
            });
        }

        return data;
    }

    async function extractTreemapDataViaTempCube(app, objModel, maxRows) {
        var props = await objModel.getProperties();
        var hcDef = props && props.qHyperCubeDef;

        if (!hcDef) {
            return [['No data']];
        }

        var qDimensions = (hcDef.qDimensions || []).map(function(dim) {
            return cloneObject(dim);
        });

        var qMeasures = (hcDef.qMeasures || []).map(function(meas) {
            return cloneObject(meas);
        });

        var headers = [];
        qDimensions.forEach(function(dim, index) {
            var title =
                dim.qDef && dim.qDef.qFieldLabels && dim.qDef.qFieldLabels[0] ||
                dim.qDef && dim.qDef.qFieldDefs && dim.qDef.qFieldDefs[0] ||
                dim.qDef && dim.qDef.qLabel ||
                'Dim ' + (index + 1);
            headers.push(title);
        });

        qMeasures.forEach(function(meas, index) {
            var title =
                meas.qDef && meas.qDef.qLabel ||
                'Measure ' + (index + 1);
            headers.push(title);
        });

        if (!headers.length) {
            return [['No data']];
        }

        var width = qDimensions.length + qMeasures.length;
        var pageSize = 1000;
        var sessionObject;
        var data = [headers];

        try {
            sessionObject = await app.model.enigmaModel.createSessionObject({
                qInfo: {
                    qType: 'qxport-temp-straight'
                },
                qHyperCubeDef: {
                    qDimensions: qDimensions,
                    qMeasures: qMeasures,
                    qInitialDataFetch: [{
                        qTop: 0,
                        qLeft: 0,
                        qWidth: Math.max(1, width),
                        qHeight: Math.min(pageSize, maxRows)
                    }],
                    qSuppressZero: false,
                    qSuppressMissing: false
                }
            });

            var tempLayout = await sessionObject.getLayout();
            var totalRows = Math.min(tempLayout.qHyperCube && tempLayout.qHyperCube.qSize && tempLayout.qHyperCube.qSize.qcy || 0, maxRows);

            for (var top = 0; top < totalRows; top += pageSize) {
                var pages = await sessionObject.getHyperCubeData('/qHyperCubeDef', [{
                    qTop: top,
                    qLeft: 0,
                    qWidth: Math.max(1, width),
                    qHeight: Math.min(pageSize, totalRows - top)
                }]);

                (pages[0] && pages[0].qMatrix || []).forEach(function(row) {
                    var values = row.map(function(cell) {
                        return getCellValue(cell);
                    });

                    while (values.length < headers.length) {
                        values.push('');
                    }

                    data.push(values);
                });
            }

            return data;
        } finally {
            if (sessionObject) {
                try {
                    await app.model.enigmaModel.destroySessionObject(sessionObject.id);
                } catch (destroyErr) {
                    console.warn('Failed to destroy temp session object', destroyErr);
                }
            }
        }
    }

    function getCellValue(cell) {
        if (!cell) return '';

        if (cell.qText !== undefined && cell.qText !== '') {
            return cell.qText;
        }

        if (cell.qNum !== undefined && cell.qNum !== null && !isNaN(cell.qNum)) {
            return cell.qNum;
        }

        return '';
    }

    function cloneObject(obj) {
        return JSON.parse(JSON.stringify(obj || {}));
    }

    function dedupeObjectsById(objects) {
        var seen = {};
        var result = [];

        for (var i = 0; i < objects.length; i++) {
            var id = objects[i] && objects[i].qInfo && objects[i].qInfo.qId;
            if (!id || seen[id]) {
                continue;
            }
            seen[id] = true;
            result.push(objects[i]);
        }

        return result;
    }

    function getSheetTitle(sheetLayout, sheetInfo, index) {
        return (
            sheetLayout && sheetLayout.title ||
            sheetLayout && sheetLayout.qMeta && sheetLayout.qMeta.title ||
            sheetInfo && sheetInfo.qMeta && sheetInfo.qMeta.title ||
            sheetInfo && sheetInfo.qData && sheetInfo.qData.title ||
            sheetInfo && sheetInfo.qData && sheetInfo.qData.cells && sheetInfo.qData.cells[0] && sheetInfo.qData.cells[0].name ||
            ('Sheet ' + index)
        );
    }

    function findSheetGroupById(groups, sheetId) {
        for (var i = 0; i < groups.length; i++) {
            if (groups[i].sheetId === sheetId) {
                return groups[i];
            }
        }
        return null;
    }

    async function ensureExcelJsLoaded() {
		if (window.ExcelJS) {
			return window.ExcelJS;
		}

		if (excelJsLoadPromise) {
			return excelJsLoadPromise;
		}

		excelJsLoadPromise = new Promise(function(resolve, reject) {
			requirejs(['./exceljs.min'], function(ExcelJS) {
				if (ExcelJS) {
					window.ExcelJS = ExcelJS;
					resolve(ExcelJS);
					return;
				}

				if (window.ExcelJS) {
					resolve(window.ExcelJS);
					return;
				}

				reject(new Error('ExcelJS loaded but no module/global export was found'));
			}, function(err) {
				reject(new Error('Failed to load ExcelJS via RequireJS: ' + (err && err.message ? err.message : err)));
			});
		});

		return excelJsLoadPromise;
	}

    async function exportSheetsToXlsx(sheets, filename) {
        if (!window.ExcelJS) {
            throw new Error('ExcelJS is not available');
        }

        var ExcelJS = await ensureExcelJsLoaded();
		var workbook = new ExcelJS.Workbook();
        workbook.creator = 'QXport';
        workbook.lastModifiedBy = 'QXport';
        workbook.created = new Date();
        workbook.modified = new Date();

        var usedNames = {};

        sheets.forEach(function(sheet) {
            var worksheetName = getUniqueWorksheetName(sheet.name || 'Sheet', usedNames);
            var worksheet = workbook.addWorksheet(worksheetName);

            var rows = Array.isArray(sheet.data) && sheet.data.length ? sheet.data : [['No data']];
            rows.forEach(function(row) {
                worksheet.addRow(Array.isArray(row) ? row : [row]);
            });

            styleWorksheet(worksheet, rows);
        });

        var buffer = await workbook.xlsx.writeBuffer();
        downloadArrayBufferAsXlsx(buffer, filename);
    }

    function styleWorksheet(worksheet, rows) {
        if (!rows || !rows.length) {
            return;
        }

        var headerRow = worksheet.getRow(1);
        headerRow.font = {
            bold: true
        };
        headerRow.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            wrapText: true
        };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'FFD9EAF7'
            }
        };
        headerRow.border = {
            bottom: {
                style: 'thin',
                color: {
                    argb: 'FFB7C9D6'
                }
            }
        };

        worksheet.views = [{
            state: 'frozen',
            ySplit: 1
        }];

        var headerLength = Array.isArray(rows[0]) ? rows[0].length : 0;
        if (headerLength > 0) {
            worksheet.autoFilter = {
                from: {
                    row: 1,
                    column: 1
                },
                to: {
                    row: 1,
                    column: headerLength
                }
            };
        }

        worksheet.eachRow(function(row, rowNumber) {
            row.alignment = {
                vertical: 'top',
                wrapText: true
            };

            if (rowNumber === 1) {
                row.height = 22;
            }
        });

        autoFitWorksheetColumns(worksheet, rows);
    }

    function autoFitWorksheetColumns(worksheet, rows) {
        var maxColumnCount = 0;

        rows.forEach(function(row) {
            if (Array.isArray(row) && row.length > maxColumnCount) {
                maxColumnCount = row.length;
            }
        });

        for (var colIndex = 1; colIndex <= maxColumnCount; colIndex++) {
            var maxLength = 10;

            rows.forEach(function(row) {
                if (!Array.isArray(row)) {
                    return;
                }

                var value = row[colIndex - 1];
                var text = value === null || value === undefined ? '' : String(value);
                var lines = text.split(/\r\n|\r|\n/);

                lines.forEach(function(line) {
                    if (line.length > maxLength) {
                        maxLength = line.length;
                    }
                });
            });

            worksheet.getColumn(colIndex).width = Math.min(Math.max(maxLength + 2, 10), 60);
        }
    }

    function getUniqueWorksheetName(name, usedNames) {
        var baseName = sanitizeSheetName(name || 'Sheet') || 'Sheet';
        var finalName = baseName;
        var counter = 1;

        while (usedNames[finalName]) {
            var suffix = '_' + counter;
            finalName = baseName.substring(0, Math.max(1, 31 - suffix.length)) + suffix;
            counter++;
        }

        usedNames[finalName] = true;
        return finalName;
    }

    function downloadArrayBufferAsXlsx(buffer, filename) {
        var blob = new Blob(
            [buffer],
            {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        );

        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);

        setTimeout(function() {
            URL.revokeObjectURL(url);
        }, 0);
    }

    function escapeHtml(str) {
        return String(str || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function escapeHtmlAttr(str) {
        return escapeHtml(str).replace(/`/g, '&#96;');
    }

    function sanitizeSheetName(name) {
        return String(name || 'Sheet')
            .replace(/[\\\/\?\*\[\]:;|=,<>]/g, '_')
            .substring(0, 31) || 'Sheet';
    }

    function sanitizeFileName(name) {
        return String(name || 'QlikSense_Export')
            .replace(/[\\\/:*?"<>|]+/g, '_')
            .trim() || 'QlikSense_Export';
    }

    function delay(ms) {
        return new Promise(function(resolve) {
            setTimeout(resolve, ms);
        });
    }
});
