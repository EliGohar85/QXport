define([
    'qlik',
    'jquery',
    'text!./qxport.css'
], function(qlik, $, cssContent) {
    'use strict';

    $('<style>').html(cssContent).appendTo('head');

    function capitalizeWords(str) {
        return String(str)
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/[_\-]/g, ' ')
            .replace(/\b\w/g, c => c.toUpperCase());
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
                            label: 'Version: 1.0.0',
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
            const app = qlik.currApp();
            const buttonText = layout.buttonText || 'Export to Excel';

            $element.empty();

            const $container = $(`
                <div class="export-excel-container">
                    <div class="export-list-title">Available charts to be exported:</div>
                    <div class="chart-list-container">
                        <ul class="export-object-list"><li>Loading charts...</li></ul>
                    </div>
                    <button class="export-excel-btn" type="button">
                        <span class="export-icon">📊</span>
                        ${buttonText}
                    </button>
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

            const $objectList = $element.find('.export-object-list');
            const sheetId = qlik.navigation.getCurrentSheetId().sheetId;

            const excludedVisualizations = [
                'qxport',
                'action-button',
                'navigation-button',
                'filterpane',
                'textbox',
                'button',
                'qlik-date-picker',
                'tcmenu',
                'variable-input'
            ];

            app.getObject(sheetId).then(sheet => {
                sheet.getLayout().then(async layout => {
                    const objects = await getAllExportableObjects(layout, app);

                    $objectList.empty();

                    if (!objects.length) {
                        $objectList.append('<li>No charts found on this sheet.</li>');
                        return;
                    }

                    const promises = objects.map(async (obj) => {
                        try {
                            const model = await app.getObject(obj.qInfo.qId);
                            const layout = await model.getLayout();
                            const title = layout.title?.trim();
                            const type = layout.visualization || 'Chart';
                            const displayName = title || capitalizeWords(type);

                            if (excludedVisualizations.includes(layout.visualization)) return '';
                            if (layout.visualization === 'excel-export' || displayName === 'QXport') return '';

                            return `<li><label><input type="checkbox" class="chart-checkbox" value="${obj.qInfo.qId}" checked /> 📊 ${displayName}</label></li>`;
                        } catch (err) {
                            console.warn(`Could not load title for object ${obj.qInfo.qId}`, err);
                            return `<li><label><input type="checkbox" class="chart-checkbox" value="${obj.qInfo.qId}" checked /> 📊 Chart</label></li>`;
                        }
                    });

                    Promise.all(promises).then(items => {
                        $objectList.html(items.filter(Boolean).join(''));
                    });
                });
            });

            $element.find('.export-excel-btn').on('click', function() {
                exportFromSheet(app, layout, $element);
            });

            return qlik.Promise.resolve();
        }
    };

    async function getAllExportableObjects(sheetLayout, app) {
        const allObjects = [];

        for (const obj of sheetLayout.qChildList.qItems) {
            if (obj.qExtendsId === 'excel-export') continue;

            const model = await app.getObject(obj.qInfo.qId);
            const layout = await model.getLayout();

            if (layout.visualization === 'container' && model.getChildInfos) {
                const children = await model.getChildInfos();
                for (const child of children) {
                    allObjects.push({
                        qInfo: child
                    });
                }
            } else {
                allObjects.push(obj);
            }
        }

        return allObjects;
    }

    async function exportFromSheet(app, layout, $element) {
        const $status = $element.find('.export-status');
        const $button = $element.find('.export-excel-btn');
        const $progress = $element.find('.export-progress');
        const $progressFill = $element.find('.progress-fill');
        const $progressText = $element.find('.progress-text');
        const selectedIds = $element.find('.chart-checkbox:checked').map(function() {
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

            const exportData = [];
            const maxRows = layout.maxRows || 1000;

            for (let i = 0; i < selectedIds.length; i++) {
                const objId = selectedIds[i];
                const progress = Math.round(((i + 1) / selectedIds.length) * 100);

                $status.text(`Processing object ${i + 1} of ${selectedIds.length}...`);
                $progressFill.css('width', progress + '%');
                $progressText.text(progress + '%');

                try {
                    const objModel = await app.getObject(objId);
                    const objLayout = await objModel.getLayout();
                    const data = await extractTableData(objModel, objLayout, maxRows);

                    if (data && data.length > 0) {
                        exportData.push({
                            name: sanitizeSheetName(objLayout.title || objLayout.visualization || `Object_${i + 1}`),
                            data: data
                        });
                    }
                } catch (err) {
                    console.warn(`Failed to extract data from ${objId}:`, err);
                }

                await new Promise(r => setTimeout(r, 30));
            }

            if (!exportData.length) throw new Error('No data to export');

            const fileName = layout.fileName || 'QlikSense_Export';
            const xml = generateExcelXML(exportData);
            downloadXMLAsXLS(xml, fileName + '.xls');
            $status.text(`✅ Exported ${exportData.length} sheet(s) to Excel`).css('color', '#28a745');

            $progress.delay(1500).fadeOut();
            $status.delay(3000).fadeOut();
        } catch (err) {
            console.error('Export failed:', err);
            $status.text(`❌ Export failed: ${err.message}`).css('color', '#dc3545');
            $progress.hide();
        } finally {
            $button.prop('disabled', false);
        }
    }

    async function extractTableData(objModel, layout, maxRows) {
        const data = [];
        const hc = layout.qHyperCube;
        if (!hc) return data;

        const headers = [];
        hc.qDimensionInfo.forEach(dim => headers.push(dim.qFallbackTitle || 'Dim'));
        hc.qMeasureInfo.forEach(meas => headers.push(meas.qFallbackTitle || 'Meas'));
        data.push(headers);

        const totalRows = Math.min(hc.qSize.qcy, maxRows);
        const pageSize = 1000;

        for (let i = 0; i < totalRows; i += pageSize) {
            const pages = await objModel.getHyperCubeData('/qHyperCubeDef', [{
                qTop: i,
                qLeft: 0,
                qWidth: hc.qSize.qcx,
                qHeight: Math.min(pageSize, totalRows - i)
            }]);

            pages[0]?.qMatrix.forEach(row => {
                const values = row.map(cell => cell.qText ?? cell.qNum ?? '');
                if (values.some(v => v !== '')) {
                    data.push(values);
                }
            });
        }

        return data;
    }

    function generateExcelXML(sheets) {
        const xmlHeader = `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">`;

        const styles = `
<Styles>
  <Style ss:ID="Header">
    <Font ss:Bold="1"/>
    <Interior ss:Color="#CFE2F3" ss:Pattern="Solid"/>
  </Style>
</Styles>`;

        const nameCount = {};
        const safeSheets = sheets.map(sheet => {
            let baseName = escapeXml(sheet.name.substring(0, 31)) || 'Sheet';
            if (!nameCount[baseName]) {
                nameCount[baseName] = 1;
            } else {
                nameCount[baseName]++;
                baseName = `${baseName}_${nameCount[baseName]}`;
            }
            return {
                ...sheet,
                name: baseName
            };
        });

        const sheetsXML = safeSheets.map(sheet => {
            const rows = sheet.data.map((row, rowIndex) => {
                const cells = row.map(val =>
                    `<Cell${rowIndex === 0 ? ' ss:StyleID="Header"' : ''}><Data ss:Type="String">${escapeXml(val)}</Data></Cell>`
                ).join('');
                return `<Row>${cells}</Row>`;
            }).join('');

            return `<Worksheet ss:Name="${sheet.name}">
<Table>${rows}</Table>
</Worksheet>`;
        }).join('');

        return `${xmlHeader}${styles}${sheetsXML}</Workbook>`;
    }

    function downloadXMLAsXLS(content, filename) {
        const blob = new Blob([content], {
            type: 'application/vnd.ms-excel'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function escapeXml(str) {
        return String(str || '')
            .replace(/[^\x09\x0A\x0D\x20-\x7E\xA0-\uFFFF]/g, '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function sanitizeSheetName(name) {
        return name.replace(/[\\\/\?\*\[\]:;|=,<>]/g, '_').substring(0, 31) || 'Sheet';
    }
});