import { enableRipple } from '@syncfusion/ej2-base';
import { fixInitialScaleForTile } from '@syncfusion/ej2-maps/src/maps';
enableRipple(true);

import { SheetModel, Spreadsheet } from '@syncfusion/ej2-spreadsheet';
import * as dataSource from './default-data.json';

/**
 * Default Spreadsheet sample
 */
let documentStartIndex = 0;
let documentStartRow = documentStartIndex + 1;

let headerStartIndex = documentStartIndex + 2;
let headerStartRow = headerStartIndex + 1;
let headerSize = 2;

let startDataIndex = headerStartIndex + headerSize;
let startDataRow = startDataIndex + 1;
let dataSize = (dataSource as any).defaultData.length;

//Initialize Spreadsheet component
let spreadsheet: Spreadsheet = new Spreadsheet({
  sheets: [
    {
      name: 'GeekHub [V.2.0] Invoice',
      ranges: [
        {
          startCell: 'A' + startDataRow,
          dataSource: (dataSource as any).defaultData,
          showFieldAsHeader: false,
        },
        {
          startCell: 'A' + 8,
          dataSource: {
            dataSource: {
              json: dataSource as any,
              data: dataSource as JSON,
              accept: true,
            },
          },
          showFieldAsHeader: false,
        },
      ],
      rows: [
        {
          index: documentStartIndex,
          height: 30,
          cells: [
            {
              value: (dataSource as any).projectName,
              style: {
                backgroundColor: '#e56590',
                color: '#fff',
                fontWeight: 'bold',
                textAlign: 'center',
                verticalAlign: 'middle',
              },
            },
          ],
        },
        {
          index: headerStartIndex,
          cells: [
            {
              index: 0,
              value: 'Position',
              rowSpan: headerSize,
            },
            {
              index: 1,
              value: 'Employee',
              rowSpan: headerSize,
            },
            {
              index: 2,
              value: 'Rate',
              rowSpan: headerSize,
            },
            {
              index: 3,
              value: 'Planned',
              rowSpan: headerSize,
            },
            {
              index: 4,
              value: 'Reported',
              rowSpan: headerSize,
            },
            {
              index: 5,
              value: 'Total Cost',
              rowSpan: headerSize,
            },
            {
              index: 6,
              value: 'Billable',
              colSpan: 3,
            },
          ],
        },
        {
          index: headerStartIndex + 1,
          cells: [
            {
              index: 6,
              value: 'Regular',
            },
            {
              index: 7,
              value: 'Overtimes',
            },
            {
              index: 8,
              value: 'Total',
            },
          ],
        },
        {
          index: startDataIndex + dataSize,
          format: {
            style: {
              fontWeight: 'bold',
              textAlign: 'right',
            },
          },
          cells: [
            {
              index: 4,
              value: 'Total Amount:',
              style: {
                fontWeight: 'bold',
                textAlign: 'right',
              },
            },
            {
              formula:
                '=SUM(F' +
                startDataRow +
                ':F' +
                (startDataRow + dataSize - 1) +
                ')',
              style: {
                fontWeight: 'bold',
              },
            },
          ],
        },
      ],
      columns: [
        {
          width: 180,
        },
        {
          width: 130,
        },
        {
          width: 130,
        },
        {
          width: 180,
        },
        {
          width: 130,
        },
        {
          width: 120,
        },
      ],
    },
  ],
  saveUrl:
    'https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/save',
  created: (): void => {
    // Formatting cells dynamically using cellFormat method
    spreadsheet.merge('A' + documentStartRow + ':F' + documentStartRow);

    //header
    spreadsheet.cellFormat(
      { fontWeight: 'bold', textAlign: 'center' },
      'A' + headerStartRow + ':I' + (headerStartRow + 1)
    );

    //Format cost column
    spreadsheet.numberFormat('$#,##0.00', 'F2:F31');

    //set range without numbers like SUM(G#,H#)
    for (let i = startDataRow; i < startDataRow + dataSize; i++) {
      spreadsheet.updateCell(
        {
          formula: '=SUM(G' + i + ',H' + i + ')',
        },
        'I' + i
      );
    }
  },
});

let sheetModel =
  //Render initialized Spreadsheet component
  spreadsheet.appendTo('#spreadsheet');
