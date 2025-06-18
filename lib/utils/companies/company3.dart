import 'package:flutter/services.dart' show rootBundle;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:flutter/foundation.dart' show debugPrint;
import '../company_config.dart';

class Company3Config implements CompanyConfig {
@override
String get jsonPath => 'assets/company3.json';

@override
List<String> get columnHeaders => [
'Test',
'Test Method',
'Pc No. 24045849',
'', // Empty column
'', // Empty column
'', // Empty column
'', // Empty column
'Standard'
];

@override
Map<String, String> get columnMapping => {
'test': 'Test',
'method': 'Test Method',
'pc_no': 'Pc No. 24045849',
'empty1': '', // Placeholder for empty columns
'empty2': '',
'empty3': '',
'empty4': '',
'standard': 'Standard'
};

@override
Future<void> initializeSheet(xlsio.Worksheet sheet, xlsio.Workbook workbook) async {
sheet.showGridlines = false;

try {
final imageData = await rootBundle.load('assets/logo.png');
final List<int> imageBytes = imageData.buffer.asUint8List();
final xlsio.Picture picture = sheet.pictures.addStream(1, 1, imageBytes);
picture.height = 130;
picture.width = 130;
picture.row = 2;
picture.column = 1;
// picture.columnOffsetPx = 40;
// picture.rowOffsetPx = 50;
} catch (e) {
debugPrint('Error loading logo for Company 3: $e');
}
}

@override
void configureSheet(
xlsio.Worksheet sheet,
Map<String, String> reportDetails,
List<Map<String, dynamic>> testData,
xlsio.Workbook workbook,
) {
// --- Set Column Widths ---
sheet.getRangeByIndex(1, 1).columnWidth = 23.56;
sheet.getRangeByIndex(1, 2).columnWidth = 23.70;
sheet.getRangeByIndex(1, 3).columnWidth = 8.33;
sheet.getRangeByIndex(1, 4).columnWidth = 5.67;
sheet.getRangeByIndex(1, 5).columnWidth = 5.67;
sheet.getRangeByIndex(1, 6).columnWidth = 5.67;
sheet.getRangeByIndex(1, 7).columnWidth = 5.67;
sheet.getRangeByIndex(1, 8).columnWidth = 18.33;

// --- Set Row Heights ---
sheet.getRangeByIndex(1, 1).rowHeight = 1.80; // Row 1
sheet.getRangeByIndex(2, 1).rowHeight = 102.60; // Row 2 (Company Header)
sheet.getRangeByIndex(29, 2).rowHeight = 30.00; // Row 29
sheet.setRowHeightInPixels(3, 30); // Row 3 (TEST REPORT)

// --- Define Styles ---
final headerStyle = workbook.styles.add('headerStyle');
headerStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..bold = true
..wrapText = true;

final labelStyle = workbook.styles.add('labelStyle');
labelStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center
..bold = true;

final valueStyle = workbook.styles.add('valueStyle');
valueStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center;

final valueStyle1 = workbook.styles.add('valueStyle1');
valueStyle1
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center
..bold = true;

final tableHeaderStyle = workbook.styles.add('tableHeaderStyle');
tableHeaderStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..bold = true
..wrapText = true
..borders.all.lineStyle = xlsio.LineStyle.thin
..borders.all.color = '#000000';

final item = workbook.styles.add('item');
item
..bold = true
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.top
..wrapText = true;

final dataStyle = workbook.styles.add('dataStyle');
dataStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..wrapText = true
..borders.all.lineStyle = xlsio.LineStyle.thin
..borders.all.color = '#000000';

final companyTitleStyle = workbook.styles.add('companyTitleStyle');
companyTitleStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 13
..bold = true
..fontName = 'Times New Roman'
..wrapText = true;

final testNameStyle = workbook.styles.add('testNameStyle');
testNameStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..bold = true
..wrapText = true
..borders.all.lineStyle = xlsio.LineStyle.thin
..borders.all.color = '#000000';

// --- Company Header ---
sheet.getRangeByName('A2:H2').merge();
sheet.getRangeByName('A2').setText(
'''KUSUMGAR LIMITED       
An ISO 9001:2015 Certified Company
House of Synthetic Textile
101/102, Manjushree, V.M.Road, Corner of N.S.Road No. 5
JVPD Scheme, Vile Parle (West), Mumbai 400056.
Tel. No. 2618 4341/2618 4350 Fax No. 26115651''');
sheet.getRangeByName('A2:H2').cellStyle = companyTitleStyle;

// --- Test Report Header ---
sheet.getRangeByName('A3:H3').merge();
sheet.getRangeByName('A3').setText('TEST REPORT');
sheet.getRangeByName('A3:H3').cellStyle = headerStyle;

// Report Details
sheet.getRangeByName('A4').setText('Report No.');
sheet.getRangeByName('A4').cellStyle = labelStyle;
sheet.getRangeByName('B4').setText(reportDetails['report_no'] ?? 'Q250001330');
sheet.getRangeByName('B4').cellStyle = valueStyle1;

sheet.getRangeByName('C4').setText('Date');
sheet.getRangeByName('C4').cellStyle = labelStyle;
sheet.getRangeByName('D4').setText(reportDetails['date'] ?? '21-04-2025');
sheet.getRangeByName('D4').cellStyle = valueStyle1;

sheet.getRangeByName('A5:A15').merge();
sheet.getRangeByName('A5').setText('Item');
sheet.getRangeByName('A5').cellStyle = item;

sheet.getRangeByName('B5:B15').merge();
sheet.getRangeByName('B5').setText(reportDetails['item'] ?? '166452\nCLOTH,NYL,65",\nFG 24165\nINTERNATIONAL-\nP44378 T4-NB\n');
sheet.getRangeByName('B5').cellStyle = item;

sheet.getRangeByName('C5:C15').merge();
sheet.getRangeByName('C5').setText('Customer');
sheet.getRangeByName('C5').cellStyle = item;

sheet.getRangeByName('D5:H5').merge();
sheet.getRangeByName('D5').setText(reportDetails['customer'] ?? 'Airborne Systems NA of CA Inc.');
sheet.getRangeByName('D5').cellStyle = valueStyle1;

// Lot Numbers
int lotRow = 6;
final lots = reportDetails['lots']?.split('\n') ?? [
'24P32725 (391.52 Yard)',
'24P33570 (195.76 Yard)',
'24P33937 (2,090.98 Yard)',
'24P35293 (317.15 Yard)',
'24P41229 (1,735.56 Yard)',
'24P41230 (1,863.51 Yard)',
'24P41231 (2,063.65 Yard)',
'24P41764 (2,061.45 Yard)',
'24P41765 (2,230.97 Yard)',
'24P41768 (2,035.21 Yard)',
];
for (var lot in lots) {
sheet.getRangeByName('D$lotRow').setText('Lot:');
sheet.getRangeByName('D$lotRow').cellStyle = valueStyle1;
sheet.getRangeByName('E$lotRow:H$lotRow').merge();
sheet.getRangeByName('E$lotRow').setText(lot);
sheet.getRangeByName('E$lotRow').cellStyle = valueStyle1;
sheet.setRowHeightInPixels(lotRow, 20);
lotRow++;
}

// Additional Details
sheet.getRangeByName('A$lotRow').setText('Q. No.');
sheet.getRangeByName('A$lotRow').cellStyle = labelStyle;
sheet.getRangeByName('B$lotRow').setText(reportDetails['q_no'] ?? '4201 (GN0088)');
sheet.getRangeByName('B$lotRow').cellStyle = valueStyle1;

sheet.getRangeByName('C16:E16').merge();
sheet.getRangeByName('C16').setText(reportDetails['qty'] ?? 'Qty.: 14,985.73 Yards');
sheet.getRangeByName('C16').cellStyle = item;

sheet.getRangeByName('F16:H16').merge();
sheet.getRangeByName('F16').setText(reportDetails['rolls'] ?? 'Rolls: 66');
sheet.getRangeByName('F16').cellStyle = item;

sheet.getRangeByName('A17').setText('Width');
sheet.getRangeByName('A17').cellStyle = item;
sheet.getRangeByName('B17').setText(reportDetails['width'] ?? '165.0 CMS (65.0”)');
sheet.getRangeByName('B17').cellStyle = item;

sheet.getRangeByName('C17:E17').merge();
sheet.getRangeByName('C17').setText(reportDetails['invoice_no'] ?? 'Invoice No. ES25260039');
sheet.getRangeByName('C17').cellStyle = item;

sheet.getRangeByName('F17:H17').merge();
sheet.getRangeByName('F17').setText(reportDetails['invoice_date'] ?? '21-04-2025');
sheet.getRangeByName('F17').cellStyle = item;

// Cell Merges (as per original code)
sheet.getRangeByName('A18:H18').merge();
sheet.getRangeByName('C19:G19').merge();
sheet.getRangeByName('C20:G20').merge();
sheet.getRangeByName('C21:G21').merge();
sheet.getRangeByName('C22:G22').merge();
sheet.getRangeByName('C23:G23').merge();
sheet.getRangeByName('C24:G24').merge();
sheet.getRangeByName('C25:G25').merge();
sheet.getRangeByName('C26:G26').merge();
sheet.getRangeByName('C27:G27').merge();
sheet.getRangeByName('C28:G28').merge();
sheet.getRangeByName('D4:H4').merge();
sheet.getRangeByName('D31:G31').merge();
sheet.getRangeByName('D33:G33').merge();
sheet.getRangeByName('D35:G35').merge();
sheet.getRangeByName('D37:G37').merge();
sheet.getRangeByName('D39:G39').merge();
sheet.getRangeByName('D41:G41').merge();
sheet.getRangeByName('D43:G43').merge();
sheet.getRangeByName('D45:G45').merge();
sheet.getRangeByName('D47:G47').merge();
sheet.getRangeByName('D49:G49').merge();
sheet.getRangeByName('D51:G51').merge();
sheet.getRangeByName('D53:G53').merge();
sheet.getRangeByName('D55:G55').merge();
sheet.getRangeByName('D57:G57').merge();
sheet.getRangeByName('C60:G60').merge();
sheet.getRangeByName('C61:G61').merge();
sheet.getRangeByName('C62:D62').merge();
sheet.getRangeByName('E62:G62').merge();
sheet.getRangeByName('D63:G63').merge();
sheet.getRangeByName('C64:G64').merge();
sheet.getRangeByName('D65:E65').merge();
sheet.getRangeByName('F65:G65').merge();
sheet.getRangeByName('D66:G66').merge();
sheet.getRangeByName('D59:G59').merge();
sheet.getRangeByName('C29:G29').merge();
sheet.getRangeByName('A30:A31').merge();
sheet.getRangeByName('A32:A33').merge();
sheet.getRangeByName('A34:A35').merge();
sheet.getRangeByName('A36:A37').merge();
sheet.getRangeByName('A38:A39').merge();
sheet.getRangeByName('A40:A41').merge();
sheet.getRangeByName('A42:A43').merge();
sheet.getRangeByName('A44:A45').merge();
sheet.getRangeByName('A46:A47').merge();
sheet.getRangeByName('A48:A49').merge();
sheet.getRangeByName('A50:A51').merge();
sheet.getRangeByName('A52:A53').merge();
sheet.getRangeByName('A54:A55').merge();
sheet.getRangeByName('A56:A57').merge();
sheet.getRangeByName('A58:A59').merge();
sheet.getRangeByName('A65:A66').merge();
sheet.getRangeByName('A62:A63').merge();
sheet.getRangeByName('B23:B25').merge();
sheet.getRangeByName('B26:B28').merge();
sheet.getRangeByName('B30:B33').merge();
sheet.getRangeByName('B34:B37').merge();
sheet.getRangeByName('B38:B39').merge();
sheet.getRangeByName('B40:B41').merge();
sheet.getRangeByName('B44:B45').merge();
sheet.getRangeByName('B46:B49').merge();
sheet.getRangeByName('B50:B53').merge();
sheet.getRangeByName('B54:B55').merge();
sheet.getRangeByName('B56:B59').merge();
sheet.getRangeByName('B62:B63').merge();
sheet.getRangeByName('B65:B66').merge();
sheet.getRangeByName('H30:H33').merge();
sheet.getRangeByName('H34:H37').merge();
sheet.getRangeByName('H38:H39').merge();
sheet.getRangeByName('H42:H45').merge();
sheet.getRangeByName('H46:H49').merge();
sheet.getRangeByName('H50:H53').merge();
sheet.getRangeByName('H54:H55').merge();
sheet.getRangeByName('H56:H57').merge();
sheet.getRangeByName('H58:H59').merge();

// --- Test Results Table ---
int tableStartRow = lotRow + 3;
sheet.getRangeByName('A$tableStartRow').setText(columnHeaders[0]);
sheet.getRangeByName('B$tableStartRow').setText(columnHeaders[1]);
sheet.getRangeByName('C$tableStartRow').setText('Pc No. 24045849');
sheet.getRangeByName('D$tableStartRow').setText(columnHeaders[3]);
sheet.getRangeByName('E$tableStartRow').setText(columnHeaders[4]);
sheet.getRangeByName('F$tableStartRow').setText(columnHeaders[5]);
sheet.getRangeByName('G$tableStartRow').setText(columnHeaders[6]);
sheet.getRangeByName('H$tableStartRow').setText(columnHeaders[7]);
sheet.getRangeByName('A$tableStartRow:H$tableStartRow').cellStyle = tableHeaderStyle;
sheet.setRowHeightInPixels(tableStartRow, 40);
tableStartRow++;

// Table Data
for (var test in testData) {
sheet.getRangeByName('A$tableStartRow').setText(test['test']?.toString() ?? '');
sheet.getRangeByName('B$tableStartRow').setText(test['method']?.toString() ?? '');
sheet.getRangeByName('C$tableStartRow').setText(test['pc_no']?.toString() ?? '');
sheet.getRangeByName('D$tableStartRow').setText(test['empty1']?.toString() ?? '');
sheet.getRangeByName('E$tableStartRow').setText(test['empty2']?.toString() ?? '');
sheet.getRangeByName('F$tableStartRow').setText(test['empty3']?.toString() ?? '');
sheet.getRangeByName('G$tableStartRow').setText(test['empty4']?.toString() ?? '');
sheet.getRangeByName('H$tableStartRow').setText(test['standard']?.toString() ?? '');

sheet.getRangeByName('A$tableStartRow').cellStyle = testNameStyle;
sheet.getRangeByName('B$tableStartRow:H$tableStartRow').cellStyle = dataStyle;

if ([20, 29, 64].contains(tableStartRow)) {
sheet.setRowHeightInPixels(tableStartRow, 50);
} else {
sheet.setRowHeightInPixels(tableStartRow, 25);
}
tableStartRow++;
}

// --- Footer ---
final tableRange = sheet.getRangeByName('A2:H66');
final borderStyle = tableRange.cellStyle.borders;
borderStyle.all.lineStyle = xlsio.LineStyle.thin;
borderStyle.all.color = '#000000';
}
}
