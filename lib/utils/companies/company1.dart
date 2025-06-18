
import 'dart:ui';
import 'package:flutter/services.dart' show rootBundle;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import '../company_config.dart';

class Company1Config implements CompanyConfig {
@override
String get jsonPath => 'assets/company1.json';

@override
List<String> get columnHeaders => [
'Test',
'Test Method No.',
'Result',
'Standard',
'Minimum',
'Maximum',
'Remarks'
];

@override
Map<String, String> get columnMapping => {
'test': 'Test',
'method_no': 'Test Method No.',
'result': 'Result',
'standard': 'Standard',
'minimum': 'Minimum',
'maximum': 'Maximum',
'remarks': 'Remarks',
};

@override
Future<void> initializeSheet(xlsio.Worksheet sheet, xlsio.Workbook workbook) async {
sheet.showGridlines = false;

// Add logo
try {
final imageData = await rootBundle.load('assets/company1/logo.png');
final List<int> imageBytes = imageData.buffer.asUint8List();
final xlsio.Picture picture = sheet.pictures.addStream(1, 1, imageBytes);
picture.height = 130;
picture.width = 130;
picture.row = 2;
picture.column = 1;
// picture.columnOffset = 40;
// picture.rowOffset = 50;
} catch (e) {
print('Error loading logo for Company 1: $e');
}
}

@override
void configureSheet(
xlsio.Worksheet sheet,
Map<String, String> reportDetails,
List<Map<String, dynamic>> testData,
xlsio.Workbook workbook,
) {
sheet.getRangeByIndex(1, 1).columnWidth = 2.56;
sheet.getRangeByIndex(1, 2).columnWidth = 17.00;
sheet.getRangeByIndex(1, 3).columnWidth = 8.33;
sheet.getRangeByIndex(1, 4).columnWidth = 5.67;
sheet.getRangeByIndex(1, 5).columnWidth = 5.67;
sheet.getRangeByIndex(1, 6).columnWidth = 5.67;
sheet.getRangeByIndex(1, 7).columnWidth = 5.67;
sheet.getRangeByIndex(1, 8).columnWidth = 18.33;
sheet.getRangeByIndex(1, 9).columnWidth = 27.57;

sheet.getRangeByIndex(1, 1).rowHeight = 1.80;
sheet.getRangeByIndex(2, 1).rowHeight = 102.60;

final companyTitleStyle = workbook.styles.add('companyTitleStyle');
companyTitleStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 13
..bold = true
..fontName = 'Arial'
..wrapText = true;

final companyCertStyle = workbook.styles.add('companyCertStyle');
companyCertStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..bold = true;

final labelStyle = workbook.styles.add('labelStyle');
labelStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center
..bold = true;

final valueStyle = workbook.styles.add('valueStyle');
valueStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center;

final yellowValueStyle = workbook.styles.add('yellowValueStyle');
yellowValueStyle
..hAlign = xlsio.HAlignType.left
..vAlign = xlsio.VAlignType.center
..backColorRgb = const Color(0xFFFFFF00);

final tableHeaderStyle = workbook.styles.add('tableHeaderStyle');
tableHeaderStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 11
..bold = true
..borders.all.lineStyle = xlsio.LineStyle.thin
..borders.all.color = '#000000';

final dataStyle = workbook.styles.add('dataStyle');
dataStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..borders.all.lineStyle = xlsio.LineStyle.thin
..borders.all.color = '#000000';

sheet.getRangeByName('A2:H2').merge();
sheet.getRangeByName('A2').setText(reportDetails['Company'] ?? '');
sheet.getRangeByName('A2:H2').cellStyle = companyTitleStyle;

sheet.getRangeByName('A3:H3').merge();
sheet.getRangeByName('A3').setText(reportDetails['Certification'] ?? '');
sheet.getRangeByName('A3:H3').cellStyle = companyCertStyle;

int row = 6;
List<String> keys = [
'Report No',
'Rec Dt of Sample',
'Date of Testing',
'MFG Quality No',
'Material',
'SO No',
'MFG Unit',
'Colour',
'Warp Yarn',
'Warp Yarn1',
'Weft Yarn',
'Weft Yarn1',
'Reed',
'Pick'
];
for (var key in keys) {
sheet.getRangeByName('B$row').setText(key);
sheet.getRangeByName('B$row').cellStyle = labelStyle;
sheet.getRangeByName('C$row:F$row').merge();
sheet.getRangeByName('C$row').setText(reportDetails[key] ?? '');
sheet.getRangeByName('C$row').cellStyle =
(key == 'Report No' || key == 'Customer Name') ? yellowValueStyle : valueStyle;
row++;
}

int row1 = 6;
List<String> keys1 = [
'Report No',
'Rec Dt of Sample',
'Date of Testing',
'Customer Name',
'Material',
'SO No',
'Dimensions',
'Colour',
'Warp Yarn',
'Warp Yarn1',
'Weft Yarn',
'Weft Yarn1',
'Reed',
'Pick'
];
for (var key in keys1) {
sheet.getRangeByName('H$row1').setText(key);
sheet.getRangeByName('H$row1').cellStyle = labelStyle;
if (key == 'Dimensions') {
List<String> dimensionsParts = (reportDetails[key] ?? '').split(' ');
for (int i = 0; i < 4; i++) {
String column = String.fromCharCode('I'.codeUnitAt(0) + i);
sheet.getRangeByName('$column$row1').setText(dimensionsParts.length > i ? dimensionsParts[i] : '');
sheet.getRangeByName('$column$row1').cellStyle = valueStyle;
}
} else {
sheet.getRangeByName('I$row1:L$row1').merge();
sheet.getRangeByName('I$row1').setText(reportDetails[key] ?? '');
sheet.getRangeByName('I$row1').cellStyle =
(key == 'Report No' || key == 'Customer Name') ? yellowValueStyle : valueStyle;
}
row1++;
}

sheet.getRangeByName('B21').setText('Test');
sheet.getRangeByName('C21:D21').merge();
sheet.getRangeByName('C21').setText('Test Method No.');
sheet.getRangeByName('E21').setText('Result');
sheet.getRangeByName('F21').setText('Standard');
sheet.getRangeByName('G21').setText('Minimum');
sheet.getRangeByName('H21').setText('Maximum');
sheet.getRangeByName('I21').setText('Remarks');
sheet.getRangeByName('B21:I21').cellStyle = tableHeaderStyle;

for (int i = 0; i < testData.length; i++) {
int dataRow = 22 + i;
final rowData = testData[i];
columnMapping.forEach((jsonKey, header) {
final colIndex = columnHeaders.indexOf(header) + 1;
final value = rowData[jsonKey]?.toString() ?? '';
final column = String.fromCharCode(66 + colIndex - 1); // B=66
sheet.getRangeByName('$column$dataRow').setText(value);
sheet.getRangeByName('$column$dataRow').cellStyle = dataStyle;
});
sheet.getRangeByName('C$dataRow:D$dataRow').merge(); // Merge C and D for Test Method No.
}
}
}
