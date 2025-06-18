
import 'package:flutter/services.dart' show rootBundle;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:flutter/foundation.dart' show debugPrint; // Added for debugPrint
import '../company_config.dart';

class Company2Config implements CompanyConfig {
@override
String get jsonPath => 'assets/company2.json';

@override
List<String> get columnHeaders => [
'Test',
'Test Method',
'Result',
'Standard',
'Minimum',
'Maximum',
'Remarks'
];

@override
Map<String, String> get columnMapping => {
'test': 'Test',
'method': 'Test Method',
'result': 'Result',
'standard': 'Standard',
'minimum': 'Minimum',
'maximum': 'Maximum',
'remarks': 'Remarks',
};

@override
Future<void> initializeSheet(xlsio.Worksheet sheet, xlsio.Workbook workbook) async {
sheet.showGridlines = true;

try {
final imageData = await rootBundle.load('assets/company2/logo.png');
final List<int> imageBytes = imageData.buffer.asUint8List();
final xlsio.Picture picture = sheet.pictures.addStream(1, 1, imageBytes);
picture.height = 100;
picture.width = 100;
picture.row = 1;
picture.column = 1;
// picture.columnOffsetPx = 20; // Changed from columnOffset
// picture.rowOffsetPx = 20;    // Changed from rowOffset
} catch (e) {
debugPrint('Error loading logo for Company 2: $e');
}
}

@override
void configureSheet(
xlsio.Worksheet sheet,
Map<String, String> reportDetails,
List<Map<String, dynamic>> testData,
xlsio.Workbook workbook,
) {
for (int i = 2; i <= 8; i++) {
sheet.getRangeByIndex(1, i).columnWidth = 15.0;
}

final titleStyle = workbook.styles.add('titleStyle');
titleStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 16
..bold = true;

final subtitleStyle = workbook.styles.add('subtitleStyle');
subtitleStyle
..hAlign = xlsio.HAlignType.center
..vAlign = xlsio.VAlignType.center
..fontSize = 12;

final headerStyle = workbook.styles.add('headerStyle');
headerStyle
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

sheet.getRangeByName('B2:G2').merge();
sheet.getRangeByName('B2').setText(reportDetails['Company'] ?? '');
sheet.getRangeByName('B2:G2').cellStyle = titleStyle;

sheet.getRangeByName('B3:G3').merge();
sheet.getRangeByName('B3').setText(reportDetails['Certification'] ?? '');
sheet.getRangeByName('B3:G3').cellStyle = subtitleStyle;

int row = 5;
reportDetails.forEach((key, value) {
if (key != 'Company' && key != 'Certification') {
sheet.getRangeByName('B$row').setText(key);
sheet.getRangeByName('C$row:G$row').merge();
sheet.getRangeByName('C$row').setText(value);
sheet.getRangeByName('B$row:G$row').cellStyle = dataStyle;
row++;
}
});

row += 2;
sheet.getRangeByName('B$row').setText(columnHeaders[0]);
sheet.getRangeByName('C$row').setText(columnHeaders[1]);
sheet.getRangeByName('D$row').setText(columnHeaders[2]);
sheet.getRangeByName('E$row').setText(columnHeaders[3]);
sheet.getRangeByName('F$row').setText(columnHeaders[4]);
sheet.getRangeByName('G$row').setText(columnHeaders[5]);
sheet.getRangeByName('H$row').setText(columnHeaders[6]);
sheet.getRangeByName('B$row:H$row').cellStyle = headerStyle;

for (int i = 0; i < testData.length; i++) {
int dataRow = row + i + 1;
final rowData = testData[i];
columnMapping.forEach((jsonKey, header) {
final colIndex = columnHeaders.indexOf(header);
final value = rowData[jsonKey]?.toString() ?? '';
sheet.getRangeByName('${String.fromCharCode(66 + colIndex)}$dataRow').setText(value);
sheet.getRangeByName('${String.fromCharCode(66 + colIndex)}$dataRow').cellStyle = dataStyle;
});
}
}
}
