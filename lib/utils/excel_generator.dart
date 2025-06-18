
import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart' show debugPrint, kIsWeb, ChangeNotifier;
import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'package:path_provider/path_provider.dart';
import 'package:open_filex/open_filex.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:universal_html/html.dart' as html;
import 'companies/company1.dart';
import 'companies/company2.dart';
import 'companies/company3.dart';
import 'company_config.dart';

class ExcelGenerator extends ChangeNotifier {
List<Map<String, dynamic>> testData = [];
List<Map<String, dynamic>> originalTestData = [];
List<List<TextEditingController>> testDataControllers = [];
List<Map<String, dynamic>> changeLog = [];
Map<String, String> reportDetails = {};
bool _dataLoaded = false;
bool _isLoading = false;
String? _message;
String _selectedCompany = 'Company 1';
static int _fileCounter = 0;

final Map<String, CompanyConfig> companyConfigs = {
'Company 1': Company1Config(),
'Company 2': Company2Config(),
'Company 3': Company3Config(),
};

bool get dataLoaded => _dataLoaded;
bool get isLoading => _isLoading;
String? get message => _message;
String get selectedCompany => _selectedCompany;

void setSelectedCompany(String company) {
_selectedCompany = company;
_dataLoaded = false;
testData = [];
originalTestData = [];
reportDetails = {};
_message = null;
changeLog = [];
for (var row in testDataControllers) {
for (var controller in row) {
controller.dispose();
}
}
testDataControllers = [];
notifyListeners();
}

Future<void> loadData() async {
_isLoading = true;
_message = null;
for (var row in testDataControllers) {
for (var controller in row) {
controller.dispose();
}
}
testDataControllers = [];
changeLog = [];
notifyListeners();

try {
final CompanyConfig config = companyConfigs[_selectedCompany]!;
final String jsonPath = config.jsonPath;
final String jsonString = await rootBundle.loadString(jsonPath);
final Map<String, dynamic> jsonData = jsonDecode(jsonString);

List<Map<String, dynamic>> loadedTestData = [];
List<String> headers = config.columnHeaders;

if (jsonData['test_data'] is List) {
if (jsonData['test_data'].isNotEmpty && jsonData['test_data'][0] is List) {
loadedTestData = (jsonData['test_data'] as List)
    .asMap()
    .map((i, row) {
final rowData = List<String>.from(row);
if (rowData.length != config.columnHeaders.length) {
throw Exception('Inconsistent column count in test_data for $_selectedCompany');
}
final map = <String, dynamic>{};
for (int j = 0; j < rowData.length; j++) {
map[config.columnMapping.keys.elementAt(j)] = rowData[j];
}
return MapEntry(i, map);
})
    .values
    .toList();
} else {
loadedTestData = (jsonData['test_data'] as List)
    .map((row) => Map<String, dynamic>.from(row))
    .toList();

final expectedKeys = config.columnMapping.keys.toSet();
for (var row in loadedTestData) {
if (!expectedKeys.containsAll(row.keys)) {
throw Exception('Invalid keys in test_data for $_selectedCompany');
}
}

if (jsonData['headers'] != null) {
headers = List<String>.from(jsonData['headers']);
if (headers.length != config.columnHeaders.length) {
throw Exception('Header count mismatch for $_selectedCompany');
}
}
}
} else {
throw Exception('Invalid test_data format for $_selectedCompany');
}

List<List<TextEditingController>> controllers = [];
for (int i = 0; i < loadedTestData.length; i++) {
List<TextEditingController> rowControllers = [];
for (int j = 0; j < headers.length; j++) {
final jsonKey = config.columnMapping.keys.elementAt(j);
final value = loadedTestData[i][jsonKey]?.toString() ?? '';
final controller = TextEditingController(text: value);
final rowIndex = i;
final colIndex = j;
controller.addListener(() {
final newValue = controller.text;
final originalValue = originalTestData[rowIndex][jsonKey]?.toString() ?? '';
if (newValue != originalValue) {
changeLog.removeWhere((log) =>
log['row'] == rowIndex + 1 && log['column'] == headers[colIndex]);
changeLog.add({
'row': rowIndex + 1,
'column': headers[colIndex],
'originalValue': originalValue,
'newValue': newValue,
'timestamp': DateTime.now().toUtc().toIso8601String(),
});
} else {
changeLog.removeWhere((log) =>
log['row'] == rowIndex + 1 && log['column'] == headers[colIndex]);
}
loadedTestData[rowIndex][jsonKey] = newValue;
notifyListeners();
});
rowControllers.add(controller);
}
controllers.add(rowControllers);
}

reportDetails = Map<String, String>.from(jsonData['report_details']);
testData = loadedTestData;
originalTestData = loadedTestData.map((row) => Map<String, dynamic>.from(row)).toList();
testDataControllers = controllers;
_dataLoaded = true;
_message = 'Data loaded successfully!';
} catch (e) {
_message = 'Error loading data: $e';
} finally {
_isLoading = false;
notifyListeners();
}
}

Future<void> createExcel(BuildContext context) async {
_isLoading = true;
_message = null;
notifyListeners();

try {
final logData = {
'company': _selectedCompany,
'changes': changeLog,
};
final logString = jsonEncode(logData);
final logFileName = 'changes_${_selectedCompany.toLowerCase().replaceAll(' ', '_')}.json';
String logMessage;

if (kIsWeb) {
final blob = html.Blob([logString], 'text/json');
final url = html.Url.createObjectUrlFromBlob(blob);
final anchor = html.AnchorElement(href: url)
..setAttribute('download', logFileName)
..click();
html.Url.revokeObjectUrl(url);
logMessage = 'Change log downloaded as $logFileName';
} else {
final directory = await getApplicationSupportDirectory();
final filePath = '${directory.path}/$logFileName';
final file = File(filePath);
await file.writeAsString(logString);
logMessage = 'Change log saved at $filePath';
}

_fileCounter++;
final String fileNameBase = 'Test_Report_$_fileCounter';
final xlsio.Workbook workbook = xlsio.Workbook();
final xlsio.Worksheet sheet = workbook.worksheets[0];

final CompanyConfig config = companyConfigs[_selectedCompany]!;
await config.initializeSheet(sheet, workbook);
config.configureSheet(sheet, reportDetails, testData, workbook);

List<int>? bytes;
try {
bytes = workbook.saveAsStream();
} catch (e) {
throw Exception('Failed to save Excel file: $e');
} finally {
workbook.dispose();
}

if (bytes == null || bytes.isEmpty) {
throw Exception('Excel file generation failed: No data generated');
}

if (kIsWeb) {
final blob = html.Blob([bytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
final url = html.Url.createObjectUrlFromBlob(blob);
final anchor = html.AnchorElement(href: url)
..setAttribute('download', '$fileNameBase.xlsx')
..click();
html.Url.revokeObjectUrl(url);
_message = 'Excel file downloaded! $logMessage';
} else {
final String path = (await getApplicationSupportDirectory()).path;
final String fileName = '$path/$fileNameBase.xlsx';
final File file = File(fileName);
await file.writeAsBytes(bytes, flush: true);
final result = await OpenFilex.open(fileName);
_message = result.type == ResultType.done
? 'Excel file created and opened! $logMessage'
    : 'Error opening file: ${result.message}';
}
} catch (e, stackTrace) {
debugPrint('Error in createExcel: $e\nStack trace: $stackTrace');
_message = 'Error generating Excel file: $e';
} finally {
_isLoading = false;
notifyListeners();
}
}
}
