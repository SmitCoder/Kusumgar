
import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:flutter/material.dart';
import 'package:archive/archive.dart';
import 'package:intl/intl.dart';
import 'package:open_filex/open_filex.dart';
import 'package:path_provider/path_provider.dart';
import 'package:kusumgar_final/models/docx_models.dart';
import 'package:kusumgar_final/utils/companies/docx_company_config.dart';
import 'package:kusumgar_final/utils/companies/company1_docx.dart';
import 'package:kusumgar_final/utils/companies/company2_docx.dart';
import 'package:kusumgar_final/utils/companies/company3_docx.dart';

class DocxGenerator extends ChangeNotifier {
int _fileCounter = 0;
bool _isGenerating = false;
String _selectedCompany = 'Company 3'; // Default to Company 3

final Map<String, CompanyDocxConfig> companyConfigs = {
'Company 1': Company1DocxConfig(),
'Company 2': Company2DocxConfig(),
'Company 3': Company3DocxConfig(),
};

bool get isGenerating => _isGenerating;
String get selectedCompany => _selectedCompany;

void setSelectedCompany(String company) {
_selectedCompany = company;
notifyListeners();
}

void setIsGenerating(bool value) {
_isGenerating = value;
notifyListeners();
}

Future<String> _generateUniqueFileName(String baseName, String directoryPath) async {
String fileName;
String filePath;
int counter = _fileCounter;

do {
counter++;
fileName = counter == 1 ? '$baseName.docx' : '$baseName($counter).docx';
filePath = '$directoryPath/$fileName';
} while (await File(filePath).exists());

_fileCounter = counter;
return fileName;
}

Future<bool> _isDirectoryWritable(Directory dir) async {
try {
final tempFile = File('${dir.path}/temp_test_${DateTime.now().millisecondsSinceEpoch}.txt');
await tempFile.writeAsString('Test');
await tempFile.delete();
return true;
} catch (e) {
return false;
}
}

SnackBar _buildSnackBar(String message, {bool isError = false, bool isSuccess = false}) {
return SnackBar(
content: Row(
children: [
Icon(
isError
? Icons.error_outline
    : (isSuccess ? Icons.check_circle_outline : Icons.info_outline),
color: Colors.white,
size: 24,
),
const SizedBox(width: 12),
Expanded(
child: Text(
message,
style: const TextStyle(
color: Colors.white,
fontSize: 14,
fontWeight: FontWeight.w500,
),
),
),
],
),
backgroundColor: isError
? Colors.red.shade700
    : (isSuccess ? Colors.green.shade700 : Colors.blue.shade700),
behavior: SnackBarBehavior.floating,
shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
margin: const EdgeInsets.all(16),
duration: const Duration(seconds: 4),
);
}

Future<void> generateAndOpenDocx(BuildContext context) async {
final scaffoldMessenger = ScaffoldMessenger.of(context);

try {
if (kIsWeb) {
scaffoldMessenger.showSnackBar(
_buildSnackBar('Web platform not supported for .docx generation.', isError: true),
);
return;
}

Directory? saveDir;
try {
if (Platform.isAndroid) {
Directory downloadsDir = Directory('/storage/emulated/0/Download');
if (await downloadsDir.exists() && await _isDirectoryWritable(downloadsDir)) {
saveDir = downloadsDir;
} else {
saveDir = await getTemporaryDirectory();
}
} else if (Platform.isIOS) {
saveDir = await getApplicationDocumentsDirectory();
} else if (Platform.isWindows || Platform.isMacOS || Platform.isLinux) {
final homeDir = await getApplicationSupportDirectory();
final downloadsDir = Directory('${homeDir.path}/Downloads');
if (!await downloadsDir.exists()) {
await downloadsDir.create(recursive: true);
}
if (await _isDirectoryWritable(downloadsDir)) {
saveDir = downloadsDir;
} else {
saveDir = homeDir;
}
}
} catch (e) {
saveDir = await getTemporaryDirectory();
}

if (saveDir == null || !await _isDirectoryWritable(saveDir)) {
throw Exception("Couldn't access a writable directory.");
}

if (!await saveDir.exists()) {
await saveDir.create(recursive: true);
}

final CompanyDocxConfig config = companyConfigs[_selectedCompany]!;
final doc = DocxDocument(paragraphs: config.paragraphs, styles: config.styles);

final archive = Archive();

const contentTypesXml = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>
''';
archive.addFile(ArchiveFile('[Content_Types].xml', contentTypesXml.length, utf8.encode(contentTypesXml)));

const rootRels = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="word/settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="word/fontTable.xml"/>
</Relationships>
''';
archive.addFile(ArchiveFile('_rels/.rels', rootRels.length, utf8.encode(rootRels)));

const documentRels = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="word/fontTable.xml"/>
</Relationships>
''';
archive.addFile(
ArchiveFile('word/_rels/document.xml.rels', documentRels.length, utf8.encode(documentRels)));

const settingsXml = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:defaultTabStop w:val="720"/>
  <w:compat/>
</w:settings>
''';
archive.addFile(ArchiveFile('word/settings.xml', settingsXml.length, utf8.encode(settingsXml)));

const fontTableXml = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C0000000" w:usb2="00000000" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
</w:fonts>
''';
archive.addFile(ArchiveFile('word/fontTable.xml', fontTableXml.length, utf8.encode(fontTableXml)));

final stylesXml = doc.toStylesXml();
archive.addFile(ArchiveFile('word/styles.xml', stylesXml.length, utf8.encode(stylesXml)));

final documentXml = doc.toDocumentXml();
archive.addFile(ArchiveFile('word/document.xml', documentXml.length, utf8.encode(documentXml)));

final zipEncoder = ZipEncoder();
final bytes = zipEncoder.encode(archive);
if (bytes == null) {
throw Exception('Failed to encode .docx');
}

final dateFormat = DateFormat('yyyyMMdd_HHmm');
final baseFileName = 'certificate_${dateFormat.format(DateTime.now())}_${config.companyName}';
final fileName = await _generateUniqueFileName(baseFileName, saveDir.path);
final filePath = '${saveDir.path}/$fileName';
final file = File(filePath);

await file.writeAsBytes(bytes);

final result = await OpenFilex.open(
filePath, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
if (result.type != ResultType.done) {
scaffoldMessenger.showSnackBar(
_buildSnackBar('Failed to open file: ${result.message}', isError: true),
);
} else {
scaffoldMessenger.showSnackBar(
_buildSnackBar('File saved and opened successfully.', isSuccess: true),
);
}
} catch (e) {
scaffoldMessenger.showSnackBar(_buildSnackBar('Error: $e', isError: true));
} finally {
setIsGenerating(false); // Ensure state is reset even on error
}
}
}