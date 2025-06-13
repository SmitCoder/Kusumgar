import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/foundation.dart' show debugPrint, kIsWeb, ChangeNotifier;
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_filex/open_filex.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:universal_html/html.dart' as html;

class ExcelGenerator extends ChangeNotifier {
  List<List<String>> testData = [];
  Map<String, String> reportDetails = {};
  bool _dataLoaded = false;
  bool _isLoading = false;
  String? _message;
  static int _fileCounter = 0; // Static counter for file naming

  bool get dataLoaded => _dataLoaded;
  bool get isLoading => _isLoading;
  String? get message => _message;

  get i => null;

  Future<void> loadData() async {
    _isLoading = true;
    _message = null;
    notifyListeners();

    try {
      reportDetails = {
        'Company': 'Kusumgar Limited (Formerly Known As Kusumgar Private Limited)',
        'Certification': 'An ISO 9001:2015 Company',
        'Report No': 'Q250001330',
        'Date': '05-04-2025',
        'Rec Dt of Sample': '05-04-2025',
        'Project Leader': 'Sandeep',
        'Date of Testing': '05-04-2025',
        'Reg/Dev': 'Regular',
        'MFG Quality No': 'PWNN4201',
        'Sales Quality No': 'CLOTH,NYL,65",FG 24165,INTERNATIONAL-P44378 T4-NB',
        'Product No': 'INT-P443',
        'Dimensions': 'ACL02 GN0088 WRSI01 170',
        'Customer Name': 'Airborne Systems Limited.',
        'Material': 'Nylon 66',
        'Type of Finish': 'WRSI01 No Finish',
        'SO No': 'SO2425-001457',
        'Batch No': '24P41768',
        'MFG Unit': 'Kusumgar Limited',
        'Colour': 'GN0088 LD No. 4448 - Blue',
        'No of Samples': '1',
        'P/c Roll No Result': '24045849',
        'Final Flag': 'OK',
        'Tested By': 'Anil Pandit',
        'Prepared By': 'Chandan Verma',
        'Remarked By': 'All Parameters are ok',
        'Verified By': 'qa.vapi',
        'Verified DateTime': '26-04-2025 13:01:51',
      };

      testData = [
        ['Width Inch', 'ASTM D 3774', '67.0000', '65.00', '35.00', '75.00', ''],
        ['Weight Oz/Yd2', 'ASTM D 3776', '1.1200', '1.20', '1.00', '1.20', ''],
        ['Thrd Count-Wp inch', 'ASTM D 3775', '132.0000', '126.00', '126.00', '145.00', ''],
        ['Thrd Count-Wt inch', 'ASTM D 3775', '136.0000', '132.00', '132.00', '150.00', ''],
        ['BS-Wp lbf', 'ASTM D 5035', '48.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wp lbf', 'ASTM D 5035', '48.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wp lbf', 'ASTM D 5035', '48.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wp lbf', 'ASTM D 5035', '48.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wp lbf', 'ASTM D 5035', '48.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wt lbf', 'ASTM D 5035', '46.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wt lbf', 'ASTM D 5035', '46.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wt lbf', 'ASTM D 5035', '46.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wt lbf', 'ASTM D 5035', '46.0000', '45.00', '45.00', '80.00', ''],
        ['BS-Wt lbf', 'ASTM D 5035', '46.0000', '45.00', '45.00', '80.00', ''],
        ['%Elong-Wp', 'ASTM D 5035', '23.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wp', 'ASTM D 5035', '23.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wp', 'ASTM D 5035', '23.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wp', 'ASTM D 5035', '23.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wp', 'ASTM D 5035', '23.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wt', 'ASTM D 5035', '21.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wt', 'ASTM D 5035', '21.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wt', 'ASTM D 5035', '21.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wt', 'ASTM D 5035', '21.0000', '20.00', '20.00', '50.00', ''],
        ['%Elong-Wt', 'ASTM D 5035', '21.0000', '20.00', '20.00', '50.00', ''],
        ['TS-Wp lbf', 'ASTM D 2261', '7.0000', '5.00', '5.00', '15.00', ''],
        ['TS-Wp lbf', 'ASTM D 2261', '7.0000', '5.00', '5.00', '15.00', ''],
        ['TS-Wp lbf', 'ASTM D 2261', '7.0000', '5.00', '5.00', '15.00', ''],
        ['TS-Wp lbf', 'ASTM D 2261', '7.0000', '5.00', '5.00', '15.00', ''],
        ['TS-Wp lbf', 'ASTM D 2261', '7.0000', '5.00', '5.00', '15.00', ''],
        ['TS-Wt lbf', 'ASTM D 2261', '6.5000', '5.00', '5.00', '15.00', ''],
        ['TS-Wt lbf', 'ASTM D 2261', '6.5000', '5.00', '5.00', '15.00', ''],
        ['TS-Wt lbf', 'ASTM D 2261', '6.5000', '5.00', '5.00', '15.00', ''],
        ['TS-Wt lbf', 'ASTM D 2261', '6.5000', '5.00', '5.00', '15.00', ''],
        ['TS-Wt lbf', 'ASTM D 2261', '6.5000', '5.00', '5.00', '15.00', ''],
        ['Thick inch', 'ASTM D 1777', '0.0015', '0.00', '0.00', '0.00', ''],
        ['AP@1/2 Inch (ft3/ft2/m)Av', 'ASTM D 737', '0.7700', '1.75', '0.50', '3.00', ''],
        ['AP@1/2ynos (ft3/ft2/m)Mn', 'ASTM D 737', '0.6400', '1.75', '0.50', '3.00', ''],
        ['AP@1/2 Inch (ft3/ft2/m)Mx', 'ASTM D 737', '0.8600', '1.75', '0.50', '3.00', ''],
        ['Water Repellency', 'AATCC 22', '90.0000', '80.00', '80.00', '100.00', ''],
        ['pH Value', 'AATCC 61', '6.2000', '7.30', '5.50', '9.00', ''],
        ['CF-Rub-Dry', 'AATCC 8', '4.5000', '3.50', '3.50', '5.00', ''],
        ['CF-Rub-Wet', 'AATCC 8', '4.5000', '3.50', '3.50', '5.00', ''],
      ];

      _dataLoaded = true;
      _message = 'Data loaded successfully!';
    } catch (e) {
      _message = 'Error loading data: $e';
    } finally {
      _isLoading = false;
      notifyListeners();
    }
  }

  Future<void> createExcel() async {
    _isLoading = true;
    _message = null;
    notifyListeners();

    try {
      // Increment file counter for unique file naming
      _fileCounter++;
      final String fileNameBase = 'Test_Report_$_fileCounter';

      // Create a new Excel workbook and get the first sheet
      final xlsio.Workbook workbook = xlsio.Workbook();
      final xlsio.Worksheet sheet = workbook.worksheets[0];

      // Disable gridlines to make the sheet appear completely blank outside the table area
      sheet.showGridlines = false;

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
      sheet.getRangeByIndex(2, 1).rowHeight = 102.60;
      sheet.getRangeByIndex(29, 2).rowHeight = 30.00;// Row 2 (Company Header)


      // --- Define Styles ---
      final companyTitleStyle = workbook.styles.add('companyTitleStyle');
      companyTitleStyle
        ..hAlign = xlsio.HAlignType.center
        ..vAlign = xlsio.VAlignType.center
        ..fontSize = 12
        ..bold = true
        ..wrapText = true;

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
      sheet.setRowHeightInPixels(3, 30); // Set height for row 3 (TEST REPORT)

      // Report Details
      sheet.getRangeByName('A4').setText('Report No.');
      sheet.getRangeByName('A4').cellStyle = labelStyle;
      sheet.getRangeByName('B4').setText('Q250001330');
      sheet.getRangeByName('B4').cellStyle = valueStyle;

      sheet.getRangeByName('C4').setText('Date');
      sheet.getRangeByName('C4').cellStyle = labelStyle;
      sheet.getRangeByName('D4').setText('21-04-2025');
      sheet.getRangeByName('D4').cellStyle = valueStyle;

      sheet.getRangeByName('A5:A15').merge();
      sheet.getRangeByName('A5').setText('Item');
      sheet.getRangeByName('A5').cellStyle = item;
      // sheet.getRangeByName('A3:H3').cellStyle = headerStyle;
      // sheet.setRowHeightInPixels(3, 30); // Set height for row 3 (TEST REPORT)

      sheet.getRangeByName('B5:B15').merge();

      sheet.getRangeByName('B5').setText('166452\nCLOTH,NYL,65",\nFG 24165\nINTERNATIONAL-\nP44378 T4-NB\n');
      sheet.getRangeByName('B5').cellStyle = item;

      sheet.getRangeByName('C5:C15').merge();
      sheet.getRangeByName('C5').setText('Customer');
      sheet.getRangeByName('C5').cellStyle = item;


      sheet.getRangeByName('D5:H5').merge();
      sheet.getRangeByName('D5').setText('Airborne Systems NA of CA Inc.');
      sheet.getRangeByName('D5').cellStyle = valueStyle;

      // Lot Numbers
      int lotRow = 6;
      final lots = [
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
        sheet.getRangeByName('E$lotRow:H$lotRow').merge();
        sheet.getRangeByName('D$lotRow').setText('Lot:');
        sheet.getRangeByName('E$lotRow').setText('$lot');
        sheet.getRangeByName('E$lotRow').cellStyle = valueStyle;
        sheet.setRowHeightInPixels(lotRow, 20);
        lotRow++;
      }



      // Additional Details
      sheet.getRangeByName('A$lotRow').setText('Q. No.');
      sheet.getRangeByName('A$lotRow').cellStyle = labelStyle;
      sheet.getRangeByName('B$lotRow').setText('4201 (GN0088)');
      sheet.getRangeByName('B$lotRow').cellStyle = valueStyle;

      sheet.getRangeByName('C16:E16').merge();
      sheet.getRangeByName('C16').setText('Qty.: 14,985.73 Yards');
      sheet.getRangeByName('C16').cellStyle = item;



      sheet.getRangeByName('F16:H16').merge();
      sheet.getRangeByName('F16').setText('Rolls: 66');
      sheet.getRangeByName('F16').cellStyle = item;



      sheet.getRangeByName('A17').setText('Width');
      sheet.getRangeByName('A17').cellStyle = item;


      sheet.getRangeByName('B17').setText('165.0 CMS (65.0”)');
      sheet.getRangeByName('B17').cellStyle = item;

      sheet.getRangeByName('C17:E17').merge();
      sheet.getRangeByName('C17').setText('Invoice No. ES25260039');
      sheet.getRangeByName('C17').cellStyle = item;

      sheet.getRangeByName('F17:H17').merge();
      sheet.getRangeByName('F17').setText('21-04-2025');
      sheet.getRangeByName('F17').cellStyle = item;

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


      ////
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


      //
      sheet.getRangeByName('H30:H33').merge();
      sheet.getRangeByName('H34:H37').merge();
      sheet.getRangeByName('H38:H39').merge();
      sheet.getRangeByName('H42:H45').merge();
      sheet.getRangeByName('H46:H49').merge();
      sheet.getRangeByName('H50:H53').merge();
      sheet.getRangeByName('H54:H55').merge();
      sheet.getRangeByName('H56:H57').merge();
      sheet.getRangeByName('H58:H59').merge();
      sheet.getRangeByName('H38:H39').merge();



















      // --- Test Results Table ---
      int tableStartRow = lotRow + 3;
      sheet.getRangeByName('A$tableStartRow').setText('Test');
      sheet.getRangeByName('B$tableStartRow').setText('Test Method');
      sheet.getRangeByName('C$tableStartRow').setText('Pc No. 24045849');
      sheet.getRangeByName('D$tableStartRow').setText('');
      sheet.getRangeByName('E$tableStartRow').setText('');
      sheet.getRangeByName('F$tableStartRow').setText('');
      sheet.getRangeByName('G$tableStartRow').setText('');
      sheet.getRangeByName('H$tableStartRow').setText('Standard');
      sheet.getRangeByName('A$tableStartRow:H$tableStartRow').cellStyle = tableHeaderStyle;
      sheet.setRowHeightInPixels(tableStartRow, 40);
      tableStartRow++;

      // Table Data
      final tests = [
        ['Yarn', 'ASTM D276', 'Nylon 6-6 H.T. light and heat resistance', '', '', '', '', 'Nylon 6-6 H.T. light and heat resistance'],
        ['Color', 'Visual', 'Foliage Green', '', '', '', '', 'Foliage Green'],
        ['Colorfastness Light', 'AATCC 16.3', '3-4', '', '', '', '', '3-4 min.'],
        ['Colorfastness Laundering:', 'AATCC 61, Test 1A\nAATCC 8', '', '', '', '', '', ''],
        ['Color Change-', '', '4', '', '', '', '', '3-4 min.'],
        ['Staining-', '', '4', '', '', '', '', '3-4 min.'],
        ['Crocking:', 'AATCC 8', '', '', '', '', '', ''],
        ['Dry-', '', '4.5', '', '', '', '', '3.5 min.'],
        ['Wet-', '', '4.5', '', '', '', '', '3.5 min.'],
        ['Bleeding in damp air', '4.9.3 As Mention in\nPIA-C-44378E', '4', '', '', '', '', '3-4 min.'],
        ['Light resistance: Warp-', '4.9.2.1 As Mention in\nPIA-C-44378E', '17', '20', '18', '19', '20', 'Not lose more than 25% of original strength'],
        ['(Light Source – Xenon)', '', 'Avg: ', '18.8', '', '', '', ''],
        ['Light resistance: Filling', '', '19', '19', '20', '18', '16', ''],
        ['(Light Source – Xenon)', '', 'Avg: ', '18.4', '', '', '', ''],
        ['Heat resistance: Warp-', '4.9.2.2 As Mention in\nPIA-C-44378E', '20', '21', '18', '21', '15', 'Not lose more than 25% of original strength'],
        ['', '', 'Avg: ', '19.0', '', '', '', ''],
        ['Heat resistance: Filling-', '', '17', '20', '20', '17', '17', ''],
        ['', '', 'Avg: ', '18.2', '', '', '', ''],
        ['Weight', 'ASTM D3776', '1.12', '1.11', '1.12', '1.12', '1.11', 'OSY 1.20 Max.'],
        ['', '', 'Avg: ', '1.12', '', '', '', ''],
        ['Thickness', 'ASTM D1777', '0.0015', '0.0016', '0.0016', '0.0015', '0.0015', '0.003” Max.'],
        ['', '', 'Avg: ', ' 0.0015', '', '', '', ''],
        ['Breaking strength: Warp-', 'ASTM D5035', '48.0', '48.5', '48.2', '48.9', '49.0', 'Min. 45 lbs/inch'],
        ['', '', 'Avg: ', ' 48.5', '', '', '', ''],
        ['Breaking strength: Filling-', '', '46.0', '47.8', '47.2', '47.3', '46.7', ''],
        ['', '', 'Avg: ', ' 47.0', '', '', '', ''],
        ['% Elongation: Warp-', 'ASTM D5035', '23.0', '24.4', '24.9', '24.1', '24.3', 'Min. 20%'],
        ['', '', 'Avg: ', '24.1', '', '', '', ''],
        ['% Elongation: Filling-', '', '21.0', '21.5', '22.5', '22.6', '22.9', ''],
        ['', '', 'Avg: ', '22.1', '', '', '', ''],
        ['Tearing strength: Warp-', 'ASTM D2261', '7.0', '8.0', '7.6', '7.7', '8.1', 'Min. 5 lbs'],
        ['', '', 'Avg: ', '7.7', '', '', '', ''],
        ['Tearing strength: Filling-', '', '6.5', '7.9', '6.9', '7.5', '7.2', ''],
        ['', '', 'Avg: ', ' 7.2', '', '', '', ''],
        ['Air permeability', 'ASTM D737', '0.77', '0.64', '0.86', '1.05', '1.13', '0.5 to 3.0 CFM'],
        ['', '', 'Avg: ', '0.89', '', '', '', ''],
        ['Yarn: Warp-', 'ASTM D3775', '132', '132', '132', '132', '132', 'Min. 126 per Inch'],
        ['', '', 'Avg: ', ' 132', '', '', '', ''],
        ['Yarn: Filling-', '', '136', '136', '136', '136', '136', 'Min. 132 per Inch'],
        ['', '', 'Avg: ', '136', '', '', '', ''],
        ['Weave (pattern)', 'Visual', 'Rip stop Figure1', '', '', '', '', 'Rip stop Figure1'],
        ['Width', 'ASTM D3774', '65.0”', '', '', '', '', '65-1/2"+/-1/2"'],
        ['pH Value', 'AATCC 81', '6.2', '', '6.5', '', '', '5.5 to 9.0'],
        ['', '', 'Avg: ', '6.35', '', '', '', ''],
        ['Fluorocarbon', '3.6.2.1 As Mention in PIA-C-44378E', 'Applied to fabric', '', '', '', '', 'Applied to fabric'],
        ['Spray rating', 'AATCC 22', '90', '90', '', '90', '', '80, 80, 70 Min.'],
        ['', '', 'Avg: 90', '', '', '', '', ''],
      ];

      for (var test in tests) {
        sheet.getRangeByName('A$tableStartRow').setText(test[0]);
        sheet.getRangeByName('B$tableStartRow').setText(test[1]);
        sheet.getRangeByName('C$tableStartRow').setText(test[2]);
        sheet.getRangeByName('D$tableStartRow').setText(test[3]);
        sheet.getRangeByName('E$tableStartRow').setText(test[4]);
        sheet.getRangeByName('F$tableStartRow').setText(test[5]);
        sheet.getRangeByName('G$tableStartRow').setText(test[6]);
        sheet.getRangeByName('H$tableStartRow').setText(test[7]);

        sheet.getRangeByName('A$tableStartRow:H$tableStartRow').cellStyle = dataStyle;

        // ✅ Direct row match instead of index
        if ([20, 29, 64].contains(tableStartRow)) {
          sheet.setRowHeightInPixels(tableStartRow, 50); // Increased height
        } else {
          sheet.setRowHeightInPixels(tableStartRow, 25); // Normal height
        }

        tableStartRow++;



        Future<void> insertLogo(xlsio.Worksheet sheet) async {
          final ByteData imageData = await rootBundle.load('assets/images/logo.png'); // your actual path
          final List<int> imageBytes = imageData.buffer.asUint8List();

          // Insert image
          final picture = sheet.pictures.addStream(1, 1, imageBytes); // 1,1 = cell A2 (0-based index)
          picture.height = 80; // Adjust height in pixels
          picture.width = 80;  // Adjust width in pixels
        }
      }



      // --- Footer ---


      final tableRange = sheet.getRangeByName('A2:H66');
      final borderStyle = tableRange.cellStyle.borders;

      borderStyle.all.lineStyle = xlsio.LineStyle.thin;
      borderStyle.all.color = '#000000'; // Black border


// Insert image from assets into A2 cell
      final ByteData imageData = await rootBundle.load('assets/logo.png');
      final List<int> imageBytes = imageData.buffer.asUint8List();

// Insert picture in A2 (row=1, col=1 in 1-based index)
      final xlsio.Picture picture = sheet.pictures.addStream(1, 1, imageBytes);

// Optional: Adjust image position and size to fit nicely
      picture.height = 130;   // Adjust height as needed
      picture.width = 130;   // Adjust width as needed
      picture.row = 2;       // Padding from top
      picture.column = 1;
      picture.columnOffset = 40; // horizontal shift inside A2 (~center)
      picture.rowOffset = 20;     // Shift image down within merged cell// Padding from left

























      // --- Save and Export the Excel File ---
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
        try {
          final blob = html.Blob(
              [bytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          final url = html.Url.createObjectUrlFromBlob(blob);
          final anchor = html.AnchorElement(href: url)
            ..setAttribute('download', '$fileNameBase.xlsx')
            ..click();
          html.Url.revokeObjectUrl(url);

          _message = 'Excel file downloaded! Check your downloads folder.';
        } catch (e) {
          _message = 'Error downloading file on web: $e';
        }
      } else {
        try {
          final String path = (await getApplicationSupportDirectory()).path;
          final String fileName = '$path/$fileNameBase.xlsx';
          final File file = File(fileName);
          try {
            await file.writeAsBytes(bytes, flush: true);
          } catch (e) {
            throw Exception('Failed to save file on mobile: $e');
          }

          try {
            final result = await OpenFilex.open(fileName);
            _message = result.type == ResultType.done
                ? 'Excel file created and opened!'
                : 'Error opening file: ${result.message}';
          } catch (e) {
            _message = 'Error opening file on mobile: $e';
          }
        } catch (e) {
          _message = 'Error saving file on mobile: $e';
        }
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

extension on xlsio.Picture {
  set rowOffset(int rowOffset) {}

  set columnOffset(int columnOffset) {}
}