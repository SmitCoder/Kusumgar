import 'dart:typed_data';
import 'package:excel/excel.dart' as ex;
import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'excel_page.dart';
import 'docx_page.dart';
import 'package:dropdown_button2/dropdown_button2.dart';
// import 'package:excel/excel.dart' as ex;
import 'dart:io';
import 'dart:convert';
import 'package:path_provider/path_provider.dart';


class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  List<String> _companyList = [];
  String? _selectedCompany;


  Map<String, Set<String>> _companyBatchMap = {};
  List<String> _batchList = [];
  String? _selectedBatch;


  dynamic _sanitizeExcelValue(dynamic value) {
    if (value == null) return null;
    if (value is String || value is num || value is bool) return value;
    return value.toString(); // Handles SharedString, DateTime, etc.
  }



  Map<String, Map<String, Set<String>>> _companyBatchSerialMap = {};
  List<String> _serialList = [];
  String? _selectedSerial;

  late ex.Excel _excel;




  Map<String, Map<String, Map<String, String>>> _qualityOrderMap = {};
  String? get _selectedQualityOrder {
    if (_selectedCompany != null &&
        _selectedBatch != null &&
        _selectedSerial != null &&
        _qualityOrderMap[_selectedCompany!]?[_selectedBatch!]?[_selectedSerial!] != null) {
      return _qualityOrderMap[_selectedCompany!]![_selectedBatch!]![_selectedSerial!];
    }
    return null;
  }


// This is for Json files
  Map<String, dynamic> _formatRow(Map<String, dynamic> row) {
    return row.map((key, value) => MapEntry(
      key.toString().trim().toLowerCase().replaceAll(' ', '_'), // Optional formatting
      _sanitizeExcelValue(value),
    ));
  }
  Map<String, dynamic> _convertRowToMap(List<ex.Data?> header, List<ex.Data?> row) {
    Map<String, dynamic> rowMap = {};
    for (int i = 0; i < header.length; i++) {
      final key = header[i]?.value?.toString() ?? 'col_$i';
      final value = i < row.length ? row[i]?.value : null;
      rowMap[key] = value;
    }
    return rowMap;
  }





  Map<String, dynamic>? _getQualityDataRow(String qualityOrder) {
    final sheet = _excel.tables['QualityData'];
    if (sheet == null) return null;

    final rows = sheet.rows;
    if (rows.isEmpty) return null;

    final headerRow = rows[0];
    int qualityOrderColumnIndex = headerRow.indexWhere((cell) =>
    cell?.value.toString().toLowerCase().trim() == 'quality order');

    if (qualityOrderColumnIndex == -1) return null;

    for (var i = 1; i < rows.length; i++) {
      final row = rows[i];
      if (row.length > qualityOrderColumnIndex) {
        final cell = row[qualityOrderColumnIndex];
        if (cell != null &&
            cell.value.toString().trim().toLowerCase() ==
                qualityOrder.trim().toLowerCase()) {
          return _convertRowToMap(headerRow, row);
        }
      }
    }
    return null;
  }

  Map<String, dynamic> _convertRowToMap1(List<ex.Data?> header, List<ex.Data?> row) {
    Map<String, dynamic> map = {};
    for (int i = 0; i < header.length; i++) {
      final key = header[i]?.value.toString().trim();
      final value = (i < row.length) ? row[i]?.value : null;
      if (key != null && key.isNotEmpty) {
        map[key] = _sanitizeExcelValue(value); // 👈 convert all values safely
      }
    }

    return map;
  }



  // Map<String, dynamic> _generateJsonFromQualityData(
  //      qualityorder,
  //      company
  //     // String company,
  //     ) {
  //   print("qualiuty order while generating json ${qualityorder}");
  //   final sheet = _excel.tables['QualityData'];
  //   if (sheet == null) {
  //     print("❌ QualityData sheet not found");
  //     return {};
  //   }
  //
  //   final rows = sheet.rows;
  //   if (rows.isEmpty) {
  //     print("❌ No rows in QualityData sheet");
  //     return {};
  //   }
  //
  //   final headerRow = rows[0];
  //   int qualityOrderColumnIndex = headerRow.indexWhere(
  //         (cell) => cell?.value.toString().toLowerCase().trim() == 'quality order',
  //   );
  //
  //   if (qualityOrderColumnIndex == -1) {
  //     print("❌ 'Quality Order' column not found");
  //     return {};
  //   }
  //
  //   List<Map<String, dynamic>> matchingRows = [];
  //
  //   for (var i = 1; i < rows.length; i++) {
  //     final row = rows[i];
  //     if (row.length <= qualityOrderColumnIndex) continue;
  //
  //     final cell = row[qualityOrderColumnIndex];
  //     if (cell != null &&
  //         cell.value.toString().trim().toLowerCase() ==
  //             _selectedQualityOrder!.trim().toLowerCase()) {
  //       final rowMap = _convertRowToMap1(headerRow, row);
  //       matchingRows.add(_formatRow(rowMap)); // ✅ Format each row cleanly
  //     }
  //   }
  //   print('matchin rows ${matchingRows}');
  //   print('✅ Found ${matchingRows.length} matching rows for Quality Order: $_selectedQualityOrder');
  //
  //   return {
  //     // "report_info": {
  //     //   "company_name": company,
  //     //   "certification": data["Certification"] ?? "",
  //     //   "report_number": data["Report No."] ?? "",
  //     //   "report_date": data["Date"] ?? "",
  //     //   "item_name": data["Item"] ?? "",
  //     //   "quality_order_number": _selectedQualityOrder,
  //     // },
  //     "summary": {
  //       "total_matching_rows": matchingRows.length,
  //     },
  //     "rows": matchingRows,
  //   };
  // }

  final List<String> requiredKeys = ["Test","Test method" , "Test result sub line"];

  // Map<String, dynamic> _generateJsonFromQualityData(qualityorder,  company) {
  //   print("qualityorder while generating json: $qualityorder");
  //
  //   final sheet = _excel.tables['QualityData'];
  //   if (sheet == null) {
  //     print("❌ QualityData sheet not found");
  //     return {};
  //   }
  //
  //   final rows = sheet.rows;
  //   if (rows.length < 2) {
  //     print("❌ Not enough rows in QualityData sheet");
  //     return {};
  //   }
  //
  //   final headerRow = rows[0];
  //
  //   // Find the index of "Quality Order" column
  //   int qualityOrderIndex = headerRow.indexWhere((cell) =>
  //   cell?.value.toString().trim().toLowerCase() == 'quality order');
  //   if (qualityOrderIndex == -1) {
  //     print('❌ "Quality Order" column not found in header');
  //   } else {
  //     print('✅ Quality Order Index: $qualityOrderIndex');
  //   }
  //
  //   List<Map<String, dynamic>> matchedRows = [];
  //
  //   for (int i = 1; i < rows.length; i++) {
  //     final row = rows[i];
  //
  //     // Make sure row has enough columns
  //     if (row.length <= qualityOrderIndex) continue;
  //
  //     final cellValue = row[qualityOrderIndex]?.value?.toString()?.trim();
  //
  //     if (cellValue == qualityorder) {
  //       Map<String, dynamic> rowData = {};
  //
  //       for (int j = 0; j < headerRow.length; j++) {
  //         final key = headerRow[j]?.value?.toString()?.trim();
  //         final cell = j < row.length ? row[j] : null;
  //         final value = cell?.value?.toString(); // Converts SharedString to string
  //
  //         if (key != null && requiredKeys.contains(key)) {
  //           rowData[key] = value;
  //         }
  //       }
  //
  //
  //       matchedRows.add(rowData);
  //     }
  //   }
  //
  //   if (matchedRows.isEmpty) {
  //     print("⚠️ No rows matched the qualityorder: $qualityorder");
  //     return {};
  //   }
  //
  //   return {
  //     "company": company,
  //     "qualityOrder": qualityorder,
  //     "matchedRows": matchedRows,
  //   };
  // }


  Map<String, dynamic> _generateJsonFromQualityData(qualityorder, company) {
    print("qualityorder while generating json: $qualityorder");

    final sheet = _excel.tables['QualityData'];
    if (sheet == null) {
      print("❌ QualityData sheet not found");
      return {};
    }

    final rows = sheet.rows;
    if (rows.length < 2) {
      print("❌ Not enough rows in QualityData sheet");
      return {};
    }

    final headerRow = rows[0];

    // Find the index of "Quality Order" column
    int qualityOrderIndex = headerRow.indexWhere((cell) =>
    cell?.value.toString().trim().toLowerCase() == 'quality order');

    if (qualityOrderIndex == -1) {
      print('❌ "Quality Order" column not found in header');
      return {};
    } else {
      print('✅ Quality Order Index: $qualityOrderIndex');
    }

    // STEP 1: Collect matched rows
    List<Map<String, dynamic>> matchedRows = [];

    for (int i = 1; i < rows.length; i++) {
      final row = rows[i];

      if (row.length <= qualityOrderIndex) continue;

      final cellValue = row[qualityOrderIndex]?.value?.toString()?.trim();

      if (cellValue == qualityorder) {
        Map<String, dynamic> rowData = {};

        for (int j = 0; j < headerRow.length; j++) {
          final key = headerRow[j]?.value?.toString()?.trim();
          final cell = j < row.length ? row[j] : null;
          final value = cell?.value?.toString();

          if (key != null && requiredKeys.contains(key)) {
            rowData[key] = value;
          }
        }

        matchedRows.add(rowData);
      }
    }

    if (matchedRows.isEmpty) {
      print("⚠️ No rows matched the qualityorder: $qualityorder");
      return {};
    }

    // STEP 2: Group by Test
    Map<String, List<Map<String, dynamic>>> grouped = {};
    for (var row in matchedRows) {
      final test = row['Test'];
      if (test == null) continue;

      if (!grouped.containsKey(test)) {
        grouped[test] = [];
      }
      grouped[test]!.add(row);
    }

    // STEP 3: Transform grouped data into your custom format
    List<Map<String, dynamic>> finalRows = [];

    grouped.forEach((testName, groupRows) {
      String method = groupRows.first['Test method'] ?? '';
      List<String> subValues = groupRows
          .map((r) => r['Test result sub line']?.toString() ?? '')
          .toList();

      finalRows.add({
        'test': testName,
        'method': method,
        'pc_no': '',
        'empty1': subValues.length > 0 ? subValues[0] : '',
        'empty2': subValues.length > 1 ? subValues[1] : '',
        'empty3': subValues.length > 2 ? subValues[2] : '',
        'empty4': subValues.length > 3 ? subValues[3] : '',
        'empty5': subValues.length > 4 ? subValues[4] : '',
        'standard': '',
      });
    });

    // STEP 4: Return final transformed JSON
    return {
      "company": company,
      "qualityOrder": qualityorder,
      "matchedRows": finalRows,
    };
  }








  Future<void> _saveJsonFile(Map<String, dynamic> jsonData, String fileName) async {
    final directory = await getApplicationDocumentsDirectory();
    final path = '${directory.path}/$fileName.json';
    final file = File(path);
    await file.writeAsString(jsonEncode(jsonData));
    print('✅ JSON saved at: $path');
  }








  // Main function
  Future<void> generateJsonForSelectedOrder() async {
    if (_selectedQualityOrder == null || _selectedCompany == null) return;
    // print("quality order ${_selectedQualityOrder}");

    // final data = _getQualityDataRow(_selectedQualityOrder!);
    // if (data == null) {
    //   print("❌ No match in Quality Data sheet.");
    //   return;
    // }

    final jsonData = _generateJsonFromQualityData(_selectedQualityOrder , _selectedCompany);
    await _saveJsonFile(jsonData, "quality_report_${_selectedQualityOrder!}");
  }
  String _hoveredLabel = '';

  void _setHoveredLabel(String label) {
    setState(() {
      _hoveredLabel = label;
    });
  }








  @override
  void initState() {
    super.initState();
    _loadCompaniesFromExcel();
  }

  Future<void> _loadCompaniesFromExcel() async {
    try {
      final ByteData data = await rootBundle.load('assets/DataQuality.xlsx');
      final List<int> bytes = data.buffer.asUint8List();
      _excel = ex.Excel.decodeBytes(bytes); // Store for Quality Data use later

      final ex.Sheet? sheet = _excel.tables['InvoiceData'];
      if (sheet == null) throw Exception("Sheet 'InvoiceData' not found");

      _companyBatchMap.clear();
      _companyBatchSerialMap.clear();
      _qualityOrderMap.clear();

      const int startRowIndex = 1;
      for (int i = startRowIndex; i < sheet.rows.length; i++) {
        final row = sheet.rows[i];
        if (row.length < 14) continue;

        final customerCell = row[3];
        final batchCell = row[11];
        final serialCell = row[13];
        final qualityOrderCell = row.length > 21 ? row[21] : null; // Column V

        if (customerCell == null || customerCell.value == null) continue;
        final customerName = customerCell.value.toString().trim();
        if (customerName.isEmpty || customerName.toLowerCase() == 'customer name') continue;

        final batchNo = (batchCell != null && batchCell.value != null)
            ? batchCell.value.toString().trim()
            : '';
        final serialNo = (serialCell != null && serialCell.value != null)
            ? serialCell.value.toString().trim()
            : '';
        const int qualityOrderColumnIndex = 21;
        final qualityOrder = (row.length > qualityOrderColumnIndex && row[qualityOrderColumnIndex] != null)
            ? row[qualityOrderColumnIndex]!.value.toString()
            : '';

        _companyBatchMap.putIfAbsent(customerName, () => <String>{});
        _companyBatchMap[customerName]!.add(batchNo);

        _companyBatchSerialMap.putIfAbsent(customerName, () => {});
        _companyBatchSerialMap[customerName]!.putIfAbsent(batchNo, () => <String>{});
        if (serialNo.isNotEmpty) {
          _companyBatchSerialMap[customerName]![batchNo]!.add(serialNo);
        }

        _qualityOrderMap.putIfAbsent(customerName, () => {});
        _qualityOrderMap[customerName]!.putIfAbsent(batchNo, () => {});
        _qualityOrderMap[customerName]![batchNo]![serialNo] = qualityOrder;
      }

      _companyList = _companyBatchMap.keys.toList()..sort();
      setState(() {
        _batchList = [];
        _selectedCompany = null;
        _selectedBatch = null;
        _selectedSerial = null;
      });
    } catch (e) {
      print('❌ Error reading Excel: $e');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(
            colors: [Colors.blue.shade50, Colors.white],
            begin: Alignment.topCenter,
            end: Alignment.bottomCenter,
          ),
        ),
        padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 32),
        child: Center(
          child: SingleChildScrollView(
            child: Column(
              children: [
                Image.asset(
                  'assets/logo.png',
                  height: 150,
                ),
                const SizedBox(height: 32),
                Text(
                  'Document Generator',
                  style: Theme.of(context).textTheme.headlineLarge,
                  textAlign: TextAlign.center,
                ),
                const SizedBox(height: 16),
                Text(
                  'Select a company to generate Excel or Word documents.',
                  style: Theme.of(context).textTheme.bodyLarge?.copyWith(
                    color: Colors.grey.shade600,
                  ),
                  textAlign: TextAlign.center,
                ),
                const SizedBox(height: 32),

                _companyList.isEmpty
                    ? const CircularProgressIndicator()
                    :Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    MouseRegion(
                      onEnter: (_) => _setHoveredLabel('company'),
                      onExit: (_) => _setHoveredLabel(''),
                      child: GestureDetector(
                        onTap: () => _setHoveredLabel('company'), // simulate hover on tap
                        child: Text(
                          "Company Name",
                          style: TextStyle(
                            fontSize: 16,
                            fontWeight: FontWeight.bold,
                            color: _hoveredLabel == 'company' ? Colors.blue : Colors.black,
                            decoration: _hoveredLabel == 'company' ? TextDecoration.underline : TextDecoration.none,
                          ),
                        ),
                      ),
                    ),

                    // Text("Company Name", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                    SizedBox(height: 8),
                    DropdownButton2<String>(
                      isExpanded: true,
                      hint: Text('Select a company'),
                      value: _selectedCompany,
                      items: _companyList.map((c) => DropdownMenuItem(value: c, child: Text(c))).toList(),
                      onChanged: (value) {
                        setState(() {
                          _selectedCompany = value;
                          if (value != null && _companyBatchMap.containsKey(value)) {
                            _batchList = _companyBatchMap[value]!.toList();
                            _batchList.sort();
                          } else {
                            _batchList = [];
                          }
                          _selectedBatch = null;
                        });
                      },

                    ),
                    SizedBox(height: 16),
                    if (_selectedCompany != null)
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          MouseRegion(
                            onEnter: (_) => _setHoveredLabel('batch'),
                            onExit: (_) => _setHoveredLabel(''),
                            child: GestureDetector(
                              onTap: () => _setHoveredLabel('batch'),
                              child: Text(
                                "Batch Number",
                                style: TextStyle(
                                  fontSize: 16,
                                  fontWeight: FontWeight.bold,
                                  color: _hoveredLabel == 'batch' ? Colors.blue : Colors.black,
                                  decoration: _hoveredLabel == 'batch' ? TextDecoration.underline : TextDecoration.none,
                                ),
                              ),
                            ),
                          ),

                          // Text("Batch Number", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                          SizedBox(height: 8),
                          DropdownButton2<String>(
                            isExpanded: true,
                            hint: const Text('Select batch number'),
                            value: _selectedBatch,
                            items: _batchList
                                .map((batch) => DropdownMenuItem<String>(
                              value: batch,
                              child: Text(batch),
                            ))
                                .toList(),
                            onChanged: (value) {
                              setState(() {
                                _selectedBatch = value;
                                _serialList = (_selectedCompany != null &&
                                    value != null &&
                                    _companyBatchSerialMap[_selectedCompany!]
                                        ?.containsKey(value) == true)
                                    ? _companyBatchSerialMap[_selectedCompany!]![value]!.toList()
                                    : [];
                                _serialList.sort();
                                _selectedSerial = null;
                              });
                            },
                            buttonStyleData: ButtonStyleData(
                              height: 40,
                              padding: const EdgeInsets.symmetric(horizontal: 12),
                              decoration: BoxDecoration(
                                borderRadius: BorderRadius.circular(4),
                              ),
                            ),
                            dropdownStyleData: DropdownStyleData(
                              maxHeight: 250,
                              width: MediaQuery.of(context).size.width - 48,
                              decoration: BoxDecoration(
                                borderRadius: BorderRadius.circular(8),
                                color: Colors.white,
                              ),
                            ),
                          ),
                          SizedBox(height: 16),

                          // 👇 Add this block to show Serial Number dropdown
                          if (_serialList.isNotEmpty)
                            Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                MouseRegion(
                                  onEnter: (_) => _setHoveredLabel('serial'),
                                  onExit: (_) => _setHoveredLabel(''),
                                  child: GestureDetector(
                                    onTap: () => _setHoveredLabel('serial'),
                                    child: Text(
                                      "Serial Number",
                                      style: TextStyle(
                                        fontSize: 16,
                                        fontWeight: FontWeight.bold,
                                        color: _hoveredLabel == 'serial' ? Colors.blue : Colors.black,
                                        decoration: _hoveredLabel == 'serial' ? TextDecoration.underline : TextDecoration.none,
                                      ),
                                    ),
                                  ),
                                ),

                                // Text("Serial Number", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                                SizedBox(height: 8),
                                DropdownButton2<String>(
                                  isExpanded: true,
                                  hint: const Text('Select serial number'),
                                  value: _selectedSerial,
                                  items: _serialList
                                      .map((serial) => DropdownMenuItem<String>(
                                    value: serial,
                                    child: Text(serial),
                                  ))
                                      .toList(),
                                  onChanged: (value) {
                                    setState(() {
                                      _selectedSerial = value;
                                    });
                                  },
                                  buttonStyleData: ButtonStyleData(
                                    height: 40,
                                    padding: const EdgeInsets.symmetric(horizontal: 12),
                                    decoration: BoxDecoration(
                                      borderRadius: BorderRadius.circular(4),
                                    ),
                                  ),
                                  dropdownStyleData: DropdownStyleData(
                                    maxHeight: 250,
                                    width: MediaQuery.of(context).size.width - 48,
                                    decoration: BoxDecoration(
                                      borderRadius: BorderRadius.circular(8),
                                      color: Colors.white,
                                    ),
                                  ),
                                ),
                              ],
                            ),
                        ],
                      ),
                  ],),

                const SizedBox(height: 32),

                if (_selectedCompany != null && _selectedBatch != null && _selectedSerial != null)


                // if (_selectedCompany != null)
                  Column(
                    children: [
                      Text(
                        'Selected: ${_selectedCompany ?? 'None'}  ,  ${_selectedBatch ?? 'None'}  , ${_selectedSerial ?? 'None'}  ,  ${_selectedQualityOrder ?? 'null'}',
                        style: const TextStyle(
                          fontSize: 18,
                          fontWeight: FontWeight.bold,
                        ),
                      ),
                      const SizedBox(height: 24),

                      // Row of buttons
                      Row(
                        mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                        children: [
                          Expanded(
                            child: ElevatedButton.icon(
                              icon: const Icon(Icons.table_chart),
                              label: const Text('Excel'),
                              onPressed: () {
                                Navigator.push(
                                  context,
                                  MaterialPageRoute(
                                    builder: (context) => ExcelPage(number: _selectedCompany!),
                                  ),
                                );
                              },
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: ElevatedButton.icon(
                              icon: const Icon(Icons.description),
                              label: const Text('Docx'),
                              onPressed: () {
                                Navigator.push(
                                  context,
                                  MaterialPageRoute(
                                    builder: (context) => DocxPage(number: _selectedCompany!),
                                  ),
                                );
                              },
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: ElevatedButton.icon(
                              icon: const Icon(Icons.save),
                              label: const Text('JSON'),
                              onPressed: () async {
                                await generateJsonForSelectedOrder();
                              },
                            ),
                          ),
                        ],
                      ),
                    ],
                  )



              ],
            ),
          ),
        ),
      ),
    );
  }
}