import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:animate_do/animate_do.dart';
import 'package:provider/provider.dart';
import '../utils/excel_generator.dart';
import 'package:data_table_2/data_table_2.dart';


// class ExcelPage extends StatelessWidget {
//   const ExcelPage({super.key});
class ExcelPage extends StatelessWidget {
  final String number;

  const ExcelPage({required this.number, super.key});


  @override
  Widget build(BuildContext context) {
    return ChangeNotifierProvider(
      create: (_) => ExcelGenerator(),
      child: Scaffold(
        appBar: AppBar(
          title: const Text('Excel Generator'),
          leading: IconButton(
            icon: const Icon(Icons.arrow_back),
            onPressed: () => Navigator.pop(context),
          ),
        ),
        body: Consumer<ExcelGenerator>(
          builder: (context, excelGenerator, child) {
            final config = excelGenerator.companyConfigs[excelGenerator.selectedCompany]!;
            return Container(
              decoration: BoxDecoration(
                gradient: LinearGradient(
                  begin: Alignment.topLeft,
                  end: Alignment.bottomRight,
                  colors: [
                    Colors.blue[100]!,
                    Colors.blue[300]!,
                  ],
                ),
              ),
              child: Padding(
                padding: const EdgeInsets.all(16.0),
                child: Column(
                  children: [
                    FadeInUp(
                      child: Text(
                        'Excel Creator',
                        style: Theme.of(context).textTheme.headlineMedium?.copyWith(
                          color: Colors.white,
                          shadows: const [
                            Shadow(
                              blurRadius: 10.0,
                              color: Colors.black26,
                              offset: Offset(2.0, 2.0),
                            ),
                          ],
                        ),
                      ),
                    ),
                    const SizedBox(height: 16),
                    FadeInUp(
                      delay: const Duration(milliseconds: 200),
                      child: const Text(
                        'View your fabric test report in the app or generate an Excel file!',
                        style: TextStyle(
                          fontSize: 16,
                          color: Colors.white70,
                          fontStyle: FontStyle.italic,
                        ),
                        textAlign: TextAlign.center,
                      ),
                    ),
                    const SizedBox(height: 20),
                    FadeInUp(
                      delay: const Duration(milliseconds: 400),
                      child: DropdownButton<String>(
                        value: excelGenerator.selectedCompany,
                        onChanged: (String? newValue) {
                          excelGenerator.setSelectedCompany(newValue!);
                        },
                        items: excelGenerator.companyConfigs.keys
                            .map<DropdownMenuItem<String>>((String value) {
                          return DropdownMenuItem<String>(
                            value: value,
                            child: Text(value),
                          );
                        }).toList(),
                      ),
                    ),
                    const SizedBox(height: 20),
                    FadeInUp(
                      delay: const Duration(milliseconds: 600),
                      child: Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          ElevatedButton(
                            onPressed: excelGenerator.isLoading
                                ? null
                                : () async {
                              await excelGenerator.loadData();
                            },
                            style: ElevatedButton.styleFrom(
                              backgroundColor: Colors.blue[600],
                              foregroundColor: Colors.white,
                              padding: const EdgeInsets.symmetric(
                                horizontal: 32,
                                vertical: 16,
                              ),
                              textStyle: const TextStyle(
                                fontSize: 18,
                                fontWeight: FontWeight.bold,
                              ),
                              shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(12),
                              ),
                              elevation: 8,
                              shadowColor: Colors.blueGrey[300],
                            ),
                            child: const Text('Load Data'),
                          ),
                          const SizedBox(width: 16),
                          if (excelGenerator.dataLoaded)
                            ElevatedButton(
                              onPressed: excelGenerator.isLoading
                                  ? null
                                  : () async {
                                await excelGenerator.createExcel(context);
                              },
                              style: ElevatedButton.styleFrom(
                                backgroundColor: Colors.green[600],
                                foregroundColor: Colors.white,
                                padding: const EdgeInsets.symmetric(
                                  horizontal: 32,
                                  vertical: 16,
                                ),
                                textStyle: const TextStyle(
                                  fontSize: 18,
                                  fontWeight: FontWeight.bold,
                                ),
                                shape: RoundedRectangleBorder(
                                  borderRadius: BorderRadius.circular(12),
                                ),
                                elevation: 8,
                                shadowColor: Colors.blueGrey[300],
                              ),
                              child: const Text('Generate Excel'),
                            ),
                        ],
                      ),
                    ),
                    const SizedBox(height: 20),
                    if (excelGenerator.isLoading)
                      const CircularProgressIndicator(
                        valueColor: AlwaysStoppedAnimation<Color>(Colors.white),
                      ),
                    if (excelGenerator.message != null)
                      Padding(
                        padding: const EdgeInsets.only(top: 24),
                        child: Container(
                          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                          decoration: BoxDecoration(
                            color: excelGenerator.message!.contains('Error')
                                ? Colors.red[400]!.withOpacity(0.9)
                                : Colors.green[400]!.withOpacity(0.9),
                            borderRadius: BorderRadius.circular(8),
                          ),
                          child: Text(
                            excelGenerator.message!,
                            style: const TextStyle(
                              fontSize: 16,
                              color: Colors.white,
                              fontWeight: FontWeight.w500,
                            ),
                            textAlign: TextAlign.center,
                          ),
                        ),
                      ),
                    if (excelGenerator.dataLoaded && !excelGenerator.isLoading)
                      Expanded(
                        child: SingleChildScrollView(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                'Company: ${excelGenerator.reportDetails['Company'] ?? ''}',
                                style: const TextStyle(
                                  fontSize: 20,
                                  fontWeight: FontWeight.bold,
                                  color: Colors.white,
                                ),
                              ),
                              Text(
                                'Certification: ${excelGenerator.reportDetails['Certification'] ?? ''}',
                                style: const TextStyle(
                                  fontSize: 16,
                                  color: Colors.white70,
                                ),
                              ),
                              const SizedBox(height: 10),
                              const Text(
                                'Report Details',
                                style: TextStyle(
                                  fontSize: 20,
                                  fontWeight: FontWeight.bold,
                                  color: Colors.white,
                                ),
                              ),
                              const SizedBox(height: 10),
                              Container(
                                padding: const EdgeInsets.all(8.0),
                                decoration: BoxDecoration(
                                  color: Colors.white.withOpacity(0.9),
                                  borderRadius: BorderRadius.circular(8),
                                ),
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: excelGenerator.reportDetails.entries
                                      .map((entry) {
                                    return Padding(
                                      padding: const EdgeInsets.symmetric(vertical: 4.0),
                                      child: Text(
                                        '${entry.key}: ${entry.value}',
                                        style: const TextStyle(
                                          fontSize: 14,
                                          color: Colors.black87,
                                        ),
                                      ),
                                    );
                                  }).toList(),
                                ),
                              ),
                              const SizedBox(height: 20),
                              const Text(
                                'Test Results',
                                style: TextStyle(
                                  fontSize: 20,
                                  fontWeight: FontWeight.bold,
                                  color: Colors.white,
                                ),
                              ),
                              const SizedBox(height: 20),
                              SingleChildScrollView(
                                scrollDirection: Axis.horizontal,
                                child: DataTable(
                                  columnSpacing: 16,
                                  headingRowColor: MaterialStateColor.resolveWith((states) => Colors.blue[600]!),
                                  dataRowColor: MaterialStateColor.resolveWith((states) => Colors.white.withOpacity(0.9)),
                                  columns: config.columnHeaders.asMap().entries.map((entry) {
                                    return DataColumn(
                                      label: Text(
                                        entry.value,
                                        style: const TextStyle(color: Colors.white),
                                      ),
                                    );
                                  }).toList(),
                                  rows: excelGenerator.testData.asMap().entries.map((entry) {
                                    int rowIndex = entry.key;
                                    return DataRow(
                                      cells: config.columnHeaders.asMap().entries.map((colEntry) {
                                        int colIndex = colEntry.key;
                                        return DataCell(
                                          Padding(
                                            padding: const EdgeInsets.symmetric(vertical: 4.0),
                                            child: SizedBox(
                                              width: 150,
                                              // height: 200,

                                              child: TextFormField(
                                                controller: excelGenerator.testDataControllers[rowIndex][colIndex],
                                                style: const TextStyle(color: Colors.black87),
                                                decoration: const InputDecoration(
                                                  border: InputBorder.none,
                                                  contentPadding: EdgeInsets.symmetric(horizontal: 12, vertical: 9),
                                                  isDense: true,
                                                ),
                                                maxLines: null,
                                                minLines: 1,
                                                keyboardType: TextInputType.multiline,
                                                textAlignVertical: TextAlignVertical.top,
                                                onChanged: (value) {
                                                  (context as Element).markNeedsBuild();
                                                },
                                                onFieldSubmitted: (value) {
                                                  FocusScope.of(context).unfocus();
                                                },
                                              ),
                                            ),
                                          ),
                                        );
                                      }).toList(),
                                    );
                                  }).toList(),
                                ),
                              ),
                              // SingleChildScrollView(
                              //   scrollDirection: Axis.horizontal,
                              //   child: DataTable(
                              //     columnSpacing: 16,
                              //     headingRowColor: MaterialStateColor.resolveWith(
                              //             (states) => Colors.blue[600]!),
                              //     dataRowColor: MaterialStateColor.resolveWith(
                              //             (states) => Colors.white.withOpacity(0.9)),
                              //     dataRowHeight: 60.0, // Increased row height to accommodate more space
                              //     // verticalMargin: 8.0, // Added vertical margin between rows
                              //
                              //
                              //
                              //     columns: config.columnHeaders
                              //         .asMap()
                              //         .entries
                              //         .map((entry) => DataColumn(
                              //       label: Text(
                              //         entry.value,
                              //         style: const TextStyle(color: Colors.white),
                              //       ),
                              //     ))
                              //         .toList(),
                              //     rows: excelGenerator.testData.asMap().entries.map(
                              //           (entry) {
                              //         int rowIndex = entry.key;
                              //         return DataRow(
                              //           cells: config.columnHeaders.asMap().entries.map(
                              //                 (colEntry) {
                              //               int colIndex = colEntry.key;
                              //               return DataCell(
                              //                   Padding(
                              //                     padding: const EdgeInsets.symmetric(vertical: 4.0), // Small vertical spacing between rows
                              //                     child: SizedBox(
                              //                   width: 150,
                              //                   child: TextFormField(
                              //                     controller: excelGenerator.testDataControllers[rowIndex][colIndex],
                              //                     style: const TextStyle(color: Colors.black87),
                              //                     decoration: const InputDecoration(
                              //                       border: InputBorder.none,
                              //                       contentPadding: EdgeInsets.symmetric(horizontal: 8, vertical: 6),
                              //                     ),
                              //                     maxLines: null, // Allow unlimited lines
                              //                     minLines: 1,   // Minimum one line, expands as needed
                              //                     keyboardType: TextInputType.multiline, // Enable multi-line input
                              //                     onFieldSubmitted: (value) {
                              //                       FocusScope.of(context).unfocus();
                              //                     },
                              //                   ),
                              //                 ),
                              //                   )
                              //               );
                              //             },
                              //           ).toList(),
                              //         );
                              //       },
                              //     ).toList(),
                              //   ),
                              // ),
                              const SizedBox(height: 20),
                              ExpansionTile(
                                title: const Text(
                                  'Change Log',
                                  style: TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.bold,
                                    color: Colors.white,
                                  ),
                                ),
                                backgroundColor: Colors.white.withOpacity(0.9),
                                collapsedBackgroundColor: Colors.blue[200]!.withOpacity(0.9),
                                childrenPadding: const EdgeInsets.all(16.0),
                                children: excelGenerator.changeLog.isEmpty
                                    ? [
                                  const Text(
                                    'No changes made yet.',
                                    style: TextStyle(
                                      fontSize: 16,
                                      color: Colors.black87,
                                    ),
                                  ),
                                ]
                                    : excelGenerator.changeLog.map((log) {
                                  return Padding(
                                    padding: const EdgeInsets.symmetric(vertical: 4.0),
                                    child: Text(
                                      'Row ${log['row']}, ${log['column']}: Changed from "${log['originalValue']}" to "${log['newValue']}" at ${log['timestamp']}',
                                      style: const TextStyle(
                                        fontSize: 14,
                                        color: Colors.black87,
                                      ),
                                    ),
                                  );
                                }).toList(),
                              ),
                            ],
                          ),
                        ),
                      ),
                  ],
                ),
              ),
            );
          },
        ),
      ),
    );
  }
}