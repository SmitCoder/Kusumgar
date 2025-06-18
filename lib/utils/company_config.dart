import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;

abstract class CompanyConfig {
  String get jsonPath; // Path to JSON file
  List<String> get columnHeaders; // Column headers for UI and Excel
  Map<String, String> get columnMapping; // Maps JSON keys to Excel column headers

  // Initialize the worksheet (e.g., logo, gridlines)
  Future<void> initializeSheet(xlsio.Worksheet sheet, xlsio.Workbook workbook);

  // Configure the worksheet with data and formatting
  void configureSheet(
      xlsio.Worksheet sheet,
      Map<String, String> reportDetails,
      List<Map<String, dynamic>> testData,
      xlsio.Workbook workbook,
      );
}