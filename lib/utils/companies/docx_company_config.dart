import 'package:kusumgar_final/models/docx_models.dart';

abstract class CompanyDocxConfig {
  String get companyName;
  List<DocxStyle> get styles;
  List<DocxParagraph> get paragraphs;
}