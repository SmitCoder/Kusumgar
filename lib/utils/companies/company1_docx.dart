import 'package:kusumgar_final/models/docx_models.dart';
import 'package:kusumgar_final/utils/companies/docx_company_config.dart';

class Company1DocxConfig implements CompanyDocxConfig {
  @override
  String get companyName => 'Company 1';

  @override
  List<DocxStyle> get styles => [
    DocxStyle(id: 'Normal', name: 'Normal', fontSize: 22),
    DocxStyle(id: 'Heading', name: 'Heading', fontSize: 28, bold: true, alignment: 'center', spacingBefore: 360),
    DocxStyle(id: 'Signature', name: 'Signature', fontSize: 27, bold: true),
  ];

  @override
  List<DocxParagraph> get paragraphs => [
    DocxParagraph([TextRun('Certificate of Conformance')], style: 'Heading'),
    DocxParagraph([TextRun('Company: Company 1'), TextRun('Customer: Customer A')], style: 'Normal'),
    DocxParagraph([TextRun('Authorized Representative')], style: 'Signature'),
  ];
}