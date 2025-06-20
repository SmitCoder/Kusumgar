import 'package:kusumgar_final/models/docx_models.dart';
import 'package:kusumgar_final/utils/companies/docx_company_config.dart';

class Company2DocxConfig implements CompanyDocxConfig {
  @override
  String get companyName => 'Company 2';

  @override
  List<DocxStyle> get styles => [
    DocxStyle(id: 'Normal', name: 'Normal', fontSize: 20),
    DocxStyle(id: 'Heading', name: 'Heading', fontSize: 26, bold: true, alignment: 'center', spacingBefore: 240),
    DocxStyle(id: 'Signature', name: 'Signature', fontSize: 25, bold: true),
  ];

  @override
  List<DocxParagraph> get paragraphs => [
    DocxParagraph([TextRun('Certificate of Compliance')], style: 'Heading'),
    DocxParagraph([TextRun('Company: Company 2'), TextRun('Customer: Customer B')], style: 'Normal'),
    DocxParagraph([TextRun('Authorized Signatory')], style: 'Signature'),
  ];
}