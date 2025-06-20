
import 'package:kusumgar_final/models/docx_models.dart';
import 'package:kusumgar_final/utils/companies/docx_company_config.dart';

class Company3DocxConfig implements CompanyDocxConfig {
@override
String get companyName => 'Company 3';

@override
List<DocxStyle> get styles => [
DocxStyle(id: 'Normal', name: 'Normal', fontSize: 22),
DocxStyle(id: 'Normal2', name: 'Normal1', fontSize: 20, bold: true),
DocxStyle(id: 'NormalBold', name: 'NormalBold', fontSize: 22, bold: true),
DocxStyle(id: 'Heading', name: 'Heading', fontSize: 28, bold: true, alignment: 'center', spacingBefore: 360),
DocxStyle(id: 'Justified', name: 'Justified', fontSize: 26, bold: true, alignment: 'both', spacingBefore: 360),
DocxStyle(id: 'Signature', name: 'Signature', fontSize: 27, bold: true),
DocxStyle(id: 'Normal1', name: 'Normal1', fontSize: 22, bold: true),
];

@override
List<DocxParagraph> get paragraphs => [
DocxParagraph([TextRun('Supplier:					                                       Ship to:')], style: 'Normal2'),
DocxParagraph([TextRun('KUSUMGAR LIMITED		                                                                Airborne Systems NA of CA Inc.')], style: 'Normal1'),
DocxParagraph([TextRun('Certificate of Conformance/Compliance')], style: 'Heading'),
DocxParagraph([TextRun('COC No.		      ', isBold: true), TextRun('KL/QA/ASNA/2025-2026/021I', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Customer PO No.:                 ', isBold: true), TextRun('56273, Date- 10th Dec. 2024', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Product Number:                  ', isBold: true), TextRun('CLOTH, NYL,65, FG, 24165, INT-P44378 T4-NB, (Part No.-166452)', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Color:                                       ', isBold: true), TextRun('Foliage Green', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Quality No.:                            ', isBold: true), TextRun('4201', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Fabric Lot No.:		       ', isBold: true), TextRun('24P41768 (2,035.21 Yard)', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Test report No.: 	       ', isBold: true), TextRun('Q250001330', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Date of Manufacture:           ', isBold: true), TextRun('April- 2025', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Country of Origin:                  ', isBold: true), TextRun('India', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Total Quantity:                       ', isBold: true), TextRun('4,985.73 Yard.', isBold: false)], style: 'Normal'),
DocxParagraph([TextRun('Width: 		    	        ', isBold: true), TextRun('65.0', isBold: false)], style: 'Normal'),
DocxParagraph([
TextRun(
'We hereby certify that the above material been processed in conformance to all specified requirements(PIA-C-44378E T4), including those stated on the purchase order, drawings and in specifications. Melting point is 244 Celsius min., the yarn has not been bleached. The quality control arrangements adopted in respect of these supplies have accorded with the conditions of our quality approval/registration.'),
], style: 'Justified'),
DocxParagraph([TextRun('Authorized Supplier Representative')], style: 'Signature'),
DocxParagraph([TextRun('Sign and Stamp:')], style: 'Signature', spacingAfter: 1700),
DocxParagraph([TextRun('                                                                                                                             Date:', isBold: true), TextRun('21-04-2025', isBold: false)]),
DocxParagraph([TextRun('Name:', isBold: true), TextRun(' Anubhav Shukla                                                                                   ', isBold: false), TextRun('Title:', isBold: true), TextRun(' Q.A. Sr. Manager', isBold: false)]),
];
}