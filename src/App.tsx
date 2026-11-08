import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Globe,
  ArrowLeftRight,
  FileCode,
  Download,
  CheckCircle2,
  AlertTriangle,
  Play,
  SkipForward,
  PauseCircle,
  Search,
  ShieldCheck,
  Sparkles,
  Eye,
  EyeOff,
  Mic,
  ScanSearch
} from 'lucide-react';
import { PythonProjectHub } from './components/PythonProjectHub';

interface ScannedFieldSpec {
  order: number;
  section: string;
  label: string;
  columnName: string;
  type: string;
  required: boolean;
  validationRule: string;
  options: string[];
  selector: string;
}

const SCANNED_PORTAL_FIELDS: ScannedFieldSpec[] = [
  {
    order: 1,
    section: 'Step 1: Personal & Location Details',
    label: 'Applicant Full Name (आवेदक का नाम) *',
    columnName: 'Applicant_Name',
    type: 'text',
    required: true,
    validationRule: 'Mandatory Text',
    options: [],
    selector: '#applicant_name'
  },
  {
    order: 2,
    section: 'Step 1: Personal & Location Details',
    label: 'Father / Husband Name (पिता/पति का नाम) *',
    columnName: 'Father_Name',
    type: 'text',
    required: true,
    validationRule: 'Mandatory Text',
    options: [],
    selector: '#father_name'
  },
  {
    order: 3,
    section: 'Step 1: Personal & Location Details',
    label: 'Date of Birth (जन्म तिथि - DD/MM/YYYY) *',
    columnName: 'DOB',
    type: 'text',
    required: true,
    validationRule: 'Date DD/MM/YYYY',
    options: [],
    selector: '#dob'
  },
  {
    order: 4,
    section: 'Step 1: Personal & Location Details',
    label: 'Gender (लिंग) *',
    columnName: 'Gender',
    type: 'radio',
    required: true,
    validationRule: 'Allowed Options',
    options: ['Male', 'Female', 'Other'],
    selector: 'input[type="radio"][name="gender"]'
  },
  {
    order: 5,
    section: 'Step 1: Personal & Location Details',
    label: 'Mobile Number (मोबाइल नंबर - 10 Digits) *',
    columnName: 'Mobile',
    type: 'text',
    required: true,
    validationRule: '10-Digit Mobile',
    options: [],
    selector: '#mobile'
  },
  {
    order: 6,
    section: 'Step 1: Personal & Location Details',
    label: 'Aadhaar / Jan-Aadhaar No (12 Digits) *',
    columnName: 'Aadhaar',
    type: 'text',
    required: true,
    validationRule: '12-Digit Aadhaar',
    options: [],
    selector: '#aadhaar'
  },
  {
    order: 7,
    section: 'Step 1: Personal & Location Details',
    label: 'Social Category (श्रेणी) *',
    columnName: 'Category',
    type: 'select',
    required: true,
    validationRule: 'Dropdown Options',
    options: ['General', 'OBC', 'SC', 'ST'],
    selector: '#category'
  },
  {
    order: 8,
    section: 'Step 1: Personal & Location Details',
    label: 'PIN Code (पिन कोड - 6 Digits) *',
    columnName: 'PIN_Code',
    type: 'text',
    required: true,
    validationRule: '6-Digit PIN',
    options: [],
    selector: '#pincode'
  },
  {
    order: 9,
    section: 'Step 1: Personal & Location Details',
    label: 'District (जिला) *',
    columnName: 'District',
    type: 'select',
    required: true,
    validationRule: 'Parent Dropdown',
    options: ['Jaipur', 'Jodhpur', 'Udaipur'],
    selector: '#district'
  },
  {
    order: 10,
    section: 'Step 1: Personal & Location Details',
    label: 'Block / Tehsil (ब्लॉक/तहसील) *',
    columnName: 'Block',
    type: 'select',
    required: true,
    validationRule: 'Cascading (Depends on District)',
    options: ['Sanganer', 'Amer', 'Mandore', 'Osian', 'Girwa', 'Mavli'],
    selector: '#block'
  },
  {
    order: 11,
    section: 'Step 1: Personal & Location Details',
    label: 'Gram Panchayat / Village (गाँव/वार्ड) *',
    columnName: 'Village',
    type: 'select',
    required: true,
    validationRule: 'Cascading (Depends on Block)',
    options: ['Muhana', 'Bagru', 'Achrol', 'Boranada', 'Bedla', 'Dabok'],
    selector: '#village'
  },
  {
    order: 12,
    section: 'Step 2: Income, Address & Upload',
    label: 'Annual Family Income (वार्षिक आय ₹)',
    columnName: 'Annual_Income',
    type: 'number',
    required: false,
    validationRule: 'Numeric',
    options: [],
    selector: '#annual_income'
  },
  {
    order: 13,
    section: 'Step 2: Income, Address & Upload',
    label: 'Permanent Address (स्थायी पता)',
    columnName: 'Address',
    type: 'text',
    required: false,
    validationRule: 'Text',
    options: [],
    selector: '#address'
  },
  {
    order: 14,
    section: 'Step 2: Income, Address & Upload',
    label: 'Applicant Passport Photo (फोटो अपलोड ≤ 50KB)',
    columnName: 'Photo_File',
    type: 'file',
    required: false,
    validationRule: 'Image <= 50KB (JPG)',
    options: [],
    selector: '#applicant_photo'
  }
];

interface ApplicantRow {
  Serial_No: string;
  Application_No: string;
  Applicant_Name: string;
  Father_Name: string;
  DOB: string;
  Gender: string;
  Mobile: string;
  Aadhaar: string;
  Category: string;
  PIN_Code: string;
  District: string;
  Block: string;
  Village: string;
  Annual_Income: number;
  Photo_File: string;
  Status: string;
  Portal_Ref_No?: string;
  Portal_Amount?: string;
  Portal_Date?: string;
  Portal_Remarks?: string;
}

const INITIAL_APPLICANTS: ApplicantRow[] = [
  {
    Serial_No: '101',
    Application_No: 'RAJ-2025-101',
    Applicant_Name: 'Ramesh Kumar',
    Father_Name: 'Mohan Lal',
    DOB: '15/08/1992',
    Gender: 'Male',
    Mobile: '9876543210',
    Aadhaar: '456789012345',
    Category: 'OBC',
    PIN_Code: '302029',
    District: 'Jaipur',
    Block: 'Sanganer',
    Village: 'Muhana',
    Annual_Income: 45000,
    Photo_File: '101_Ramesh_Kumar.jpg (38 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '12500'
  },
  {
    Serial_No: '102',
    Application_No: 'RAJ-2025-102',
    Applicant_Name: 'Sunita Devi',
    Father_Name: 'Sohan Lal',
    DOB: '20/11/1995',
    Gender: 'Female',
    Mobile: '9123456789',
    Aadhaar: '789012345678',
    Category: 'General',
    PIN_Code: '342001',
    District: 'Jodhpur',
    Block: 'Mandore',
    Village: 'Boranada',
    Annual_Income: 38000,
    Photo_File: '102_Sunita_Devi.jpg (41 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '15000'
  },
  {
    Serial_No: '103',
    Application_No: 'RAJ-2025-103',
    Applicant_Name: 'Mahesh Sharma',
    Father_Name: 'Kailash Sharma',
    DOB: '05/03/1989',
    Gender: 'Male',
    Mobile: '9988776655',
    Aadhaar: '654321098765',
    Category: 'General',
    PIN_Code: '313001',
    District: 'Udaipur',
    Block: 'Girwa',
    Village: 'Bedla',
    Annual_Income: 52000,
    Photo_File: '103_Mahesh_Sharma.jpg (36 KB)',
    Status: 'Pending Fill',
    Portal_Amount: '10000'
  },
  {
    Serial_No: '104',
    Application_No: 'RAJ-2025-104',
    Applicant_Name: 'Kavita Meena',
    Father_Name: 'Ramkishan Meena',
    DOB: '01/12/1997',
    Gender: 'Female',
    Mobile: '98765', // Intentional 5-digit validation error
    Aadhaar: '12345678', // Intentional 8-digit validation error
    Category: 'ST',
    PIN_Code: '3020', // Intentional 4-digit validation error
    District: 'Jaipur',
    Block: 'Amer',
    Village: 'Achrol',
    Annual_Income: 40000,
    Photo_File: '104_Kavita_Meena.jpg',
    Status: 'Validation Error',
    Portal_Amount: '12500'
  }
];

const PORTAL_LIVE_DATABASE: Record<string, { status: string; date: string; amount: string; remarks: string }> = {
  'RAJ-2025-101': { status: 'Approved', date: '14/02/2025', amount: '12500', remarks: 'Verified by Tehsildar Sanganer' },
  'RAJ-2025-102': { status: 'Approved', date: '18/02/2025', amount: '15000', remarks: 'Direct Benefit Transfer Completed' },
  'RAJ-2025-103': { status: 'Pending Verification', date: '21/02/2025', amount: '11500', remarks: 'Amount revised by Gram Sevak' },
  'RAJ-2025-104': { status: 'Rejected', date: '25/02/2025', amount: '0', remarks: 'Income certificate expired' }
};

function maskPii(val: string): string {
  const digits = String(val || '').replace(/\D/g, '');
  if (digits.length <= 4) return val;
  return 'X'.repeat(digits.length - 4) + digits.slice(-4);
}

export default function App() {
  const [activeMode, setActiveMode] = useState<
    'scan_make_excel' | 'fill_from_excel' | 'web_to_excel' | 'fill_verify' | 'python_code'
  >('scan_make_excel');
  const [scanUrl, setScanUrl] = useState<string>('file:///samples/demo_portal.html');
  const [rows, setRows] = useState<ApplicantRow[]>(INITIAL_APPLICANTS);
  const [currentRowIdx, setCurrentRowIdx] = useState<number>(0);
  const [maskSensitive, setMaskSensitive] = useState<boolean>(true);
  const [filledInBrowser, setFilledInBrowser] = useState<boolean>(false);
  const [pausedForOtp, setPausedForOtp] = useState<boolean>(false);
  const [voiceCommand, setVoiceCommand] = useState<string>('');
  const [statusBanner, setStatusBanner] = useState<{ en: string; hi: string }>({
    en: 'URL Scan Report Ready: Form fields found: 14, Tables found: 1, Uploads: 1. Download the generated Excel format or test filling below.',
    hi: 'URL स्कैन रिपोर्ट तैयार: 14 फॉर्म फील्ड्स, 1 टेबल और 1 फोटो अपलोड मिला। रेडी-टू-फिल एक्सेल फॉर्मेट डाउनलोड करें।'
  });

  const currentRow = rows[currentRowIdx] || rows[0];

  const getValidationIssues = (row: ApplicantRow) => {
    const issues: string[] = [];
    const mobLen = row.Mobile.replace(/\D/g, '').length;
    const aadLen = row.Aadhaar.replace(/\D/g, '').length;
    const pinLen = row.PIN_Code.replace(/\D/g, '').length;
    if (mobLen !== 10) issues.push(`mobile number ${mobLen} digit ka hai, 10 hona chahiye`);
    if (aadLen !== 12) issues.push(`Aadhaar ${aadLen} digit ka hai, 12 hona chahiye`);
    if (pinLen !== 6) issues.push(`PIN code ${pinLen} digit ka hai, 6 hona chahiye`);
    return issues;
  };

  // Step 2: Generate 3-Sheet Excel Template (.xlsx) matching website fields
  const handleDownloadGeneratedTemplate = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Form_Data_Template (Header row + Hint row + Blank rows)
    const headers = [
      ...SCANNED_PORTAL_FIELDS.map(f => f.columnName),
      'Status',
      'Application_Acknowledgement_No',
      'Submission_Time',
      'Error_Message'
    ];
    const hintRow = [
      ...SCANNED_PORTAL_FIELDS.map(
        f => `${f.required ? 'Required *' : 'Optional'} | ${f.validationRule}${f.options.length ? ' [' + f.options.slice(0, 3).join(', ') + ']' : ''}`
      ),
      'Auto-filled Status',
      'Auto-filled Application No',
      'Auto-filled Timestamp',
      'Auto-filled Error'
    ];
    const wsTemplate = XLSX.utils.aoa_to_sheet([headers, hintRow]);
    XLSX.utils.book_append_sheet(wb, wsTemplate, 'Form_Data_Template');

    // Sheet 2: Instructions (Hindi + English)
    const instructions = [
      ['URL-First Excel <-> Website Automation Assistant — Instructions / निर्देश'],
      [`Scanned Website URL: ${scanUrl}`],
      [''],
      ['हिंदी निर्देश:'],
      ['1. Form_Data_Template शीट में पंक्ति 3 से अपना डेटा भरना शुरू करें।'],
      ['2. मोबाइल नंबर 10 अंक, आधार 12 अंक, पिन कोड 6 अंक और जन्म तिथि DD/MM/YYYY फॉर्मेट में लिखें।'],
      ['3. ड्रॉपडाउन वाले कॉलम के मान्य विकल्प Allowed Values शीट में दिए गए हैं।'],
      ['4. आखिरी 4 कॉलम (Status, Application_Acknowledgement_No, Submission_Time, Error_Message) खाली छोड़ दें।']
    ];
    const wsInst = XLSX.utils.aoa_to_sheet(instructions);
    XLSX.utils.book_append_sheet(wb, wsInst, 'Instructions');

    // Sheet 3: Allowed Values
    const dropdownFields = SCANNED_PORTAL_FIELDS.filter(f => f.options.length > 0);
    const allowedHeaders = dropdownFields.map(f => f.columnName);
    const maxOpts = Math.max(...dropdownFields.map(f => f.options.length), 1);
    const allowedRows: string[][] = [allowedHeaders];
    for (let r = 0; r < maxOpts; r++) {
      allowedRows.push(dropdownFields.map(f => f.options[r] || ''));
    }
    const wsAllowed = XLSX.utils.aoa_to_sheet(allowedRows);
    XLSX.utils.book_append_sheet(wb, wsAllowed, 'Allowed Values');

    XLSX.writeFile(wb, 'SSO_New_Registration_Template.xlsx');
    setStatusBanner({
      en: 'Downloaded 3-sheet Excel Template (Form_Data_Template + Instructions + Allowed Values)!',
      hi: '3-शीट वाला रेडी-टू-फिल एक्सेल फॉर्मेट (Form_Data_Template + Instructions + Allowed Values) डाउनलोड हो गया!'
    });
  };

  const handleFixValidationIssues = () => {
    setRows(prev =>
      prev.map(r =>
        r.Serial_No === '104'
          ? { ...r, Mobile: '9876501234', Aadhaar: '890123456789', PIN_Code: '302028', Status: 'Pending Fill' }
          : r
      )
    );
    setStatusBanner({
      en: 'Fixed Row #4 validation errors (Mobile 10 digits, Aadhaar 12 digits, PIN 6 digits). Ready to fill!',
      hi: 'पंक्ति #4 की सभी त्रुटियाँ (मोबाइल, आधार, पिन कोड) सुधार दी गई हैं!'
    });
  };

  const handleFillRow = () => {
    if (pausedForOtp) return;
    const issues = getValidationIssues(currentRow);
    if (issues.length > 0) {
      setStatusBanner({
        en: `Row ${currentRowIdx + 1}: ${issues.join(' | ')}`,
        hi: `Row ${currentRowIdx + 1}: ${issues[0]}`
      });
      return;
    }
    setFilledInBrowser(true);
    setStatusBanner({
      en: `Filled Row #${currentRowIdx + 1} (${currentRow.Applicant_Name}) across Step 1 & Step 2 including cascading District → Block → Village & compressed photo. Stopped before Submit.`,
      hi: `पंक्ति #${currentRowIdx + 1} (${currentRow.Applicant_Name}) फॉर्म में भर दी गई है। कृपया जांचें और 'Approve & Submit' दबाएं।`
    });
  };

  const handleApproveAndSubmit = () => {
    const ackNo = `RAJ-2025-${8800 + currentRowIdx + 1}`;
    setRows(prev =>
      prev.map((r, i) => (i === currentRowIdx ? { ...r, Status: 'Submitted', Portal_Ref_No: ackNo } : r))
    );
    setFilledInBrowser(false);
    setStatusBanner({
      en: `Submitted Row #${currentRowIdx + 1}! Captured Acknowledgement No: ${ackNo} and saved proof screenshot.`,
      hi: `पंक्ति #${currentRowIdx + 1} सफलतापूर्वक सबमिट हो गई! एप्लीकेशन नंबर: ${ackNo} सेव कर लिया गया।`
    });
    if (currentRowIdx < rows.length - 1) {
      setCurrentRowIdx(currentRowIdx + 1);
    }
  };

  const handleRunSearchAndFetch = () => {
    setRows(prev =>
      prev.map(r => {
        const live = PORTAL_LIVE_DATABASE[r.Application_No];
        if (!live) return r;
        return {
          ...r,
          Status: live.status,
          Portal_Date: live.date,
          Portal_Amount: live.amount,
          Portal_Remarks: live.remarks
        };
      })
    );
    setStatusBanner({
      en: 'Search-and-Fetch completed! Queried all 4 Application IDs on the portal and wrote Status, Date, Amount, and Remarks into Excel.',
      hi: 'Search-and-Fetch पूरा हुआ! सभी 4 एप्लीकेशन नंबर पोर्टल पर सर्च करके उनका स्टेटस, तारीख और राशि एक्सेल में जोड़ दी गई है।'
    });
  };

  const handleExportCurrentExcel = (filename: string) => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Automation_Results');
    XLSX.writeFile(wb, filename);
  };

  const handleVoiceCommandSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const cmd = voiceCommand.toLowerCase();
    if (cmd.includes('scan')) {
      setActiveMode('scan_make_excel');
      setStatusBanner({
        en: 'Scanned URL: Form fields found: 14, Tables found: 1, Uploads: 1.',
        hi: 'URL स्कैन पूरा हुआ: 14 फॉर्म फील्ड्स, 1 टेबल और 1 फोटो अपलोड मिला।'
      });
    } else if (cmd.includes('format') || cmd.includes('template')) {
      handleDownloadGeneratedTemplate();
    } else if (cmd.includes('agli') || cmd.includes('next') || cmd.includes('bharo') || cmd.includes('pehli')) {
      setActiveMode('fill_from_excel');
      handleFillRow();
    } else if (cmd.includes('submit') || cmd.includes('approve')) {
      handleApproveAndSubmit();
    } else if (cmd.includes('status') || cmd.includes('table') || cmd.includes('nikalo')) {
      setActiveMode('web_to_excel');
      handleRunSearchAndFetch();
    }
    setVoiceCommand('');
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* Top Bar following 3-Zone Contract */}
      <header className="flex items-center justify-between gap-8 px-6 py-3.5 border-b border-slate-800 bg-slate-950 sticky top-0 z-50">
        <a
          href="#top"
          onClick={e => {
            e.preventDefault();
            setActiveMode('scan_make_excel');
          }}
          className="text-base font-bold tracking-tight text-white whitespace-nowrap shrink-0"
        >
          URL-First Excel ↔ Website Assistant
        </a>

        <nav className="hidden lg:flex items-center gap-6 text-xs font-medium text-slate-300">
          <button
            onClick={() => setActiveMode('scan_make_excel')}
            className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeMode === 'scan_make_excel' ? 'text-emerald-400 font-semibold underline underline-offset-4' : ''
            }`}
          >
            Scan & Make Excel Format
          </button>
          <button
            onClick={() => setActiveMode('fill_from_excel')}
            className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeMode === 'fill_from_excel' ? 'text-emerald-400 font-semibold underline underline-offset-4' : ''
            }`}
          >
            Fill Website from Excel
          </button>
          <button
            onClick={() => setActiveMode('web_to_excel')}
            className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeMode === 'web_to_excel' ? 'text-emerald-400 font-semibold underline underline-offset-4' : ''
            }`}
          >
            Website to Excel
          </button>
          <button
            onClick={() => setActiveMode('fill_verify')}
            className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeMode === 'fill_verify' ? 'text-emerald-400 font-semibold underline underline-offset-4' : ''
            }`}
          >
            Fill + Verify
          </button>
          <button
            onClick={() => setActiveMode('python_code')}
            className={`hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeMode === 'python_code' ? 'text-amber-400 font-semibold underline underline-offset-4' : ''
            }`}
          >
            Python Code (19 Files)
          </button>
        </nav>

        <div className="flex items-center gap-3 shrink-0">
          <a
            href="/excel_website_automation_assistant.zip"
            download="excel_website_automation_assistant.zip"
            className="px-4 py-2 text-xs font-semibold text-slate-950 bg-amber-400 hover:bg-amber-300 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download Python .ZIP</span>
          </a>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 space-y-5">
        {/* Workspace Toolbar: 4-Mode Segmented Control + Voice Bar + PII Toggle */}
        <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl flex flex-wrap items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-1.5 bg-slate-950 p-1 rounded-lg border border-slate-800">
            <button
              onClick={() => setActiveMode('scan_make_excel')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 flex items-center gap-1.5 ${
                activeMode === 'scan_make_excel' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <ScanSearch className="w-3.5 h-3.5" />
              <span>1. Scan & Make Excel Format</span>
            </button>

            <button
              onClick={() => setActiveMode('fill_from_excel')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 flex items-center gap-1.5 ${
                activeMode === 'fill_from_excel' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>2. Fill Website from Excel</span>
            </button>

            <button
              onClick={() => setActiveMode('web_to_excel')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 flex items-center gap-1.5 ${
                activeMode === 'web_to_excel' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Globe className="w-3.5 h-3.5" />
              <span>3. Website to Excel</span>
            </button>

            <button
              onClick={() => setActiveMode('fill_verify')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 flex items-center gap-1.5 ${
                activeMode === 'fill_verify' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <ArrowLeftRight className="w-3.5 h-3.5" />
              <span>4. Fill + Verify</span>
            </button>

            <button
              onClick={() => setActiveMode('python_code')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0 flex items-center gap-1.5 ${
                activeMode === 'python_code' ? 'bg-amber-400 text-slate-950 font-semibold' : 'text-amber-300 hover:text-amber-200'
              }`}
            >
              <FileCode className="w-3.5 h-3.5" />
              <span>Python Source & Hindi Guide</span>
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <form onSubmit={handleVoiceCommandSubmit} className="flex items-center gap-2">
              <div className="relative">
                <Mic className="w-3.5 h-3.5 text-emerald-400 absolute left-2.5 top-2" />
                <input
                  type="text"
                  value={voiceCommand}
                  onChange={e => setVoiceCommand(e.target.value)}
                  placeholder="Voice: 'is URL ko scan karo', 'Excel format banao'..."
                  className="pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 w-64 focus:outline-none focus:border-emerald-500"
                />
              </div>
              <button
                type="submit"
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-medium cursor-pointer whitespace-nowrap shrink-0"
              >
                Run Command
              </button>
            </form>

            <button
              onClick={() => setMaskSensitive(v => !v)}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-950 hover:bg-slate-800 border border-slate-800 rounded-lg text-xs text-slate-300 cursor-pointer whitespace-nowrap shrink-0"
            >
              {maskSensitive ? <EyeOff className="w-3.5 h-3.5 text-emerald-400" /> : <Eye className="w-3.5 h-3.5 text-amber-400" />}
              <span>{maskSensitive ? 'PII Masked' : 'PII Visible'}</span>
            </button>
          </div>
        </div>

        {/* Status Feedback Bar */}
        <div className="px-4 py-3 bg-slate-900 border border-slate-800 rounded-xl flex flex-wrap items-center justify-between gap-4">
          <div className="space-y-0.5">
            <div className="text-xs font-semibold text-emerald-400">{statusBanner.en}</div>
            <div className="text-xs text-slate-400">{statusBanner.hi}</div>
          </div>
          <div className="text-xs text-slate-400 font-mono tabular-nums">
            Form fields found: 14 · Tables found: 1 · Uploads: 1
          </div>
        </div>

        {activeMode === 'python_code' ? (
          <PythonProjectHub />
        ) : activeMode === 'scan_make_excel' ? (
          /* =================================================================
             MODE 1: SCAN & MAKE EXCEL FORMAT (STEPS 1 & 2)
             ================================================================= */
          <div className="space-y-5">
            <div className="p-5 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div>
                  <h2 className="text-base font-semibold text-white">
                    Step 1 & Step 2: Scan Website URL & Auto-Generate Excel Format (`scanner.py` + `template_maker.py`)
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Scans all inputs, dropdowns (including cascading District → Block → Village), radios, checkboxes, uploads & tables, then builds a 3-sheet Excel format with DataValidation dropdown lists.
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2.5">
                  <button
                    onClick={handleDownloadGeneratedTemplate}
                    className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
                  >
                    <Download className="w-4 h-4" />
                    <span>Download Generated Excel Format (.xlsx)</span>
                  </button>
                  <button
                    onClick={() => setActiveMode('fill_from_excel')}
                    className="flex items-center gap-1.5 px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-medium cursor-pointer whitespace-nowrap shrink-0"
                  >
                    <span>Next: Fill Website from Excel →</span>
                  </button>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-3 pt-1">
                <input
                  type="text"
                  value={scanUrl}
                  onChange={e => setScanUrl(e.target.value)}
                  className="flex-1 min-w-[260px] px-3 py-2 bg-slate-950 border border-slate-800 rounded-lg text-xs font-mono text-slate-200"
                />
                <button
                  onClick={() =>
                    setStatusBanner({
                      en: `Scanned ${scanUrl} — Form fields found: 14, Tables found: 1, Uploads: 1. Gemini grouped fields & inferred validation rules.`,
                      hi: `${scanUrl} स्कैन किया गया — 14 फॉर्म फील्ड्स, 1 टेबल और 1 फोटो अपलोड मिला।`
                    })
                  }
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
                >
                  Scan This Page
                </button>
              </div>

              <div className="overflow-x-auto border border-slate-800 rounded-lg">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                    <tr>
                      <th className="p-2.5">#</th>
                      <th className="p-2.5">Form Section</th>
                      <th className="p-2.5">Website Field Label (Hindi/English)</th>
                      <th className="p-2.5">Generated Excel Column</th>
                      <th className="p-2.5">Type</th>
                      <th className="p-2.5">Required</th>
                      <th className="p-2.5">Validation & Allowed Options</th>
                      <th className="p-2.5">CSS Selector</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/80">
                    {SCANNED_PORTAL_FIELDS.map(f => (
                      <tr key={f.order} className="hover:bg-slate-950/60">
                        <td className="p-2.5 font-mono tabular-nums text-slate-400">{f.order}</td>
                        <td className="p-2.5 text-slate-400">{f.section}</td>
                        <td className="p-2.5 text-slate-200 font-medium">{f.label}</td>
                        <td className="p-2.5 font-mono text-emerald-400 font-semibold">{f.columnName}</td>
                        <td className="p-2.5 font-mono text-slate-300">{f.type}</td>
                        <td className="p-2.5 text-slate-300">{f.required ? 'Required *' : 'Optional'}</td>
                        <td className="p-2.5 text-slate-300">
                          {f.validationRule}
                          {f.options.length > 0 && (
                            <span className="text-slate-400 font-mono ml-1.5">
                              [{f.options.slice(0, 4).join(', ')}]
                            </span>
                          )}
                        </td>
                        <td className="p-2.5 font-mono text-slate-400">{f.selector}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : activeMode === 'fill_from_excel' ? (
          /* =================================================================
             MODE 2: FILL WEBSITE FROM EXCEL (STEP 3)
             ================================================================= */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
            <div className="lg:col-span-7 space-y-4">
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-xl space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h2 className="text-sm font-semibold text-white">
                      Step 3: Input Excel / Word Data & Pre-Flight Validation (`validator.py`)
                    </h2>
                    <p className="text-xs text-slate-400">
                      Checks mandatory fields, 10-digit Mobile, 12-digit Aadhaar, 6-digit PIN, DD/MM/YYYY dates & allowed dropdown values.
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={handleFixValidationIssues}
                      className="px-3 py-1.5 bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 rounded-lg text-xs font-medium cursor-pointer whitespace-nowrap shrink-0"
                    >
                      Fix Row #4 Errors
                    </button>
                    <button
                      onClick={() => handleExportCurrentExcel('applicants_results_updated.xlsx')}
                      className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
                    >
                      Export Updated Excel Copy
                    </button>
                  </div>
                </div>

                <div className="overflow-x-auto border border-slate-800 rounded-lg">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                      <tr>
                        <th className="p-2.5">#</th>
                        <th className="p-2.5">Applicant_Name</th>
                        <th className="p-2.5">Mobile</th>
                        <th className="p-2.5">Aadhaar</th>
                        <th className="p-2.5">District → Block → Village</th>
                        <th className="p-2.5">Status / Ack No</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/70">
                      {rows.map((r, idx) => {
                        const issues = getValidationIssues(r);
                        const isCurrent = idx === currentRowIdx;
                        return (
                          <tr
                            key={r.Serial_No}
                            onClick={() => {
                              setCurrentRowIdx(idx);
                              setFilledInBrowser(false);
                            }}
                            className={`cursor-pointer transition-colors ${
                              isCurrent ? 'bg-emerald-950/30' : 'hover:bg-slate-950/60'
                            }`}
                          >
                            <td className="p-2.5 font-mono tabular-nums text-slate-400">{idx + 1}</td>
                            <td className="p-2.5 font-medium text-slate-200">{r.Applicant_Name}</td>
                            <td className={`p-2.5 font-mono tabular-nums ${r.Mobile.length !== 10 ? 'text-rose-400 font-bold' : 'text-slate-300'}`}>
                              {maskSensitive ? maskPii(r.Mobile) : r.Mobile}
                            </td>
                            <td className={`p-2.5 font-mono tabular-nums ${r.Aadhaar.length !== 12 ? 'text-rose-400 font-bold' : 'text-slate-300'}`}>
                              {maskSensitive ? maskPii(r.Aadhaar) : r.Aadhaar}
                            </td>
                            <td className="p-2.5 text-slate-300">
                              {r.District} → {r.Block} → {r.Village}
                            </td>
                            <td className="p-2.5">
                              {issues.length > 0 ? (
                                <span className="inline-flex items-center gap-1 text-rose-400 font-medium">
                                  <AlertTriangle className="w-3.5 h-3.5" /> {issues[0]}
                                </span>
                              ) : r.Status === 'Submitted' ? (
                                <span className="inline-flex items-center gap-1 text-emerald-400 font-mono tabular-nums">
                                  <CheckCircle2 className="w-3.5 h-3.5" /> {r.Portal_Ref_No}
                                </span>
                              ) : (
                                <span className="text-slate-400">{r.Status}</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Right Column: Simulated Headed Chromium Browser */}
            <div className="lg:col-span-5">
              <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden h-full flex flex-col justify-between">
                <div className="bg-slate-950 px-4 py-2.5 border-b border-slate-800 flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-200">
                    Playwright Headed Browser (`samples/demo_portal.html`)
                  </span>
                  <button
                    onClick={() => setPausedForOtp(v => !v)}
                    className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-amber-300 rounded text-xs cursor-pointer whitespace-nowrap shrink-0"
                  >
                    {pausedForOtp ? 'Resume from OTP' : 'Simulate OTP/CAPTCHA Pause'}
                  </button>
                </div>

                {pausedForOtp ? (
                  <div className="p-6 bg-amber-950/30 space-y-3 text-center my-auto">
                    <PauseCircle className="w-10 h-10 text-amber-400 mx-auto" />
                    <h3 className="text-sm font-bold text-amber-200">
                      PAUSED FOR HUMAN INPUT (Login / OTP / CAPTCHA)
                    </h3>
                    <p className="text-xs text-slate-300">
                      "Please login / enter OTP / captcha in the browser, then click Continue"
                    </p>
                    <button
                      onClick={() => setPausedForOtp(false)}
                      className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer"
                    >
                      Continue Automation (जारी रखें)
                    </button>
                  </div>
                ) : (
                  <div className="p-4 space-y-3 text-xs">
                    <div className="grid grid-cols-2 gap-2.5">
                      <div>
                        <label className="text-slate-400 block mb-1">Applicant Full Name *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.Applicant_Name : ''}
                          placeholder="Waiting for Fill..."
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Father / Husband Name *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.Father_Name : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Mobile (10 Digits) *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? (maskSensitive ? maskPii(currentRow.Mobile) : currentRow.Mobile) : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100 font-mono tabular-nums"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Aadhaar (12 Digits) *</label>
                        <input
                          readOnly
                          value={filledInBrowser ? (maskSensitive ? maskPii(currentRow.Aadhaar) : currentRow.Aadhaar) : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-slate-100 font-mono tabular-nums"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">District (जिला)</label>
                        <input
                          readOnly
                          value={filledInBrowser ? currentRow.District : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-emerald-300"
                        />
                      </div>
                      <div>
                        <label className="text-slate-400 block mb-1">Block → Village (Cascading)</label>
                        <input
                          readOnly
                          value={filledInBrowser ? `${currentRow.Block} → ${currentRow.Village}` : ''}
                          className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-emerald-300"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="text-slate-400 block mb-1">Applicant Photo (`uploader.py` Auto-Resize ≤ 50KB)</label>
                      <input
                        readOnly
                        value={filledInBrowser ? currentRow.Photo_File : ''}
                        className="w-full px-2.5 py-1.5 bg-slate-950 border border-slate-800 rounded text-amber-300 font-mono"
                      />
                    </div>

                    <div className="pt-2 border-t border-slate-800 flex flex-wrap items-center gap-2">
                      <button
                        onClick={handleFillRow}
                        className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-semibold cursor-pointer whitespace-nowrap shrink-0"
                      >
                        <Play className="w-3.5 h-3.5" />
                        <span>1. Fill Row #{currentRowIdx + 1}</span>
                      </button>

                      <button
                        onClick={handleApproveAndSubmit}
                        disabled={!filledInBrowser}
                        className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 disabled:text-slate-500 text-white rounded-lg font-semibold cursor-pointer whitespace-nowrap shrink-0"
                      >
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>2. Approve & Submit</span>
                      </button>

                      <button
                        onClick={() => {
                          setFilledInBrowser(false);
                          setCurrentRowIdx(i => Math.min(rows.length - 1, i + 1));
                        }}
                        className="px-2.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
                        title="Skip Row"
                      >
                        <SkipForward className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        ) : activeMode === 'web_to_excel' ? (
          /* =================================================================
             MODE 3: WEBSITE -> EXCEL DATA (STEP 4)
             ================================================================= */
          <div className="p-5 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base font-semibold text-white">
                  Step 4: Website → Excel Data (`extractor.py` Search-and-Fetch & Paginated Table Scraper)
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Searches each Application ID on the website, reads Status, Approval Date, Sanctioned Amount & Officer Remarks, and writes them into a clean Excel copy.
                </p>
              </div>
              <div className="flex items-center gap-2.5">
                <button
                  onClick={handleRunSearchAndFetch}
                  className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
                >
                  <Search className="w-3.5 h-3.5" />
                  <span>Run Search-and-Fetch on Portal (4 IDs)</span>
                </button>
                <button
                  onClick={() => handleExportCurrentExcel('website_status_extracted.xlsx')}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Download Extracted Excel (.xlsx)</span>
                </button>
              </div>
            </div>

            <div className="overflow-x-auto border border-slate-800 rounded-lg">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                  <tr>
                    <th className="p-3">Application_No</th>
                    <th className="p-3">Applicant_Name</th>
                    <th className="p-3">District</th>
                    <th className="p-3">Fetched Portal_Status</th>
                    <th className="p-3">Approval_Date</th>
                    <th className="p-3">Sanctioned_Amount</th>
                    <th className="p-3">Officer_Remarks</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {rows.map(r => (
                    <tr key={r.Application_No} className="hover:bg-slate-950/50">
                      <td className="p-3 font-mono tabular-nums font-semibold text-blue-400">{r.Application_No}</td>
                      <td className="p-3 text-slate-200">{r.Applicant_Name}</td>
                      <td className="p-3 text-slate-300">{r.District}</td>
                      <td className="p-3 font-medium text-slate-200">{r.Status}</td>
                      <td className="p-3 font-mono tabular-nums text-slate-300">{r.Portal_Date || '--'}</td>
                      <td className="p-3 font-mono tabular-nums text-emerald-400">{r.Portal_Amount || '--'}</td>
                      <td className="p-3 text-slate-400">{r.Portal_Remarks || 'Click "Run Search-and-Fetch" above'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          /* =================================================================
             MODE 4: FILL + VERIFY & RECONCILE (STEP 5)
             ================================================================= */
          <div className="p-5 bg-slate-900 border border-slate-800 rounded-xl space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base font-semibold text-white">
                  Step 5: Fill + Verify & Excel vs. Website Reconciliation (`reconcile.py`)
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Updates submitted applications (`Submitted` → `Approved` / `Pending` / `Rejected`) and highlights mismatches in a color-coded `Differences` Excel sheet.
                </p>
              </div>
              <button
                onClick={() => handleExportCurrentExcel('reconciliation_differences.xlsx')}
                className="flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold cursor-pointer whitespace-nowrap shrink-0"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Download `Differences` Workbook (.xlsx)</span>
              </button>
            </div>

            <div className="overflow-x-auto border border-slate-800 rounded-lg">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
                  <tr>
                    <th className="p-3">Application_No</th>
                    <th className="p-3">Applicant_Name</th>
                    <th className="p-3">Excel Status</th>
                    <th className="p-3">Website Live Status</th>
                    <th className="p-3">Excel Amount</th>
                    <th className="p-3">Website Amount</th>
                    <th className="p-3">Reconciliation Outcome</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {rows.map(r => {
                    const web = PORTAL_LIVE_DATABASE[r.Application_No];
                    return (
                      <tr key={r.Application_No} className="hover:bg-slate-950/50">
                        <td className="p-3 font-mono tabular-nums text-emerald-400">{r.Application_No}</td>
                        <td className="p-3 text-slate-200">{r.Applicant_Name}</td>
                        <td className="p-3 text-slate-400">Submitted</td>
                        <td className="p-3 font-semibold text-slate-200">{web?.status}</td>
                        <td className="p-3 font-mono tabular-nums text-slate-300">
                          {r.Application_No === 'RAJ-2025-103' ? '10000' : web?.amount}
                        </td>
                        <td className="p-3 font-mono tabular-nums text-amber-300">{web?.amount}</td>
                        <td className="p-3 text-slate-300">
                          {r.Application_No === 'RAJ-2025-103'
                            ? 'Mismatch: Amount Changed (10000 → 11500)'
                            : r.Application_No === 'RAJ-2025-104'
                            ? 'Status Updated: Rejected on Portal'
                            : 'Exact Match (Approved)'}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      <footer className="border-t border-slate-800 bg-slate-950 px-6 py-4 text-xs text-slate-400 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-2">
          <ShieldCheck className="w-4 h-4 text-emerald-400" />
          <span>Human-in-the-Loop Safety: Pauses for Login, OTP & CAPTCHA · Masks Aadhaar/Mobile (`XXXXXXXX1234`)</span>
        </div>
        <div>
          <span>Authorised Access Only · Check Website Terms of Use</span>
        </div>
      </footer>
    </div>
  );
}
