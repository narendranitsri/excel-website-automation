import React, { useState } from 'react';
import JSZip from 'jszip';
import {
  FileCode,
  FolderArchive,
  Copy,
  Check,
  BookOpen,
  FileSpreadsheet,
  Sparkles,
  ShieldCheck
} from 'lucide-react';
import { GENERATED_PYTHON_FILES } from '../generatedPythonBundle';

export const PythonProjectHub: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<string>('app.py');
  const [copied, setCopied] = useState(false);
  const [isZipping, setIsZipping] = useState(false);
  const [activeTab, setActiveTab] = useState<'code' | 'guide'>('code');

  const fileData = GENERATED_PYTHON_FILES[selectedFile] || GENERATED_PYTHON_FILES['app.py'];

  const handleCopyCode = () => {
    navigator.clipboard.writeText(fileData.content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadZip = async () => {
    setIsZipping(true);
    try {
      const link = document.createElement('a');
      link.href = '/excel_website_automation_assistant.zip';
      link.download = 'excel_website_automation_assistant.zip';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('Fallback JSZip generation:', err);
      const zip = new JSZip();
      Object.entries(GENERATED_PYTHON_FILES).forEach(([filename, file]) => {
        zip.file(`excel_website_automation_assistant/${filename}`, file.content);
      });
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'excel_website_automation_assistant.zip';
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      setTimeout(() => setIsZipping(false), 500);
    }
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
      {/* Header with 1-Click ZIP Download */}
      <div className="px-6 py-4 bg-slate-950 border-b border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-base font-semibold text-slate-100">
            Complete Python Project Bundle (`excel_website_automation_assistant.zip`)
          </h2>
          <p className="text-xs text-slate-400 mt-0.5">
            All 14 Python modules (`app.py`, `scanner.py`, `template_maker.py`, `importer.py`, `validator.py`, `browser.py`, `llm.py`, `uploader.py`, `filler.py`, `runner.py`, `extractor.py`, `reconcile.py`, `report.py`, `utils.py`), `requirements.txt`, `README.md`, `templates/` & `samples/`
          </p>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={() => setActiveTab(t => (t === 'code' ? 'guide' : 'code'))}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-lg text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <BookOpen className="w-3.5 h-3.5 text-emerald-400" />
            <span>{activeTab === 'code' ? 'View Setup & Hindi Guide' : 'View Full Python Code'}</span>
          </button>

          <button
            onClick={handleDownloadZip}
            disabled={isZipping}
            className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 text-white rounded-lg text-xs font-semibold transition-all cursor-pointer whitespace-nowrap shrink-0"
          >
            <FolderArchive className="w-4 h-4" />
            <span>Download Full Python .ZIP</span>
          </button>
        </div>
      </div>

      {activeTab === 'guide' ? (
        <div className="p-6 space-y-6 text-sm text-slate-300">
          {/* Quickstart Command Cards */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="p-4 bg-slate-950 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">1. Install Python Packages</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                pip install -r requirements.txt
              </pre>
            </div>

            <div className="p-4 bg-slate-950 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">2. Install Chromium</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                playwright install chromium
              </pre>
            </div>

            <div className="p-4 bg-slate-950 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">3. Set Gemini Key</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                export GEMINI_API_KEY="your_key"
              </pre>
            </div>

            <div className="p-4 bg-slate-950 border border-slate-800 rounded-lg">
              <span className="text-xs font-semibold text-emerald-400">4. Run Streamlit App</span>
              <pre className="mt-2 p-2.5 bg-slate-900 border border-slate-800 rounded text-xs font-mono text-slate-200 overflow-x-auto">
                streamlit run app.py
              </pre>
            </div>
          </div>

          {/* Safety & Privacy Rules */}
          <div className="p-5 bg-slate-950 border border-slate-800 rounded-xl space-y-2.5">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              <h3 className="font-semibold text-slate-100 text-sm">Built-In Safety, Privacy & Human-in-the-Loop Rules</h3>
            </div>
            <ul className="list-disc pl-5 space-y-1 text-xs text-slate-300">
              <li><strong>No Login/OTP/CAPTCHA Bypass:</strong> Whenever the portal asks for Login, OTP, or CAPTCHA, the browser pauses with <em>"Please login / enter OTP / captcha in the browser, then click Continue"</em>.</li>
              <li><strong>Semi-Automatic Mode (Default):</strong> Fills all fields and uploads compressed photos, then stops before the Submit button so you can verify and click <strong>Approve & Submit</strong>.</li>
              <li><strong>PII Masking (`XXXXXXXX1234`):</strong> Aadhaar, Mobile, and Bank Account numbers are masked in logs and never sent in full to Gemini (`llm.py` sends only column headers, field labels, and 1–2 masked sample values).</li>
            </ul>
          </div>

          {/* Hindi 5-Step Guide */}
          <div className="p-5 bg-slate-950 border border-slate-800 rounded-xl space-y-3">
            <div className="flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-amber-400" />
              <h3 className="font-semibold text-slate-100 text-sm">हिंदी में स्टेप-बाय-स्टेप गाइड (5 आसान चरण)</h3>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs text-slate-300 leading-relaxed">
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-lg space-y-2">
                <div className="font-semibold text-emerald-400">(1) URL Scan & (2) Excel Format Download</div>
                <ol className="list-decimal pl-4 space-y-1.5">
                  <li><strong>URL Scan:</strong> ऊपर <strong>"Scan & Make Excel Format"</strong> चुनें, वेबसाइट का URL डालें और <em>"Open URL & Scan This Page"</em> दबाएं। ऐप फॉर्म के सभी फील्ड्स, ड्रॉपडाउन (`District → Block → Village`), फोटो अपलोड और टेबल्स को स्कैन करके रिपोर्ट दिखाएगा।</li>
                  <li><strong>Excel Format Download:</strong> <em>"Download Auto-Generated Website Excel Format (.xlsx)"</em> दबाएं। इस फाइल में 3 शीट होती हैं: <code>Form_Data_Template</code> (जिसमें वेबसाइट के ड्रॉपडाउन विकल्प पहले से सेट होते हैं), <code>Instructions</code> और <code>Allowed Values</code>।</li>
                  <li><strong>Data Bharna:</strong> डाउनलोड किए गए फॉर्मेट में अपना डेटा भरें (या अपनी पुरानी एक्सेल अपलोड करके उसे वेबसाइट के फॉर्मेट में Reshape कर लें)।</li>
                </ol>
              </div>
              <div className="p-4 bg-slate-900 border border-slate-800 rounded-lg space-y-2">
                <div className="font-semibold text-blue-400">(4) Website Mein Fill & (5) Website Se Data Excel Mein</div>
                <ol className="list-decimal pl-4 space-y-1.5" start={4}>
                  <li><strong>Website Mein Fill:</strong> <strong>"Fill Website from Excel"</strong> मोड में फाइल अपलोड करें। ऐप डेटा चेक करेगा (मोबाइल 10 अंक, आधार 12 अंक, पिन 6 अंक)। फिर <em>"Fill Current Row"</em> दबाएं—ऐप फॉर्म भरकर Submit से पहले रुक जाएगा। जांच करके <strong>"Approve & Submit"</strong> दबाएं।</li>
                  <li><strong>Website Se Data Excel Mein:</strong> <strong>"Website to Excel"</strong> मोड में किसी भी वेब टेबल को कई पेजों (`max_pages`) तक एक्सेल में निकालें, या <strong>Search-and-Fetch</strong> में एप्लीकेशन नंबर की एक्सेल देकर सभी का ताज़ा <em>Status, Date, Amount</em> अपनी एक्सेल में सेव करें।</li>
                </ol>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-12 divide-y md:divide-y-0 md:divide-x divide-slate-800">
          {/* File Tree Selector */}
          <div className="md:col-span-4 p-4 bg-slate-950/60 space-y-2">
            <div className="px-2 py-1 text-xs font-semibold text-slate-400">
              Project Source Files ({Object.keys(GENERATED_PYTHON_FILES).length})
            </div>
            <div className="space-y-1 max-h-[540px] overflow-y-auto pr-1">
              {Object.entries(GENERATED_PYTHON_FILES).map(([key, file]) => {
                const isSelected = selectedFile === key;
                return (
                  <button
                    key={key}
                    onClick={() => setSelectedFile(key)}
                    className={`w-full text-left px-3 py-2 rounded-lg text-xs font-mono transition-all flex items-start gap-2.5 cursor-pointer ${
                      isSelected
                        ? 'bg-slate-800 text-emerald-300 font-semibold border border-slate-700'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900/60'
                    }`}
                  >
                    <FileCode className={`w-4 h-4 mt-0.5 shrink-0 ${isSelected ? 'text-emerald-400' : 'text-slate-500'}`} />
                    <div className="min-w-0">
                      <div className="truncate">{file.name}</div>
                      <div className="text-[11px] text-slate-500 font-sans truncate mt-0.5">
                        {file.description}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>

            <div className="pt-3 border-t border-slate-800/80">
              <div className="px-2 py-1 text-xs font-semibold text-slate-400">
                Bundled Binary Samples in ZIP
              </div>
              <div className="space-y-1 mt-1 text-xs text-slate-400 pl-2">
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
                  <span>samples/sample_applicants.xlsx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-blue-400" />
                  <span>samples/sample_applicants.docx</span>
                </div>
                <div className="flex items-center gap-2 py-0.5">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-amber-400" />
                  <span>samples/photos/*.jpg (Applicant Photos)</span>
                </div>
              </div>
            </div>
          </div>

          {/* Code Viewer */}
          <div className="md:col-span-8 flex flex-col min-w-0">
            <div className="px-5 py-2.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-mono text-slate-300 truncate">
                <FileCode className="w-4 h-4 text-emerald-400 shrink-0" />
                <span>{fileData.name}</span>
                <span className="text-slate-500 font-sans truncate">· {fileData.description}</span>
              </div>

              <button
                onClick={handleCopyCode}
                className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-800 rounded text-xs transition-colors cursor-pointer shrink-0 whitespace-nowrap"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copied ? 'Copied!' : 'Copy Full File'}</span>
              </button>
            </div>

            <pre className="p-5 text-xs font-mono text-slate-300 overflow-x-auto max-h-[600px] bg-slate-950/90 leading-relaxed">
              <code>{fileData.content}</code>
            </pre>
          </div>
        </div>
      )}
    </div>
  );
};
