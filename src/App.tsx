/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { Download, Upload, FileText, CheckCircle2, Files, AlertCircle, Trash2, FileCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import JSZip from 'jszip';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  AlignmentType, 
  HeadingLevel,
} from 'docx';
import { saveAs } from 'file-saver';

interface RecordData {
  [key: string]: any;
}

export default function App() {
  const [excelData, setExcelData] = useState<RecordData[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [atmTemplateFile, setAtmTemplateFile] = useState<File | null>(null);
  const [memoTemplateFile, setMemoTemplateFile] = useState<File | null>(null);
  const [deploymentTemplateFile, setDeploymentTemplateFile] = useState<File | null>(null);
  const [excelFileName, setExcelFileName] = useState<string>('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const atmTemplateInputRef = useRef<HTMLInputElement>(null);
  const memoTemplateInputRef = useRef<HTMLInputElement>(null);
  const deploymentTemplateInputRef = useRef<HTMLInputElement>(null);

  const handleExcelUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelFileName(file.name);
    setError(null);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        if (json.length === 0) {
          setError('The Excel file appears to be empty.');
          setExcelData([]);
        } else {
          setExcelData(json as RecordData[]);
        }
      } catch (err) {
        setError('Failed to parse Excel file. Please ensure it is a valid .xlsx or .csv file.');
        console.error(err);
      }
    };
    reader.onerror = () => setError('Error reading the file.');
    reader.readAsArrayBuffer(file);
    
    // Reset input
    if (excelInputRef.current) excelInputRef.current.value = '';
  };

  const handleTemplateUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.docx')) {
        setTemplateFile(file);
        setError(null);
      } else {
        setError('Please upload a valid .docx template file.');
      }
    }
    // Reset input
    if (templateInputRef.current) templateInputRef.current.value = '';
  };

  const handleAtmTemplateUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.docx')) {
        setAtmTemplateFile(file);
        setError(null);
      } else {
        setError('Please upload a valid .docx template for the ATM Endorsement.');
      }
    }
    if (atmTemplateInputRef.current) atmTemplateInputRef.current.value = '';
  };

  const handleMemoTemplateUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.docx')) {
        setMemoTemplateFile(file);
        setError(null);
      } else {
        setError('Please upload a valid .docx template for the Memo.');
      }
    }
    if (memoTemplateInputRef.current) memoTemplateInputRef.current.value = '';
  };

  const handleDeploymentTemplateUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.docx')) {
        setDeploymentTemplateFile(file);
        setError(null);
      } else {
        setError('Please upload a valid .docx template for the Deployment Letter.');
      }
    }
    if (deploymentTemplateInputRef.current) deploymentTemplateInputRef.current.value = '';
  };

  const clearExcel = () => {
    setExcelData([]);
    setExcelFileName('');
  };

  const clearTemplate = () => {
    setTemplateFile(null);
  };

  const clearAtmTemplate = () => {
    setAtmTemplateFile(null);
  };

  const clearMemoTemplate = () => {
    setMemoTemplateFile(null);
  };

  const clearDeploymentTemplate = () => {
    setDeploymentTemplateFile(null);
  };

  const downloadExcelTemplate = () => {
    const headers = [[
      'Date', 
      'Name', 
      'Address', 
      'SalutationName', 
      'Position', 
      'DailyRate', 
      'WorkDays', 
      'DailyHours', 
      'Company', 
      'AssumptionDate'
    ]];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "Staff_Contract_Population_Template.xlsx");
  };

  const downloadSampleWordTemplate = async () => {
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "LSERV", bold: true, size: 32, color: "008A45" }),
              new TextRun({ text: " CORPORATION", bold: true, size: 32, color: "003366" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "EMPLOYMENT CONTRACT", bold: true, size: 28 })],
            spacing: { before: 400, after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "{Date}", bold: true })],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "{Name}", bold: true, break: 1 }),
              new TextRun({ text: "{Address}", break: 1 }),
            ],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "Dear Mr./Ms. {SalutationName}," })],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun("Please be informed that we are hiring you for the position of "),
              new TextRun({ text: "{Position}", bold: true }),
              new TextRun(" with a daily rate of "),
              new TextRun({ text: "{DailyRate}", bold: true }),
              new TextRun(" subject to the completion of the pre-employment requirements. You are required to report to work for "),
              new TextRun({ text: "{WorkDays}", bold: true }),
              new TextRun(" per week "),
              new TextRun({ text: "{DailyHours}", bold: true }),
              new TextRun(" per day."),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun("Your initial place of assignment will be at "),
              new TextRun({ text: "{Company}", bold: true }),
              new TextRun(" effective upon your assumption to duty on "),
              new TextRun({ text: "{AssumptionDate}", bold: true }),
              new TextRun("."),
            ],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "The TERMS AND CONDITIONS of your employment are as follows:", bold: true })],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [new TextRun("1. You are an employee of LSERV Corporation...")],
            spacing: { before: 100 },
          }),
          // Simplified for sample purposes
          new Paragraph({
            children: [new TextRun({ text: "Very truly yours,", break: 3 })],
          }),
          new Paragraph({
            children: [new TextRun({ text: "JOSEPH V. ANGELES", bold: true, break: 2 }), new TextRun({ text: "President", break: 1 })],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "LSERV_Contract_Template.docx");
  };

  const downloadSampleATMTemplate = async () => {
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "LSERV", bold: true, size: 32, color: "008A45" }),
              new TextRun({ text: " CORPORATION", bold: true, size: 32, color: "003366" }),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: "{Date}", bold: true, break: 2 })],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "THE HEAD", bold: true }),
              new TextRun({ text: "LandBank of the Philippines", bold: true, break: 1 }),
            ],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "Dear Sir/Madam:", break: 1 })],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun("Please facilitate issuance of ATM card to the following contractual/s of LSERV CORPORATION whose salary would be through "),
              new TextRun({ text: "ATM payroll", bold: true }),
              new TextRun(" arrangement with your Branch:"),
            ],
            spacing: { after: 400 },
          }),
          // In a real template, you'd add a table, but for sample tags:
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "NAME: {Name}", bold: true, break: 1 }),
              new TextRun({ text: "POSITION: {Position}", bold: true, break: 1 }),
              new TextRun({ text: "OFFICE: {Company}", bold: true, break: 1 }),
            ],
            spacing: { before: 200, after: 400 },
          }),
          new Paragraph({
            children: [new TextRun("Per our arrangement / agreement, the account/s of our employee/s may be opened with your branch.")],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [new TextRun("Thank you very much for your usual assistance.")],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "Very truly yours,", break: 3 })],
          }),
          new Paragraph({
            children: [new TextRun({ text: "LESLIE ANNE A. ALCOBER", bold: true, break: 2 }), new TextRun({ text: "Senior Manager", break: 1 })],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "LSERV_ATM_Endorsement_Template.docx");
  };

  const generateWordFiles = async () => {
    if ((!templateFile && !atmTemplateFile && !memoTemplateFile && !deploymentTemplateFile) || excelData.length === 0) return;

    setIsGenerating(true);
    setProgress(0);
    setError(null);

    try {
      const zip = new JSZip();
      
      const templates = [
        { file: templateFile, suffix: 'Contract' },
        { file: atmTemplateFile, suffix: 'ATM_Endorsement' },
        { file: memoTemplateFile, suffix: 'Memo_Random' },
        { file: deploymentTemplateFile, suffix: 'Deployment_Letter' }
      ];

      // Read all active templates first
      const loadedTemplates = await Promise.all(
        templates.map(async (t) => {
          if (t.file) {
            return { buffer: await t.file.arrayBuffer(), suffix: t.suffix };
          }
          return null;
        })
      );

      for (let i = 0; i < excelData.length; i++) {
        const record = excelData[i];
        const baseName = (record.Name || record.name || record['Full Name'] || record.Names || `Staff_${i + 1}`).toString().replace(/[/\\?%*:|"<>]/g, '-').trim();

        for (const t of loadedTemplates) {
          if (t) {
            try {
              const zipContent = new PizZip(t.buffer);
              const doc = new Docxtemplater(zipContent, { paragraphLoop: true, linebreaks: true });
              doc.setData(record);
              doc.render();
              const out = doc.getZip().generate({ type: 'blob' });
              zip.file(`${baseName}_${t.suffix}.docx`, out);
            } catch (err) {
              console.error(`Error processing ${t.suffix} for ${baseName}`, err);
            }
          }
        }
        
        setProgress(Math.round(((i + 1) / excelData.length) * 100));
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, `LSERV_Batch_${new Date().toISOString().slice(0, 10)}.zip`);
      
    } catch (err: any) {
      setError(err.message || 'An error occurred during generation.');
      console.error(err);
    } finally {
      setIsGenerating(false);
    }
  };

  const columns = excelData.length > 0 ? Object.keys(excelData[0]) : [];

  return (
    <div className="w-full h-screen bg-[#F8FAFC] font-sans text-slate-900 flex flex-col overflow-hidden">
      {/* Navigation Bar */}
      <nav className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0 z-20">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center shadow-lg shadow-blue-100">
            <FileText className="w-6 h-6 text-white" />
          </div>
          <div className="flex flex-col">
            <span className="font-bold text-lg tracking-tight text-slate-800 leading-none">DocBatch</span>
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">Automator</span>
          </div>
        </div>
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-3 bg-slate-50 px-3 py-1.5 rounded-full border border-slate-200">
            <div className="w-6 h-6 rounded-full bg-blue-500 flex items-center justify-center text-[10px] font-bold text-white">
              GR
            </div>
            <span className="text-sm font-semibold text-slate-600">Admin Console</span>
          </div>
        </div>
      </nav>

      {/* Main Workspace */}
      <main className="flex-1 p-8 grid grid-cols-12 gap-8 overflow-hidden">
        {/* Left Side: Upload & Config */}
        <div className="col-span-4 flex flex-col gap-6 overflow-y-auto pr-2 custom-scrollbar">
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 hover:border-blue-200 transition-colors group">
            <h3 className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] mb-4">Step 01: Setup & Data</h3>
            
            <div className="grid grid-cols-2 gap-3 mb-6">
              <button 
                onClick={downloadExcelTemplate}
                className="px-4 py-3 bg-white hover:bg-slate-50 text-slate-700 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all flex flex-col items-center justify-center gap-2 border border-slate-200 shadow-sm hover:border-blue-300"
              >
                <div className="w-8 h-8 bg-green-50 rounded-lg flex items-center justify-center text-green-600">
                  <Download className="w-4 h-4" />
                </div>
                Excel Template
              </button>
              <button 
                onClick={downloadSampleWordTemplate}
                className="px-4 py-3 bg-white hover:bg-slate-50 text-slate-700 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all flex flex-col items-center justify-center gap-2 border border-slate-200 shadow-sm hover:border-blue-300"
              >
                <div className="w-8 h-8 bg-blue-50 rounded-lg flex items-center justify-center text-blue-600">
                  <FileCheck className="w-4 h-4" />
                </div>
                Contract Template
              </button>
              <button 
                onClick={downloadSampleATMTemplate}
                className="px-4 py-3 bg-white hover:bg-slate-50 text-slate-700 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all flex flex-col items-center justify-center gap-2 border border-slate-200 shadow-sm hover:border-blue-300"
              >
                <div className="w-8 h-8 bg-purple-50 rounded-lg flex items-center justify-center text-purple-600">
                  <FileCheck className="w-4 h-4" />
                </div>
                ATM Endorsement
              </button>
            </div>

            <div 
              onClick={() => excelInputRef.current?.click()}
              className="border-2 border-dashed border-slate-200 rounded-2xl p-8 flex flex-col items-center text-center bg-slate-50/50 hover:bg-blue-50/30 hover:border-blue-300 cursor-pointer transition-all duration-300 relative group"
            >
              <div className="w-16 h-16 bg-white rounded-2xl shadow-sm border border-slate-100 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-300">
                <Upload className="w-8 h-8 text-blue-500" />
              </div>
              <p className="text-sm font-semibold text-slate-700">Drop Excel or CSV here</p>
              <p className="text-xs text-slate-400 mt-1">Click to browse files</p>
              <input 
                type="file" 
                ref={excelInputRef} 
                onChange={handleExcelUpload} 
                accept=".xlsx,.xls,.csv" 
                className="hidden" 
              />
            </div>
            {excelFileName && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-4 flex items-center gap-3 p-3 bg-blue-50 border border-blue-100 rounded-xl"
              >
                <Files className="w-5 h-5 text-blue-600 shrink-0" />
                <div className="flex-1 min-w-0">
                  <span className="text-xs font-bold text-blue-900 truncate block">{excelFileName}</span>
                  <span className="text-[10px] text-blue-600 font-medium">{excelData.length} entries detected</span>
                </div>
                <button onClick={clearExcel} className="p-1.5 hover:bg-blue-100 rounded-lg text-blue-400 hover:text-red-500 transition-colors">
                  <Trash2 className="w-4 h-4" />
                </button>
              </motion.div>
            )}
          </section>

          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 hover:border-blue-200 transition-colors">
            <h3 className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] mb-4">Step 02: Template Design</h3>
            <div className="space-y-4">
              {/* PRIMARY CONTRACT TEMPLATE */}
              <div 
                onClick={() => templateInputRef.current?.click()}
                className={`flex items-center gap-4 p-4 border rounded-2xl cursor-pointer transition-all duration-300 ${
                  templateFile 
                    ? 'border-blue-600 bg-blue-50 shadow-sm shadow-blue-100' 
                    : 'border-slate-100 bg-slate-50/50 hover:bg-slate-100 hover:border-slate-300'
                }`}
              >
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${templateFile ? 'bg-blue-600 text-white' : 'bg-white border border-slate-200 text-slate-400'}`}>
                  <FileText className="w-5 h-5" />
                </div>
                <div className="flex-1 min-w-0">
                  <p className={`text-[10px] font-bold uppercase tracking-widest ${templateFile ? 'text-blue-400' : 'text-slate-400'}`}>Template 01</p>
                  <p className={`text-sm font-bold truncate ${templateFile ? 'text-blue-900' : 'text-slate-500'}`}>
                    {templateFile ? templateFile.name : 'Choose Contract'}
                  </p>
                </div>
                {templateFile && (
                  <button onClick={(e) => { e.stopPropagation(); clearTemplate(); }} className="p-2 hover:bg-blue-100 rounded-lg text-blue-400">
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
                <input type="file" ref={templateInputRef} onChange={handleTemplateUpload} accept=".docx" className="hidden" />
              </div>

              {/* SECONDARY ATM TEMPLATE */}
              <div 
                onClick={() => atmTemplateInputRef.current?.click()}
                className={`flex items-center gap-4 p-4 border rounded-2xl cursor-pointer transition-all duration-300 ${
                  atmTemplateFile 
                    ? 'border-purple-600 bg-purple-50 shadow-sm shadow-purple-100' 
                    : 'border-slate-100 bg-slate-50/50 hover:bg-slate-100 hover:border-slate-300'
                }`}
              >
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${atmTemplateFile ? 'bg-purple-600 text-white' : 'bg-white border border-slate-200 text-slate-400'}`}>
                  <FileCheck className="w-5 h-5" />
                </div>
                <div className="flex-1 min-w-0" title={atmTemplateFile?.name}>
                  <p className={`text-[10px] font-bold uppercase tracking-widest ${atmTemplateFile ? 'text-purple-400' : 'text-slate-400'}`}>Template 02</p>
                  <p className={`text-sm font-bold truncate ${atmTemplateFile ? 'text-purple-900' : 'text-slate-500'}`}>
                    {atmTemplateFile ? atmTemplateFile.name : 'ATM Endorsement'}
                  </p>
                </div>
                {atmTemplateFile && (
                  <button onClick={(e) => { e.stopPropagation(); clearAtmTemplate(); }} className="p-2 hover:bg-purple-100 rounded-lg text-purple-400">
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
                <input type="file" ref={atmTemplateInputRef} onChange={handleAtmTemplateUpload} accept=".docx" className="hidden" />
              </div>

              {/* MEMO RANDOM TEMPLATE */}
              <div 
                onClick={() => memoTemplateInputRef.current?.click()}
                className={`flex items-center gap-4 p-4 border rounded-2xl cursor-pointer transition-all duration-300 ${
                  memoTemplateFile 
                    ? 'border-emerald-600 bg-emerald-50 shadow-sm shadow-emerald-100' 
                    : 'border-slate-100 bg-slate-50/50 hover:bg-slate-100 hover:border-slate-300'
                }`}
              >
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${memoTemplateFile ? 'bg-emerald-600 text-white' : 'bg-white border border-slate-200 text-slate-400'}`}>
                  <FileText className="w-5 h-5" />
                </div>
                <div className="flex-1 min-w-0" title={memoTemplateFile?.name}>
                  <p className={`text-[10px] font-bold uppercase tracking-widest ${memoTemplateFile ? 'text-emerald-400' : 'text-slate-400'}`}>Template 03</p>
                  <p className={`text-sm font-bold truncate ${memoTemplateFile ? 'text-emerald-900' : 'text-slate-500'}`}>
                    {memoTemplateFile ? memoTemplateFile.name : 'Memo Random'}
                  </p>
                </div>
                {memoTemplateFile && (
                  <button onClick={(e) => { e.stopPropagation(); clearMemoTemplate(); }} className="p-2 hover:bg-emerald-100 rounded-lg text-emerald-400">
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
                <input type="file" ref={memoTemplateInputRef} onChange={handleMemoTemplateUpload} accept=".docx" className="hidden" />
              </div>

              {/* DEPLOYMENT TEMPLATE */}
              <div 
                onClick={() => deploymentTemplateInputRef.current?.click()}
                className={`flex items-center gap-4 p-4 border rounded-2xl cursor-pointer transition-all duration-300 ${
                  deploymentTemplateFile 
                    ? 'border-amber-600 bg-amber-50 shadow-sm shadow-amber-100' 
                    : 'border-slate-100 bg-slate-50/50 hover:bg-slate-100 hover:border-slate-300'
                }`}
              >
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${deploymentTemplateFile ? 'bg-amber-600 text-white' : 'bg-white border border-slate-200 text-slate-400'}`}>
                  <FileCheck className="w-5 h-5" />
                </div>
                <div className="flex-1 min-w-0" title={deploymentTemplateFile?.name}>
                  <p className={`text-[10px] font-bold uppercase tracking-widest ${deploymentTemplateFile ? 'text-amber-400' : 'text-slate-400'}`}>Template 04</p>
                  <p className={`text-sm font-bold truncate ${deploymentTemplateFile ? 'text-amber-900' : 'text-slate-500'}`}>
                    {deploymentTemplateFile ? deploymentTemplateFile.name : 'Deployment Letter'}
                  </p>
                </div>
                {deploymentTemplateFile && (
                  <button onClick={(e) => { e.stopPropagation(); clearDeploymentTemplate(); }} className="p-2 hover:bg-amber-100 rounded-lg text-amber-400">
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
                <input type="file" ref={deploymentTemplateInputRef} onChange={handleDeploymentTemplateUpload} accept=".docx" className="hidden" />
              </div>
              
              <AnimatePresence>
                {(!templateFile || !atmTemplateFile || !memoTemplateFile || !deploymentTemplateFile) && (
                  <motion.div 
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    exit={{ opacity: 0, height: 0 }}
                    className="overflow-hidden"
                  >
                    <div className="p-4 bg-amber-50/50 border border-amber-100 rounded-2xl flex items-start gap-3">
                      <AlertCircle className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
                      <div className="space-y-1">
                        <p className="text-xs font-bold text-amber-800 leading-tight">Batch Hint</p>
                        <p className="text-[10px] text-amber-700 leading-relaxed font-medium">
                          You can upload either or both templates. The system will skip whichever is missing but process what's available for each record.
                        </p>
                      </div>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </section>

          <AnimatePresence>
            {error && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="p-4 bg-red-50 border border-red-100 rounded-2xl flex items-start gap-3"
              >
                <div className="w-8 h-8 bg-red-100 rounded-lg flex items-center justify-center shrink-0">
                  <AlertCircle className="w-5 h-5 text-red-600" />
                </div>
                <div className="flex-1">
                  <p className="text-xs font-bold text-red-900">System Error</p>
                  <p className="text-[10px] text-red-700 mt-0.5 font-medium">{error}</p>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Right Side: Preview & Action */}
        <div className="col-span-8 flex flex-col bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden relative">
          <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-white sticky top-0 z-10">
            <div className="flex flex-col">
              <h3 className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">Live Data Preview</h3>
              <p className="text-[11px] font-semibold text-slate-500 mt-1 shrink-0">
                {excelData.length > 0 ? `${Math.min(10, excelData.length)} of ${excelData.length} records detected` : 'Waiting for data source...'}
              </p>
            </div>
            {excelData.length > 0 && (
              <div className="flex gap-2">
                {columns.slice(0, 3).map(col => (
                  <span key={col} className="px-2 py-1 bg-slate-100 rounded-md text-[9px] font-bold text-slate-500 border border-slate-200">
                    {col}
                  </span>
                ))}
              </div>
            )}
          </div>
          
          <div className="flex-1 overflow-auto custom-scrollbar relative">
            {excelData.length > 0 ? (
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-white shadow-sm z-10">
                  <tr>
                    {columns.map((header) => (
                      <th key={header} className="px-6 py-4 text-[10px] font-bold text-slate-400 uppercase tracking-wider border-b border-slate-100">
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {excelData.map((row, idx) => (
                    <tr key={idx} className="hover:bg-blue-50/30 transition-colors group">
                      {columns.map((col, i) => (
                        <td key={i} className="px-6 py-4 text-sm text-slate-600 font-medium truncate max-w-[200px]">
                          {row[col]?.toString() || <span className="text-slate-300 italic">null</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-slate-300 p-12 text-center bg-slate-50/30">
                <div className="relative mb-6">
                  <Files className="w-20 h-20 opacity-10" />
                  <div className="absolute inset-0 flex items-center justify-center">
                    <Download className="w-8 h-8 opacity-20" />
                  </div>
                </div>
                <h4 className="text-slate-500 font-bold text-lg">No data to display</h4>
                <p className="text-sm max-w-xs mt-2 font-medium">Upload your Excel population file to map fields and verify your batch before generation.</p>
              </div>
            )}
            {excelData.length > 0 && (
              <div className="p-12 text-center bg-gradient-to-t from-slate-50/80 to-transparent">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">End of Record Set</p>
              </div>
            )}
          </div>

          <div className="p-8 bg-slate-50/80 border-t border-slate-100 flex items-center justify-between gap-8">
            <div className="flex items-center gap-8">
              <div className="flex items-center gap-2">
                <div className={`w-5 h-5 rounded-lg flex items-center justify-center ${excelData.length > 0 ? 'bg-green-100 text-green-600' : 'bg-slate-200 text-slate-400'}`}>
                  <CheckCircle2 className="w-3.5 h-3.5" />
                </div>
                <span className={`text-[11px] font-bold uppercase tracking-wider ${excelData.length > 0 ? 'text-slate-700' : 'text-slate-400'}`}>
                  Data Validated
                </span>
              </div>
              <div className="flex items-center gap-2">
                <div className={`w-5 h-5 rounded-lg flex items-center justify-center ${templateFile || atmTemplateFile || memoTemplateFile || deploymentTemplateFile ? 'bg-green-100 text-green-600' : 'bg-slate-200 text-slate-400'}`}>
                  <CheckCircle2 className="w-3.5 h-3.5" />
                </div>
                <span className={`text-[11px] font-bold uppercase tracking-wider ${templateFile || atmTemplateFile || memoTemplateFile || deploymentTemplateFile ? 'text-slate-700' : 'text-slate-400'}`}>
                  Template Ready
                </span>
              </div>
            </div>

            <div className="flex items-center gap-6">
              {isGenerating && (
                <div className="flex flex-col items-end gap-1.5">
                  <span className="text-[10px] font-bold text-blue-600 uppercase tracking-widest">Processing... {progress}%</span>
                  <div className="w-48 bg-blue-100 rounded-full h-1.5 overflow-hidden">
                    <motion.div 
                      key="progress-bar"
                      initial={{ width: 0 }}
                      animate={{ width: `${progress}%` }}
                      className="h-full bg-blue-600 rounded-full"
                    />
                  </div>
                </div>
              )}
              <button 
                onClick={generateWordFiles}
                disabled={isGenerating || (!templateFile && !atmTemplateFile) || excelData.length === 0}
                className={`group relative bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 px-10 rounded-2xl shadow-xl shadow-blue-200 flex items-center gap-3 transition-all duration-300 active:scale-95 disabled:grayscale disabled:opacity-40 disabled:shadow-none disabled:cursor-not-allowed`}
              >
                {isGenerating ? (
                  <>
                    <div className="w-5 h-5 border-3 border-white/30 border-t-white rounded-full animate-spin" />
                    <span className="tracking-wide">Building Documents...</span>
                  </>
                ) : (
                  <>
                    <Download className="w-5 h-5 group-hover:translate-y-0.5 transition-transform" />
                    <span className="tracking-wide">Generate Batch Files</span>
                    <div className="absolute -top-2 -right-2 bg-blue-400 text-white text-[10px] py-0.5 px-2 rounded-full shadow-md border-2 border-white scale-0 group-hover:scale-100 transition-transform">
                      ZIP
                    </div>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-10 bg-[#0F172A] text-slate-400 px-8 flex items-center justify-between shrink-0 z-20">
        <div className="flex items-center gap-4">
          <div className="text-[10px] uppercase font-black tracking-[0.3em] flex items-center gap-2">
            System Status: 
            <span className={isGenerating ? 'text-amber-400' : 'text-emerald-400'}>
               {isGenerating ? '• BUSY' : '• ONLINE'}
            </span>
          </div>
          <div className="h-3 w-px bg-slate-700"></div>
          <div className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">
            Queue: <span className="text-slate-300">{excelData.length} Items</span>
          </div>
        </div>
        <div className="text-[10px] font-bold text-slate-500 uppercase tracking-tight flex items-center gap-4">
          <span>{new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false })} IST</span>
          <span className="opacity-40">Build 0426.26</span>
        </div>
      </footer>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #E2E8F0;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #CBD5E1;
        }
      `}} />
    </div>
  );
}
