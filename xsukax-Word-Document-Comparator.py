#!/usr/bin/env python3
"""
xsukax Word Document Comparator
A single-file web application for side-by-side Word document comparison.
"""

import os
import io
import re
from datetime import datetime
from flask import Flask, request, render_template_string, jsonify
from docx import Document
import docx2txt
import pythoncom
from difflib import SequenceMatcher

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>xsukax Word Document Comparator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans Arabic', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #24292e; background-color: #f6f8fa; padding: 20px; }
        .container { max-width: 100%; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 30px; padding: 25px; background: white; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .header h1 { color: #24292e; margin-bottom: 8px; font-size: 28px; font-weight: 600; }
        .header .brand { color: #0366d6; font-weight: 700; }
        .header p { color: #586069; font-size: 14px; }
        .upload-section { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
        .upload-box { background: white; padding: 25px; border-radius: 6px; border: 2px dashed #d1d5da; text-align: center; transition: all 0.3s ease; }
        .upload-box:hover { border-color: #0366d6; background: #f6f8fa; }
        .upload-box h3 { margin-bottom: 15px; color: #24292e; font-size: 16px; font-weight: 600; }
        .file-input { width: 100%; margin: 15px 0; padding: 8px; border: 1px solid #d1d5da; border-radius: 6px; font-size: 14px; }
        .file-name { margin-top: 10px; color: #586069; font-size: 13px; min-height: 20px; }
        .btn { background: #2ea44f; color: white; border: none; padding: 12px 28px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: background 0.2s ease; }
        .btn:hover { background: #2c974b; }
        .btn:disabled { background: #94d3a2; cursor: not-allowed; }
        .comparison-section { display: none; background: white; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; margin-top: 20px; }
        .analytics-panel { background: #f6f8fa; padding: 20px; border-bottom: 1px solid #e1e4e8; display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; }
        .analytics-card { background: white; padding: 15px; border-radius: 6px; border: 1px solid #e1e4e8; text-align: center; }
        .analytics-card h4 { color: #586069; font-size: 12px; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 0.5px; }
        .analytics-card .value { color: #24292e; font-size: 24px; font-weight: 700; }
        .analytics-card .label { color: #586069; font-size: 11px; margin-top: 4px; }
        .analytics-card.added .value { color: #22863a; }
        .analytics-card.removed .value { color: #cb2431; }
        .analytics-card.modified .value { color: #b08800; }
        .analytics-card.similarity .value { color: #0366d6; }
        .differences-summary { background: white; padding: 20px; border-radius: 6px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); display: none; }
        .differences-summary h3 { color: #24292e; font-size: 16px; font-weight: 600; margin-bottom: 15px; display: flex; align-items: center; gap: 8px; }
        .differences-summary h3::before { content: '📍'; }
        .diff-lines-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .diff-lines-section { background: #f6f8fa; padding: 15px; border-radius: 6px; border: 1px solid #e1e4e8; }
        .diff-lines-section h4 { font-size: 13px; font-weight: 600; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.5px; }
        .diff-lines-section.added h4 { color: #22863a; }
        .diff-lines-section.removed h4 { color: #cb2431; }
        .diff-lines-section.modified h4 { color: #b08800; }
        .line-numbers { display: flex; flex-wrap: wrap; gap: 6px; }
        .line-number-badge { background: white; color: #586069; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; border: 1px solid #e1e4e8; font-family: 'Courier New', monospace; }
        .diff-lines-section.added .line-number-badge { border-color: #85e89d; color: #22863a; }
        .diff-lines-section.removed .line-number-badge { border-color: #fdaeb7; color: #cb2431; }
        .diff-lines-section.modified .line-number-badge { border-color: #ffd33d; color: #b08800; }
        .no-differences { color: #586069; font-size: 13px; font-style: italic; }
        .comparison-header { display: grid; grid-template-columns: 1fr 1fr; background: #f6f8fa; border-bottom: 1px solid #e1e4e8; padding: 12px 20px; }
        .document-title { font-weight: 600; color: #24292e; font-size: 14px; display: flex; align-items: center; }
        .document-title::before { content: '📄'; margin-right: 8px; }
        .comparison-container { display: grid; grid-template-columns: 1fr 1fr; height: 70vh; overflow: hidden; border-top: 1px solid #e1e4e8; }
        .document-panel { border-right: 1px solid #e1e4e8; overflow: auto; position: relative; background: #ffffff; }
        .document-panel:last-child { border-right: none; }
        .document-panel::-webkit-scrollbar { width: 12px; height: 12px; }
        .document-panel::-webkit-scrollbar-track { background: #f6f8fa; }
        .document-panel::-webkit-scrollbar-thumb { background: #d1d5da; border-radius: 6px; border: 2px solid #f6f8fa; }
        .document-panel::-webkit-scrollbar-thumb:hover { background: #959da5; }
        .document-content { padding: 20px; font-family: 'Courier New', 'Courier', 'Noto Sans Arabic', monospace; white-space: pre-wrap; line-height: 1.8; font-size: 13px; direction: ltr; }
        .document-content.rtl { direction: rtl; text-align: right; }
        .content-line { display: flex; align-items: flex-start; min-height: 24px; }
        .rtl .content-line { flex-direction: row-reverse; }
        .line-number { display: inline-block; min-width: 45px; color: #6a737d; text-align: right; padding-right: 12px; user-select: none; font-size: 12px; flex-shrink: 0; }
        .rtl .line-number { text-align: left; padding-right: 0; padding-left: 12px; }
        .line-content { flex: 1; }
        .word { padding: 2px 4px; border-radius: 3px; margin: 1px; display: inline-block; transition: all 0.2s ease; }
        .word:hover { opacity: 0.8; }
        .same { background-color: transparent; }
        .different { background-color: #fff2c5; border: 1px solid #ffd33d; }
        .missing { background-color: #ffdce0; border: 1px solid #fdaeb7; color: #cb2431; text-decoration: line-through; }
        .added { background-color: #dcffe4; border: 1px solid #85e89d; color: #22863a; }
        .legend { display: flex; justify-content: center; gap: 20px; margin: 20px 0; flex-wrap: wrap; padding: 15px; background: white; border-radius: 6px; }
        .legend-item { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #586069; }
        .color-box { width: 18px; height: 18px; border-radius: 3px; border: 1px solid; }
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
        .modal-content { background: white; padding: 30px; border-radius: 6px; max-width: 500px; width: 90%; text-align: center; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
        .modal h3 { margin-bottom: 15px; color: #24292e; font-size: 18px; font-weight: 600; }
        .modal p { margin-bottom: 20px; color: #586069; font-size: 14px; line-height: 1.5; }
        .modal-btn { background: #0366d6; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: background 0.2s ease; }
        .modal-btn:hover { background: #0256c7; }
        .loading { display: none; text-align: center; padding: 30px; background: white; border-radius: 6px; margin-top: 20px; }
        .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #0366d6; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 15px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .loading p { color: #586069; font-size: 14px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><span class="brand">xsukax</span> Word Document Comparator</h1>
            <p>Upload two Word documents (.docx) to compare them side-by-side</p>
        </div>

        <form id="uploadForm" enctype="multipart/form-data">
            <div class="upload-section">
                <div class="upload-box">
                    <h3>📄 Document 1</h3>
                    <input type="file" name="doc1" accept=".docx" class="file-input" required>
                    <div class="file-name" id="fileName1">No file selected</div>
                </div>
                <div class="upload-box">
                    <h3>📄 Document 2</h3>
                    <input type="file" name="doc2" accept=".docx" class="file-input" required>
                    <div class="file-name" id="fileName2">No file selected</div>
                </div>
            </div>
            
            <div style="text-align: center;">
                <button type="submit" class="btn" id="compareBtn">🔍 Compare Documents</button>
            </div>
        </form>

        <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Processing documents...</p>
        </div>

        <div class="legend" id="legend" style="display: none;">
            <div class="legend-item"><div class="color-box same" style="border-color: #e1e4e8; background: #ffffff;"></div> Unchanged</div>
            <div class="legend-item"><div class="color-box different" style="background-color: #fff2c5; border-color: #ffd33d;"></div> Modified</div>
            <div class="legend-item"><div class="color-box missing" style="background-color: #ffdce0; border-color: #fdaeb7;"></div> Removed</div>
            <div class="legend-item"><div class="color-box added" style="background-color: #dcffe4; border-color: #85e89d;"></div> Added</div>
        </div>

        <div class="differences-summary" id="differencesSummary">
            <h3>Lines with Differences</h3>
            <div class="diff-lines-container" id="diffLinesContainer"></div>
        </div>

        <div class="comparison-section" id="comparisonSection">
            <div class="analytics-panel" id="analyticsPanel"></div>
            <div class="comparison-header">
                <div class="document-title" id="doc1Title">Document 1</div>
                <div class="document-title" id="doc2Title">Document 2</div>
            </div>
            <div class="comparison-container">
                <div class="document-panel" id="doc1Panel">
                    <div class="document-content" id="doc1Content"></div>
                </div>
                <div class="document-panel" id="doc2Panel">
                    <div class="document-content" id="doc2Content"></div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal" id="errorModal">
        <div class="modal-content">
            <h3>⚠️ Error</h3>
            <p id="errorMessage">An error occurred while processing the documents.</p>
            <button class="modal-btn" onclick="closeModal()">OK</button>
        </div>
    </div>

    <script>
        document.querySelectorAll('input[type="file"]').forEach((input, index) => {
            input.addEventListener('change', function(e) {
                const fileName = e.target.files[0] ? e.target.files[0].name : 'No file selected';
                document.getElementById(`fileName${index + 1}`).textContent = fileName;
            });
        });

        document.getElementById('uploadForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            const compareBtn = document.getElementById('compareBtn');
            const loading = document.getElementById('loading');
            
            compareBtn.disabled = true;
            loading.style.display = 'block';
            document.getElementById('comparisonSection').style.display = 'none';
            document.getElementById('legend').style.display = 'none';
            
            try {
                const response = await fetch('/compare', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    displayComparison(result);
                } else {
                    showError(result.error || 'An error occurred during comparison');
                }
            } catch (error) {
                showError('Network error: ' + error.message);
            } finally {
                compareBtn.disabled = false;
                loading.style.display = 'none';
            }
        });

        function detectArabic(text) {
            const arabicPattern = /[\u0600-\u06FF]/;
            return arabicPattern.test(text);
        }

        function displayComparison(result) {
            document.getElementById('doc1Title').textContent = result.doc1_name;
            document.getElementById('doc2Title').textContent = result.doc2_name;
            
            const doc1Content = document.getElementById('doc1Content');
            const doc2Content = document.getElementById('doc2Content');
            
            doc1Content.innerHTML = result.doc1_html;
            doc2Content.innerHTML = result.doc2_html;
            
            if (detectArabic(result.doc1_html) || detectArabic(result.doc2_html)) {
                doc1Content.classList.add('rtl');
                doc2Content.classList.add('rtl');
            } else {
                doc1Content.classList.remove('rtl');
                doc2Content.classList.remove('rtl');
            }
            
            const analyticsHTML = `
                <div class="analytics-card">
                    <h4>Total Words</h4>
                    <div class="value">${result.analytics.total_words}</div>
                    <div class="label">analyzed</div>
                </div>
                <div class="analytics-card similarity">
                    <h4>Similarity</h4>
                    <div class="value">${result.analytics.similarity}%</div>
                    <div class="label">match rate</div>
                </div>
                <div class="analytics-card added">
                    <h4>Added</h4>
                    <div class="value">${result.analytics.added_words}</div>
                    <div class="label">words</div>
                </div>
                <div class="analytics-card removed">
                    <h4>Removed</h4>
                    <div class="value">${result.analytics.removed_words}</div>
                    <div class="label">words</div>
                </div>
                <div class="analytics-card modified">
                    <h4>Modified</h4>
                    <div class="value">${result.analytics.modified_words}</div>
                    <div class="label">words</div>
                </div>
            `;
            
            document.getElementById('analyticsPanel').innerHTML = analyticsHTML;
            document.getElementById('legend').style.display = 'flex';
            document.getElementById('comparisonSection').style.display = 'block';
            
            displayDifferencesSummary(result.line_differences);
            
            document.getElementById('differencesSummary').scrollIntoView({ 
                behavior: 'smooth',
                block: 'start'
            });
            
            const doc1Panel = document.getElementById('doc1Panel');
            const doc2Panel = document.getElementById('doc2Panel');
            doc1Panel.addEventListener('scroll', () => {
                doc2Panel.scrollTop = doc1Panel.scrollTop;
            });
            doc2Panel.addEventListener('scroll', () => {
                doc1Panel.scrollTop = doc2Panel.scrollTop;
            });
        }

        function showError(message) {
            document.getElementById('errorMessage').textContent = message;
            document.getElementById('errorModal').style.display = 'flex';
        }

        function closeModal() {
            document.getElementById('errorModal').style.display = 'none';
        }

        window.onclick = function(event) {
            const modal = document.getElementById('errorModal');
            if (event.target === modal) {
                closeModal();
            }
        }

        function displayDifferencesSummary(lineDiffs) {
            const container = document.getElementById('diffLinesContainer');
            
            const addedLines = lineDiffs.added || [];
            const removedLines = lineDiffs.removed || [];
            const modifiedLines = lineDiffs.modified || [];
            
            let html = '';
            
            if (addedLines.length > 0) {
                html += `
                    <div class="diff-lines-section added">
                        <h4>✅ Added (${addedLines.length} lines)</h4>
                        <div class="line-numbers">
                            ${addedLines.map(line => `<span class="line-number-badge">Line ${line}</span>`).join('')}
                        </div>
                    </div>
                `;
            }
            
            if (removedLines.length > 0) {
                html += `
                    <div class="diff-lines-section removed">
                        <h4>❌ Removed (${removedLines.length} lines)</h4>
                        <div class="line-numbers">
                            ${removedLines.map(line => `<span class="line-number-badge">Line ${line}</span>`).join('')}
                        </div>
                    </div>
                `;
            }
            
            if (modifiedLines.length > 0) {
                html += `
                    <div class="diff-lines-section modified">
                        <h4>✏️ Modified (${modifiedLines.length} lines)</h4>
                        <div class="line-numbers">
                            ${modifiedLines.map(line => `<span class="line-number-badge">Line ${line}</span>`).join('')}
                        </div>
                    </div>
                `;
            }
            
            if (html === '') {
                html = '<div class="no-differences">No differences found between the documents.</div>';
            }
            
            container.innerHTML = html;
            document.getElementById('differencesSummary').style.display = 'block';
        }
    </script>
</body>
</html>
"""

def extract_text_from_docx(file_stream):
    """Extract text from .docx file"""
    try:
        doc = Document(file_stream)
        text = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                text.append(paragraph.text)
        return '\n'.join(text)
    except Exception as e:
        try:
            file_stream.seek(0)
            return docx2txt.process(file_stream)
        except:
            raise Exception(f"Failed to parse DOCX file: {str(e)}")

def tokenize_text(text):
    """Tokenize text into words while preserving line breaks and Arabic characters"""
    lines = text.split('\n')
    tokenized_lines = []
    for line in lines:
        tokens = re.split(r'(\s+)', line) if line.strip() else ['']
        tokens = [token for token in tokens if token]
        tokenized_lines.append(tokens)
    return tokenized_lines

def compare_documents(doc1_tokens, doc2_tokens):
    """Compare tokenized documents and return alignment information"""
    flat_doc1 = [token for line in doc1_tokens for token in line]
    flat_doc2 = [token for line in doc2_tokens for token in line]
    
    matcher = SequenceMatcher(None, flat_doc1, flat_doc2)
    aligned_doc1 = []
    aligned_doc2 = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            aligned_doc1.extend([(token, 'same') for token in flat_doc1[i1:i2]])
            aligned_doc2.extend([(token, 'same') for token in flat_doc2[j1:j2]])
        elif tag == 'replace':
            aligned_doc1.extend([(token, 'different') for token in flat_doc1[i1:i2]])
            aligned_doc2.extend([(token, 'different') for token in flat_doc2[j1:j2]])
        elif tag == 'delete':
            aligned_doc1.extend([(token, 'missing') for token in flat_doc1[i1:i2]])
            aligned_doc2.extend([('', 'missing') for _ in range(i2 - i1)])
        elif tag == 'insert':
            aligned_doc1.extend([('', 'missing') for _ in range(j2 - j1)])
            aligned_doc2.extend([(token, 'added') for token in flat_doc2[j1:j2]])
    
    def reconstruct_lines(aligned_tokens, original_lines):
        result_lines = []
        token_index = 0
        
        for original_line in original_lines:
            line_tokens = []
            for original_token in original_line:
                if token_index < len(aligned_tokens):
                    token, status = aligned_tokens[token_index]
                    line_tokens.append((token, status))
                    token_index += 1
                else:
                    line_tokens.append(('', 'missing'))
            
            while (token_index < len(aligned_tokens) and 
                   aligned_tokens[token_index][0] in [' ', '\t']):
                token, status = aligned_tokens[token_index]
                line_tokens.append((token, status))
                token_index += 1
                
            result_lines.append(line_tokens)
        
        if token_index < len(aligned_tokens):
            result_lines.append([(token, status) for token, status in aligned_tokens[token_index:]])
        
        return result_lines
    
    doc1_aligned = reconstruct_lines(aligned_doc1, doc1_tokens)
    doc2_aligned = reconstruct_lines(aligned_doc2, doc2_tokens)
    
    return doc1_aligned, doc2_aligned

def generate_html_content(aligned_lines):
    """Generate HTML content with line numbers and colored tokens"""
    html_lines = []
    
    for line_num, line_tokens in enumerate(aligned_lines, 1):
        line_html = f'<div class="content-line"><span class="line-number">{line_num}</span><div class="line-content">'
        
        for token, status in line_tokens:
            if token.strip():
                escaped_token = token.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                line_html += f'<span class="word {status}">{escaped_token}</span>'
            else:
                line_html += token.replace(' ', '&nbsp;').replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')
        
        line_html += '</div></div>'
        html_lines.append(line_html)
    
    return ''.join(html_lines)

def analyze_line_differences(doc1_aligned, doc2_aligned):
    """Analyze which lines contain differences"""
    added_lines = []
    removed_lines = []
    modified_lines = []
    
    max_lines = max(len(doc1_aligned), len(doc2_aligned))
    
    for line_num in range(max_lines):
        doc1_line = doc1_aligned[line_num] if line_num < len(doc1_aligned) else []
        doc2_line = doc2_aligned[line_num] if line_num < len(doc2_aligned) else []
        
        doc1_has_diff = any(status in ['different', 'missing'] for token, status in doc1_line if token.strip())
        doc2_has_diff = any(status in ['different', 'added'] for token, status in doc2_line if token.strip())
        
        doc1_has_missing = any(status == 'missing' for token, status in doc1_line if token.strip())
        doc2_has_added = any(status == 'added' for token, status in doc2_line if token.strip())
        
        if doc1_has_missing and doc2_has_added:
            if not any(status == 'different' for token, status in doc1_line if token.strip()):
                if doc1_has_missing:
                    removed_lines.append(line_num + 1)
                if doc2_has_added:
                    added_lines.append(line_num + 1)
            else:
                modified_lines.append(line_num + 1)
        elif doc1_has_diff or doc2_has_diff:
            modified_lines.append(line_num + 1)
    
    return {
        'added': sorted(set(added_lines)),
        'removed': sorted(set(removed_lines)),
        'modified': sorted(set(modified_lines))
    }

def calculate_analytics(doc1_aligned, doc2_aligned):
    """Calculate detailed comparison analytics"""
    total_words_doc1 = sum(len([t for t, s in line if t.strip()]) for line in doc1_aligned)
    total_words_doc2 = sum(len([t for t, s in line if t.strip()]) for line in doc2_aligned)
    
    removed_words = sum(len([t for t, s in line if t.strip() and s == 'missing']) for line in doc1_aligned)
    added_words = sum(len([t for t, s in line if t.strip() and s == 'added']) for line in doc2_aligned)
    modified_words = sum(len([t for t, s in line if t.strip() and s == 'different']) for line in doc1_aligned)
    
    same_words = sum(len([t for t, s in line if t.strip() and s == 'same']) for line in doc1_aligned)
    total_words = max(total_words_doc1, total_words_doc2)
    
    similarity = round((same_words / total_words * 100) if total_words > 0 else 100, 1)
    
    return {
        'total_words': total_words,
        'added_words': added_words,
        'removed_words': removed_words,
        'modified_words': modified_words,
        'similarity': similarity
    }

@app.route('/')
def index():
    """Main page"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/compare', methods=['POST'])
def compare():
    """Handle document comparison"""
    try:
        if 'doc1' not in request.files or 'doc2' not in request.files:
            return jsonify({'success': False, 'error': 'Please upload both documents'})
        
        doc1_file = request.files['doc1']
        doc2_file = request.files['doc2']
        
        if doc1_file.filename == '' or doc2_file.filename == '':
            return jsonify({'success': False, 'error': 'Please select both files'})
        
        if not doc1_file.filename.lower().endswith('.docx') or not doc2_file.filename.lower().endswith('.docx'):
            return jsonify({'success': False, 'error': 'Only .docx files are supported'})
        
        doc1_text = extract_text_from_docx(doc1_file.stream)
        doc2_text = extract_text_from_docx(doc2_file.stream)
        
        doc1_tokens = tokenize_text(doc1_text)
        doc2_tokens = tokenize_text(doc2_text)
        
        doc1_aligned, doc2_aligned = compare_documents(doc1_tokens, doc2_tokens)
        
        doc1_html = generate_html_content(doc1_aligned)
        doc2_html = generate_html_content(doc2_aligned)
        
        analytics = calculate_analytics(doc1_aligned, doc2_aligned)
        line_differences = analyze_line_differences(doc1_aligned, doc2_aligned)
        
        return jsonify({
            'success': True,
            'doc1_name': doc1_file.filename,
            'doc2_name': doc2_file.filename,
            'doc1_html': doc1_html,
            'doc2_html': doc2_html,
            'analytics': analytics,
            'line_differences': line_differences
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    try:
        pythoncom.CoInitialize()
    except:
        pass
    
    print("=" * 60)
    print("xsukax Word Document Comparator")
    print("=" * 60)
    print("Server running at: http://localhost:5000")
    print("Press CTRL+C to stop the server")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)