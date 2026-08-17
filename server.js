const express = require('express');
const axios = require('axios');
const FormData = require('form-data');
const Groq = require('groq-sdk');
const ExcelJS = require('exceljs');
const auth = require('basic-auth');
const { createClient } = require('@supabase/supabase-js');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
app.use(express.json());

// 🔒 1. PASSWORD PROTECTION MIDDLEWARE GATE
app.use((req, res, next) => {
    const credentials = auth(req);
    const adminUser = process.env.ADMIN_USER || 'admin';
    const adminPass = process.env.ADMIN_PASS || 'password123';
    if (!credentials || credentials.name !== adminUser || credentials.pass !== adminPass) {
        res.setHeader('WWW-Authenticate', 'Basic realm="Secure Scraper"');
        return res.status(401).send('Authentication Required!');
    }
    next();
});

// 📂 2. STATIC INLINE USER INTERFACE (No index.html file needed on disk!)
app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Mobile Subtitle Scraper</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f4f6f9; padding: 15px; margin: 0; }
        .container { max-width: 100%; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); box-sizing: border-box; }
        h2 { color: #333; margin: 0 0 15px 0; font-size: 20px; text-align: center; }
        .system-controls { display: flex; gap: 10px; margin-bottom: 15px; }
        .system-controls button { flex: 1; padding: 10px; font-size: 12px; font-weight: bold; border-radius: 6px; border: none; color: white;}
        .btn-danger { background: #dc3545 !important; }
        .btn-secondary { background: #6c757d !important; }
        label { font-size: 13px; font-weight: 600; color: #444; }
        input[type="text"], select { width: 100%; padding: 12px; margin: 8px 0 15px 0; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; font-size: 14px; }
        .attr-box { background: #f9f9f9; padding: 10px; border-radius: 8px; border: 1px dashed #bbb; margin-bottom: 15px; min-height: 40px; }
        .chip { display: inline-flex; align-items: center; background: #e0e0e0; padding: 6px 12px; border-radius: 20px; margin: 4px; font-size: 12px; font-weight: 500; }
        .chip-remove { margin-left: 8px; color: #888; cursor: pointer; font-weight: bold; }
        .btn-main { background: #28a745; color: white; border: none; padding: 14px 20px; border-radius: 6px; cursor: pointer; font-weight: 600; width: 100%; font-size: 15px; }
        .player-section { margin-top: 15px; background: #111; padding: 10px; border-radius: 8px; text-align: center; color: white; }
        video { width: 100%; max-height: 280px; border-radius: 4px; background: #000; }
        video::cue { background: rgba(0, 0, 0, 0.8); color: #fff; font-size: 14px; font-weight: bold; }
        .media-box { display: none; background: #eef7ff; padding: 15px; border-radius: 8px; border: 1px solid #bce0ff; margin-top: 15px; }
        .loading { display: none; color: #007bff; font-weight: bold; text-align: center; margin-top: 15px; font-size: 14px; font-style: italic; }
    </style>
</head>
<body>
<div class="container">
    <h2>📊 Mobile AI Subtitle Scraper</h2>
    <div class="system-controls">
        <button class="btn-danger" onclick="clearSystemHistory()">🗑️ Wipe Data</button>
        <button class="btn-secondary" onclick="createNewCampaignSheet()">📁 New Sheet</button>
    </div>
    <label><b>Active Sheet Tab Target:</b></label>
    <select id="sheetSelector"><option value="AI Master Data Log">AI Master Data Log (Default)</option></select>
    <label><b>Instagram URL Link:</b></label>
    <input type="text" id="targetUrl" placeholder="Paste Reel link here...">
    <label><b>Custom Column Attributes:</b></label>
    <div style="display: flex; gap: 10px; margin-bottom: 8px;">
        <input type="text" id="newAttr" placeholder="e.g., Video_Hook" style="margin:0;">
        <button type="button" id="addBtn" style="width: auto; margin:0; background:#6c757d; font-weight: bold; border-radius: 6px; color: white; border: none; padding: 0 15px;">➕ Add</button>
    </div>
    <div class="attr-box" id="chipsContainer"></div>
    <button class="btn-main" onclick="processDataPipeline()">🚀 Scan & Generate Subtitles</button>
    <div class="loading" id="loader">🧠 Groq AI is transcribing audio and compiling columns in RAM...</div>
    <div class="media-box" id="resultPanel">
        <h4 style="margin:0 0 10px 0; color:#0056b3; text-align:center;">✅ Process Complete</h4>
        <div class="player-section">
            <video id="videoPlayer" controls playsinline>
                <source id="videoSource" src="" type="video/mp4">
                <track id="subtitleTrack" label="English CC" kind="subtitles" srclang="en" src="" default>
            </video>
        </div>
        <p style="font-size:11px; color:#666; text-align:center; margin: 5px 0 12px 0;">📺 Click 'CC' on your mobile screen to view live subtitles.</p>
        <button style="background:#007bff;" class="btn-main" id="masterExcelBtn">📥 Download Excel Database File</button>
    </div>
</div>
<script>
    let dynamicAttributes = ['Creator_Name', 'Video_Hook', 'Product_Promoted'];
    const chipsContainer = document.getElementById('chipsContainer');
    function renderChips() {
        chipsContainer.innerHTML = '';
        dynamicAttributes.forEach((attr, idx) => {
            chipsContainer.innerHTML += '<div class="chip">' + attr + ' <span class="chip-remove" onclick="removeAttribute(' + idx + ')">×</span></div>';
        });
    }
    document.getElementById('addBtn').addEventListener('click', () => {
        const val = document.getElementById('newAttr').value.trim().replace(/[^a-zA-Z0-9_]/g, '');
        if(val && !dynamicAttributes.includes(val)) { dynamicAttributes.push(val); document.getElementById('newAttr').value = ''; renderChips(); }
    });
    function removeAttribute(idx) { dynamicAttributes.splice(idx, 1); renderChips(); }
    async function clearSystemHistory() {
        if(!confirm("Permanently clear logs?")) return;
        await fetch('/api/clear-history', { method: 'POST' });
        location.reload();
    }
    async function createNewCampaignSheet() {
        const name = prompt("Enter sheet name:");
        if(!name) return;
        const res = await fetch('/api/new-sheet', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ sheetName: name }) });
        const d = await res.json();
        if(d.success) {
            const selector = document.getElementById('sheetSelector'); selector.innerHTML = '';
            d.currentSheets.forEach(s => { selector.innerHTML += '<option value="' + s + '">' + s + '</option>'; });
            selector.value = name;
        }
    }
    async function processDataPipeline() {
        const url = document.getElementById('targetUrl').value.trim();
        if(!url) return alert('Input an Instagram URL.');
        document.getElementById('loader').style.display = 'block';
        document.getElementById('resultPanel').style.display = 'none';
        try {
            const response = await fetch('/api/scrape', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ url: url, attributes: dynamicAttributes, targetSheet: document.getElementById('sheetSelector').value }) });
            const data = await response.json();
            if (data.error) throw new Error(data.error);
            document.getElementById('videoSource').src = data.videoLink;
            document.getElementById('subtitleTrack').src = data.vttLink;
            document.getElementById('videoPlayer').load();
            document.getElementById('masterExcelBtn').onclick = () => { window.location.href = '/api/download-excel'; };
            document.getElementById('resultPanel').style.display = 'block';
        } catch (err) { alert('Error: ' + err.message); } finally { document.getElementById('loader').style.display = 'none'; }
    }
    renderChips();
</script>
</body>
</html>
    `);
});

// Dynamic routing for generated media subtitle streams
const mediaDir = path.join(__dirname, 'media');
app.use('/media', express.static(mediaDir));

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

// 🗑️ 3. ROUTE: WIPE HISTORY IN SUPABASE
app.post('/api/clear-history', async (req, res) => {
    try {
        await supabase.from('instagram_campaigns').delete().neq('id', 0);
        res.json({ success: true, message: "History cleared." });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 📁 4. ROUTE: DYNAMIC CAMPAIGN TRACKER LIST GENERATION
app.post('/api/new-sheet', async (req, res) => {
    const { sheetName } = req.body;
    try {
        let safeName = sheetName.replace(/[^a-zA-Z0-9 ]/g, '') || "Campaign_Log";
        const { data: rows } = await supabase.from('instagram_campaigns').select('extracted_data');
        let sheetsList = new Set(['AI Master Data Log (Default)']);
        rows.forEach(r => {
            if (r.extracted_data && r.extracted_data._sheet_target) sheetsList.add(r.extracted_data._sheet_target);
        });
        sheetsList.add(safeName);
        res.json({ success: true, currentSheets: Array.from(sheetsList) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 📥 5. ROUTE: GENERATE EXCEL AND STREAM STRAIGHT TO SMARTPHONE
app.get('/api/download-excel', async (req, res) => {
    try {
        const { data: rows } = await supabase.from('instagram_campaigns').select('*').order('scraped_at', { ascending: false });
        const workbook = new ExcelJS.Workbook();
        
