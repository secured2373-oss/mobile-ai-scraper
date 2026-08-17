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

// 🔒 PASSWORD PROTECTION GATING MIDDLEWARE
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

app.use(express.static('public'));
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

app.post('/api/clear-history', async (req, res) => {
    try {
        await supabase.from('instagram_campaigns').delete().neq('id', 0);
        res.json({ success: true, message: "History cleared." });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/download-excel', async (req, res) => {
    try {
        const { data: rows } = await supabase.from('instagram_campaigns').select('*').order('scraped_at', { ascending: false });
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Master Data');
        let dynamicKeys = new Set();
        rows.forEach(r => r.extracted_data && Object.keys(r.extracted_data).forEach(k => dynamicKeys.add(k)));
        
        let headers = ['Scraped At', 'Source URL', 'AI Summary Text', ...Array.from(dynamicKeys)];
        worksheet.columns = headers.map(h => ({ header: h, key: h, width: 25 }));
        rows.forEach(r => worksheet.addRow({ 'Scraped At': new Date(r.scraped_at).toLocaleString(), 'Source URL': r.source_url, 'AI Summary Text': r.ai_summary, ...r.extracted_data }));
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');
        await workbook.xlsx.write(res);
    } catch (err) { res.status(500).send(err.message); }
});

app.post('/api/scrape', async (req, res) => {
    const { url, attributes } = req.body;
    try {
        const shortcode = url.match(/(?:\/p\/|\/reel\/)([A-Za-z0-9_-]+)/)[1];
        const instaRes = await axios.get(`https://instagram.com{shortcode}/?__a=1&__d=dis`, { headers: { 'User-Agent': 'Mozilla/5.0' } });
        const item = instaRes.data?.items?.[0];
        const writtenCaption = item?.caption?.text || "No caption";
        const videoUrl = item?.video_versions?.[0]?.url;
        
        let spokenTranscript = "No speech detected.";
        let vttSubtitleContent = "WEBVTT\n\n00:00:00.000 --> 00:00:05.000\n[No Spoken Audio]";
        
        const videoFilename = `video_${shortcode}.mp4`;
        const vttFilename = `subs_${shortcode}.vtt`;
        
        if (videoUrl) {
            const videoBuf = await axios({ url: videoUrl, method: 'GET', responseType: 'arraybuffer' });
            fs.writeFileSync(path.join(__dirname, 'public', 'media', videoFilename), videoBuf.data);
            fs.writeFileSync(path.join(__dirname, 'public', 'media', 'temp.mp3'), videoBuf.data);

            const form = new FormData();
            form.append('file', fs.createReadStream(path.join(__dirname, 'public', 'media', 'temp.mp3')));
            form.append('model', 'whisper-large-v3-turbo');
            form.append('response_format', 'verbose_json');

            const transRes = await axios.post('https://groq.com', form, {
                headers: { ...form.getHeaders(), 'Authorization': `Bearer ${process.env.GROQ_API_KEY}` }
            });
            spokenTranscript = transRes.data.text || "";
            
            if(transRes.data.segments) {
                let vttLines = ["WEBVTT\n"];
                transcriptionRes.data.segments.forEach((seg) => {
                    const formatTime = (secs) => {
                        let date = new Date(0);
                        date.setSeconds(secs);
                        let ms = Math.floor((secs % 1) * 1000).toString().padStart(3, '0');
                        return date.toISOString().substr(11, 8) + '.' + ms;
                    };
                    vttLines.push(`${formatTime(seg.start)} --> ${formatTime(seg.end)}`);
                    vttLines.push(`${seg.text.trim()}\n`);
                });
                vttSubtitleContent = vttLines.join('\n');
            }
            fs.writeFileSync(path.join(__dirname, 'public', 'media', vttFilename), vttSubtitleContent);
            fs.unlinkSync(path.join(__dirname, 'public', 'media', 'temp.mp3'));
        }

        const sysPrompt = `Extract ONLY keys: ${JSON.stringify(attributes)} alongside an "ai_summary" key. Return raw JSON.`;
        const completion = await groq.chat.completions.create({
            model: "llama-3.3-70b-specdec",
            messages: [{ role: "system", content: sysPrompt }, { role: "user", content: `Cap: ${writtenCaption}\nTr: ${spokenTranscript}` }],
            response_format: { type: "json_object" }
        });
        
        const aiResult = JSON.parse(completion.choices.message.content);
        const { ai_summary, ...customAttributes } = aiResult;

        await supabase.from('instagram_campaigns').insert([{ source_url: url, ai_summary: ai_summary || "Processed", extracted_data: customAttributes }]);
        res.json({ success: true, videoLink: `/media/${videoFilename}`, vttLink: `/media/${vttFilename}` });
    } catch (error) { res.status(500).json({ error: error.message }); }
});

app.listen(process.env.PORT || 3000);
  
