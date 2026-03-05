/**
 * DTC Automation Script
 * Base Version: 3.4.0
 * Update: Added Vehicle Data Validation + Fast Wait + Timing Fix
 * Changes:
 * - Replaced Hard Waits with `checkAndWait` that looks for actual table rows, not just buttons.
 * - Added a strict 5-second Safety Buffer after data appears before clicking export.
 * - Improved `prepareBeforeSearch` to wipe old table rows entirely to prevent false positives.
 * - Checks specifically for '-' in the license plate column to ensure valid data.
 * - Retries searching and downloading up to 3 times if data is missing or corrupted.
 */

const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// --- Helper Functions ---

// 1. ฟังก์ชันรอโหลดไฟล์ และแปลงเป็น CSV
async function waitForDownloadAndRename(downloadPath, newFileName, maxWaitMs = 120000) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;
    const checkInterval = 5000; 
    let waittime = 0;

    while (waittime < maxWaitMs) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_') &&
            !f.startsWith('Converted_')
        );
        
        if (downloadedFile) {
            console.log(`   ✅ File detected: ${downloadedFile} (${waittime/1000}s)`);
            break; 
        }
        
        await new Promise(resolve => setTimeout(resolve, checkInterval));
        waittime += checkInterval;
    }

    if (!downloadedFile) throw new Error(`Download timeout for ${newFileName}`);

    await new Promise(resolve => setTimeout(resolve, 5000)); // รอไฟล์เขียนเสร็จ

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) console.warn(`   ⚠️ Downloaded file is empty!`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    const csvFileName = `Converted_${newFileName.replace('.xls', '.csv')}`;
    const csvPath = path.join(downloadPath, csvFileName);
    await convertToCsv(newPath, csvPath);
    
    return csvPath;
}

// 2. ฟังก์ชันแปลงไฟล์ (XLSX/HTML -> CSV)
async function convertToCsv(sourcePath, destPath) {
    try {
        console.log(`   🔄 Converting to CSV...`);
        const buffer = fs.readFileSync(sourcePath);
        let rows = [];

        const isXLSX = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;

        if (isXLSX) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1);
            
            worksheet.eachRow((row) => {
                const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
                rows.push(rowValues.map(v => {
                    if (v === null || v === undefined) return '';
                    if (typeof v === 'object') return v.text || v.result || '';
                    return String(v).trim();
                }));
            });
        } else {
            const content = buffer.toString('utf8');
            const dom = new JSDOM(content);
            const table = dom.window.document.querySelector('table');
            if (table) {
                const trs = Array.from(table.querySelectorAll('tr'));
                rows = trs.map(tr => 
                    Array.from(tr.querySelectorAll('td, th')).map(td => td.textContent.replace(/\s+/g, ' ').trim())
                );
            }
        }

        if (rows.length > 0) {
            let csvContent = '\uFEFF'; 
            rows.forEach(row => {
                const escapedRow = row.map(cell => {
                    if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
                        return `"${cell.replace(/"/g, '""')}"`;
                    }
                    return cell;
                });
                csvContent += escapedRow.join(',') + '\n';
            });
            fs.writeFileSync(destPath, csvContent, 'utf8');
            console.log(`   ✅ CSV Created: ${path.basename(destPath)}`);
        }
    } catch (e) {
        console.warn(`   ⚠️ CSV Conversion error: ${e.message}`);
    }
}

// --- NEW Helper: Validating Vehicle Data ---
// ตรวจสอบว่าในไฟล์มีข้อมูลทะเบียนรถ (เช็คจากเครื่องหมาย "-") หรือไม่
function validateVehicleData(filePath, colIndex) {
    try {
        if (!filePath || filePath === '') return false;
        if (!fs.existsSync(filePath)) return false;
        if (fs.statSync(filePath).size === 0) return false;

        const fileContent = fs.readFileSync(filePath, 'utf8');
        const rows = parse(fileContent, {
            columns: false,
            skip_empty_lines: true,
            relax_column_count: true, 
            bom: true
        });

        // 1. หาหัวตาราง
        let headerIndex = -1;
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            if (rows[i].some(cell => cell.includes('ลำดับ'))) {
                headerIndex = i;
                break;
            }
        }

        if (headerIndex === -1) {
            console.warn(`   ⚠️ Validation Failed: ไม่พบหัวตาราง 'ลำดับ'`);
            return false;
        }

        // 2. เช็คข้อมูลคอลัมน์ทะเบียนรถว่ามี '-' โผล่มาหรือไม่
        const dataRows = rows.slice(headerIndex + 1);
        for (const row of dataRows) {
            const license = row[colIndex] ? String(row[colIndex]).trim() : '';
            if (license && license.includes('-')) {
                return true; 
            }
        }

        console.warn(`   ⚠️ Validation Failed: ไม่พบข้อมูลทะเบียนรถที่มี '-' ในคอลัมน์ ${colIndex}`);
        return false;
    } catch (err) {
        console.error(`   ❌ Validation Error:`, err.message);
        return false;
    }
}

// --- NEW Helper: Fast Wait & Empty Check (TIMING FIX) ---
// เช็คสถานะหน้าเว็บเพื่อลดเวลาการรอ และกันการกดเร็วเกินไป
async function checkAndWait(page, maxWaitMs) {
    console.log(`   ⏳ Waiting for system to process (Max ${maxWaitMs/1000}s)...`);
    const startTime = Date.now();
    let isEmpty = false;

    // เผื่อเวลาให้เว็บส่ง Request และเปิด Loader 
    await page.waitForFunction(() => {
        const loaders = document.querySelectorAll('.blockUI, #loading, #loader, .loading, .spinner');
        return Array.from(loaders).some(el => el.offsetParent !== null && window.getComputedStyle(el).display !== 'none');
    }, { timeout: 5000 }).catch(() => {});

    while (Date.now() - startTime < maxWaitMs) {
        const status = await page.evaluate(() => {
            // 1. เช็คว่ามี Spinner/Loader กำลังหมุนอยู่ไหม
            const loaders = document.querySelectorAll('.blockUI, #loading, #loader, .loading, .spinner, .modal-backdrop');
            for(let el of loaders) {
                if(el.offsetParent !== null && window.getComputedStyle(el).display !== 'none') return { ready: false };
            }

            // 2. เช็คกรณีไม่มีข้อมูล
            const bodyText = document.body.innerText;
            if(bodyText.includes('ไม่พบข้อมูล') || bodyText.includes('No data found')) {
                return { ready: true, empty: true };
            }

            // 3. เช็คว่าตารางมีข้อมูลแถวใหม่เข้ามาแล้วจริงๆ (กรองหัวตาราง th ทิ้ง)
            const rows = document.querySelectorAll('table tr');
            let dataRowCount = 0;
            for(let i=0; i<rows.length; i++) {
                if(!rows[i].querySelector('th')) dataRowCount++;
            }
            if (dataRowCount > 0) return { ready: true, empty: false };

            // 4. กรณีหาตารางไม่เจอจริงๆ แต่ปุ่ม Export มาชัวร์ๆ (แผนสำรอง)
            const btn = document.getElementById('btnexport');
            if(btn && btn.offsetParent !== null && !btn.disabled) return { ready: true, empty: false };

            return { ready: false };
        });

        if (status.ready) {
            const timeTaken = ((Date.now() - startTime) / 1000).toFixed(1);
            if (status.empty) {
                console.log(`   ℹ️ System ready in ${timeTaken}s. Result: 'ไม่พบข้อมูล' (No Data).`);
                isEmpty = true;
            } else {
                console.log(`   ⚡ Data visually ready in ${timeTaken}s!`);
            }
            break;
        }
        await new Promise(r => setTimeout(r, 2000));
    }
    
    // *** SAFETY BUFFER: บังคับหน่วงเวลาอีก 5 วินาที เพื่อให้เว็บเรนเดอร์ชัวร์ๆ ก่อนกดปุ่ม Export ***
    if (!isEmpty) {
        console.log(`   ⏳ Applying 5s safety buffer before attempting to export...`);
        await new Promise(r => setTimeout(r, 5000));
    }

    return isEmpty;
}

// --- Helper: Clear Old Data (ROBUST DOM CLEARING) ---
// ล้างตารางเก่าทิ้งป้องกันการโดนหลอกในรอบถัดไป
async function prepareBeforeSearch(page) {
    await page.evaluate(() => {
        // ลบแถวข้อมูลตารางทิ้ง (เก็บไว้แค่บรรทัดหัวตาราง <th>)
        const rows = document.querySelectorAll('table tr');
        rows.forEach(tr => {
            if (!tr.querySelector('th')) {
                tr.remove();
            }
        });

        // ล้างข้อความ 'ไม่พบข้อมูล'
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        let node;
        const nodesToRemove = [];
        while (node = walker.nextNode()) {
            if (node.nodeValue.includes('ไม่พบข้อมูล') || node.nodeValue.includes('No data found')) {
                nodesToRemove.push(node);
            }
        }
        nodesToRemove.forEach(n => n.nodeValue = '');
    });
}

// --- Helper: Parse Date ---
function parseDateTimeToSeconds(dateStr) {
    if (!dateStr) return 0;
    const parts = dateStr.split(/[ /:-]/);
    if (parts.length < 6) return 0;
    
    let day, month, year, hour, minute, second;
    if (parts[0].length === 4) {
        year = parseInt(parts[0]);
        month = parseInt(parts[1]) - 1; 
        day = parseInt(parts[2]);
    } else {
        day = parseInt(parts[0]);
        month = parseInt(parts[1]) - 1;
        year = parseInt(parts[2]);
    }
    hour = parseInt(parts[3]);
    minute = parseInt(parts[4]);
    second = parseInt(parts[5]);

    const date = new Date(year, month, day, hour, minute, second);
    return date.getTime() / 1000;
}

// --- Helper: Format Seconds to HH:MM:SS ---
function formatSeconds(totalSeconds) {
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = Math.floor(totalSeconds % 60);
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

// --- FUNCTION: Process CSV V3 ---
function processCSV_V3(filePath, config) {
    try {
        if (!filePath || filePath === '') return [];
        if (!fs.existsSync(filePath)) return [];

        const fileContent = fs.readFileSync(filePath, 'utf8');
        const rows = parse(fileContent, {
            columns: false,
            skip_empty_lines: true,
            relax_column_count: true,
            bom: true
        });

        let headerIndex = -1;
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            if (rows[i].some(cell => cell.includes('ลำดับ'))) {
                headerIndex = i;
                break;
            }
        }

        if (headerIndex === -1) return [];

        const dataRows = rows.slice(headerIndex + 1);
        const results = [];

        dataRows.forEach(row => {
            const license = row[config.colLicense] ? row[config.colLicense].trim() : '';

            if (license && license.includes('-')) {
                const item = { license };

                if (config.useTimeCalc && config.colStart !== undefined && config.colEnd !== undefined) {
                    const t1 = parseDateTimeToSeconds(row[config.colStart]); 
                    const t2 = parseDateTimeToSeconds(row[config.colEnd]);   
                    item.durationSec = (t2 > t1) ? (t2 - t1) : 0;
                    item.durationStr = formatSeconds(item.durationSec);
                }
                
                if (config.colDate !== undefined) item.date = row[config.colDate]; 
                if (config.colStation !== undefined) item.station = row[config.colStation];
                if (config.colSpeedStart !== undefined) item.v_start = row[config.colSpeedStart];
                if (config.colSpeedEnd !== undefined) item.v_end = row[config.colSpeedEnd];

                results.push(item);
            }
        });

        return results;

    } catch (err) {
        console.error(`Error processing ${filePath}:`, err.message);
        return [];
    }
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

function zipFiles(sourceDir, outPath, filesToZip) {
    return new Promise((resolve, reject) => {
        const output = fs.createWriteStream(outPath);
        const archive = archiver('zip', { zlib: { level: 9 } });
        output.on('close', () => resolve(outPath));
        archive.on('error', (err) => reject(err));
        archive.pipe(output);
        filesToZip.forEach(file => archive.file(path.join(sourceDir, file), { name: file }));
        archive.finalize();
    });
}

// --- Main Script ---

(async () => {
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing Secrets.');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Base 3.4.0 + Validated + Timing Fix)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(3600000); 
    page.setDefaultTimeout(3600000);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    const MAX_RETRIES = 3; 

    try {
        // Step 1: Login
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`🕒 Global Time Settings: ${startDateTime} to ${endDateTime}`);

        // --- Step 2 to 6: DOWNLOAD REPORTS ---
        
        // REPORT 1: Over Speed
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        
        await new Promise(r => setTimeout(r, 10000));
        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) {
                document.getElementById('ddlMinute').value = '1';
                document.getElementById('ddlMinute').dispatchEvent(new Event('change'));
            }
            var selectElement = document.getElementById('ddl_truck'); 
            var options = selectElement.options; 
            for (var i = 0; i < options.length; i++) { 
                if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
            } 
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
        }, startDateTime, endDateTime);

        let file1 = '';
        for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            console.log(`   Searching Report 1 (Attempt ${attempt}/${MAX_RETRIES})...`);
            
            // เคลียร์ตารางของรอบที่แล้วก่อนกดปุ่ม
            await prepareBeforeSearch(page);
            
            await page.evaluate(() => {
                if(typeof sertch_data === 'function') sertch_data();
                else document.querySelector("span[onclick='sertch_data();']").click();
            });

            const isEmpty = await checkAndWait(page, 300000); 
            if (isEmpty) {
                console.log(`   ⏩ Skipping Report 1 (No data found).`);
                break;
            }
            
            console.log('   Exporting Report 1...');
            await page.evaluate(() => {
                const btn = document.getElementById('btnexport');
                if (btn) btn.click();
            });
            
            file1 = await waitForDownloadAndRename(downloadPath, `Report1_OverSpeed_Att${attempt}.xls`);
            
            if (validateVehicleData(file1, 1)) {
                console.log(`   ✅ Report 1 is valid (Found vehicle data).`);
                break;
            } else {
                console.warn(`   ⚠️ Report 1 invalid or empty. Deleting...`);
                if (fs.existsSync(file1)) fs.unlinkSync(file1);
                file1 = '';
            }
        }

        // REPORT 2: Idling
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);
        
        let file2 = '';
        for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            console.log(`   Searching Report 2 (Attempt ${attempt}/${MAX_RETRIES})...`);
            
            await prepareBeforeSearch(page);
            
            await page.evaluate(() => {
                const btn = document.querySelector('td:nth-of-type(6) > span');
                if (btn) btn.click();
            });
            
            const isEmpty = await checkAndWait(page, 180000); 
            if (isEmpty) {
                console.log(`   ⏩ Skipping Report 2 (No data found).`);
                break;
            }

            console.log('   Exporting Report 2...');
            await page.evaluate(() => {
                const btn = document.getElementById('btnexport');
                if (btn) btn.click();
            });
            
            file2 = await waitForDownloadAndRename(downloadPath, `Report2_Idling_Att${attempt}.xls`);
            
            if (validateVehicleData(file2, 1)) {
                console.log(`   ✅ Report 2 is valid (Found vehicle data).`);
                break;
            } else {
                console.warn(`   ⚠️ Report 2 invalid or empty. Deleting...`);
                if (fs.existsSync(file2)) fs.unlinkSync(file2);
                file2 = '';
            }
        }

        // REPORT 3: Sudden Brake
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);
        
        let file3 = '';
        for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            console.log(`   Searching Report 3 (Attempt ${attempt}/${MAX_RETRIES})...`);
            
            await prepareBeforeSearch(page);
            
            await page.evaluate(() => {
                const btn = document.querySelector('td:nth-of-type(6) > span');
                if (btn) btn.click();
            });
            
            const isEmpty = await checkAndWait(page, 180000); 
            if (isEmpty) {
                console.log(`   ⏩ Skipping Report 3 (No data found).`);
                break;
            }

            console.log('   Exporting Report 3...');
            await page.evaluate(() => {
                const btns = Array.from(document.querySelectorAll('button'));
                const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
                if (b) {
                    b.click();
                } else {
                    const fallbackBtn = document.querySelector('#table button:nth-of-type(3)');
                    if (fallbackBtn) fallbackBtn.click();
                }
            });
            
            file3 = await waitForDownloadAndRename(downloadPath, `Report3_SuddenBrake_Att${attempt}.xls`);
            
            if (validateVehicleData(file3, 1)) {
                console.log(`   ✅ Report 3 is valid (Found vehicle data).`);
                break;
            } else {
                console.warn(`   ⚠️ Report 3 invalid or empty. Deleting...`);
                if (fs.existsSync(file3)) fs.unlinkSync(file3);
                file3 = '';
            }
        }

        // REPORT 4: Harsh Start
        console.log('📊 Processing Report 4: Harsh Start...');
        let file4 = '';
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9', { visible: true, timeout: 60000 });
            await new Promise(r => setTimeout(r, 10000));
            await page.evaluate((start, end) => {
                document.getElementById('date9').value = start;
                document.getElementById('date10').value = end;
                document.getElementById('date9').dispatchEvent(new Event('change'));
                document.getElementById('date10').dispatchEvent(new Event('change'));
                const select = document.getElementById('ddl_truck');
                if (select) {
                    let found = false;
                    for (let i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                            select.selectedIndex = i; found = true; break;
                        }
                    }
                    if (!found && select.options.length > 0) select.selectedIndex = 0;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof $ !== 'undefined' && $(select).data('select2')) { $(select).trigger('change'); }
                }
            }, startDateTime, endDateTime);
            
            for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
                console.log(`   Searching Report 4 (Attempt ${attempt}/${MAX_RETRIES})...`);
                
                await prepareBeforeSearch(page);
                
                await page.evaluate(() => {
                    if (typeof sertch_data === 'function') { 
                        sertch_data(); 
                    } else { 
                        const btn = document.querySelector('td:nth-of-type(6) > span');
                        if (btn) btn.click(); 
                    }
                });
                
                const isEmpty = await checkAndWait(page, 180000); 
                if (isEmpty) {
                    console.log(`   ⏩ Skipping Report 4 (No data found).`);
                    break;
                }
                
                console.log('   Exporting Report 4...');
                await page.evaluate(() => {
                    const xpathResult = document.evaluate('//*[@id="table"]/div[1]/button[3]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                    const btn = xpathResult.singleNodeValue;
                    if (btn) {
                        btn.click();
                    } else {
                        const allBtns = Array.from(document.querySelectorAll('button'));
                        const excelBtn = allBtns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
                        if (excelBtn) excelBtn.click();
                    }
                });
                
                file4 = await waitForDownloadAndRename(downloadPath, `Report4_HarshStart_Att${attempt}.xls`);
                
                if (validateVehicleData(file4, 1)) {
                    console.log(`   ✅ Report 4 is valid (Found vehicle data).`);
                    break;
                } else {
                    console.warn(`   ⚠️ Report 4 invalid or empty. Deleting...`);
                    if (fs.existsSync(file4)) fs.unlinkSync(file4);
                    file4 = '';
                }
            }
        } catch (error) {
            console.error('❌ Report 4 Failed:', error.message);
        }

        // REPORT 5: Forbidden
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { 
                for(var i=0; i<s.options.length; i++) { 
                    const txt = s.options[i].text;
                    if(txt.includes('พิ้น')) { 
                        s.value = s.options[i].value; 
                        s.dispatchEvent(new Event('change', { bubbles: true })); 
                        break; 
                    } 
                } 
            }
        }, startDateTime, endDateTime);
        await new Promise(r => setTimeout(r, 10000));
        await page.evaluate(() => {
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { for(var i=0; i<s.options.length; i++) { if(s.options[i].text.includes('สถานีทั้งหมด')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } } }
        });
        
        let file5 = '';
        for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            console.log(`   Searching Report 5 (Attempt ${attempt}/${MAX_RETRIES})...`);
            
            await prepareBeforeSearch(page);

            await page.evaluate(() => {
                const btn = document.querySelector('td:nth-of-type(7) > span');
                if (btn) btn.click();
            });
            
            const isEmpty = await checkAndWait(page, 180000); 
            if (isEmpty) {
                console.log(`   ⏩ Skipping Report 5 (No data found).`);
                break;
            }

            console.log('   Exporting Report 5...');
            try { await page.waitForSelector('#btnexport', { visible: true, timeout: 10000 }); } catch(e) {}
            await page.evaluate(() => {
                const btn = document.getElementById('btnexport');
                if (btn) btn.click();
            });
            
            file5 = await waitForDownloadAndRename(downloadPath, `Report5_ForbiddenParking_Att${attempt}.xls`);
            
            // ใน V3.4 Report 5 นำข้อมูลคอลัมน์ Index 2 ไปใช้งาน จึงตรวจจับที่ colIndex 2
            if (validateVehicleData(file5, 2)) {
                console.log(`   ✅ Report 5 is valid (Found vehicle data).`);
                break;
            } else {
                console.warn(`   ⚠️ Report 5 invalid or empty. Deleting...`);
                if (fs.existsSync(file5)) fs.unlinkSync(file5);
                file5 = '';
            }
        }

        // =================================================================
        // STEP 7: Generate PDF Summary
        // =================================================================
        console.log('📑 Step 7: Generating PDF Summary...');

        const FILES_CSV = {
            OVERSPEED: file1,
            IDLING: file2,
            SUDDEN_BRAKE: file3,
            HARSH_START: file4 !== '' ? file4 : '', 
            PROHIBITED: file5
        };

        // 1. Process Report 1
        const rawSpeed = processCSV_V3(FILES_CSV.OVERSPEED, { 
            colLicense: 1, 
            colStart: 2, 
            colEnd: 3, 
            useTimeCalc: true 
        });
        
        const speedStats = {};
        rawSpeed.forEach(r => {
            if (!speedStats[r.license]) speedStats[r.license] = { count: 0, time: 0, license: r.license };
            speedStats[r.license].count++;
            speedStats[r.license].time += r.durationSec;
        });
        const topSpeed = Object.values(speedStats).sort((a, b) => b.time - a.time).slice(0, 10);
        const totalOverSpeed = rawSpeed.length;

        // 2. Process Report 2
        const rawIdling = processCSV_V3(FILES_CSV.IDLING, { 
            colLicense: 1, 
            colStart: 2, 
            colEnd: 3, 
            useTimeCalc: true 
        });

        const idleStats = {};
        rawIdling.forEach(r => {
            if (!idleStats[r.license]) idleStats[r.license] = { count: 0, time: 0, license: r.license };
            idleStats[r.license].count++;
            idleStats[r.license].time += r.durationSec;
        });
        const topIdle = Object.values(idleStats).sort((a, b) => b.time - a.time).slice(0, 10);
        const maxIdleCar = topIdle.length > 0 ? topIdle[0] : { time: 0, license: '-' };

        // 3. Process Report 3 & 4 
        const rawBrake = fs.existsSync(FILES_CSV.SUDDEN_BRAKE || '') ? processCSV_V3(FILES_CSV.SUDDEN_BRAKE, {
            colLicense: 1,
            colDate: 3,
            colSpeedStart: 4,
            colSpeedEnd: 5
        }) : [];

        const rawStart = (FILES_CSV.HARSH_START && fs.existsSync(FILES_CSV.HARSH_START)) ? processCSV_V3(FILES_CSV.HARSH_START, {
            colLicense: 1,
            colDate: 3,
            colSpeedStart: 4,
            colSpeedEnd: 5
        }) : [];
        
        const criticalEvents = [
            ...rawBrake.map(r => ({ ...r, type: 'Sudden Brake', level: r.date })), 
            ...rawStart.map(r => ({ ...r, type: 'Harsh Start', level: r.date }))
        ];

        // 4. Process Report 5
        const rawForbidden = processCSV_V3(FILES_CSV.PROHIBITED, {
            colLicense: 2,
            colStation: 4,
            colStart: 5,  
            colEnd: 6,    
            useTimeCalc: true
        });

        const forbiddenList = rawForbidden
            .sort((a, b) => b.durationSec - a.durationSec)
            .slice(0, 10);
        
        // Chart Stats for Prohibited 
        const forbiddenChartStats = {};
        rawForbidden.forEach(r => {
            if(!forbiddenChartStats[r.license]) forbiddenChartStats[r.license] = 0;
            forbiddenChartStats[r.license] += r.durationSec;
        });
        const topForbiddenChart = Object.entries(forbiddenChartStats)
            .map(([license, time]) => ({ license, time }))
            .sort((a, b) => b.time - a.time).slice(0, 5);

        // --- HTML GENERATION ---
        const today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
        
        const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
            @page { size: A4; margin: 0; }
            body { font-family: 'Noto Sans Thai', sans-serif; margin: 0; padding: 0; background: #fff; color: #333; }
            .page { width: 210mm; height: 296mm; position: relative; page-break-after: always; overflow: hidden; }
            .content { padding: 40px; }
            .header-banner { background: #1E40AF; color: white; padding: 15px 40px; font-size: 24px; font-weight: bold; margin-bottom: 30px; }
            h1 { font-size: 32px; color: #1E40AF; margin-bottom: 10px; }
            .grid-2x2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 50px; }
            .card { background: #F8FAFC; border-radius: 12px; padding: 30px; text-align: center; border: 1px solid #E2E8F0; }
            .card-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }
            .card-value { font-size: 48px; font-weight: bold; margin: 10px 0; }
            .card-sub { font-size: 14px; color: #64748B; }
            .c-blue { color: #1E40AF; }
            .c-orange { color: #F59E0B; }
            .c-red { color: #DC2626; }
            .c-purple { color: #9333EA; }
            .chart-container { margin: 40px 0; }
            .bar-row { display: flex; align-items: center; margin-bottom: 15px; }
            .bar-label { width: 180px; text-align: right; padding-right: 15px; font-weight: 600; font-size: 14px; }
            .bar-track { flex-grow: 1; background: #F1F5F9; height: 30px; border-radius: 4px; overflow: hidden; }
            .bar-fill { height: 100%; display: flex; align-items: center; justify-content: flex-end; padding-right: 10px; color: white; font-size: 12px; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th { background: #1E40AF; color: white; padding: 12px; text-align: left; }
            td { padding: 10px; border-bottom: 1px solid #E2E8F0; }
            tr:nth-child(even) { background: #F8FAFC; }
            </style>
        </head>
        <body>

            <!-- Page 1: Executive Summary -->
            <div class="page">
            <div style="text-align: center; padding-top: 60px;">
                <h1 style="font-size: 48px;">รายงานสรุปพฤติกรรมการขับขี่</h1>
                <div style="font-size: 24px; color: #64748B;">Fleet Safety & Telematics Analysis Report</div>
                <div style="margin-top: 20px; font-size: 18px;">ประจำวันที่: ${today}</div>
            </div>

            <div class="content">
                <div class="header-banner" style="margin-top: 40px; text-align: center;">บทสรุปผู้บริหาร (Executive Summary)</div>
                <div class="grid-2x2">
                <div class="card">
                    <div class="card-title c-blue">Over Speed (ครั้ง)</div>
                    <div class="card-value c-blue">${totalOverSpeed}</div>
                    <div class="card-sub">เหตุการณ์ทั้งหมด</div>
                </div>
                <div class="card">
                    <div class="card-title c-orange">Max Idling (สูงสุด)</div>
                    <div class="card-value c-orange">${Math.round(maxIdleCar.time / 60)}m</div>
                    <div class="card-sub">${maxIdleCar.license}</div>
                </div>
                <div class="card">
                    <div class="card-title c-red">Critical Events</div>
                    <div class="card-value c-red">${criticalEvents.length}</div>
                    <div class="card-sub">เบรก/ออกตัว กระชาก</div>
                </div>
                <div class="card">
                    <div class="card-title c-purple">พื้นที่ห้ามจอด</div>
                    <div class="card-value c-purple">${rawForbidden.length}</div>
                    <div class="card-sub">จำนวนครั้งทั้งหมด</div>
                </div>
                </div>
            </div>
            </div>

            <!-- Page 2: Over Speed -->
            <div class="page">
            <div class="header-banner">1. การใช้ความเร็วเกินกำหนด (Over Speed Analysis)</div>
            <div class="content">
                <h3>Top 10 Over Speed by Duration</h3>
                <div class="chart-container">
                ${topSpeed.slice(0, 5).map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.time / (topSpeed[0]?.time || 1)) * 100}%; background: #1E40AF;">${formatSeconds(item.time)}</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา (Start-End)</th></tr>
                </thead>
                <tbody>
                    ${topSpeed.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.count}</td>
                        <td>${formatSeconds(item.time)}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 3: Idling -->
            <div class="page">
            <div class="header-banner">2. การจอดไม่ดับเครื่อง (Idling Analysis)</div>
            <div class="content">
                <h3>Top 10 Idling by Duration</h3>
                <div class="chart-container">
                ${topIdle.slice(0, 5).map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.time / (topIdle[0]?.time || 1)) * 100}%; background: #F59E0B;">${formatSeconds(item.time)}</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา (Start-End)</th></tr>
                </thead>
                <tbody>
                    ${topIdle.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.count}</td>
                        <td>${formatSeconds(item.time)}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 4: Critical Events -->
            <div class="page">
            <div class="header-banner">3. เหตุการณ์วิกฤต (Critical Safety Events)</div>
            <div class="content">
                <h3 style="color: #DC2626;">3.1 Sudden Brake (เบรกกะทันหัน)</h3>
                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>รายละเอียด</th><th>วันที่บันทึก</th></tr>
                </thead>
                <tbody>
                    ${criticalEvents.filter(x => x.type === 'Sudden Brake').map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>Speed: ${item.v_start} &#8594; ${item.v_end} km/h</td>
                        <td>${item.level}</td>
                    </tr>
                    `).join('')}
                    ${criticalEvents.filter(x => x.type === 'Sudden Brake').length === 0 ? '<tr><td colspan="4" style="text-align:center">ไม่มีข้อมูล</td></tr>' : ''}
                </tbody>
                </table>

                <br><br>
                <h3 style="color: #F59E0B;">3.2 Harsh Start (ออกตัวกระชาก)</h3>
                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>รายละเอียด</th><th>วันที่บันทึก</th></tr>
                </thead>
                <tbody>
                    ${criticalEvents.filter(x => x.type === 'Harsh Start').map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>Speed: ${item.v_start} &#8594; ${item.v_end} km/h</td>
                        <td>${item.level}</td>
                    </tr>
                    `).join('')}
                    ${criticalEvents.filter(x => x.type === 'Harsh Start').length === 0 ? '<tr><td colspan="4" style="text-align:center">ไม่มีข้อมูล</td></tr>' : ''}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 5: Prohibited Parking -->
            <div class="page">
            <div class="header-banner">4. รายงานพื้นที่ห้ามจอด (Prohibited Parking Area Report)</div>
            <div class="content">
                <h3>Top 5 Prohibited Area Duration</h3>
                <div class="chart-container">
                ${topForbiddenChart.map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.time / (topForbiddenChart[0]?.time || 1)) * 100}%; background: #9333EA;">${formatSeconds(item.time)}</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>ชื่อสถานี</th><th>รวมเวลา</th></tr>
                </thead>
                <tbody>
                    ${forbiddenList.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.station}</td>
                        <td>${item.durationStr}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

        </body>
        </html>
        `;

        await page.setContent(html, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true
        });
        console.log(`   ✅ PDF Generated: ${pdfPath}`);


        // =================================================================
        // STEP 8: Zip & Email
        // =================================================================
        console.log('📧 Step 8: Zipping CSVs & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const csvsToZip = allFiles.filter(f => f.startsWith('Converted_') && f.endsWith('.csv'));

        if (csvsToZip.length > 0 || fs.existsSync(pdfPath)) {
            const zipName = `DTC_Report_Data_${today.replace(/ /g, '_')}.zip`;
            const zipPath = path.join(downloadPath, zipName);
            
            if(csvsToZip.length > 0) {
                await zipFiles(downloadPath, zipPath, csvsToZip);
            }

            const attachments = [];
            if (fs.existsSync(zipPath)) attachments.push({ filename: zipName, path: zipPath });
            if (fs.existsSync(pdfPath)) attachments.push({ filename: 'Fleet_Safety_Analysis_Report.pdf', path: pdfPath });

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงานสรุปพฤติกรรมการขับขี่ (Fleet Safety Report) - ${today}`,
                text: `เรียน ผู้เกี่ยวข้อง\n\nระบบส่งรายงานประจำวันกะกลางวัน (06:00 - 18:00)\nช่วงเวลา: ${todayStr} 06:00 ถึง ${todayStr} 18:00\n\nสิ่งที่แนบมาด้วย:\n1. ไฟล์ข้อมูลดิบ CSV (อยู่ใน Zip)\n2. ไฟล์ PDF สรุปภาพรวม\n\nด้วยความนับถือ\nDTC Automation Bot`,
         		 attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully!`);
        } else {
            console.warn('⚠️ No files to send!');
        }

        console.log('🧹 Cleanup...');
        // fs.rmSync(downloadPath, { recursive: true, force: true });
        console.log('   ✅ Cleanup Complete.');

    } catch (err) {
        console.error('❌ Fatal Error:', err);
        await page.screenshot({ path: path.join(downloadPath, 'fatal_error.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
