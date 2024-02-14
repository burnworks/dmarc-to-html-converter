import fs from 'fs-extra';
import path from 'path';
import xml2js from 'xml2js';
import JSZip from 'jszip';

// レポート（.zip or .xml）を格納するディレクトリ
const reportsDir = './report';
// レポートの HTML
const outputHtml = './report.html';

// UNIX タイムスタンプを日本標準時の日付けに変換
const convertTimestampToJST = (timestamp) => {
    const date = new Date(timestamp * 1000);
    date.setHours(date.getHours() + 9);
    return date.toISOString().replace(/T/, ' ').replace(/\..+/, '');
};

const processXML = async (xmlData) => {
    const parser = new xml2js.Parser();
    return await parser.parseStringPromise(xmlData);
};

// レポート HTML を生成
const extractDataAndGenerateHTML = (result) => {
    let sectionHtml = `<section>`;
    const reportId = result.feedback?.report_metadata?.[0]?.report_id?.[0] ?? 'N/A';
    const beginDate = convertTimestampToJST(result.feedback?.report_metadata?.[0]?.date_range?.[0]?.begin?.[0] ?? '');
    const endDate = convertTimestampToJST(result.feedback?.report_metadata?.[0]?.date_range?.[0]?.end?.[0] ?? '');

    sectionHtml += `<h2 class="header">ID: ${reportId}</h2>`;
    sectionHtml += `<p class="date">${beginDate} ～ ${endDate}</p>`;
    sectionHtml += `<table class="result">`;
    sectionHtml += `<thead>
            <tr>
                <th>IP</th>
                <th>From</th>
                <th>Count</th>
                <th>Disposition</th>
                <th>DKIM</th>
                <th>SPF</th>
                <th>DKIM Domein</th>
                <th>DKIM Results</th>
                <th>SPF Domein</th>
                <th>SPF Results</th>
            </tr>
        </thead>`;
    sectionHtml += `<tbody>`;

    result.feedback?.record?.forEach(record => {
        const sourceIp = record.row?.[0]?.source_ip?.[0] ?? 'N/A';
        const count = record.row?.[0]?.count?.[0] ?? 'N/A';
        const disposition = record.row?.[0]?.policy_evaluated?.[0]?.disposition?.[0] ?? 'N/A';
        const dkim = record.row?.[0]?.policy_evaluated?.[0]?.dkim?.[0] ?? 'N/A';
        const spf = record.row?.[0]?.policy_evaluated?.[0]?.spf?.[0] ?? 'N/A';
        const from = record.identifiers?.[0]?.header_from?.[0] ?? 'N/A';
        const dkimDomein = record.auth_results?.[0]?.dkim?.[0]?.domain?.[0] ?? 'N/A';
        const dkimResult = record.auth_results?.[0]?.dkim?.[0]?.result?.[0] ?? 'N/A';
        const spfDomein = record.auth_results?.[0]?.spf?.[0]?.domain?.[0] ?? 'N/A';
        const spfResult = record.auth_results?.[0]?.spf?.[0]?.result?.[0] ?? 'N/A';

        // class 名の割り当て
        const sourceIpClass = sourceIp === 'N/A' ? 'none' : 'value';
        const countClass = count === 'N/A' ? 'none' : 'value';
        const dispositionClass = disposition === 'N/A' ? 'none' : 'value';
        const dkimDomainClass = dkimDomein === 'N/A' ? 'none' : 'value';
        const spfDomainClass = spfDomein === 'N/A' ? 'none' : 'value';
        const fromClass = from === 'N/A' ? 'none' : 'value';

        const dkimClass = dkim === 'N/A' ? 'none' : dkim === 'fail' ? 'fail' : 'pass';
        const spfClass = spf === 'N/A' ? 'none' : spf === 'fail' ? 'fail' : 'pass';
        const dkimResultClass = dkimResult === 'N/A' ? 'none' : dkimResult === 'fail' ? 'fail' : dkimResult === 'softfail' ? 'softfail' : 'pass';
        const spfResultClass = spfResult === 'N/A' ? 'none' : spfResult === 'fail' ? 'fail' : spfResult === 'softfail' ? 'softfail' : 'pass';

        sectionHtml += `<tr>
                <td class="${sourceIpClass}">${sourceIp}</td>
                <td class="${fromClass}">${from}</td>
                <td class="${countClass}">${count}</td>
                <td class="${dispositionClass}">${disposition}</td>
                <td class="${dkimClass}"><span class="${dkimClass}">${dkim}</span></td>
                <td class="${dkimClass}"><span class="${spfClass}">${spf}</span></td>
                <td class="${dkimDomainClass}">${dkimDomein}</td>
                <td class="${dkimResultClass}"><span class="${dkimResultClass}">${dkimResult}</span></td>
                <td class="${spfDomainClass}">${spfDomein}</td>
                <td class="${spfResultClass}"><span class="${spfResultClass}">${spfResult}</span></td>
            </tr>`;
    });

    sectionHtml += `</tbody></table></section>`;
    return sectionHtml;
};

const processFile = async (filePath) => {
    let fileData;
    if (path.extname(filePath) === '.zip') {
        const zipData = await fs.readFile(filePath);
        const zip = await JSZip.loadAsync(zipData);
        const xmlFileName = Object.keys(zip.files)[0];
        fileData = await zip.files[xmlFileName].async('string');
    } else if (path.extname(filePath) === '.xml') {
        fileData = await fs.readFile(filePath, 'utf8');
    } else {
        return;
    }

    const result = await processXML(fileData);
    return extractDataAndGenerateHTML(result);
};

const unzipAndConvertToHtml = async () => {
    const files = await fs.readdir(reportsDir);
    let htmlContent = `<!DOCTYPE html>
    <html lang="ja">
        <head>
            <title>DMARC レポート</title>
            <style>
                body {
                    margin: 0;
                    padding: 0;
                    font-size: 1rem;
                    color: #111827;
                    line-height: 1.6;
                }
                header {
                    margin: 0;
                    padding: 1rem;
                    background-color: #f9fafb;
                }
                h1 {
                    font-size: 1.5rem;
                    margin: 0;
                    padding: 0;
                }
                main {
                    padding: 1rem;
                    margin-top: 3rem;
                }
                footer {
                    padding: 1rem;
                    margin-top: 3rem;
                    background-color: #f9fafb;
                }
                address a {
                    color: #111827;
                    text-decoration: underline;
                }
                section + section {
                    margin-top: 4rem;
                }
                .header {
                    font-size: 1.25rem;
                    margin: 0;
                    padding: 0;
                }
                .date {
                    margin-top: 0.5rem;
                    font-size: 1rem;
                }
                .result {
                    margin-top: 0.5rem;
                    border: 1px solid #d1d5db;
                    border-collapse: collapse;
                    font-size: 0.875rem;
                }
                tbody tr:nth-child(even) {
                    background-color: #f9fafb;
                }
                th {
                    padding: 1rem;
                    border: 1px solid #d1d5db;
                    text-align: center;
                    font-weight: bold;
                }
                td {
                    padding: 1rem;
                    border: 1px solid #d1d5db;
                    text-align: center;
                }
                td span:not(.none) {
                    display: inline-block;
                    padding: 0.125rem 1rem;
                    font-size: 0.75rem;
                    border-radius: 0.25rem;
                    font-weight: bold;
                }
                span.fail {
                    background-color: #dc2626;
                    color: white;
                }
                span.softfail {
                    background-color: #fcd34d;
                }
                span.pass {
                    background-color: #15803d;
                    color: white;
                }
                td.none {
                    color: #6b7280;
                }
            </style>
        </head>
        <body>
            <header>
                <h1>DMARC レポート</h1>
            </header>
            <main>`;

    for (const file of files) {
        const filePath = path.join(reportsDir, file);
        const sectionHtml = await processFile(filePath);
        htmlContent += sectionHtml;
    }

    htmlContent += `</main>
            <footer>
                <address><a href="https://github.com/burnworks/dmarc-to-html-converter" target="_blank">@burnworks/dmarc-to-html-converter</a></address>
            </footer>
        </body>
    </html>`;
    await fs.writeFile(outputHtml, htmlContent);
    console.log('\x1b[32m%s\x1b[0m', 'HTML report has been created.');
};

unzipAndConvertToHtml().catch(console.error);
