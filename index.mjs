import fs from 'fs-extra';
import path from 'path';
import xml2js from 'xml2js';
import JSZip from 'jszip';
import zlib from 'zlib';

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
    // resultが存在しない、または予期した構造でない場合のチェック
    if (!result || !result.feedback) {
        console.error('Invalid XML structure:', JSON.stringify(result, null, 2).substring(0, 500) + '...');
        return `<section><h2 class="header">Error: Invalid XML structure</h2><p>XMLデータが有効なDMARCレポート形式ではありません。</p></section>`;
    }

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
                <th>DKIM Domain</th>
                <th>DKIM Results</th>
                <th>SPF Domain</th>
                <th>SPF Results</th>
            </tr>
        </thead>`;
    sectionHtml += `<tbody>`;

    // recordが存在しない場合のチェック
    if (!result.feedback?.record) {
        sectionHtml += `<tr><td colspan="10" class="none">レコードデータがありません</td></tr>`;
    } else {
        result.feedback.record.forEach(record => {
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
                    <td class="${spfClass}"><span class="${spfClass}">${spf}</span></td>
                    <td class="${dkimDomainClass}">${dkimDomein}</td>
                    <td class="${dkimResultClass}"><span class="${dkimResultClass}">${dkimResult}</span></td>
                    <td class="${spfDomainClass}">${spfDomein}</td>
                    <td class="${spfResultClass}"><span class="${spfResultClass}">${spfResult}</span></td>
                </tr>`;
        });
    }

    sectionHtml += `</tbody></table></section>`;
    return sectionHtml;
};

const processFile = async (filePath) => {
    try {
        let fileData;
        if (path.extname(filePath) === '.zip') {
            // .zip ファイルの処理
            try {
                const zipData = await fs.readFile(filePath);
                const zip = await JSZip.loadAsync(zipData);
                const xmlFiles = Object.keys(zip.files).filter(name =>
                    name.endsWith('.xml') || (!path.extname(name) && !zip.files[name].dir)
                );

                if (xmlFiles.length === 0) {
                    console.error(`No XML file found in ${filePath}`);
                    return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>ZIPファイル内にXMLファイルが見つかりませんでした。</p></section>`;
                }

                fileData = await zip.files[xmlFiles[0]].async('string');
            } catch (error) {
                console.error(`Error processing ZIP file ${filePath}:`, error);
                return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>ZIPファイルの処理中にエラーが発生しました: ${error.message}</p></section>`;
            }
        } else if (path.extname(filePath) === '.xml') {
            // .xml ファイルの処理
            try {
                fileData = await fs.readFile(filePath, 'utf8');
            } catch (error) {
                console.error(`Error reading XML file ${filePath}:`, error);
                return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>XMLファイルの読み込み中にエラーが発生しました: ${error.message}</p></section>`;
            }
        } else if (path.extname(filePath) === '.gz') {
            // .gz ファイルの処理
            try {
                const gzData = await fs.readFile(filePath);
                const xmlData = await zlib.gunzipSync(gzData);
                fileData = xmlData.toString();
            } catch (error) {
                console.error(`Error extracting GZ file ${filePath}:`, error);
                return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>GZファイルの展開中にエラーが発生しました: ${error.message}</p></section>`;
            }
        } else {
            console.warn(`Skipping unsupported file: ${filePath}`);
            return;
        }

        // XML解析前にデータが有効かチェック
        if (!fileData || typeof fileData !== 'string' || fileData.trim() === '') {
            console.error(`Empty or invalid file data in ${filePath}`);
            return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>ファイルが空または不正なデータです。</p></section>`;
        }

        try {
            const result = await processXML(fileData);
            // XMLが有効なDMARCレポート形式かチェック
            if (!result || !result.feedback) {
                console.error(`Invalid DMARC report format in ${filePath}`);
                // XMLの内容を一部出力してデバッグに役立てる
                console.log('XML structure:', JSON.stringify(result, null, 2).substring(0, 1000) + '...');
                return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>XMLデータが有効なDMARCレポート形式ではありません。</p></section>`;
            }
            return extractDataAndGenerateHTML(result);
        } catch (error) {
            console.error(`Error parsing XML data from ${filePath}:`, error);
            return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>XMLデータの解析中にエラーが発生しました: ${error.message}</p></section>`;
        }
    } catch (error) {
        console.error(`Unexpected error processing ${filePath}:`, error);
        return `<section><h2 class="header">Error: ${path.basename(filePath)}</h2><p>ファイル処理中に予期しないエラーが発生しました: ${error.message}</p></section>`;
    }
};

const unzipAndConvertToHtml = async () => {
    try {
        // ディレクトリが存在するか確認
        if (!(await fs.pathExists(reportsDir))) {
            console.error(`Directory not found: ${reportsDir}`);
            await fs.writeFile(outputHtml, `<!DOCTYPE html><html><body><h1>Error</h1><p>レポートディレクトリが見つかりません: ${reportsDir}</p></body></html>`);
            return;
        }

        let files = await fs.readdir(reportsDir);
        if (files.length === 0) {
            console.warn(`No files found in ${reportsDir}`);
            await fs.writeFile(outputHtml, `<!DOCTYPE html><html><body><h1>No Reports</h1><p>レポートファイルが見つかりません。</p></body></html>`);
            return;
        }

        // ファイルのメタデータを取得してソート
        let fileStats = await Promise.all(files.map(async file => {
            const filePath = path.join(reportsDir, file);
            const stats = await fs.stat(filePath);
            return { file, stats };
        }));

        // ファイルの作成日時が新しい順にソート
        fileStats.sort((a, b) => b.stats.mtime - a.stats.mtime);

        let htmlContent = `<!DOCTYPE html>
        <html lang="ja">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
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

        // 処理されたファイルの数を追跡
        let processedCount = 0;

        for (const { file } of fileStats) {
            const filePath = path.join(reportsDir, file);
            try {
                const sectionHtml = await processFile(filePath);
                if (sectionHtml) {
                    htmlContent += sectionHtml;
                    processedCount++;
                }
            } catch (error) {
                console.error(`Error processing file ${filePath}:`, error);
                htmlContent += `<section><h2 class="header">Error: ${file}</h2><p>ファイル処理中にエラーが発生しました: ${error.message}</p></section>`;
            }
        }

        if (processedCount === 0) {
            htmlContent += `<section><h2 class="header">No valid reports found</h2><p>有効なDMARCレポートファイルが見つかりませんでした。</p></section>`;
        }

        htmlContent += `</main>
                <footer>
                    <address><a href="https://github.com/burnworks/dmarc-to-html-converter" target="_blank">@burnworks/dmarc-to-html-converter</a></address>
                </footer>
            </body>
        </html>`;
        await fs.writeFile(outputHtml, htmlContent);
        console.log('\x1b[32m%s\x1b[0m', 'HTML report has been created.');
    } catch (error) {
        console.error('Unexpected error:', error);
        await fs.writeFile(outputHtml, `<!DOCTYPE html><html><body><h1>Error</h1><p>予期しないエラーが発生しました: ${error.message}</p></body></html>`);
    }
};

unzipAndConvertToHtml().catch(error => {
    console.error('Fatal error:', error);
    fs.writeFileSync(outputHtml, `<!DOCTYPE html><html><body><h1>Fatal Error</h1><p>${error.message}</p></body></html>`);
});