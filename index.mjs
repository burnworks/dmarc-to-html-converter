import fs from 'fs-extra';
import path from 'path';
import xml2js from 'xml2js';
import JSZip from 'jszip';
import zlib from 'zlib';
import stream from 'stream';
import { promisify } from 'util';

// レポート（.zip or .xml）を格納するディレクトリ
const reportsDir = './report';
// レポートの HTML
const outputHtml = './report.html';
// パイプラインの非同期化
const pipeline = promisify(stream.pipeline);

// UNIX タイムスタンプを日本標準時の日付けに変換
const convertTimestampToJST = (timestamp) => {
    if (!timestamp || isNaN(Number(timestamp))) {
        return 'N/A';
    }
    try {
        const date = new Date(Number(timestamp) * 1000);
        date.setHours(date.getHours() + 9);
        return date.toISOString().replace(/T/, ' ').replace(/\..+/, '');
    } catch (error) {
        console.error('Error converting timestamp:', error);
        return 'N/A';
    }
};

// XMLデータを検証し、必要な構造が存在するかチェック
const validateDmarcXml = (result) => {
    if (!result) {
        return { valid: false, error: 'XMLデータが空です' };
    }

    if (!result.feedback) {
        return { valid: false, error: 'XMLデータに「feedback」要素がありません' };
    }

    if (!result.feedback.report_metadata || !result.feedback.report_metadata[0]) {
        return { valid: false, error: '「report_metadata」要素が見つかりません' };
    }

    return { valid: true };
};

// XMLを解析
const processXML = async (xmlData) => {
    try {
        const parser = new xml2js.Parser();
        return await parser.parseStringPromise(xmlData);
    } catch (error) {
        throw new Error(`XML解析エラー: ${error.message}`);
    }
};

// レポートメタデータを抽出
const extractReportMetadata = (feedback) => {
    const reportId = feedback?.report_metadata?.[0]?.report_id?.[0] ?? 'N/A';
    const beginDate = convertTimestampToJST(feedback?.report_metadata?.[0]?.date_range?.[0]?.begin?.[0]);
    const endDate = convertTimestampToJST(feedback?.report_metadata?.[0]?.date_range?.[0]?.end?.[0]);

    return { reportId, beginDate, endDate };
};

// レコードデータを抽出して処理
const processRecordData = (record) => {
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

    return {
        data: { sourceIp, count, disposition, dkim, spf, from, dkimDomein, dkimResult, spfDomein, spfResult },
        classes: {
            sourceIpClass, countClass, dispositionClass, dkimDomainClass, spfDomainClass, fromClass,
            dkimClass, spfClass, dkimResultClass, spfResultClass
        }
    };
};

// レコード行のHTMLを生成
const generateRecordRowHtml = (record) => {
    const { data, classes } = processRecordData(record);

    return `<tr>
        <td class="${classes.sourceIpClass}">${data.sourceIp}</td>
        <td class="${classes.fromClass}">${data.from}</td>
        <td class="${classes.countClass}">${data.count}</td>
        <td class="${classes.dispositionClass}">${data.disposition}</td>
        <td class="${classes.dkimClass}"><span class="${classes.dkimClass}">${data.dkim}</span></td>
        <td class="${classes.spfClass}"><span class="${classes.spfClass}">${data.spf}</span></td>
        <td class="${classes.dkimDomainClass}">${data.dkimDomein}</td>
        <td class="${classes.dkimResultClass}"><span class="${classes.dkimResultClass}">${data.dkimResult}</span></td>
        <td class="${classes.spfDomainClass}">${data.spfDomein}</td>
        <td class="${classes.spfResultClass}"><span class="${classes.spfResultClass}">${data.spfResult}</span></td>
    </tr>`;
};

// レポート HTML を生成
const generateReportHtml = (result) => {
    // XMLデータを検証
    const validation = validateDmarcXml(result);
    if (!validation.valid) {
        return `<section><h2 class="header">Error: Invalid XML structure</h2><p>${validation.error}</p></section>`;
    }

    // メタデータを抽出
    const { reportId, beginDate, endDate } = extractReportMetadata(result.feedback);

    let sectionHtml = `<section>`;
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
    if (!result.feedback?.record || result.feedback.record.length === 0) {
        sectionHtml += `<tr><td colspan="10" class="none">レコードデータがありません</td></tr>`;
    } else {
        // 各レコードを処理
        result.feedback.record.forEach(record => {
            sectionHtml += generateRecordRowHtml(record);
        });
    }

    sectionHtml += `</tbody></table></section>`;
    return sectionHtml;
};

// エラーセクションのHTMLを生成
const generateErrorSectionHtml = (filename, errorMessage) => {
    return `<section><h2 class="header">Error: ${filename}</h2><p>${errorMessage}</p></section>`;
};

// XMLデータを読み込み（ストリーム対応）
const readXmlFromFile = async (filePath) => {
    try {
        // ファイルサイズを確認
        const stats = await fs.stat(filePath);
        const isLargeFile = stats.size > 10 * 1024 * 1024; // 10MB以上の場合

        if (path.extname(filePath) === '.zip') {
            return await readXmlFromZip(filePath, isLargeFile);
        } else if (path.extname(filePath) === '.xml') {
            return isLargeFile
                ? await readLargeXmlFile(filePath)
                : await fs.readFile(filePath, 'utf8');
        } else if (path.extname(filePath) === '.gz') {
            return await readXmlFromGzip(filePath, isLargeFile);
        } else {
            throw new Error(`サポートされていないファイル形式: ${path.extname(filePath)}`);
        }
    } catch (error) {
        throw new Error(`ファイル読み込みエラー: ${error.message}`);
    }
};

// 大きなXMLファイルを読み込み
const readLargeXmlFile = async (filePath) => {
    let content = '';
    const readStream = fs.createReadStream(filePath, { encoding: 'utf8' });

    return new Promise((resolve, reject) => {
        readStream.on('data', chunk => {
            content += chunk;
        });

        readStream.on('end', () => {
            resolve(content);
        });

        readStream.on('error', error => {
            reject(new Error(`XMLファイル読み込みストリームエラー: ${error.message}`));
        });
    });
};

// ZIPファイルからXMLを読み込み
const readXmlFromZip = async (filePath, isLargeFile) => {
    const zipData = isLargeFile
        ? await readLargeFile(filePath)
        : await fs.readFile(filePath);

    const zip = await JSZip.loadAsync(zipData);
    const xmlFiles = Object.keys(zip.files).filter(name =>
        name.endsWith('.xml') || (!path.extname(name) && !zip.files[name].dir)
    );

    if (xmlFiles.length === 0) {
        throw new Error('ZIPファイル内にXMLファイルが見つかりませんでした');
    }

    return await zip.files[xmlFiles[0]].async('string');
};

// 大きなファイルを読み込み
const readLargeFile = async (filePath) => {
    const chunks = [];
    const readStream = fs.createReadStream(filePath);

    return new Promise((resolve, reject) => {
        readStream.on('data', chunk => {
            chunks.push(chunk);
        });

        readStream.on('end', () => {
            resolve(Buffer.concat(chunks));
        });

        readStream.on('error', error => {
            reject(new Error(`ファイル読み込みストリームエラー: ${error.message}`));
        });
    });
};

// GZIPファイルからXMLを読み込み
const readXmlFromGzip = async (filePath, isLargeFile) => {
    if (isLargeFile) {
        const readStream = fs.createReadStream(filePath);
        const gunzipStream = zlib.createGunzip();
        const chunks = [];

        return new Promise((resolve, reject) => {
            pipeline(readStream, gunzipStream)
                .then(() => {
                    gunzipStream.on('data', chunk => {
                        chunks.push(chunk);
                    });

                    gunzipStream.on('end', () => {
                        resolve(Buffer.concat(chunks).toString());
                    });
                })
                .catch(error => {
                    reject(new Error(`GZIPファイル解凍ストリームエラー: ${error.message}`));
                });
        });
    } else {
        const gzData = await fs.readFile(filePath);
        const xmlData = zlib.gunzipSync(gzData);
        return xmlData.toString();
    }
};

// 単一のファイルを処理
const processFile = async (filePath) => {
    const filename = path.basename(filePath);

    try {
        // XMLデータの読み込み
        const xmlData = await readXmlFromFile(filePath);

        // XML解析前にデータが有効かチェック
        if (!xmlData || typeof xmlData !== 'string' || xmlData.trim() === '') {
            return generateErrorSectionHtml(filename, 'ファイルが空または不正なデータです');
        }

        // XMLデータの解析
        const result = await processXML(xmlData);

        // DMARCレポート形式の検証
        const validation = validateDmarcXml(result);
        if (!validation.valid) {
            console.error(`Invalid DMARC report format in ${filename}: ${validation.error}`);
            // XMLの内容を一部出力してデバッグに役立てる
            console.log('XML structure preview:', JSON.stringify(result, null, 2).substring(0, 500) + '...');
            return generateErrorSectionHtml(filename, `XMLデータが有効なDMARCレポート形式ではありません: ${validation.error}`);
        }

        // HTML生成
        return generateReportHtml(result);
    } catch (error) {
        console.error(`Error processing ${filename}:`, error);
        return generateErrorSectionHtml(filename, `処理中にエラーが発生しました: ${error.message}`);
    }
};

// HTMLのヘッダー部分を生成
const generateHtmlHeader = () => {
    return `<!DOCTYPE html>
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
};

// HTMLのフッター部分を生成
const generateHtmlFooter = () => {
    return `</main>
            <footer>
                <address><a href="https://github.com/burnworks/dmarc-to-html-converter" target="_blank">@burnworks/dmarc-to-html-converter</a></address>
            </footer>
        </body>
    </html>`;
};

// エラーHTMLを生成
const generateErrorHtml = (errorMessage) => {
    return `<!DOCTYPE html>
    <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>DMARC レポート - エラー</title>
            <style>
                body { font-family: sans-serif; margin: 20px; }
                h1 { color: #dc2626; }
            </style>
        </head>
        <body>
            <h1>エラー</h1>
            <p>${errorMessage}</p>
        </body>
    </html>`;
};

// ファイルリストを取得して処理
const getFileListAndSort = async () => {
    // ディレクトリが存在するか確認
    if (!(await fs.pathExists(reportsDir))) {
        throw new Error(`ディレクトリが見つかりません: ${reportsDir}`);
    }

    let files = await fs.readdir(reportsDir);
    if (files.length === 0) {
        throw new Error(`ディレクトリ内にファイルが見つかりません: ${reportsDir}`);
    }

    // ファイルのメタデータを取得してソート
    let fileStats = await Promise.all(files.map(async file => {
        const filePath = path.join(reportsDir, file);
        const stats = await fs.stat(filePath);
        return { file, stats };
    }));

    // ファイルの作成日時が新しい順にソート
    return fileStats.sort((a, b) => b.stats.mtime - a.stats.mtime);
};

// ファイルをバッチに分割して処理
const processBatch = async (fileStats, batchSize = 5) => {
    const results = [];

    for (let i = 0; i < fileStats.length; i += batchSize) {
        const batch = fileStats.slice(i, i + batchSize);
        console.log(`Processing batch ${Math.floor(i / batchSize) + 1}/${Math.ceil(fileStats.length / batchSize)}`);

        const batchResults = await Promise.all(batch.map(async ({ file }) => {
            const filePath = path.join(reportsDir, file);
            return await processFile(filePath);
        }));

        results.push(...batchResults.filter(Boolean));
    }

    return results;
};

// メインプロセス - レポートの生成
const unzipAndConvertToHtml = async () => {
    try {
        // ファイルリストの取得とソート
        const fileStats = await getFileListAndSort();

        // HTMLヘッダーの生成
        let htmlContent = generateHtmlHeader();

        // ファイルの処理（バッチ処理）
        const sections = await processBatch(fileStats);

        // 処理結果の確認
        if (sections.length === 0) {
            htmlContent += `<section><h2 class="header">No valid reports found</h2><p>有効なDMARCレポートファイルが見つかりませんでした。</p></section>`;
        } else {
            htmlContent += sections.join('');
        }

        // HTMLフッターの追加
        htmlContent += generateHtmlFooter();

        // ファイルへの出力
        await fs.writeFile(outputHtml, htmlContent);
        console.log('\x1b[32m%s\x1b[0m', 'HTML report has been created.');
    } catch (error) {
        console.error('Error creating report:', error);

        // エラーHTMLの生成と出力
        const errorHtml = generateErrorHtml(`レポート生成中にエラーが発生しました: ${error.message}`);
        await fs.writeFile(outputHtml, errorHtml);
    }
};

// プログラム実行
unzipAndConvertToHtml().catch(error => {
    console.error('Fatal error:', error);

    try {
        const errorHtml = generateErrorHtml(`致命的なエラーが発生しました: ${error.message}`);
        fs.writeFileSync(outputHtml, errorHtml);
    } catch (writeError) {
        console.error('Error writing error HTML:', writeError);
    }
});