/*
 * 依存ゼロの S3 互換クライアント (Cloudflare R2 用)。
 * Node 標準の https / crypto だけで AWS Signature V4 を生成して PutObject / DeleteObject する。
 *   - パススタイル URL: https://<host>/<bucket>/<key>
 *   - service = "s3", region = R2 は "auto"
 */
const https = require('https');
const crypto = require('crypto');

function sha256hex(data) {
    return crypto.createHash('sha256').update(data).digest('hex');
}
function hmac(key, data) {
    return crypto.createHmac('sha256', key).update(data, 'utf8').digest();
}
// 各パスセグメントを RFC3986 でエンコード ('/' は保持)
function encodePath(p) {
    return p.split('/').map(s => encodeURIComponent(s)).join('/');
}

/**
 * SigV4 署名を計算する (純粋関数。テスト可能)。
 * @returns {{authorization:string, signature:string, signedHeaders:string}}
 */
function computeAuth(o) {
    const sortedNames = Object.keys(o.signHeaders).sort();
    const canonicalHeaders = sortedNames.map(n => n + ':' + o.signHeaders[n] + '\n').join('');
    const signedHeaders = sortedNames.join(';');
    const canonicalRequest = [
        o.method, o.canonicalUri, o.query || '', canonicalHeaders, signedHeaders, o.payloadHash
    ].join('\n');
    const scope = `${o.datestamp}/${o.region}/${o.service}/aws4_request`;
    const stringToSign = [
        'AWS4-HMAC-SHA256', o.amzdate, scope, sha256hex(canonicalRequest)
    ].join('\n');
    const kDate = hmac('AWS4' + o.secretAccessKey, o.datestamp);
    const kRegion = hmac(kDate, o.region);
    const kService = hmac(kRegion, o.service);
    const kSigning = hmac(kService, 'aws4_request');
    const signature = crypto.createHmac('sha256', kSigning).update(stringToSign, 'utf8').digest('hex');
    const authorization =
        `AWS4-HMAC-SHA256 Credential=${o.accessKeyId}/${scope}, ` +
        `SignedHeaders=${signedHeaders}, Signature=${signature}`;
    return { authorization, signature, signedHeaders };
}

function amzDates(d) {
    const iso = (d || new Date()).toISOString();          // 2025-06-28T12:34:56.789Z
    const amzdate = iso.replace(/[-:]/g, '').replace(/\.\d+/, ''); // 20250628T123456Z
    return { amzdate, datestamp: amzdate.slice(0, 8) };
}

function s3Request(method, opt, key, body, extra) {
    extra = extra || {};
    return new Promise((resolve, reject) => {
        try {
            const url = new URL(opt.endpoint);
            const host = url.host;
            let buf = body || Buffer.alloc(0);
            if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
            const payloadHash = sha256hex(buf);
            const canonicalUri = encodePath('/' + opt.bucket + '/' + key);
            const region = opt.region || 'auto';
            const service = 's3';
            const { amzdate, datestamp } = amzDates();

            const signHeaders = {
                host,
                'x-amz-content-sha256': payloadHash,
                'x-amz-date': amzdate
            };
            if (extra.contentType) signHeaders['content-type'] = extra.contentType;

            const { authorization } = computeAuth({
                method, canonicalUri, query: '', signHeaders, payloadHash,
                region, service, amzdate, datestamp,
                accessKeyId: opt.accessKeyId, secretAccessKey: opt.secretAccessKey
            });

            const headers = {
                Host: host,
                'X-Amz-Date': amzdate,
                'X-Amz-Content-Sha256': payloadHash,
                Authorization: authorization,
                'Content-Length': buf.length
            };
            if (extra.contentType) headers['Content-Type'] = extra.contentType;
            if (extra.cacheControl) headers['Cache-Control'] = extra.cacheControl;

            const req = https.request({
                host, port: url.port || 443, method, path: canonicalUri, headers
            }, (res) => {
                let data = '';
                res.on('data', c => data += c);
                res.on('end', () => {
                    if (res.statusCode >= 200 && res.statusCode < 300) resolve({ status: res.statusCode });
                    else reject(new Error(`S3 ${method} ${key} -> HTTP ${res.statusCode}: ${String(data).slice(0, 300)}`));
                });
            });
            req.on('error', reject);
            if (buf.length) req.write(buf);
            req.end();
        } catch (e) { reject(e); }
    });
}

function putObject(opt, key, body, options) {
    options = options || {};
    return s3Request('PUT', opt, key, body, {
        contentType: options.contentType, cacheControl: options.cacheControl
    });
}
function deleteObject(opt, key) {
    return s3Request('DELETE', opt, key, Buffer.alloc(0), {});
}

module.exports = { putObject, deleteObject, computeAuth, sha256hex, amzDates };
