'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');

const app = express();
app.use(express.json({ limit: '20mb' }));

app.use((req, res, next) => {
	try {
		res.locals.catalyst = catalystSDK.initialize(req);
		next();
	} catch (err) {
		res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
	}
});

app.get('/', (req, res) => {
	res.status(200).json({ status: 'success', message: 'sdkpdf_function ready' });
});

function collectStream(stream) {
	return new Promise((resolve, reject) => {
		if (stream == null) {
			reject(new Error('Empty SmartBrowz response'));
			return;
		}
		if (Buffer.isBuffer(stream)) {
			resolve(stream);
			return;
		}
		if (typeof stream === 'string') {
			resolve(Buffer.from(stream));
			return;
		}
		if (stream instanceof Uint8Array) {
			resolve(Buffer.from(stream));
			return;
		}
		if (typeof stream.on !== 'function') {
			reject(new Error('Unexpected SmartBrowz response type'));
			return;
		}
		const chunks = [];
		stream.on('data', (chunk) => {
			chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
		});
		stream.on('end', () => resolve(Buffer.concat(chunks)));
		stream.on('error', reject);
	});
}

function sendPdf(res, buffer, fileName) {
	const safeName = String(fileName || 'document.pdf')
		.replace(/[^\w.\-]+/g, '_')
		.replace(/^_+|_+$/g, '') || 'document.pdf';
	const withExt = /\.pdf$/i.test(safeName) ? safeName : `${safeName}.pdf`;
	res.writeHead(200, {
		'Content-Type': 'application/pdf',
		'Content-Disposition': `attachment; filename="${withExt}"`,
		'Content-Length': buffer.length
	});
	res.end(buffer);
}

app.post('/convert', async (req, res) => {
	try {
		const catalyst = res.locals.catalyst;
		const smartbrowz = catalyst.smartbrowz();
		const body = req.body && typeof req.body === 'object' ? req.body : {};
		const html = String(body.html || '').trim();
		const url = String(body.url || '').trim();
		const templateId = body.template_id || body.templateId || '';
		const templateData = body.template_data || body.templateData || {};
		const pdf_options = body.pdf_options && typeof body.pdf_options === 'object' ? body.pdf_options : {};
		const page_options = body.page_options && typeof body.page_options === 'object' ? body.page_options : {};
		const fileName = body.fileName || body.file_name || 'document.pdf';

		let data;
		if (templateId) {
			data = await smartbrowz.generateFromTemplate(String(templateId), {
				template_data: templateData,
				output_options: { output_type: 'pdf' }
			});
		} else {
			const source = html || url;
			if (!source) {
				res.status(400).json({
					status: 'failure',
					message: 'html, url, or template_id is required'
				});
				return;
			}
			data = await smartbrowz.convertToPdf(source, {
				pdf_options,
				page_options
			});
		}

		const buffer = await collectStream(data);
		if (!buffer || buffer.length < 32) {
			res.status(502).json({ status: 'failure', message: 'SmartBrowz returned an empty PDF' });
			return;
		}
		sendPdf(res, buffer, fileName);
	} catch (err) {
		console.error('sdkpdf convert error:', err);
		if (!res.headersSent) {
			res.status(500).json({
				status: 'failure',
				message: err.message || 'PDF conversion failed'
			});
		}
	}
});

app.post('/screenshot', async (req, res) => {
	try {
		const catalyst = res.locals.catalyst;
		const smartbrowz = catalyst.smartbrowz();
		const body = req.body && typeof req.body === 'object' ? req.body : {};
		const source = String(body.url || body.html || '').trim();
		if (!source) {
			res.status(400).json({ status: 'failure', message: 'url or html is required' });
			return;
		}
		const data = await smartbrowz.takeScreenshot(source, {
			screenshot_options: body.screenshot_options || { type: 'jpeg' },
			page_options: body.page_options || {}
		});
		const buffer = await collectStream(data);
		const type = String(body.screenshot_options?.type || 'jpeg').toLowerCase() === 'png'
			? 'image/png'
			: 'image/jpeg';
		res.writeHead(200, {
			'Content-Type': type,
			'Content-Length': buffer.length
		});
		res.end(buffer);
	} catch (err) {
		console.error('sdkpdf screenshot error:', err);
		if (!res.headersSent) {
			res.status(500).json({
				status: 'failure',
				message: err.message || 'Screenshot failed'
			});
		}
	}
});

module.exports = app;
