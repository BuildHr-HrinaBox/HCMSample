'use strict';

const { IncomingMessage, ServerResponse } = require("http");
const { Readable } = require("stream");
const { URL } = require("url");
const zcatalyst = require("zcatalyst-sdk-node");

// In-memory storage (replace with Catalyst datastore for production)
let quizzes = {};
let players = {};

const BACKGROUND_FOLDER_NAME = 'Pictures';
const BACKGROUND_VIDEO_FILE_ID = process.env.BACKGROUND_VIDEO_FILE_ID || '9934000000126345';

// Helper function to generate a random quiz code
function generateQuizCode() {
	return Math.random().toString(36).substring(2, 8).toUpperCase();
}

// Helper function to parse request body
function parseBody(req) {
	return new Promise((resolve, reject) => {
		let body = '';
		req.on('data', chunk => {
			body += chunk.toString();
		});
		req.on('end', () => {
			try {
				resolve(body ? JSON.parse(body) : {});
			} catch (e) {
				reject(e);
			}
		});
		req.on('error', reject);
	});
}

// Helper function to send JSON response
function sendJSON(res, statusCode, data) {
	res.writeHead(statusCode, { 'Content-Type': 'application/json' });
	res.write(JSON.stringify(data));
	res.end();
}

function getServerBaseUrl(req) {
	const protocol = req.headers['x-forwarded-proto'] || 'https';
	const host = req.headers['host'] || 'quizversion1-60033944640.development.catalystserverless.in';
	return `${protocol}://${host}/server/quizversion_1_function`;
}

// Helper function to parse multipart form data
function parseMultipartFormData(req) {
	return new Promise((resolve, reject) => {
		const chunks = [];
		req.on('data', chunk => {
			chunks.push(chunk);
		});
		req.on('end', async () => {
			try {
				const buffer = Buffer.concat(chunks);
				const contentType = req.headers['content-type'] || '';
				const boundaryMatch = contentType.match(/boundary=([^;]+)/);
				
				if (!boundaryMatch) {
					reject(new Error('No boundary found in Content-Type'));
					return;
				}

				const boundary = Buffer.from('--' + boundaryMatch[1].trim());
				const fields = {};
				
				// Find all boundary positions
				const boundaryPositions = [];
				let searchPos = 0;
				while (true) {
					const pos = buffer.indexOf(boundary, searchPos);
					if (pos === -1) break;
					boundaryPositions.push(pos);
					searchPos = pos + boundary.length;
				}
				
				// Process each part between boundaries
				for (let i = 0; i < boundaryPositions.length - 1; i++) {
					const startPos = boundaryPositions[i] + boundary.length;
					const endPos = boundaryPositions[i + 1];
					
					// Skip if part is too small
					if (endPos - startPos < 10) continue;
					
					const partBuffer = buffer.slice(startPos, endPos);
					
					// Find header-body separator (look for \r\n\r\n or \n\n)
					const separator1 = Buffer.from('\r\n\r\n');
					const separator2 = Buffer.from('\n\n');
					
					let headerEnd = -1;
					let separatorLength = 0;
					
					const sep1Pos = partBuffer.indexOf(separator1);
					const sep2Pos = partBuffer.indexOf(separator2);
					
					if (sep1Pos !== -1) {
						headerEnd = sep1Pos;
						separatorLength = separator1.length;
					} else if (sep2Pos !== -1) {
						headerEnd = sep2Pos;
						separatorLength = separator2.length;
					} else {
						continue;
					}
					
					// Extract headers (as string)
					const headersBuffer = partBuffer.slice(0, headerEnd);
					const headers = headersBuffer.toString('utf8');
					
					// Extract body (keep as buffer for files)
					const bodyStart = headerEnd + separatorLength;
					let bodyBuffer = partBuffer.slice(bodyStart);
					
					// Remove trailing \r\n or \n
					if (bodyBuffer.length >= 2 && bodyBuffer[bodyBuffer.length - 2] === 0x0D && bodyBuffer[bodyBuffer.length - 1] === 0x0A) {
						bodyBuffer = bodyBuffer.slice(0, bodyBuffer.length - 2);
					} else if (bodyBuffer.length >= 1 && bodyBuffer[bodyBuffer.length - 1] === 0x0A) {
						bodyBuffer = bodyBuffer.slice(0, bodyBuffer.length - 1);
					}
					
					const nameMatch = headers.match(/name="([^"]+)"/);
					if (!nameMatch) continue;
					
					const name = nameMatch[1];
					const filenameMatch = headers.match(/filename="([^"]+)"/);
					
					if (filenameMatch) {
						// It's a file
						const filename = filenameMatch[1];
						const contentTypeMatch = headers.match(/Content-Type:\s*([^\r\n]+)/);
						
						fields[name] = {
							filename,
							contentType: contentTypeMatch ? contentTypeMatch[1].trim() : 'application/octet-stream',
							data: bodyBuffer
						};
					} else {
						// It's a regular field
						fields[name] = bodyBuffer.toString('utf8');
					}
				}
				
				resolve(fields);
			} catch (e) {
				reject(e);
			}
		});
		req.on('error', reject);
	});
}

/**
 * 
 * @param {IncomingMessage} req 
 * @param {ServerResponse} res 
 */
module.exports = async (req, res) => {
	var url = req.url;
	var method = req.method;
	const parsedUrl = new URL(req.url, 'http://localhost');
	const pathname = parsedUrl.pathname;

	const normalizePath = (path) => {
		const prefixes = [
			'/server/quizversion_1_function',
			'/quizversion_1_function',
			'/server/server/quizversion_1_function'
		];

		for (const prefix of prefixes) {
			if (path.startsWith(prefix)) {
				const stripped = path.substring(prefix.length);
				return stripped.startsWith('/') ? stripped : `/${stripped}`;
			}
		}

		return path;
	};

	const normalizedPath = normalizePath(pathname);

	// Enable CORS
	res.setHeader('Access-Control-Allow-Origin', '*');
	res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
	res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

	if (method === 'OPTIONS') {
		res.writeHead(200);
		res.end();
		return;
	}

	try {
		// Handle /api/background-video endpoint - Get background video URL
		if (normalizedPath === '/api/background-video' && method === 'GET') {
			try {
				const fileId = parsedUrl.searchParams.get('fileId') || BACKGROUND_VIDEO_FILE_ID;
				
				if (!fileId) {
					return sendJSON(res, 404, { error: 'Background video file ID is not configured' });
				}

				const catalyst = zcatalyst.initialize(req);
				const filestore = catalyst.filestore();
				const folder = filestore.folder(BACKGROUND_FOLDER_NAME);

				await folder.getFileDetails(fileId);

				const baseUrl = getServerBaseUrl(req);
				const proxyUrl = `${baseUrl}/api/background-video/download?fileId=${encodeURIComponent(fileId)}`;

				return sendJSON(res, 200, { videoUrl: proxyUrl, fileId });
			} catch (videoError) {
				console.error('Error fetching background video information:', videoError);
				return sendJSON(res, 500, { error: 'Failed to fetch background video details: ' + videoError.message });
			}
		}

		// Handle /api/background-video/download endpoint - Stream video
		if (normalizedPath === '/api/background-video/download' && method === 'GET') {
			try {
				const fileId = parsedUrl.searchParams.get('fileId') || BACKGROUND_VIDEO_FILE_ID;

				if (!fileId) {
					return sendJSON(res, 404, { error: 'Background video file ID is not configured' });
				}

				const catalyst = zcatalyst.initialize(req);
				const filestore = catalyst.filestore();
				const folder = filestore.folder(BACKGROUND_FOLDER_NAME);

				const fileBuffer = await folder.downloadFile(fileId);
				const fileDetails = await folder.getFileDetails(fileId);
				const fileName = fileDetails.file_name || fileDetails.File_Name || fileDetails.name || 'background-video.mp4';
				const contentType = fileName.endsWith('.webm') ? 'video/webm' : 'video/mp4';

				res.writeHead(200, {
					'Content-Type': contentType,
					'Content-Length': fileBuffer.length,
					'Cache-Control': 'public, max-age=3600'
				});
				res.write(fileBuffer);
				res.end();

				console.log(`Served background video: ${fileName}, size: ${fileBuffer.length} bytes`);
				return;
			} catch (downloadError) {
				console.error('Error downloading background video:', downloadError);
				return sendJSON(res, 500, { error: 'Failed to download background video: ' + downloadError.message });
			}
		}

		// Handle /api/image/:fileId endpoint - Get image URL from file ID
		if (url.match(/^\/api\/image\/([^\/]+)$/) && method === 'GET') {
			try {
				const match = url.match(/^\/api\/image\/([^\/]+)$/);
				const fileId = match[1];
				
				console.log('Fetching image for file ID:', fileId);
				
				if (!fileId) {
					return sendJSON(res, 400, { error: 'File ID is required' });
				}

				// Initialize Catalyst and get file from File Store
				const catalyst = zcatalyst.initialize(req);
				const filestore = catalyst.filestore();
				const folder = filestore.folder('quizimages');
				
				// Get file details to retrieve download URL and folder ID
				console.log('Getting file details for ID:', fileId);
				const fileDetails = await folder.getFileDetails(fileId);
				console.log('File details received:', JSON.stringify(fileDetails, null, 2));
				
				// Extract folder ID and project ID from file details
				const folderId = fileDetails.folder_details || fileDetails.folder_id || fileDetails.folderId;
				const projectId = fileDetails.project_details?.id || fileDetails.project_id || fileDetails.projectId || '9934000000097019';
				
				console.log('Extracted folder ID:', folderId);
				console.log('Extracted project ID:', projectId);
				
				// Get download URL - check multiple possible fields
				let fileUrl = fileDetails.downloadUrl || fileDetails.url || fileDetails.download_url || 
				             fileDetails.DownloadUrl || fileDetails.DOWNLOAD_URL;
				
				console.log('Initial fileUrl from fileDetails:', fileUrl);
				
				// If URL not in file details, try to get download request
				if (!fileUrl) {
					try {
						console.log('Getting download request for file ID:', fileId);
						const downloadRequest = await folder.getDownloadRequest(fileId);
						console.log('Download request received:', JSON.stringify(downloadRequest, null, 2));
						
						// Try different ways to construct the URL
						if (downloadRequest && downloadRequest.path) {
							// Check if path uses folder name instead of ID (incorrect format)
							if (downloadRequest.path.includes('/folder/quizimages/')) {
								// Path is incorrect - construct proper URL using folder ID from file details
								if (folderId) {
									fileUrl = `https://api.catalyst.zoho.in/baas/v1/project/${projectId}/folder/${folderId}/file/${fileId}/download`;
									console.log('Fixed URL (replaced folder name with ID):', fileUrl);
								}
							} else {
								// Path looks correct, use it
								if (downloadRequest.path.startsWith('/')) {
									fileUrl = `https://api.catalyst.zoho.in${downloadRequest.path}`;
								} else {
									fileUrl = `https://api.catalyst.zoho.in/${downloadRequest.path}`;
								}
							}
						} else if (downloadRequest && downloadRequest.url) {
							fileUrl = downloadRequest.url;
						} else if (downloadRequest && typeof downloadRequest === 'string') {
							// Sometimes downloadRequest might be a string URL
							fileUrl = downloadRequest;
						}
						
						console.log('Constructed fileUrl from download request:', fileUrl);
					} catch (downloadError) {
						console.error('Error getting download request:', downloadError);
					}
				}
				
				// If still no URL, construct it manually using folder ID and file ID from file details
				if (!fileUrl) {
					if (folderId) {
						// Use our proxy endpoint instead of direct Catalyst URL (which requires auth)
						// Get the base URL from the request
						const protocol = req.headers['x-forwarded-proto'] || 'https';
						const host = req.headers['host'] || 'quizversion1-60033944640.development.catalystserverless.in';
						fileUrl = `${protocol}://${host}/server/quizversion_1_function/api/image/${fileId}/download`;
						console.log('Using proxy URL:', fileUrl);
					} else {
						console.error('Cannot construct URL: folder ID missing from file details');
					}
				} else {
					// Even if we have a URL, use our proxy endpoint to avoid auth issues
					const protocol = req.headers['x-forwarded-proto'] || 'https';
					const host = req.headers['host'] || 'quizversion1-60033944640.development.catalystserverless.in';
					fileUrl = `${protocol}://${host}/server/quizversion_1_function/api/image/${fileId}/download`;
					console.log('Using proxy URL (overriding direct URL):', fileUrl);
				}
				
				if (!fileUrl) {
					console.error('Failed to get file URL. File details:', JSON.stringify(fileDetails, null, 2));
					return sendJSON(res, 404, { error: 'Image not found or URL unavailable. File ID: ' + fileId });
				}

				console.log('Returning image URL:', fileUrl);
				return sendJSON(res, 200, { imageUrl: fileUrl, fileId });
			} catch (imageError) {
				console.error('Error fetching image:', imageError);
				console.error('Error stack:', imageError.stack);
				return sendJSON(res, 500, { error: 'Failed to fetch image: ' + imageError.message });
			}
		}

		// Handle /api/image/:fileId/download endpoint - Proxy image file directly
		if (url.match(/^\/api\/image\/([^\/]+)\/download$/) && method === 'GET') {
			try {
				const match = url.match(/^\/api\/image\/([^\/]+)\/download$/);
				const fileId = match[1];
				
				console.log('Downloading image file for ID:', fileId);
				
				if (!fileId) {
					return sendJSON(res, 400, { error: 'File ID is required' });
				}

				// Initialize Catalyst and get file from File Store
				const catalyst = zcatalyst.initialize(req);
				const filestore = catalyst.filestore();
				const folder = filestore.folder('quizimages');
				
				// Download file as buffer
				const fileBuffer = await folder.downloadFile(fileId);
				
				// Get file details to determine content type
				const fileDetails = await folder.getFileDetails(fileId);
				const fileName = fileDetails.file_name || 'image.jpg';
				const contentType = fileName.endsWith('.png') ? 'image/png' : 
				                   fileName.endsWith('.gif') ? 'image/gif' : 
				                   fileName.endsWith('.webp') ? 'image/webp' : 'image/jpeg';
				
				// Set headers and send file
				res.writeHead(200, {
					'Content-Type': contentType,
					'Content-Length': fileBuffer.length,
					'Cache-Control': 'public, max-age=31536000'
				});
				res.write(fileBuffer);
				res.end();
				
				console.log(`Served image file: ${fileName}, size: ${fileBuffer.length} bytes`);
				return;
			} catch (downloadError) {
				console.error('Error downloading image file:', downloadError);
				return sendJSON(res, 500, { error: 'Failed to download image: ' + downloadError.message });
			}
		}

		// Handle /api/upload-image endpoint
		if (url === '/api/upload-image' && method === 'POST') {
			try {
				const formData = await parseMultipartFormData(req);
				const imageFile = formData.image;
				
				if (!imageFile || !imageFile.data) {
					return sendJSON(res, 400, { error: 'No image file provided' });
				}

				// Validate file type
				if (!imageFile.contentType || !imageFile.contentType.startsWith('image/')) {
					return sendJSON(res, 400, { error: 'File must be an image' });
				}

				// Validate file size (max 5MB)
				if (imageFile.data.length > 5 * 1024 * 1024) {
					return sendJSON(res, 400, { error: 'Image size must be less than 5MB' });
				}

				// Initialize Catalyst and upload to File Store
				const catalyst = zcatalyst.initialize(req);
				const filestore = catalyst.filestore();
				
				const folderName = 'quizimages';
				const folder = filestore.folder(folderName);
				
				// Generate unique filename
				const timestamp = Date.now();
				const randomStr = Math.random().toString(36).substring(2, 8);
				const filename = `question-${timestamp}-${randomStr}-${imageFile.filename}`;
				
				// Upload file to Catalyst File Store
				let uploadResponse;
				const fileStream = Readable.from(imageFile.data);
				fileStream.path = filename;
				try {
					uploadResponse = await folder.uploadFile({
						code: fileStream,
						name: filename
					});
				} catch (uploadError) {
					// If folder is missing, create it once and retry
					if (uploadError?.errorInfo?.code === 'filestore/folder-not-found') {
						await filestore.createFolder(folderName);
						const retryStream = Readable.from(imageFile.data);
						retryStream.path = filename;
						uploadResponse = await folder.uploadFile({
							code: retryStream,
							name: filename
						});
					} else {
						throw uploadError;
					}
				}

				// Log the full response to debug
				console.log('Upload response:', JSON.stringify(uploadResponse, null, 2));
				
				// Get the file ID - check multiple possible field names
				const fileId = uploadResponse.id || uploadResponse.file_id || uploadResponse.fileId || uploadResponse.ROWID || uploadResponse.fileId;
				
				console.log('Upload response structure:', JSON.stringify(uploadResponse, null, 2));
				console.log('Extracted file ID:', fileId);
				
				if (!fileId) {
					console.error('No file ID in upload response:', uploadResponse);
					return sendJSON(res, 500, { error: 'Failed to get file ID after upload' });
				}
				
				// Get download URL - try multiple sources
				let fileUrl = uploadResponse.downloadUrl || uploadResponse.url || uploadResponse.download_url;
				console.log('Initial fileUrl from upload response:', fileUrl);
				
				// If URL not in response, get it from file details
				if (!fileUrl && fileId) {
					try {
						console.log('Getting file details for uploaded file ID:', fileId);
						const fileDetails = await folder.getFileDetails(fileId);
						console.log('File details from getFileDetails:', JSON.stringify(fileDetails, null, 2));
						fileUrl = fileDetails.downloadUrl || fileDetails.url || fileDetails.download_url || 
						          fileDetails.DownloadUrl || fileDetails.DOWNLOAD_URL;
						console.log('FileUrl from file details:', fileUrl);
					} catch (fileDetailsError) {
						console.error('Error getting file details:', fileDetailsError);
					}
				}
				
				// If still no URL, try to get download request
				if (!fileUrl && fileId) {
					try {
						console.log('Getting download request for file ID:', fileId);
						const downloadRequest = await folder.getDownloadRequest(fileId);
						console.log('Download request:', JSON.stringify(downloadRequest, null, 2));
						
						if (downloadRequest && downloadRequest.path) {
							// Construct full URL from download request
							if (downloadRequest.path.startsWith('/')) {
								fileUrl = `https://api.catalyst.zoho.in${downloadRequest.path}`;
							} else {
								fileUrl = `https://api.catalyst.zoho.in/${downloadRequest.path}`;
							}
						} else if (downloadRequest && downloadRequest.url) {
							fileUrl = downloadRequest.url;
						}
						console.log('Constructed fileUrl from download request:', fileUrl);
					} catch (downloadError) {
						console.error('Error getting download request:', downloadError);
					}
				}
				
				if (!fileUrl) {
					console.error('Failed to get file URL. Upload response:', JSON.stringify(uploadResponse, null, 2));
					return sendJSON(res, 500, { error: 'Failed to get file URL after upload. File ID: ' + fileId });
				}

				console.log(`Image uploaded successfully: ${filename}, File ID: ${fileId}, URL: ${fileUrl}`);
				
				return sendJSON(res, 200, { imageUrl: fileUrl, fileId: fileId.toString() });
			} catch (uploadError) {
				console.error('Error uploading image:', uploadError);
				return sendJSON(res, 500, { error: 'Failed to upload image: ' + uploadError.message });
			}
		}

		// Handle /api/quiz/create endpoint
		if (url === '/api/quiz/create' && method === 'POST') {
			const body = await parseBody(req);
			const { title, description, questions } = body;

			// Validate input
			if (!title || !questions || questions.length === 0) {
				return sendJSON(res, 400, { 
					error: 'Missing required fields: title and questions are required' 
				});
			}

			// Generate unique quiz code
			let code;
			do {
				code = generateQuizCode();
			} while (quizzes[code]);

			// Create quiz object
			const quiz = {
				id: Date.now().toString(),
				title,
				description: description || '',
				code,
				questions,
				createdAt: new Date().toISOString(),
				isActive: true,
				currentQuestionIndex: 0,
				isStarted: false,
				isFinished: false,
				theme: 'christmas'
			};

			// Store quiz in memory
			quizzes[code] = quiz;
			players[code] = [];

			// Persist quiz metadata + questions in DataStore
			try {
				const catalyst = zcatalyst.initialize(req);
				const datastore = catalyst.datastore();
				const createQuizTable = datastore.table('CreateQuiz');
				const questionTable = datastore.table('Question');

				// Store each question in both CreateQuiz (legacy schema) and Question tables
				for (let index = 0; index < questions.length; index += 1) {
					const question = questions[index];
					const options = question.options || [];

					// Insert into CreateQuiz table using available columns
					await createQuizTable.insertRow({
						QuizTitle: title,
						Description: description || '',
						Question: question.text || '',
						option1: options[0] || '',
						option2: options[1] || '',
						option3: options[2] || '',
						option4: options[3] || '',
						CorrectAnswer: typeof question.correctAnswer === 'number' ? question.correctAnswer.toString() : '0',
						AddQuestion: quiz.id,
						StartQuiz: code,
						Timelimit: question.timeLimit ? question.timeLimit.toString() : '30',
						ImageFileId: question.imageFileId || '',
						MODIFIEDTIME: new Date()
					});

					// Insert into Question table (structured storage). Run in a best-effort block so
					// CreateQuiz rows are still written even if this schema is different.
					try {
						await questionTable.insertRow({
							QuizId: quiz.id,
							GameCode: code,
							QuestionId: question.id || `${quiz.id}-${index + 1}`,
							Question: question.text || '',
							option1: options[0] || '',
							option2: options[1] || '',
							option3: options[2] || '',
							option4: options[3] || '',
						CorrectAnswer: typeof question.correctAnswer === 'number' ? question.correctAnswer.toString() : '0',
						Points: typeof question.points === 'number' ? question.points.toString() : '1000',
						Timelimit: question.timeLimit ? question.timeLimit.toString() : '30',
						QuestionOrder: (index + 1).toString(),
							ImageUrl: question.imageUrl || '',
							MODIFIEDTIME: new Date()
					});
					} catch (questionTableError) {
						console.error('Error inserting into Question table:', questionTableError);
					}
				}
				console.log(`Saved ${questions.length} question(s) to CreateQuiz and Question tables`);
			} catch (dbError) {
				console.error('Error saving to DataStore:', dbError);
				// Continue even if database save fails - quiz is still created in memory
			}

			return sendJSON(res, 200, { quiz });
		}

		// Handle /api/quiz/:code endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)$/) && method === 'GET') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)$/);
			const code = match[1];
			
			// Try to get from memory first
			let quiz = quizzes[code];
			
			// If not in memory or missing questions, load from database
			if (!quiz || !quiz.questions || quiz.questions.length === 0) {
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const createQuizTable = datastore.table('CreateQuiz');
					const rows = await createQuizTable.getAllRows();
					
					// Find quiz by code
					const quizRows = rows.filter(row => {
						const rowCode = (row.StartQuiz || row.GameCode || row.gameCode || row.Code || '').toString();
						return rowCode === code;
					});
					
					if (quizRows.length > 0) {
						// Group by quiz ID
						const quizId = (quizRows[0].AddQuestion || quizRows[0].QuizId || quizRows[0].quizId || '').toString();
						if (quizId) {
							const allQuizRows = rows.filter(row => {
								const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
								return rowQuizId === quizId;
							});
							
							const options = [
								allQuizRows[0].option1 || allQuizRows[0].Option1 || '',
								allQuizRows[0].option2 || allQuizRows[0].Option2 || '',
								allQuizRows[0].option3 || allQuizRows[0].Option3 || '',
								allQuizRows[0].option4 || allQuizRows[0].Option4 || ''
							];
							const correctAnswerRaw = allQuizRows[0].CorrectAnswer ?? allQuizRows[0].correctAnswer ?? allQuizRows[0].Correctanswer ?? allQuizRows[0].correctanswer;
							const correctAnswer = Number.parseInt(correctAnswerRaw, 10);
							const timeLimitRaw = allQuizRows[0].Timelimit ?? allQuizRows[0].timelimit ?? allQuizRows[0].TimeLimit ?? allQuizRows[0].timeLimit;
							const timeLimit = Number.parseInt(timeLimitRaw, 10);
							const pointsRaw = allQuizRows[0].points ?? allQuizRows[0].Points;
							const points = Number.parseInt(pointsRaw, 10);
							
							quiz = {
								id: quizId,
								title: allQuizRows[0].QuizTitle || allQuizRows[0].quizTitle || 'Untitled Quiz',
								description: allQuizRows[0].Description || allQuizRows[0].description || '',
								code: code,
								questions: allQuizRows.map((row, index) => {
									const rowOptions = [
										row.option1 || row.Option1 || '',
										row.option2 || row.Option2 || '',
										row.option3 || row.Option3 || '',
										row.option4 || row.Option4 || ''
									];
									const rowCorrectAnswerRaw = row.CorrectAnswer ?? row.correctAnswer ?? row.Correctanswer ?? row.correctanswer;
									const rowCorrectAnswer = Number.parseInt(rowCorrectAnswerRaw, 10);
									const rowTimeLimitRaw = row.Timelimit ?? row.timelimit ?? row.TimeLimit ?? row.timeLimit;
									const rowTimeLimit = Number.parseInt(rowTimeLimitRaw, 10);
									const rowPointsRaw = row.points ?? row.Points;
									const rowPoints = Number.parseInt(rowPointsRaw, 10);
									
									return {
										id: (row.QuestionId || row.questionId || row.QuestionID || row.ROWID || (index + 1)).toString(),
										text: row.Question || row.question || '',
										options: rowOptions,
										correctAnswer: Number.isNaN(rowCorrectAnswer) ? 0 : rowCorrectAnswer,
										timeLimit: Number.isNaN(rowTimeLimit) ? 30 : rowTimeLimit,
										points: Number.isNaN(rowPoints) ? 1000 : rowPoints || 1000,
										imageUrl: row.ImageUrl || row.imageUrl || row.image_url || row.ImageURL || row.imageURL || row.imageurl || undefined,
										imageFileId: row.ImageFileId || row.imageFileId || row.image_file_id || row.IMAGEFILEID || row.imagefileid || undefined
									};
								}),
								createdAt: allQuizRows[0].CREATEDTIME || allQuizRows[0].createdTime || new Date().toISOString(),
								isActive: true,
								currentQuestionIndex: 0, // Always reset when loading from DB
								isStarted: false, // Always false when loading from DB (game state is not persisted)
								isFinished: false // Always false when loading from DB
							};
							
							// Update memory cache
							quizzes[code] = quiz;
						}
					}
				} catch (dbError) {
					console.error('Error loading quiz from DataStore:', dbError);
				}
			}

			if (!quiz) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			return sendJSON(res, 200, { quiz });
		}

		// Handle /api/quiz/:id/update endpoint
		if (url.match(/^\/api\/quiz\/([^\/]+)\/update$/) && method === 'PUT') {
			const match = url.match(/^\/api\/quiz\/([^\/]+)\/update$/);
			const quizId = match[1];
			const body = await parseBody(req);
			const { title, description, questions } = body;

			// Validate input
			if (!title || !questions || questions.length === 0) {
				return sendJSON(res, 400, { 
					error: 'Missing required fields: title and questions are required' 
				});
			}

			// Find quiz by ID in memory
			let quiz = null;
			let quizCode = null;
			for (const code in quizzes) {
				if (quizzes[code].id === quizId) {
					quiz = quizzes[code];
					quizCode = code;
					break;
				}
			}

			if (!quiz) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			// Update quiz object
			quiz.title = title;
			quiz.description = description || '';
			quiz.questions = questions;

			// Update questions in DataStore
			try {
				const catalyst = zcatalyst.initialize(req);
				const datastore = catalyst.datastore();
				const createQuizTable = datastore.table('CreateQuiz');

				// Delete existing rows for this quiz
				const existingRows = await createQuizTable.getAllRows();
				for (const row of existingRows) {
					const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
					const rowGameCode = (row.StartQuiz || row.GameCode || row.gameCode || row.Code || '').toString();
					if (rowQuizId === quizId || rowGameCode === (quizCode || quiz.code)) {
						await createQuizTable.deleteRow(row.ROWID);
					}
				}

				// Insert updated questions
				for (let index = 0; index < questions.length; index += 1) {
					const question = questions[index];
					const options = question.options || [];

					// Update Question table
					await createQuizTable.insertRow({
						QuizTitle: title,
						Description: description || '',
						Question: question.text || '',
						option1: options[0] || '',
						option2: options[1] || '',
						option3: options[2] || '',
						option4: options[3] || '',
						CorrectAnswer: typeof question.correctAnswer === 'number' ? question.correctAnswer.toString() : '0',
						AddQuestion: quiz.id,
						StartQuiz: quizCode || quiz.code,
						Timelimit: question.timeLimit ? question.timeLimit.toString() : '30',
						ImageFileId: question.imageFileId || '',
						MODIFIEDTIME: new Date()
					});
				}
				console.log(`Successfully updated ${questions.length} question(s) in CreateQuiz table`);
			} catch (dbError) {
				console.error('Error updating DataStore:', dbError);
				// Continue even if database update fails - quiz is still updated in memory
			}

			return sendJSON(res, 200, { quiz });
		}

		// Handle /api/quizzes endpoint
		if (url === '/api/quizzes' && method === 'GET') {
			let quizList = Object.values(quizzes);
			try {
				const catalyst = zcatalyst.initialize(req);
				const datastore = catalyst.datastore();
				const createQuizTable = datastore.table('CreateQuiz');
				const rows = await createQuizTable.getAllRows();
				console.log('Loaded rows from CreateQuiz table:', rows.length);

				const groupedQuizzes = new Map();
				for (const row of rows) {
					const quizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || row.ROWID || '').toString();
					if (!quizId) {
						continue;
					}

					if (!groupedQuizzes.has(quizId)) {
						const quizCode = (row.StartQuiz || row.GameCode || row.gameCode || row.Code || '').toString();
						groupedQuizzes.set(quizId, {
							id: quizId,
							title: row.QuizTitle || row.quizTitle || 'Untitled Quiz',
							description: row.Description || row.description || '',
							code: quizCode || generateQuizCode(),
							questions: [],
							createdAt: row.CREATEDTIME || row.createdTime || new Date().toISOString(),
							isActive: true,
							currentQuestionIndex: 0,
							isStarted: false,
							isFinished: false
						});
					}

					const quizEntry = groupedQuizzes.get(quizId);
					const options = [
						row.option1 || row.Option1 || '',
						row.option2 || row.Option2 || '',
						row.option3 || row.Option3 || '',
						row.option4 || row.Option4 || ''
					];
					const correctAnswerRaw = row.CorrectAnswer ?? row.correctAnswer ?? row.Correctanswer ?? row.correctanswer;
					const correctAnswer = Number.parseInt(correctAnswerRaw, 10);
					const timeLimitRaw = row.Timelimit ?? row.timelimit ?? row.TimeLimit ?? row.timeLimit;
					const timeLimit = Number.parseInt(timeLimitRaw, 10);
					const pointsRaw = row.points ?? row.Points;
					const points = Number.parseInt(pointsRaw, 10);

					quizEntry.questions.push({
						id: (row.QuestionId || row.questionId || row.QuestionID || row.ROWID || quizEntry.questions.length + 1).toString(),
						text: row.Question || row.question || '',
						options,
						correctAnswer: Number.isNaN(correctAnswer) ? 0 : correctAnswer,
						timeLimit: Number.isNaN(timeLimit) ? 30 : timeLimit,
						points: Number.isNaN(points) ? 1000 : points || 1000,
						imageUrl: row.ImageUrl || row.imageUrl || row.image_url || row.ImageURL || row.imageURL || row.imageurl || undefined,
						imageFileId: row.ImageFileId || row.imageFileId || row.image_file_id || row.IMAGEFILEID || row.imagefileid || undefined
					});
				}

				const datastoreQuizzes = Array.from(groupedQuizzes.values());
				datastoreQuizzes.forEach((quiz) => {
					if (quiz.code && !quizzes[quiz.code]) {
						quizzes[quiz.code] = quiz;
						if (!players[quiz.code]) {
							players[quiz.code] = [];
						}
					}
				});

				if (datastoreQuizzes.length > 0) {
					quizList = datastoreQuizzes;
				}
			} catch (dbError) {
				console.error('Error loading quizzes from DataStore:', dbError);
			}

			return sendJSON(res, 200, { quizzes: quizList });
		}

		// Handle /api/quiz/:code/join endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/join$/) && method === 'POST') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/join$/);
			const code = match[1];
			const body = await parseBody(req);
			const { name, avatar, deviceId } = body;

			// If quiz not in memory, try loading from database (e.g. quiz created earlier or server restarted)
			if (!quizzes[code]) {
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const createQuizTable = datastore.table('CreateQuiz');
					const rows = await createQuizTable.getAllRows();
					const quizRows = rows.filter(row => {
						const rowCode = (row.StartQuiz || row.GameCode || row.gameCode || row.Code || '').toString();
						return rowCode === code;
					});
					if (quizRows.length > 0) {
						const quizId = (quizRows[0].AddQuestion || quizRows[0].QuizId || quizRows[0].quizId || '').toString();
						if (quizId) {
							const allQuizRows = rows.filter(row => {
								const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
								return rowQuizId === quizId;
							});
							const quiz = {
								id: quizId,
								title: allQuizRows[0].QuizTitle || allQuizRows[0].quizTitle || 'Untitled Quiz',
								description: allQuizRows[0].Description || allQuizRows[0].description || '',
								code: code,
								questions: allQuizRows.map((row, index) => {
									const rowOptions = [
										row.option1 || row.Option1 || '',
										row.option2 || row.Option2 || '',
										row.option3 || row.Option3 || '',
										row.option4 || row.Option4 || ''
									];
									const rowCorrectAnswerRaw = row.CorrectAnswer ?? row.correctAnswer ?? row.Correctanswer ?? row.correctanswer;
									const rowCorrectAnswer = Number.parseInt(rowCorrectAnswerRaw, 10);
									const rowTimeLimitRaw = row.Timelimit ?? row.timelimit ?? row.TimeLimit ?? row.timeLimit;
									const rowTimeLimit = Number.parseInt(rowTimeLimitRaw, 10);
									const rowPointsRaw = row.points ?? row.Points;
									const rowPoints = Number.parseInt(rowPointsRaw, 10);
									return {
										id: (row.QuestionId || row.questionId || row.QuestionID || row.ROWID || (index + 1)).toString(),
										text: row.Question || row.question || '',
										options: rowOptions,
										correctAnswer: Number.isNaN(rowCorrectAnswer) ? 0 : rowCorrectAnswer,
										timeLimit: Number.isNaN(rowTimeLimit) ? 30 : rowTimeLimit,
										points: Number.isNaN(rowPoints) ? 1000 : rowPoints || 1000,
										imageUrl: row.ImageUrl || row.imageUrl || row.image_url || row.ImageURL || row.imageURL || row.imageurl || undefined,
										imageFileId: row.ImageFileId || row.imageFileId || row.image_file_id || row.IMAGEFILEID || row.imagefileid || undefined
									};
								}),
								createdAt: allQuizRows[0].CREATEDTIME || allQuizRows[0].createdTime || new Date().toISOString(),
								isActive: true,
								currentQuestionIndex: 0,
								isStarted: false,
								isFinished: false
							};
							quizzes[code] = quiz;
							if (!players[code]) {
								players[code] = [];
							}
						}
					}
				} catch (dbError) {
					console.error('Error loading quiz from DataStore on join:', dbError);
				}
			}

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			if (!name) {
				return sendJSON(res, 400, { error: 'Name is required' });
			}

			// Check if quiz has already started
			if (quizzes[code].isStarted) {
				return sendJSON(res, 400, { error: 'Quiz has already started. Cannot join now.' });
			}

			if (!players[code]) {
				players[code] = [];
			}

			// Check for duplicate player name (case-insensitive)
			const normalizedName = name.trim().toLowerCase();
			const duplicateName = players[code].find(p => p.name.trim().toLowerCase() === normalizedName);
			if (duplicateName) {
				return sendJSON(res, 409, { error: `A player with the name "${name}" is already in this quiz. Please choose a different name.` });
			}

			// Check for duplicate device ID (if provided) - prevent same device joining multiple times
			if (deviceId) {
				const duplicateDevice = players[code].find(p => p.deviceId === deviceId);
				if (duplicateDevice) {
					return sendJSON(res, 409, { error: `You have already joined this quiz as "${duplicateDevice.name}". Only one player per device is allowed.` });
				}
			}

			const player = {
				id: Date.now().toString() + '_' + Math.random().toString(36).substring(2, 9),
				name: name.trim(),
				avatar: avatar || 'default',
				score: 0,
				streak: 0,
				answers: [],
				badges: [],
				xp: 0,
				deviceId: deviceId || null // Store device ID if provided
			};

			players[code].push(player);

			console.log(`Player "${name}" joined quiz ${code}. Total players: ${players[code].length}`);

			// Note: Player data will be stored in PlayerName table when they answer questions
			// This ensures we store complete question-answer data for each response

			return sendJSON(res, 200, { player });
		}

		// Handle /api/quiz/:code/players endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/players$/) && method === 'GET') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/players$/);
			const code = match[1];

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			const playerList = players[code] || [];
			return sendJSON(res, 200, { players: playerList });
		}

		// Handle /api/quiz/:code/players/:playerId endpoint - Remove player
		if (normalizedPath.match(/^\/api\/quiz\/([A-Z0-9]+)\/players\/([^\/]+)$/) && method === 'DELETE') {
			const match = normalizedPath.match(/^\/api\/quiz\/([A-Z0-9]+)\/players\/([^\/]+)$/);
			const code = match[1];
			const playerId = match[2];

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			if (!players[code]) {
				return sendJSON(res, 404, { error: 'No players found for this quiz' });
			}

			// Find and remove the player
			const playerIndex = players[code].findIndex(p => p.id === playerId);
			if (playerIndex === -1) {
				return sendJSON(res, 404, { error: 'Player not found' });
			}

			// Remove the player from the array
			players[code].splice(playerIndex, 1);

			console.log(`Player ${playerId} removed from quiz ${code}`);
			return sendJSON(res, 200, { success: true, message: 'Player removed successfully' });
		}

		// Handle /api/quiz/:code/start endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/start$/) && method === 'POST') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/start$/);
			const code = match[1];

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			quizzes[code].isStarted = true;
			quizzes[code].currentQuestionIndex = 0;

			return sendJSON(res, 200, { success: true });
		}

		// Handle /api/quiz/:code/end endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/end$/) && method === 'POST') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/end$/);
			const code = match[1];

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			// Mark quiz as finished and stop the game
			quizzes[code].isFinished = true;
			quizzes[code].isStarted = false;

			// Save final results to DataStore when quiz ends
			try {
				const catalyst = zcatalyst.initialize(req);
				const datastore = catalyst.datastore();
				const resultsTable = datastore.table('QuizResults');
				
				// Sort players by score to get rankings
				const sortedPlayers = (players[code] || []).sort((a, b) => b.score - a.score);
				
				// Save each player's results
				for (let i = 0; i < sortedPlayers.length; i++) {
					const player = sortedPlayers[i];
					const ranking = i + 1;
					const correctAnswers = player.answers.filter(a => a.isCorrect).length;
					const totalQuestions = quizzes[code].questions.length;
					
					// Create question-by-question results as a simple string format
					// Format: "Q1:✓,Q2:✗,Q3:-,Q4:✓,Q5:✗" (✓=correct, ✗=wrong, -=not answered)
					const questionResults = quizzes[code].questions.map((question, qIndex) => {
						let answer = player.answers.find(a => a.questionId === question.id);
						if (!answer && player.answers[qIndex]) {
							answer = player.answers[qIndex];
						}
						
						let result = '-'; // Not answered
						if (answer) {
							result = answer.isCorrect ? '✓' : '✗';
						}
						
						return `Q${qIndex + 1}:${result}`;
					}).join(',');
					
					await resultsTable.insertRow({
						QuizId: quizzes[code].id,
						GameCode: code,
						QuizTitle: quizzes[code].title,
						PlayerId: player.id,
						PlayerName: player.name,
						Avatar: player.avatar || 'default',
						Ranking: ranking.toString(),
						Score: player.score.toString(),
						CorrectAnswers: correctAnswers.toString(),
						TotalQuestions: totalQuestions.toString(),
						QuestionResults: questionResults,
						CompletedDate: new Date().toISOString(),
						MODIFIEDTIME: new Date()
					});
				}
				
				console.log(`Saved results for ${sortedPlayers.length} players to QuizResults table`);
			} catch (dbError) {
				console.error('Error saving quiz results to DataStore:', dbError);
				// Continue even if database save fails
			}

			console.log(`Quiz ${code} ended by host`);
			return sendJSON(res, 200, { success: true, isFinished: true });
		}

		// Handle /api/quiz/:code/theme endpoint - host updates theme for all clients
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/theme$/) && method === 'POST') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/theme$/);
			const code = match[1];
			const body = await parseBody(req);
			const { theme } = body || {};

			const allowedThemes = ['candy', 'christmas'];
			if (!theme || !allowedThemes.includes(theme)) {
				return sendJSON(res, 400, { error: 'Invalid theme' });
			}

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			quizzes[code].theme = theme;
			return sendJSON(res, 200, { success: true, theme });
		}

		// Handle /api/quiz/:code/state endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/state$/) && method === 'GET') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/state$/);
			const code = match[1];

			// Ensure quiz is loaded with all data including images
			let quiz = quizzes[code];
			
			// If quiz not in memory or missing questions, load from database
			if (!quiz || !quiz.questions || quiz.questions.length === 0) {
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const createQuizTable = datastore.table('CreateQuiz');
					const rows = await createQuizTable.getAllRows();
					
					// Find quiz by code
					const quizRows = rows.filter(row => {
						const rowCode = (row.StartQuiz || row.GameCode || row.gameCode || row.Code || '').toString();
						return rowCode === code;
					});
					
					if (quizRows.length > 0) {
						const quizId = (quizRows[0].AddQuestion || quizRows[0].QuizId || quizRows[0].quizId || '').toString();
						if (quizId) {
							const allQuizRows = rows.filter(row => {
								const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
								return rowQuizId === quizId;
							});
							
							quiz = {
								id: quizId,
								title: allQuizRows[0].QuizTitle || allQuizRows[0].quizTitle || 'Untitled Quiz',
								description: allQuizRows[0].Description || allQuizRows[0].description || '',
								code: code,
								questions: allQuizRows.map((row, index) => {
									const rowOptions = [
										row.option1 || row.Option1 || '',
										row.option2 || row.Option2 || '',
										row.option3 || row.Option3 || '',
										row.option4 || row.Option4 || ''
									];
									const rowCorrectAnswerRaw = row.CorrectAnswer ?? row.correctAnswer ?? row.Correctanswer ?? row.correctanswer;
									const rowCorrectAnswer = Number.parseInt(rowCorrectAnswerRaw, 10);
									const rowTimeLimitRaw = row.Timelimit ?? row.timelimit ?? row.TimeLimit ?? row.timeLimit;
									const rowTimeLimit = Number.parseInt(rowTimeLimitRaw, 10);
									const rowPointsRaw = row.points ?? row.Points;
									const rowPoints = Number.parseInt(rowPointsRaw, 10);
									
									return {
										id: (row.QuestionId || row.questionId || row.QuestionID || row.ROWID || (index + 1)).toString(),
										text: row.Question || row.question || '',
										options: rowOptions,
										correctAnswer: Number.isNaN(rowCorrectAnswer) ? 0 : rowCorrectAnswer,
										timeLimit: Number.isNaN(rowTimeLimit) ? 30 : rowTimeLimit,
										points: Number.isNaN(rowPoints) ? 1000 : rowPoints || 1000,
										imageUrl: row.ImageUrl || row.imageUrl || row.image_url || row.ImageURL || row.imageURL || row.imageurl || undefined,
										imageFileId: row.ImageFileId || row.imageFileId || row.image_file_id || row.IMAGEFILEID || row.imagefileid || undefined
									};
								}),
								createdAt: allQuizRows[0].CREATEDTIME || allQuizRows[0].createdTime || new Date().toISOString(),
								isActive: true,
								currentQuestionIndex: quiz?.currentQuestionIndex || 0,
								isStarted: quiz?.isStarted || false,
								isFinished: quiz?.isFinished || false,
								theme: quiz?.theme || 'christmas'
							};
							
							// Update memory cache
							quizzes[code] = quiz;
						}
					}
				} catch (dbError) {
					console.error('Error loading quiz from DataStore in state endpoint:', dbError);
				}
			}

			if (!quiz) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			const playerList = players[code] || [];

			// Ensure theme exists
			quiz.theme = quiz.theme || 'candy';

			const gameState = {
				quiz: null,
				currentQuestionIndex: quiz.currentQuestionIndex,
				players: playerList,
				isStarted: quiz.isStarted,
				isFinished: quiz.isFinished,
				showLeaderboard: false,
				theme: quiz.theme
			};

			return sendJSON(res, 200, { 
				quiz, 
				players: playerList, 
				gameState 
			});
		}

		// Handle /api/quiz/:code/answer endpoint
		if (normalizedPath.match(/^\/api\/quiz\/([A-Z0-9]+)\/answer$/) && method === 'POST') {
			const match = normalizedPath.match(/^\/api\/quiz\/([A-Z0-9]+)\/answer$/);
			const code = match[1];
			const body = await parseBody(req);
			const { playerId, questionId, answer, timeSpent, pointsEarned, isCorrect } = body;

			if (!quizzes[code] || !players[code]) {
				return sendJSON(res, 404, { error: 'Quiz or player not found' });
			}

			const player = players[code].find(p => p.id === playerId);
			if (!player) {
				return sendJSON(res, 404, { error: 'Player not found' });
			}

			// Find the question by questionId if provided, otherwise fall back to currentQuestionIndex
			let currentQuestion = null;
			if (questionId) {
				currentQuestion = quizzes[code].questions.find(q => q.id === questionId);
			}
			if (!currentQuestion) {
				// Fallback to currentQuestionIndex if questionId not found
				currentQuestion = quizzes[code].questions[quizzes[code].currentQuestionIndex];
			}

			player.score += pointsEarned;
			if (isCorrect) {
				player.streak += 1;
			} else {
				player.streak = 0;
			}
			player.answers.push({
				questionId: questionId || currentQuestion?.id,
				answer,
				isCorrect,
				timeSpent,
				pointsEarned
			});

			// Insert a new row in PlayerName table for each question answer
			// This ensures all questions are stored, not just the latest one
			try {
				const catalyst = zcatalyst.initialize(req);
				const datastore = catalyst.datastore();
				const playerTable = datastore.table('PlayerName');
				
				// Insert a new row for this question answer
				await playerTable.insertRow({
					Name: player.name,
					GameCode: code,
					avatar: player.avatar || 'default',
					Question: currentQuestion ? (currentQuestion.text || '') : '',
					Correctanswer: isCorrect ? 'true' : 'false',
					score: String(player.score), // Cumulative score
					MODIFIEDTIME: new Date()
				});
				
				console.log(`Saved answer for player ${player.name} - Question: ${currentQuestion?.text}, Correct: ${isCorrect}`);
			} catch (dbError) {
				console.error('Error inserting answer into PlayerName table:', dbError);
			}

			return sendJSON(res, 200, { player });
		}

		// Handle /api/quiz/:id/start-session endpoint - Create new session for existing quiz
		if (normalizedPath.match(/^\/api\/quiz\/([^\/]+)\/start-session$/) && method === 'POST') {
			const match = normalizedPath.match(/^\/api\/quiz\/([^\/]+)\/start-session$/);
			const quizId = match[1];
			
			// Find the quiz by ID in memory or database
			let sourceQuiz = null;
			
			// First check memory
			for (const code in quizzes) {
				if (quizzes[code].id === quizId) {
					sourceQuiz = quizzes[code];
					break;
				}
			}
			
			// If not in memory, load from database
			if (!sourceQuiz) {
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const createQuizTable = datastore.table('CreateQuiz');
					const rows = await createQuizTable.getAllRows();
					
					// Find quiz by ID
					const quizRows = rows.filter(row => {
						const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
						return rowQuizId === quizId;
					});
					
					if (quizRows.length > 0) {
						const allQuizRows = rows.filter(row => {
							const rowQuizId = (row.AddQuestion || row.QuizId || row.quizId || row.quiz_id || row.QUIZID || row.QUIZ_ID || '').toString();
							return rowQuizId === quizId;
						});
						
						sourceQuiz = {
							id: quizId,
							title: allQuizRows[0].QuizTitle || allQuizRows[0].quizTitle || 'Untitled Quiz',
							description: allQuizRows[0].Description || allQuizRows[0].description || '',
							code: '', // Will generate new code
							questions: allQuizRows.map((row, index) => {
								const rowOptions = [
									row.option1 || row.Option1 || '',
									row.option2 || row.Option2 || '',
									row.option3 || row.Option3 || '',
									row.option4 || row.Option4 || ''
								];
								const rowCorrectAnswerRaw = row.CorrectAnswer ?? row.correctAnswer ?? row.Correctanswer ?? row.correctanswer;
								const rowCorrectAnswer = Number.parseInt(rowCorrectAnswerRaw, 10);
								const rowTimeLimitRaw = row.Timelimit ?? row.timelimit ?? row.TimeLimit ?? row.timeLimit;
								const rowTimeLimit = Number.parseInt(rowTimeLimitRaw, 10);
								const rowPointsRaw = row.points ?? row.Points;
								const rowPoints = Number.parseInt(rowPointsRaw, 10);
								
								return {
									id: (row.QuestionId || row.questionId || row.QuestionID || row.ROWID || (index + 1)).toString(),
									text: row.Question || row.question || '',
									options: rowOptions,
									correctAnswer: Number.isNaN(rowCorrectAnswer) ? 0 : rowCorrectAnswer,
									timeLimit: Number.isNaN(rowTimeLimit) ? 30 : rowTimeLimit,
									points: Number.isNaN(rowPoints) ? 1000 : rowPoints || 1000,
									imageUrl: row.ImageUrl || row.imageUrl || row.image_url || row.ImageURL || row.imageURL || row.imageurl || undefined,
									imageFileId: row.ImageFileId || row.imageFileId || row.image_file_id || row.IMAGEFILEID || row.imagefileid || undefined
								};
							}),
							createdAt: allQuizRows[0].CREATEDTIME || allQuizRows[0].createdTime || new Date().toISOString(),
							isActive: true,
							currentQuestionIndex: 0,
							isStarted: false,
							isFinished: false
						};
					}
				} catch (dbError) {
					console.error('Error loading quiz from DataStore for new session:', dbError);
				}
			}
			
			if (!sourceQuiz) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}
			
			// Generate a NEW unique code for this session
			let newCode;
			do {
				newCode = generateQuizCode();
			} while (quizzes[newCode] && !quizzes[newCode].isFinished);
			
			// If code exists but quiz is finished, we can reuse it (reset it)
			if (quizzes[newCode] && quizzes[newCode].isFinished) {
				// Reset the finished quiz
				quizzes[newCode].isStarted = false;
				quizzes[newCode].isFinished = false;
				quizzes[newCode].currentQuestionIndex = 0;
				players[newCode] = [];
			}
			
			// Create new quiz session with new code
			const newQuiz = {
				id: sourceQuiz.id, // Keep original quiz ID
				title: sourceQuiz.title,
				description: sourceQuiz.description,
				code: newCode, // NEW code for this session
				questions: sourceQuiz.questions, // Copy questions
				createdAt: new Date().toISOString(), // New session timestamp
				isActive: true,
				currentQuestionIndex: 0,
				isStarted: false,
				isFinished: false
			};
			
			// Store in memory
			quizzes[newCode] = newQuiz;
			players[newCode] = [];
			
			console.log(`Created new session for quiz ${quizId} with code ${newCode}`);
			
			return sendJSON(res, 200, { quiz: newQuiz });
		}

		// Handle /api/quiz/:code/next-question endpoint
		if (url.match(/^\/api\/quiz\/([A-Z0-9]+)\/next-question$/) && method === 'POST') {
			const match = url.match(/^\/api\/quiz\/([A-Z0-9]+)\/next-question$/);
			const code = match[1];

			if (!quizzes[code]) {
				return sendJSON(res, 404, { error: 'Quiz not found' });
			}

			const quiz = quizzes[code];
			quiz.currentQuestionIndex += 1;

			if (quiz.currentQuestionIndex >= quiz.questions.length) {
				quiz.isFinished = true;
				
				// Save final results to DataStore when quiz ends
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const resultsTable = datastore.table('QuizResults');
					
					// Sort players by score to get rankings
					const sortedPlayers = (players[code] || []).sort((a, b) => b.score - a.score);
					
					// Save each player's results
					for (let i = 0; i < sortedPlayers.length; i++) {
						const player = sortedPlayers[i];
						const ranking = i + 1;
						const correctAnswers = player.answers.filter(a => a.isCorrect).length;
						const totalQuestions = quiz.questions.length;
						
						// Create question-by-question results as a simple string format
						// Format: "Q1:✓,Q2:✗,Q3:-,Q4:✓,Q5:✗" (✓=correct, ✗=wrong, -=not answered)
						const questionResults = quiz.questions.map((question, qIndex) => {
							let answer = player.answers.find(a => a.questionId === question.id);
							if (!answer && player.answers[qIndex]) {
								answer = player.answers[qIndex];
							}
							
							let result = '-'; // Not answered
							if (answer) {
								result = answer.isCorrect ? '✓' : '✗';
							}
							
							return `Q${qIndex + 1}:${result}`;
						}).join(',');
						
						await resultsTable.insertRow({
							QuizId: quiz.id,
							GameCode: code,
							QuizTitle: quiz.title,
							PlayerId: player.id,
							PlayerName: player.name,
							Avatar: player.avatar || 'default',
							Ranking: ranking.toString(),
							Score: player.score.toString(),
							CorrectAnswers: correctAnswers.toString(),
							TotalQuestions: totalQuestions.toString(),
							QuestionResults: questionResults,
							CompletedDate: new Date().toISOString(),
							MODIFIEDTIME: new Date()
						});
					}
					
					console.log(`Saved results for ${sortedPlayers.length} players to QuizResults table`);
				} catch (dbError) {
					console.error('Error saving quiz results to DataStore:', dbError);
					// Continue even if database save fails
				}
				
				return sendJSON(res, 200, { success: true, isFinished: true });
			}

			return sendJSON(res, 200, { success: true, isFinished: false });
		}

		// Handle /api/feedback endpoint - store player feedback
		if (normalizedPath === '/api/feedback' && method === 'POST') {
			try {
				const body = await parseBody(req);
				const { quizCode, playerId, playerName, rating, comment } = body || {};

				if (!quizCode || !playerId || !playerName || !rating) {
					return sendJSON(res, 400, { error: 'Missing required fields: quizCode, playerId, playerName, rating' });
				}

				// Persist to DataStore
				try {
					const catalyst = zcatalyst.initialize(req);
					const datastore = catalyst.datastore();
					const feedbackTable = datastore.table('quiz_feedback');

					// Insert only the columns that exist in the table
					// CREATEDTIME and MODIFIEDTIME are auto-managed by Catalyst
					await feedbackTable.insertRow({
						quizCode: quizCode.toString(),
						playerId: playerId.toString(),
						playerName: playerName.toString(),
						rating: rating.toString(), // store emoji/string
						comment: comment ? comment.toString() : ''
					});

					console.log('Feedback saved successfully:', { quizCode, playerId, playerName, rating, comment: comment || 'none' });
					return sendJSON(res, 200, { success: true });
				} catch (dbError) {
					console.error('Error saving feedback to DataStore:', dbError);
					console.error('Error details:', {
						message: dbError.message,
						stack: dbError.stack,
						quizCode,
						playerId,
						playerName,
						rating,
						comment: comment ? 'provided' : 'empty'
					});
					// Return error to client so they know the feedback wasn't saved
					// This will help identify if there's a table structure mismatch
					return sendJSON(res, 500, { 
						error: 'Failed to save feedback to database: ' + (dbError.message || String(dbError)),
						details: process.env.NODE_ENV === 'development' ? dbError.stack : undefined
					});
				}
			} catch (feedbackError) {
				console.error('Error handling feedback:', feedbackError);
				return sendJSON(res, 500, { error: 'Failed to submit feedback: ' + feedbackError.message });
			}
		}

		// Default route
		if (url === '/') {
			res.writeHead(200, { 'Content-Type': 'text/html' });
			res.write('<h1>Quiz API Server</h1>');
			res.end();
			return;
		}

		// 404 for unknown routes
		res.writeHead(404, { 'Content-Type': 'application/json' });
		res.write(JSON.stringify({ error: 'Endpoint not found' }));
		res.end();

	} catch (error) {
		console.error('Error:', error);
		sendJSON(res, 500, { error: 'Internal server error: ' + error.message });
	}
};
