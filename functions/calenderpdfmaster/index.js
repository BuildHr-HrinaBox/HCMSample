'use strict';

const express = require('express');
const catalystSDK = require('zcatalyst-sdk-node');
const multer = require('multer');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());

// Attach catalyst instance per request
app.use((req, res, next) => {
  try {
    const catalyst = catalystSDK.initialize(req);
    res.locals.catalyst = catalyst;
    next();
  } catch (err) {
    res.status(500).send({ status: 'failure', message: 'Catalyst init failed' });
  }
});

// Table name for PDF files
const TABLE_NAME = 'CalenderPDFMaster';

// Get all PDF files
const getAllPDFFiles = async (catalyst) => {
  try {
    console.log('Fetching all PDF files from Data Store...');

    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);

    const result = await table.getAllRows();

    console.log(`Found ${result.length} PDF files`);

    const formattedData = result.map(record => ({
      id: record.ROWID,
      fileName: record.FileName || '',
      fileId: record.FileId || '',
      uploadedAt: record.UploadedAt || ''
    }));

    return {
      status: 'success',
      message: 'PDF files retrieved successfully',
      data: formattedData,
      count: formattedData.length
    };
  } catch (error) {
    console.error('Error fetching PDF files:', error);
    return {
      status: 'error',
      message: 'Failed to fetch PDF files',
      error: error.message
    };
  }
};

// Upload PDF file to File Store
const uploadPDFFile = async (catalyst, file, fileName) => {
  try {
    console.log('Uploading PDF file to File Store...');

    const fileStore = catalyst.filestore();
    const folder = fileStore.folder('PDFFile');

    const fileObj = await folder.uploadFile({
      code: file.buffer,
      name: fileName,
      contentType: 'application/pdf'
    });

    console.log('PDF file uploaded successfully with ID:', fileObj.id);

    return fileObj.id;
  } catch (error) {
    console.error('Error uploading PDF file:', error);
    throw error;
  }
};

// Add PDF file record to Data Store
const addPDFFileRecord = async (catalyst, fileName, fileId) => {
  try {
    console.log('Adding PDF file record to Data Store...');

    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);

    const record = {
      FileName: fileName,
      FileId: fileId,
      PdfFile: fileId, // Added PdfFile field as text
      UploadedAt: new Date().toISOString()
    };

    const result = await table.insertRow(record);

    console.log('PDF file record added successfully with ROWID:', result.ROWID);

    return result.ROWID;
  } catch (error) {
    console.error('Error adding PDF file record:', error);
    throw error;
  }
};

// Download PDF file from File Store
const downloadPDFFile = async (catalyst, fileId) => {
  try {
    console.log('Downloading PDF file from File Store...');

    const fileStore = catalyst.filestore();
    const fileObj = await fileStore.getFile(fileId);

    const fileData = await fileObj.download();

    console.log('PDF file downloaded successfully');

    return {
      buffer: fileData,
      fileName: fileObj.name,
      contentType: fileObj.content_type
    };
  } catch (error) {
    console.error('Error downloading PDF file:', error);
    throw error;
  }
};

// Delete PDF file from File Store and Data Store
const deletePDFFile = async (catalyst, id) => {
  try {
    console.log(`Deleting PDF file for ID: ${id}`);

    const dataStore = catalyst.datastore();
    const table = dataStore.table(TABLE_NAME);

    // Get the record first to get fileId
    const record = await table.getRow(id);
    if (!record) {
      return {
        status: 'error',
        message: 'PDF file record not found'
      };
    }

    const fileId = record.FileId;

    // Delete from File Store
    if (fileId) {
      const fileStore = catalyst.filestore();
      const fileObj = await fileStore.getFile(fileId);
      await fileObj.delete();
    }

    // Delete from Data Store
    await table.deleteRow(id);

    console.log('PDF file deleted successfully');

    return {
      status: 'success',
      message: 'PDF file deleted successfully'
    };
  } catch (error) {
    console.error('Error deleting PDF file:', error);
    return {
      status: 'error',
      message: 'Failed to delete PDF file',
      error: error.message
    };
  }
};

// Health check
app.get('/', (req, res) => {
  res.status(200).json({ status: 'success', message: 'calendar-pdf-master_function ready' });
});

// Get all PDF files
app.get('/calendar-pdf-master', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;

    console.log(`Calendar PDF Master API Request - Action: ${action}`);

    let result;

    switch (action) {
      case 'getAll':
        result = await getAllPDFFiles(catalyst);
        break;

      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: getAll'
        };
    }

    res.status(result.status === 'success' ? 200 : 400).json(result);

  } catch (error) {
    console.error('Calendar PDF Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// Upload PDF file
app.post('/calendar-pdf-master', upload.single('pdfFile'), async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action } = req.query;

    console.log(`Calendar PDF Master API POST Request - Action: ${action}`);

    let result;

    switch (action) {
      case 'upload':
        if (!req.file) {
          result = {
            status: 'error',
            message: 'No PDF file provided'
          };
        } else {
          const fileId = await uploadPDFFile(catalyst, req.file, req.file.originalname);
          const recordId = await addPDFFileRecord(catalyst, req.file.originalname, fileId);
          result = {
            status: 'success',
            message: 'PDF file uploaded successfully',
            data: {
              id: recordId,
              fileName: req.file.originalname,
              fileId: fileId
            }
          };
        }
        break;

      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: upload'
        };
    }

    res.status(result.status === 'success' ? 200 : 400).json(result);

  } catch (error) {
    console.error('Calendar PDF Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// Download PDF file
app.get('/calendar-pdf-master/download/:fileId', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { fileId } = req.params;

    console.log(`Downloading PDF file with ID: ${fileId}`);

    const fileData = await downloadPDFFile(catalyst, fileId);

    res.setHeader('Content-Type', fileData.contentType);
    res.setHeader('Content-Disposition', `attachment; filename="${fileData.fileName}"`);
    res.send(fileData.buffer);

  } catch (error) {
    console.error('Calendar PDF Master Download Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

// DELETE route for deletions
app.delete('/calendar-pdf-master', async (req, res) => {
  try {
    const catalyst = res.locals.catalyst;
    const { action, id } = req.query;

    console.log(`Calendar PDF Master API DELETE Request - Action: ${action}, ID: ${id}`);

    let result;

    switch (action) {
      case 'delete':
        if (!id) {
          result = {
            status: 'error',
            message: 'ID parameter is required for delete action'
          };
        } else {
          result = await deletePDFFile(catalyst, id);
        }
        break;

      default:
        result = {
          status: 'error',
          message: 'Invalid action. Supported actions: delete'
        };
    }

    res.status(result.status === 'success' ? 200 : 400).json(result);

  } catch (error) {
    console.error('Calendar PDF Master API Error:', error);
    res.status(500).json({
      status: 'error',
      message: 'Internal server error',
      error: error.message
    });
  }
});

module.exports = app;
