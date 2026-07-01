const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const pdfParse = require('pdf-parse');
const Papa = require('papaparse');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3001;

// تأكد من وجود مجلد uploads
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Middleware
app.use(cors({
  origin: [
    'http://localhost:3000',
    'http://localhost:3001',
    'https://stationery-shop-manager.netlify.app',
    'https://your-netlify-site.netlify.app'
  ],
  credentials: true
}));
app.use(express.json());
app.use(express.static('public'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads');
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
  fileFilter: (req, file, cb) => {
    const allowedMimes = [
      'image/jpeg',
      'image/png',
      'image/jpg',
      'application/pdf',
      'text/plain',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'text/csv',
    ];
    
    if (allowedMimes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type'), false);
    }
  },
});

// In-memory storage for products and bills
let products = [];
let bills = [];
let uploadedFiles = [];

// Routes

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'Server is running', timestamp: new Date() });
});

// Upload files
app.post('/api/upload', upload.array('files', 10), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const uploadedFileData = req.files.map((file) => ({
      id: Date.now() + Math.random(),
      originalName: file.originalname,
      filename: file.filename,
      path: file.path,
      mimetype: file.mimetype,
      size: file.size,
      uploadDate: new Date(),
    }));

    uploadedFiles.push(...uploadedFileData);

    res.json({
      success: true,
      files: uploadedFileData,
      message: 'Files uploaded successfully',
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Extract data from files
app.post('/api/extract', async (req, res) => {
  try {
    const { fileIds } = req.body;

    if (!fileIds || !Array.isArray(fileIds)) {
      return res.status(400).json({ error: 'Invalid file IDs' });
    }

    const extractedData = [];

    for (const fileId of fileIds) {
      const file = uploadedFiles.find((f) => f.id === fileId);
      if (!file) continue;

      let items = [];

      // Handle PDF
      if (file.mimetype === 'application/pdf') {
        try {
          const pdfBuffer = fs.readFileSync(file.path);
          const pdfData = await pdfParse(pdfBuffer);
          items = extractPdfData(pdfData.text);
        } catch (e) {
          console.error('PDF parsing error:', e);
        }
      }

      // Handle Excel
      else if (
        file.mimetype === 'application/vnd.ms-excel' ||
        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ) {
        try {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(file.path);
          items = extractExcelData(workbook);
        } catch (e) {
          console.error('Excel parsing error:', e);
        }
      }

      // Handle CSV
      else if (file.mimetype === 'text/csv') {
        try {
          const fileContent = fs.readFileSync(file.path, 'utf-8');
          items = extractCsvData(fileContent);
        } catch (e) {
          console.error('CSV parsing error:', e);
        }
      }

      // Handle Text
      else if (file.mimetype === 'text/plain') {
        try {
          const fileContent = fs.readFileSync(file.path, 'utf-8');
          items = extractTextData(fileContent);
        } catch (e) {
          console.error('Text parsing error:', e);
        }
      }

      if (items.length > 0) {
        extractedData.push({
          fileId,
          fileName: file.originalName,
          items,
          extractedAt: new Date(),
        });
      }
    }

    res.json({
      success: true,
      data: extractedData,
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Data extraction helper functions

function extractPdfData(text) {
  const lines = text.split('\n');
  const items = [];

  // Simple pattern matching for common bill/invoice formats
  const patterns = [
    /(?:product|item|description|name)[\s:]*([^\n]+)[\s:]*(?:price|cost)[\s:]*(\d+\.?\d*)/gi,
    /([a-zA-Z\s]+)\s+(?:EGP|egp|\$)?\s*(\d+\.?\d*)/gi,
  ];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const name = match[1].trim();
      const price = parseFloat(match[2]);

      if (name.length > 2 && !isNaN(price) && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: 1,
        });
      }
    }
  }

  return items.slice(0, 10);
}

function extractExcelData(workbook) {
  const items = [];
  const worksheet = workbook.worksheets[0];

  if (!worksheet) return items;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const cells = row.values;
    if (cells && cells.length >= 2) {
      const name = (cells[1] || '').toString().trim();
      const price = parseFloat(cells[2] || cells[3] || 0);
      const quantity = parseInt(cells[4] || cells[3] || 1);

      if (name && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: quantity || 1,
        });
      }
    }
  });

  return items.slice(0, 10);
}

function extractCsvData(csvContent) {
  const items = [];
  const parsed = Papa.parse(csvContent);
  const data = parsed.data;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row.length >= 2) {
      const name = (row[0] || '').toString().trim();
      const price = parseFloat(row[1] || row[2] || 0);
      const quantity = parseInt(row[3] || row[2] || 1);

      if (name && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: quantity || 1,
        });
      }
    }
  }

  return items.slice(0, 10);
}

function extractTextData(fileContent) {
  const items = [];
  const lines = fileContent.split('\n');

  for (const line of lines) {
    if (!line.trim()) continue;

    const match = line.match(/^(.+?)[\s:]+(\d+\.?\d*)\s*(?:EGP|egp|جنيه)?/);
    if (match) {
      const name = match[1].trim();
      const price = parseFloat(match[2]);

      if (name.length > 2 && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: 1,
        });
      }
    }
  }

  return items.slice(0, 10);
}

// Add product
app.post('/api/products', (req, res) => {
  try {
    const { name, wholesalePrice, quantity } = req.body;

    if (!name || !wholesalePrice || !quantity) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const product = {
      id: Date.now(),
      name,
      wholesalePrice: parseFloat(wholesalePrice),
      quantity: parseInt(quantity),
      sellingPrice: (parseFloat(wholesalePrice) * 1.2).toFixed(2),
      dateAdded: new Date(),
    };

    products.push(product);
    res.json(product);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get all products
app.get('/api/products', (req, res) => {
  res.json({
    success: true,
    count: products.length,
    products,
    totalValue: products.reduce((sum, p) => sum + p.wholesalePrice * p.quantity, 0),
  });
});

// Update product
app.put('/api/products/:id', (req, res) => {
  try {
    const productId = parseInt(req.params.id);
    const { name, wholesalePrice, quantity } = req.body;

    const product = products.find((p) => p.id === productId);
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }

    if (name) product.name = name;
    if (wholesalePrice) {
      product.wholesalePrice = parseFloat(wholesalePrice);
      product.sellingPrice = (product.wholesalePrice * 1.2).toFixed(2);
    }
    if (quantity) product.quantity = parseInt(quantity);

    res.json(product);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Delete product
app.delete('/api/products/:id', (req, res) => {
  try {
    const productId = parseInt(req.params.id);
    products = products.filter((p) => p.id !== productId);
    res.json({ success: true, message: 'Product deleted' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get uploaded files
app.get('/api/files', (req, res) => {
  res.json({
    success: true,
    count: uploadedFiles.length,
    files: uploadedFiles.map((f) => ({
      id: f.id,
      name: f.originalName,
      size: f.size,
      uploadDate: f.uploadDate,
    })),
  });
});

// Delete file
app.delete('/api/files/:id', (req, res) => {
  try {
    const file = uploadedFiles.find((f) => f.id === parseInt(req.params.id));
    if (file && fs.existsSync(file.path)) {
      fs.unlinkSync(file.path);
    }
    uploadedFiles = uploadedFiles.filter((f) => f.id !== parseInt(req.params.id));
    res.json({ success: true, message: 'File deleted' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Export products as JSON
app.get('/api/export/products', (req, res) => {
  res.json({
    exportDate: new Date(),
    totalProducts: products.length,
    products,
    totalInventoryValue: products.reduce((sum, p) => sum + p.wholesalePrice * p.quantity, 0),
  });
});

// Export products as CSV
app.get('/api/export/products/csv', (req, res) => {
  const csv = [
    ['Product Name', 'Wholesale Price (EGP)', 'Quantity', 'Selling Price (EGP)', 'Total Value'],
    ...products.map((p) => [
      p.name,
      p.wholesalePrice,
      p.quantity,
      p.sellingPrice,
      p.wholesalePrice * p.quantity,
    ]),
  ]
    .map((row) => row.join(','))
    .join('\n');

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="products.csv"');
  res.send(csv);
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: err.message });
});

// تشغيل الخادم على Railway
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
  console.log('✅ Stationery Shop Manager API is ready');
  console.log(`✅ Uploads directory: ${path.resolve('uploads')}`);
});
    const extractedData = [];

    for (const fileId of fileIds) {
      const file = uploadedFiles.find((f) => f.id === fileId);
      if (!file) continue;

      let items = [];

      // Handle PDF
      if (file.mimetype === 'application/pdf') {
        try {
          const pdfBuffer = fs.readFileSync(file.path);
          const pdfData = await pdfParse(pdfBuffer);
          items = extractPdfData(pdfData.text);
        } catch (e) {
          console.error('PDF parsing error:', e);
        }
      }

      // Handle Excel
      else if (
        file.mimetype === 'application/vnd.ms-excel' ||
        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ) {
        try {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(file.path);
          items = extractExcelData(workbook);
        } catch (e) {
          console.error('Excel parsing error:', e);
        }
      }

      // Handle CSV
      else if (file.mimetype === 'text/csv') {
        try {
          const fileContent = fs.readFileSync(file.path, 'utf-8');
          items = extractCsvData(fileContent);
        } catch (e) {
          console.error('CSV parsing error:', e);
        }
      }

      // Handle Text
      else if (file.mimetype === 'text/plain') {
        try {
          const fileContent = fs.readFileSync(file.path, 'utf-8');
          items = extractTextData(fileContent);
        } catch (e) {
          console.error('Text parsing error:', e);
        }
      }

      if (items.length > 0) {
        extractedData.push({
          fileId,
          fileName: file.originalName,
          items,
          extractedAt: new Date(),
        });
      }
    }

    res.json({
      success: true,
      data: extractedData,
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Data extraction helper functions

function extractPdfData(text) {
  const lines = text.split('\n');
  const items = [];

  // Simple pattern matching for common bill/invoice formats
  const patterns = [
    /(?:product|item|description|name)[\s:]*([^\n]+)[\s:]*(?:price|cost)[\s:]*(\d+\.?\d*)/gi,
    /([a-zA-Z\s]+)\s+(?:EGP|egp|\$)?\s*(\d+\.?\d*)/gi,
  ];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const name = match[1].trim();
      const price = parseFloat(match[2]);

      if (name.length > 2 && !isNaN(price) && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: 1,
        });
      }
    }
  }

  return items.slice(0, 10); // Limit to 10 items
}

function extractExcelData(workbook) {
  const items = [];
  const worksheet = workbook.worksheets[0];

  if (!worksheet) return items;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header

    const cells = row.values;
    if (cells && cells.length >= 2) {
      const name = (cells[1] || '').toString().trim();
      const price = parseFloat(cells[2] || cells[3] || 0);
      const quantity = parseInt(cells[4] || cells[3] || 1);

      if (name && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: quantity || 1,
        });
      }
    }
  });

  return items.slice(0, 10);
}

function extractCsvData(csvContent) {
  const items = [];
  const parsed = Papa.parse(csvContent);
  const data = parsed.data;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row.length >= 2) {
      const name = (row[0] || '').toString().trim();
      const price = parseFloat(row[1] || row[2] || 0);
      const quantity = parseInt(row[3] || row[2] || 1);

      if (name && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: quantity || 1,
        });
      }
    }
  }

  return items.slice(0, 10);
}

function extractTextData(fileContent) {
  const items = [];
  const lines = fileContent.split('\n');

  for (const line of lines) {
    if (!line.trim()) continue;

    // Match patterns like "Product Name: Price" or "Product Name Price"
    const match = line.match(/^(.+?)[\s:]+(\d+\.?\d*)\s*(?:EGP|egp|جنيه)?/);
    if (match) {
      const name = match[1].trim();
      const price = parseFloat(match[2]);

      if (name.length > 2 && price > 0) {
        items.push({
          name,
          wholesalePrice: price.toFixed(2),
          quantity: 1,
        });
      }
    }
  }

  return items.slice(0, 10);
}

// Add product
app.post('/api/products', (req, res) => {
  try {
    const { name, wholesalePrice, quantity } = req.body;

    if (!name || !wholesalePrice || !quantity) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const product = {
      id: Date.now(),
      name,
      wholesalePrice: parseFloat(wholesalePrice),
      quantity: parseInt(quantity),
      sellingPrice: (parseFloat(wholesalePrice) * 1.2).toFixed(2),
      dateAdded: new Date(),
    };

    products.push(product);
    res.json(product);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get all products
app.get('/api/products', (req, res) => {
  res.json({
    success: true,
    count: products.length,
    products,
    totalValue: products.reduce((sum, p) => sum + p.wholesalePrice * p.quantity, 0),
  });
});

// Update product
app.put('/api/products/:id', (req, res) => {
  try {
    const productId = parseInt(req.params.id);
    const { name, wholesalePrice, quantity } = req.body;

    const product = products.find((p) => p.id === productId);
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }

    if (name) product.name = name;
    if (wholesalePrice) {
      product.wholesalePrice = parseFloat(wholesalePrice);
      product.sellingPrice = (product.wholesalePrice * 1.2).toFixed(2);
    }
    if (quantity) product.quantity = parseInt(quantity);

    res.json(product);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Delete product
app.delete('/api/products/:id', (req, res) => {
  try {
    const productId = parseInt(req.params.id);
    products = products.filter((p) => p.id !== productId);
    res.json({ success: true, message: 'Product deleted' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Get uploaded files
app.get('/api/files', (req, res) => {
  res.json({
    success: true,
    count: uploadedFiles.length,
    files: uploadedFiles.map((f) => ({
      id: f.id,
      name: f.originalName,
      size: f.size,
      uploadDate: f.uploadDate,
    })),
  });
});

// Delete file
app.delete('/api/files/:id', (req, res) => {
  try {
    const file = uploadedFiles.find((f) => f.id === parseInt(req.params.id));
    if (file && fs.existsSync(file.path)) {
      fs.unlinkSync(file.path);
    }
    uploadedFiles = uploadedFiles.filter((f) => f.id !== parseInt(req.params.id));
    res.json({ success: true, message: 'File deleted' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Export products as JSON
app.get('/api/export/products', (req, res) => {
  res.json({
    exportDate: new Date(),
    totalProducts: products.length,
    products,
    totalInventoryValue: products.reduce((sum, p) => sum + p.wholesalePrice * p.quantity, 0),
  });
});

// Export products as CSV
app.get('/api/export/products/csv', (req, res) => {
  const csv = [
    ['Product Name', 'Wholesale Price (EGP)', 'Quantity', 'Selling Price (EGP)', 'Total Value'],
    ...products.map((p) => [
      p.name,
      p.wholesalePrice,
      p.quantity,
      p.sellingPrice,
      p.wholesalePrice * p.quantity,
    ]),
  ]
    .map((row) => row.join(','))
    .join('\n');

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="products.csv"');
  res.send(csv);
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: err.message });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Stationery Shop Manager API is ready');
});
