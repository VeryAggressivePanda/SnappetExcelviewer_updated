const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const excelRouter = require('./routes/excelRoutes');
const pdfRouter = require('./routes/pdfRoutes');

const app = express();
const PORT = process.env.PORT || 3001;

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '../views'));

// Middleware
app.use(express.static(path.join(__dirname, '../client')));
app.use('/assets', express.static(path.join(__dirname, '../assets')));
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Set up multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, '../uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    
    // No cleanup needed - we reuse the same filename
    
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // Reuse existing filename - no unnecessary timestamps
    cb(null, file.originalname);
  }
});



const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    // Log incoming file info for debugging
    console.log('Uploaded file:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      extension: path.extname(file.originalname).toLowerCase()
    });
    
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = path.extname(file.originalname).toLowerCase();
    
    if (validExtensions.includes(fileExtension)) {
      // If it has a valid extension, accept it
      return cb(null, true);
    } else {
      cb('Error: Only .xlsx or .xls files are allowed');
    }
  }
});

// Routes
app.use('/', excelRouter);
app.use('/pdf', pdfRouter);

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, '../uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Function to find available port
async function findAvailablePort(startPort) {
  const net = require('net');
  
  return new Promise((resolve) => {
    const server = net.createServer();
    
    server.listen(startPort, () => {
      const port = server.address().port;
      server.close(() => resolve(port));
    });
    
    server.on('error', () => {
      resolve(findAvailablePort(startPort + 1));
    });
  });
}

// Start server with automatic port resolution
async function startServer() {
  try {
    const availablePort = await findAvailablePort(PORT);
    
    if (availablePort !== PORT) {
      console.log(`Port ${PORT} is busy, using port ${availablePort} instead`);
    }
    
    app.listen(availablePort, () => {
      console.log(`Server running on http://localhost:${availablePort}`);
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

// Replace the existing app.listen call with:
startServer(); 