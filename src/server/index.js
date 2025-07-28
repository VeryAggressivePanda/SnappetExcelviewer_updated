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
    
    // Clean up old uploads with the same name
    const cleanupUploads = () => {
      try {
        const files = fs.readdirSync(uploadDir);
        const fileBaseName = file.originalname;
        
        // Find files with matching base name
        const matchingFiles = files.filter(f => f.includes(fileBaseName));
        
        // Keep only the most recent 2 uploads of this file
        if (matchingFiles.length > 1) {
          // Sort by creation time (which is embedded in the filename)
          matchingFiles.sort((a, b) => {
            const timeA = parseInt(a.split('-')[0]);
            const timeB = parseInt(b.split('-')[0]);
            return timeB - timeA; // Descending order
          });
          
          // Delete all but the most recent file
          matchingFiles.slice(1).forEach(f => {
            fs.unlinkSync(path.join(uploadDir, f));
            console.log(`Deleted old upload: ${f}`);
          });
        }
      } catch (err) {
        console.error('Error during upload cleanup:', err);
      }
    };
    
    // Run cleanup
    cleanupUploads();
    
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

// Also add a routine to periodically clean up old files
const cleanupUploadDirectory = () => {
  const uploadDir = path.join(__dirname, '../uploads');
  if (!fs.existsSync(uploadDir)) return;
  
  try {
    const files = fs.readdirSync(uploadDir);
    
    // Group files by original name (removing timestamp prefix)
    const fileGroups = {};
    files.forEach(file => {
      const parts = file.split('-');
      if (parts.length < 2) return;
      
      const originalName = parts.slice(1).join('-');
      if (!fileGroups[originalName]) {
        fileGroups[originalName] = [];
      }
      fileGroups[originalName].push(file);
    });
    
    // For each group, keep only the most recent file
    Object.values(fileGroups).forEach(group => {
      if (group.length > 1) {
        // Sort by creation time (embedded in filename)
        group.sort((a, b) => {
          const timeA = parseInt(a.split('-')[0]);
          const timeB = parseInt(b.split('-')[0]);
          return timeB - timeA; // Descending order
        });
        
        // Delete all but the most recent file
        group.slice(1).forEach(file => {
          fs.unlinkSync(path.join(uploadDir, file));
          console.log(`Cleaned up old upload: ${file}`);
        });
      }
    });
  } catch (err) {
    console.error('Error during upload directory cleanup:', err);
  }
};

// Run cleanup on startup
cleanupUploadDirectory();

// Schedule cleanup every 6 hours
setInterval(cleanupUploadDirectory, 6 * 60 * 60 * 1000);

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