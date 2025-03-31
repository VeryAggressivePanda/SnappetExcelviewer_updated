const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const excelRouter = require('./routes/excelRoutes');

const app = express();
const PORT = process.env.PORT || 3001;

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Set up multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
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

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
}); 