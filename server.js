import stringSimilarity from 'string-similarity';
import express from "express";
import multer from "multer";
import fetch from "node-fetch";
import fs from "fs";
import sharp from "sharp";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import ExcelJS from "exceljs";
import path from "path";
import dotenv from "dotenv";
import { google } from 'googleapis';
import passport from 'passport';
import { Strategy as GoogleStrategy } from 'passport-google-oauth20';
import session from 'express-session';
import os from 'os'

dotenv.config();

const app = express();
const port = 3000;

app.use(express.json());
app.use(express.static("public"));

app.use(session({
  secret: process.env.SESSION_SECRET || 'autoatt-session-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { 
    secure: false, // Set to true in production with HTTPS
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));


app.use(passport.initialize());
app.use(passport.session());

passport.use(new GoogleStrategy({
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: process.env.GOOGLE_REDIRECT_URI
}, async (accessToken, refreshToken, profile, done) => {
  try {
    console.log('Google OAuth callback for:', profile.emails[0].value);
    
    // Check if user exists
    let user = users.find(u => u.email === profile.emails[0].value);
    
    if (!user) {
      // Create new user
      user = {
        userId: Date.now().toString(),
        username: profile.emails[0].value.split('@')[0],
        email: profile.emails[0].value,
        fullName: profile.displayName,
        subject: '',
        googleId: profile.id,
        avatar: profile.photos[0]?.value,
        createdAt: new Date().toISOString(),
        workbookPath: null,
        studentDatabase: [],
        workbookStructure: null,
        googleTokens: {
          accessToken: accessToken,
          refreshToken: refreshToken
        }
      };
      
      users.push(user);
      saveUsers();
      console.log('Created new Google user:', user.email);
    } else {
      // Update existing user with Google info
      user.googleId = profile.id;
      user.avatar = profile.photos[0]?.value;
      user.googleTokens = {
        accessToken: accessToken,
        refreshToken: refreshToken
      };
      saveUsers();
      console.log('Updated existing user with Google auth:', user.email);
    }
    
    return done(null, user);
  } catch (error) {
    console.error('Google OAuth error:', error);
    return done(error, null);
  }
}));

passport.serializeUser((user, done) => {
  done(null, user.userId);
});

passport.deserializeUser((userId, done) => {
  const user = users.find(u => u.userId === userId);
  done(null, user);
});

const googleClientId = process.env.GOOGLE_CLIENT_ID;
const googleClientSecret = process.env.GOOGLE_CLIENT_SECRET;
const googleCallbackURL = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:3000/auth/google/callback';

if (googleClientId && googleClientSecret) {
  console.log('Configuring Google OAuth Strategy with persistent login...');
  
  passport.use(new GoogleStrategy({
    clientID: googleClientId,
    clientSecret: googleClientSecret,
    callbackURL: googleCallbackURL,
    accessType: 'offline', // Get refresh token
    prompt: 'select_account' // Only show account selector, not consent screen again
  }, async (accessToken, refreshToken, profile, done) => {
    try {
      console.log('Google OAuth callback for:', profile.emails[0].value);
      
      let user = users.find(u => u.email === profile.emails[0].value);
      
      if (!user) {
        // Create new user
        user = {
          userId: Date.now().toString(),
          username: profile.emails[0].value.split('@')[0],
          email: profile.emails[0].value,
          fullName: profile.displayName,
          subject: '',
          googleId: profile.id,
          avatar: profile.photos[0]?.value,
          createdAt: new Date().toISOString(),
          workbookPath: null,
          studentDatabase: [],
          workbookStructure: null,
          googleTokens: {
            accessToken: accessToken,
            refreshToken: refreshToken || user?.googleTokens?.refreshToken // Preserve old refresh token if not provided
          },
          googleConnected: true,
          lastGoogleLogin: new Date().toISOString()
        };
        
        users.push(user);
        saveUsers();
        console.log('Created new Google user:', user.email);
      } else {
        // Update existing user
        user.googleId = profile.id;
        user.avatar = profile.photos[0]?.value;
        user.fullName = profile.displayName;
        
        // Update tokens - preserve refresh token if new one not provided
        user.googleTokens = {
          accessToken: accessToken,
          refreshToken: refreshToken || user.googleTokens?.refreshToken
        };
        
        user.googleConnected = true;
        user.lastGoogleLogin = new Date().toISOString();
        
        saveUsers();
        console.log('Updated existing user with Google auth:', user.email);
      }
      
      return done(null, user);
    } catch (error) {
      console.error('Google OAuth error:', error);
      return done(error, null);
    }
  }));

  passport.serializeUser((user, done) => {
    done(null, user.userId);
  });

  passport.deserializeUser((userId, done) => {
    const user = users.find(u => u.userId === userId);
    done(null, user);
  });
} else {
  console.error('Google OAuth credentials missing');
}

// 2. Enhanced Google OAuth route - use select_account instead of consent
app.get('/auth/google', (req, res, next) => {
  if (!process.env.GOOGLE_CLIENT_ID || !process.env.GOOGLE_CLIENT_SECRET) {
    return res.status(500).send(`
      <html>
        <head><title>Configuration Error</title></head>
        <body style="font-family: Arial; padding: 50px; text-align: center;">
          <h2>Google OAuth Not Configured</h2>
          <p>The server needs Google OAuth credentials.</p>
          <button onclick="window.location.href='/'">Back to Login</button>
        </body>
      </html>
    `);
  }
  
  console.log('Starting Google OAuth flow...');
  
  // Use select_account for persistent login experience
  passport.authenticate('google', { 
    scope: [
      'profile', 
      'email', 
      'https://www.googleapis.com/auth/drive.file',
      'https://www.googleapis.com/auth/drive.readonly'
    ],
    accessType: 'offline',
    prompt: 'select_account' // Changed from 'consent' to 'select_account'
  })(req, res, next);
});

// 3. Enhanced callback with better session handling
app.get('/auth/google/callback', 
  (req, res, next) => {
    console.log('Received Google OAuth callback');
    passport.authenticate('google', { 
      failureRedirect: '/?error=auth_failed',
      failureMessage: true
    })(req, res, next);
  },
  (req, res) => {
    try {
      if (!req.user) {
        console.error('No user in callback');
        return res.redirect('/?error=no_user');
      }
      
      console.log('Google OAuth successful for:', req.user.email);
      
      // Generate long-lived JWT token (30 days)
      const token = jwt.sign(
        { 
          userId: req.user.userId, 
          username: req.user.username,
          googleConnected: true
        }, 
        JWT_SECRET, 
        { expiresIn: '30d' } // Extended expiry for persistent login
      );
      
      const userData = {
        userId: req.user.userId,
        username: req.user.username,
        fullName: req.user.fullName,
        email: req.user.email,
        avatar: req.user.avatar,
        hasWorkbook: !!req.user.workbookPath,
        googleConnected: true
      };
      
      const redirectUrl = `/?token=${token}&user=${encodeURIComponent(JSON.stringify(userData))}`;
      console.log('Redirecting with persistent token');
      
      res.redirect(redirectUrl);
    } catch (error) {
      console.error('Callback error:', error);
      res.redirect('/?error=callback_failed');
    }
  }
);

app.post('/api/auth/refresh-google-token', authenticateToken, async (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    
    if (!user || !user.googleTokens || !user.googleTokens.refreshToken) {
      return res.status(400).json({ 
        error: 'No refresh token available',
        message: 'Please re-authenticate with Google'
      });
    }
    
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      process.env.GOOGLE_REDIRECT_URI
    );
    
    oauth2Client.setCredentials({
      refresh_token: user.googleTokens.refreshToken
    });
    
    const { credentials } = await oauth2Client.refreshAccessToken();
    
    // Update user's access token
    user.googleTokens.accessToken = credentials.access_token;
    if (credentials.refresh_token) {
      user.googleTokens.refreshToken = credentials.refresh_token;
    }
    
    saveUsers();
    
    res.json({
      success: true,
      message: 'Token refreshed successfully'
    });
    
  } catch (error) {
    console.error('Token refresh error:', error);
    res.status(500).json({ 
      error: 'Failed to refresh token',
      details: error.message 
    });
  }
});

// 7. Google Drive integration functions
async function getGoogleDriveAuth(user) {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  
  if (user.googleTokens) {
    oauth2Client.setCredentials({
      access_token: user.googleTokens.accessToken,
      refresh_token: user.googleTokens.refreshToken
    });
    
    // Set up automatic token refresh
    oauth2Client.on('tokens', (tokens) => {
      if (tokens.refresh_token) {
        user.googleTokens.refreshToken = tokens.refresh_token;
      }
      user.googleTokens.accessToken = tokens.access_token;
      saveUsers();
      console.log('Google tokens auto-refreshed for:', user.email);
    });
  }
  
  return oauth2Client;
}

// 6. Check Google connection status endpoint
app.get('/api/auth/google-status', authenticateToken, (req, res) => {
  const user = users.find(u => u.userId === req.user.userId);
  
  if (!user) {
    return res.status(404).json({ error: 'User not found' });
  }
  
  res.json({
    googleConnected: !!user.googleConnected,
    hasRefreshToken: !!(user.googleTokens && user.googleTokens.refreshToken),
    email: user.email || null,
    avatar: user.avatar || null,
    lastLogin: user.lastGoogleLogin || null
  });
});

app.get('/api/drive/files', authenticateToken, async (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    if (!user || !user.googleTokens) {
      return res.status(400).json({ 
        error: 'Google Drive not connected',
        details: 'Please login with Google to access Drive files'
      });
    }
    
    const auth = await getGoogleDriveAuth(user);
    const drive = google.drive({ version: 'v3', auth });
    
    const response = await drive.files.list({
      q: "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel'",
      fields: 'files(id, name, modifiedTime, size, webViewLink)',
      orderBy: 'modifiedTime desc',
      pageSize: 50
    });
    
    // Mark which files are already uploaded
    const files = (response.data.files || []).map(file => {
      const isUploaded = user.driveFileId === file.id;
      return {
        ...file,
        isUploaded: isUploaded,
        uploadedAt: isUploaded ? user.updatedAt : null
      };
    });
    
    res.json({
      files: files,
      currentWorkbook: user.driveFileId ? {
        id: user.driveFileId,
        name: user.driveFileName
      } : null,
      message: `Found ${files.length} Excel files in your Drive`
    });
    
  } catch (error) {
    console.error('Drive files error:', error);
    res.status(500).json({ 
      error: 'Failed to fetch Drive files',
      details: error.message 
    });
  }
});

// 9. Download and process Excel file from Google Drive
app.post('/api/drive/process-file', authenticateToken, async (req, res) => {
  try {
    const { fileId, fileName } = req.body;
    
    if (!fileId) {
      return res.status(400).json({ error: 'File ID is required' });
    }
    
    const user = users.find(u => u.userId === req.user.userId);
    if (!user || !user.googleTokens) {
      return res.status(400).json({ 
        error: 'Google Drive not connected' 
      });
    }
    
    console.log(`Processing Drive file: ${fileName} (${fileId}) for ${user.username}`);
    
    const auth = await getGoogleDriveAuth(user);
    const drive = google.drive({ version: 'v3', auth });
    
    // Download file from Drive
    const response = await drive.files.get({
      fileId: fileId,
      alt: 'media'
    }, { responseType: 'stream' });
    
    // Save to local workbooks directory
    const workbooksDir = 'workbooks';
    if (!fs.existsSync(workbooksDir)) {
      fs.mkdirSync(workbooksDir, { recursive: true });
    }
    
    const localFileName = `${user.userId}_workbook_${Date.now()}.xlsx`;
    const localPath = path.join(workbooksDir, localFileName);
    
    const writeStream = fs.createWriteStream(localPath);
    
    return new Promise((resolve, reject) => {
      response.data.pipe(writeStream);
      
      writeStream.on('finish', async () => {
        try {
          console.log('Drive file downloaded successfully');
          
          // Process the workbook
          const workbookData = await parseWorkbook(localPath);
          
          // Clean up old workbook
          if (user.workbookPath && fs.existsSync(user.workbookPath)) {
            try {
              fs.unlinkSync(user.workbookPath);
            } catch (cleanupError) {
              console.warn('Could not delete old workbook:', cleanupError.message);
            }
          }
          
          // Update user data
          user.workbookPath = localPath;
          user.studentDatabase = workbookData.allStudents;
          user.workbookStructure = workbookData.sheets;
          user.updatedAt = new Date().toISOString();
          user.driveFileId = fileId; // Store for future updates
          user.driveFileName = fileName;
          
          saveUsers();
          
          res.json({
            message: 'Workbook processed successfully from Google Drive',
            totalStudents: workbookData.totalStudents,
            sheetsFound: workbookData.sheetsFound,
            sheets: Object.keys(workbookData.sheets).map(name => ({
              name,
              studentCount: workbookData.sheets[name].studentCount,
              sampleStudents: workbookData.sheets[name].students.slice(0, 3),
            })),
            source: 'googleDrive',
            driveFile: {
              id: fileId,
              name: fileName
            }
          });
          
        } catch (processError) {
          console.error('Failed to process Drive file:', processError);
          // Clean up downloaded file on error
          if (fs.existsSync(localPath)) {
            fs.unlinkSync(localPath);
          }
          res.status(500).json({ 
            error: 'Failed to process workbook from Drive',
            details: processError.message 
          });
        }
      });
      
      writeStream.on('error', (error) => {
        console.error('Write stream error:', error);
        res.status(500).json({ 
          error: 'Failed to download file from Drive',
          details: error.message 
        });
      });
    });
    
  } catch (error) {
    console.error('Drive process file error:', error);
    res.status(500).json({ 
      error: 'Failed to process Drive file',
      details: error.message 
    });
  }
});

app.get('/api/workbook/status', authenticateToken, (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    if (!user.workbookPath) {
      return res.json({
        hasWorkbook: false,
        message: 'No workbook uploaded'
      });
    }
    
    const fileExists = fs.existsSync(user.workbookPath);
    
    res.json({
      hasWorkbook: true,
      workbookPath: user.workbookPath,
      fileExists: fileExists,
      studentCount: user.studentDatabase?.length || 0,
      sheetCount: user.workbookStructure ? Object.keys(user.workbookStructure).length : 0,
      uploadedAt: user.updatedAt,
      source: user.driveFileId ? 'googleDrive' : 'localUpload',
      driveInfo: user.driveFileId ? {
        fileId: user.driveFileId,
        fileName: user.driveFileName
      } : null
    });
    
  } catch (error) {
    console.error('Workbook status error:', error);
    res.status(500).json({ 
      error: 'Failed to get workbook status',
      details: error.message 
    });
  }
});

// 10. Upload updated workbook back to Google Drive
app.post('/api/drive/upload-updated', authenticateToken, async (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    if (!user || !user.googleTokens) {
      return res.status(400).json({ 
        error: 'Google Drive not connected' 
      });
    }
    
    if (!user.workbookPath || !fs.existsSync(user.workbookPath)) {
      return res.status(400).json({ 
        error: 'No workbook to upload' 
      });
    }
    
    const auth = await getGoogleDriveAuth(user);
    const drive = google.drive({ version: 'v3', auth });
    
    const updatedFileName = `${user.driveFileName || 'attendance'}_updated_${new Date().toISOString().slice(0, 10)}.xlsx`;
    
    const fileMetadata = {
      name: updatedFileName,
    };
    
    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: fs.createReadStream(user.workbookPath),
    };
    
    const response = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id, name, webViewLink',
    });
    
    res.json({
      message: 'Updated workbook uploaded to Google Drive',
      driveFile: {
        id: response.data.id,
        name: response.data.name,
        link: response.data.webViewLink
      }
    });
    
  } catch (error) {
    console.error('Drive upload error:', error);
    res.status(500).json({ 
      error: 'Failed to upload to Drive',
      details: error.message 
    });
  }
});

app.delete('/api/workbook/delete', authenticateToken, (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    if (!user.workbookPath) {
      return res.status(404).json({ 
        error: 'No workbook to delete',
        message: 'You do not have an uploaded workbook'
      });
    }
    
    const workbookPath = user.workbookPath;
    
    // Delete the physical file
    if (fs.existsSync(workbookPath)) {
      try {
        fs.unlinkSync(workbookPath);
        console.log(`Deleted workbook file: ${workbookPath}`);
      } catch (fileError) {
        console.error('Error deleting file:', fileError);
        return res.status(500).json({ 
          error: 'Failed to delete workbook file',
          details: fileError.message
        });
      }
    }
    
    // Clear user's workbook data
    const oldData = {
      workbookPath: user.workbookPath,
      studentCount: user.studentDatabase?.length || 0,
      sheetCount: user.workbookStructure ? Object.keys(user.workbookStructure).length : 0
    };
    
    user.workbookPath = null;
    user.studentDatabase = [];
    user.workbookStructure = null;
    user.driveFileId = null;
    user.driveFileName = null;
    user.updatedAt = new Date().toISOString();
    
    saveUsers();
    
    res.json({
      success: true,
      message: 'Workbook deleted successfully',
      deletedData: oldData
    });
    
  } catch (error) {
    console.error('Delete workbook error:', error);
    res.status(500).json({ 
      error: 'Failed to delete workbook',
      details: error.message 
    });
  }
});

const JWT_SECRET = process.env.JWT_SECRET || "autoatt-secret-key-2024";

// User storage (in production, use a proper database)
let users = [];

// Load/Save users
function loadUsers() {
  try {
    if (fs.existsSync('users.json')) {
      const data = fs.readFileSync('users.json', 'utf8');
      users = JSON.parse(data);
      console.log(`👥 Loaded ${users.length} user accounts`);
    }
  } catch (error) {
    console.warn('Could not load users:', error.message);
  }
}

function initializeServer() {
  loadUsers();
  cleanupCorruptedUserData(); // Clean up any corrupted data on startup
}

function saveUsers() {
  try {
    fs.writeFileSync('users.json', JSON.stringify(users, null, 2));
  } catch (error) {
    console.error('Could not save users:', error.message);
  }
}

// Authentication middleware
function authenticateToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Access token required' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Invalid or expired token' });
    }
    req.user = user;
    next();
  });
}

// File upload setup
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadType = req.body.uploadType || req.query.uploadType;
    const dir = uploadType === 'workbook' ? 'workbooks/' : 'uploads/';
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const userId = req.user ? req.user.userId : Date.now();
    const timestamp = Date.now();
    const uploadType = req.body.uploadType || req.query.uploadType;
    
    if (uploadType === 'workbook') {
      const ext = path.extname(file.originalname);
      cb(null, `${userId}_workbook_${timestamp}${ext}`);
    } else {
      cb(null, `${userId}_photo_${timestamp}.jpg`);
    }
  },
});

// Enhanced file upload setup with better path separation
const enhancedStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadType = req.body.uploadType || req.query.uploadType || 'unknown';
    
    // Ensure we use the correct directory based on upload type
    let dir;
    if (uploadType === 'workbook' || file.fieldname === 'workbook') {
      dir = 'workbooks/';
    } else if (file.fieldname === 'photo') {
      dir = 'uploads/';
    } else {
      // Fallback based on file extension
      const ext = path.extname(file.originalname).toLowerCase();
      if (['.xlsx', '.xls'].includes(ext)) {
        dir = 'workbooks/';
      } else {
        dir = 'uploads/';
      }
    }
    
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    
    console.log(`Upload destination: ${dir} (type: ${uploadType}, field: ${file.fieldname})`);
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const userId = req.user ? req.user.userId : Date.now();
    const timestamp = Date.now();
    const uploadType = req.body.uploadType || req.query.uploadType || 'unknown';
    
    let filename;
    if (uploadType === 'workbook' || file.fieldname === 'workbook') {
      const ext = path.extname(file.originalname);
      filename = `${userId}_workbook_${timestamp}${ext}`;
    } else if (file.fieldname === 'photo') {
      filename = `${userId}_photo_${timestamp}.jpg`;
    } else {
      // Fallback based on file extension
      const ext = path.extname(file.originalname).toLowerCase();
      if (['.xlsx', '.xls'].includes(ext)) {
        filename = `${userId}_workbook_${timestamp}${ext}`;
      } else {
        filename = `${userId}_photo_${timestamp}.jpg`;
      }
    }
    
    console.log(`Generated filename: ${filename} (type: ${uploadType}, field: ${file.fieldname})`);
    cb(null, filename);
  },
});

const upload = multer({ 
  storage: enhancedStorage,
  limits: { fileSize: 15 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    console.log(`File filter: ${file.fieldname} - ${file.originalname} - ${file.mimetype}`);
    
    if (file.fieldname === 'workbook') {
      const ext = path.extname(file.originalname).toLowerCase();
      if (['.xlsx', '.xls'].includes(ext)) {
        cb(null, true);
      } else {
        cb(new Error('Only Excel files (.xlsx, .xls) are allowed for workbooks'));
      }
    } else if (file.fieldname === 'photo') {
      if (file.mimetype.startsWith('image/')) {
        cb(null, true);
      } else {
        cb(new Error('Only image files are allowed for photos'));
      }
    } else {
      cb(new Error('Unknown file field'));
    }
  }
});

function cleanupCorruptedUserData() {
  console.log('🔧 Starting user data cleanup...');
  
  let cleanupCount = 0;
  
  users.forEach(user => {
    let userModified = false;
    
    // Fix corrupted workbook paths pointing to photo files
    if (user.workbookPath && user.workbookPath.includes('_photo_')) {
      console.log(`Fixing corrupted workbook path for user ${user.username}: ${user.workbookPath}`);
      
      // Try to find the actual workbook file for this user
      try {
        const workbooksDir = 'workbooks';
        if (fs.existsSync(workbooksDir)) {
          const files = fs.readdirSync(workbooksDir);
          const userWorkbook = files.find(file => 
            file.startsWith(`${user.userId}_workbook_`) && 
            (file.endsWith('.xlsx') || file.endsWith('.xls'))
          );
          
          if (userWorkbook) {
            const correctPath = path.join(workbooksDir, userWorkbook);
            if (fs.existsSync(correctPath)) {
              user.workbookPath = correctPath;
              console.log(`Fixed workbook path: ${correctPath}`);
              userModified = true;
            }
          } else {
            // No workbook found, reset to null
            user.workbookPath = null;
            user.studentDatabase = [];
            user.workbookStructure = null;
            console.log(`No workbook found, reset data for user ${user.username}`);
            userModified = true;
          }
        }
      } catch (error) {
        console.error(`Error fixing user ${user.username}:`, error.message);
      }
    }
    
    // Clean up other potentially corrupted data
    if (user.workbookPath && !fs.existsSync(user.workbookPath)) {
      console.log(`Removing non-existent workbook path for user ${user.username}: ${user.workbookPath}`);
      user.workbookPath = null;
      user.studentDatabase = [];
      user.workbookStructure = null;
      userModified = true;
    }
    
    if (userModified) {
      cleanupCount++;
    }
  });
  
  if (cleanupCount > 0) {
    saveUsers();
    console.log(`Cleaned up ${cleanupCount} user records`);
  } else {
    console.log('No user data cleanup needed');
  }
}

app.post('/api/admin/cleanup-users', authenticateToken, (req, res) => {
  try {
    cleanupCorruptedUserData();
    res.json({ 
      success: true, 
      message: 'User data cleanup completed',
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(500).json({ 
      success: false, 
      error: 'Cleanup failed: ' + error.message 
    });
  }
});
// === AUTHENTICATION ROUTES ===

app.post('/api/auth/register', async (req, res) => {
  try {
    const { username, email, password, fullName, subject } = req.body;

    if (!username || !email || !password || !fullName) {
      return res.status(400).json({ error: 'All fields are required' });
    }

    if (users.find(u => u.username === username || u.email === email)) {
      return res.status(400).json({ error: 'Username or email already exists' });
    }

    const hashedPassword = await bcrypt.hash(password, 10);

    const user = {
      userId: Date.now().toString(),
      username,
      email,
      fullName,
      subject: subject || '',
      password: hashedPassword,
      createdAt: new Date().toISOString(),
      workbookPath: null,
      studentDatabase: [],
      workbookStructure: null
    };

    users.push(user);
    saveUsers();

    const token = jwt.sign(
      { userId: user.userId, username }, 
      JWT_SECRET, 
      { expiresIn: '7d' }
    );

    res.json({
      message: 'Registration successful',
      token,
      user: {
        userId: user.userId,
        username: user.username,
        fullName: user.fullName,
        subject: user.subject,
        hasWorkbook: false
      }
    });
  } catch (error) {
    res.status(500).json({ error: 'Registration failed: ' + error.message });
  }
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    const user = users.find(u => u.username === username || u.email === username);
    if (!user) {
      return res.status(401).json({ error: 'Invalid username or password' });
    }

    const validPassword = await bcrypt.compare(password, user.password);
    if (!validPassword) {
      return res.status(401).json({ error: 'Invalid username or password' });
    }

    const token = jwt.sign(
      { userId: user.userId, username: user.username }, 
      JWT_SECRET, 
      { expiresIn: '7d' }
    );

    res.json({
      message: 'Login successful',
      token,
      user: {
        userId: user.userId,
        username: user.username,
        fullName: user.fullName,
        subject: user.subject,
        hasWorkbook: !!user.workbookPath,
        studentCount: user.studentDatabase.length
      }
    });
  } catch (error) {
    res.status(500).json({ error: 'Login failed: ' + error.message });
  }
});

// === WORKBOOK MANAGEMENT ===

// ----------------- Replace parseWorkbook with this ExcelJS implementation -----------------
// Complete implementation - Replace these functions in your server.js file

// 1. Replace your parseWorkbook function with this enhanced version
// 1. Enhanced parseWorkbook function for better date and name column detection
async function parseWorkbook(filePath) {
  try {
    console.log(`Parsing Excel workbook (enhanced detection) with ExcelJS: ${filePath}`);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheets = {};
    const allStudents = new Set();

    workbook.eachSheet((worksheet) => {
      const sheetName = worksheet.name;
      console.log(`Processing sheet: "${sheetName}" (rows: ${worksheet.rowCount}, cols: ${worksheet.columnCount})`);

      const rowCount = worksheet.rowCount || 0;
      const colCount = worksheet.columnCount || 0;

      // Build grid with enhanced cell processing
      const grid = [];
      for (let r = 1; r <= rowCount; r++) {
        const rowArr = [];
        for (let c = 1; c <= colCount; c++) {
          const cell = worksheet.getRow(r).getCell(c);
          let val = cell.value;
          let text = "";

          if (val === null || val === undefined) {
            text = "";
          } else if (val instanceof Date) {
            text = val.toISOString().slice(0, 10);
          } else if (typeof val === "number" && cell.style?.numFmt) {
            // Check if it's a date format
            const numFmt = cell.style.numFmt.toLowerCase();
            if (numFmt.includes('m') && numFmt.includes('d') && (numFmt.includes('y') || numFmt.includes('yy'))) {
              try {
                const date = new Date((val - 25569) * 86400 * 1000);
                if (!isNaN(date.getTime()) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
                  text = date.toISOString().slice(0, 10);
                } else {
                  text = String(val);
                }
              } catch (e) {
                text = String(val);
              }
            } else {
              text = String(val);
            }
          } else if (typeof val === "object") {
            if (Array.isArray(val.richText)) {
              text = val.richText.map(rt => rt.text || "").join("");
            } else if (val.text) {
              text = String(val.text || "");
            } else if (val.result !== undefined) {
              text = String(val.result || "");
            } else {
              text = String(val).replace(/\r\n/g, " ");
            }
          } else {
            text = String(val);
          }

          text = text.trim();
          rowArr.push(text);
        }
        grid.push(rowArr);
      }

      // Enhanced keyword lists
      const headerKeywords = [
        'name', 'student', 'learner', 'full name', 'lastname', 'surname', 
        'given name', 'student no', 'student number', 'id', 'no.', 'roll',
        'first name', 'family name', 'pupil', 'learners name', 'students name'
      ];

      // Enhanced date detection function
      function looksLikeDateText(s) {
        if (!s) return false;
        s = s.toString().trim();
        if (!s) return false;
        
        const patterns = [
          /^\d{1,2}[\/\-]\d{1,2}([\/\-]\d{2,4})?$/, // MM/DD or MM/DD/YYYY
          /^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/, // YYYY/MM/DD
          /^\d{4}-\d{2}-\d{2}$/, // ISO date YYYY-MM-DD
          /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i, // Month names
          /\b\d{1,2}(st|nd|rd|th)?\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b/i, // DD MMM
          /\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, // Day names
          /^\d{1,2}$/, // Just day numbers like "1", "2", "15"
          /^(mon|tue|wed|thu|fri|sat|sun)/i // Abbreviated day names
        ];
        
        return patterns.some(pattern => pattern.test(s));
      }

      // Enhanced header row detection - check first 5 rows
      let headerRowIndex = -1;
      let bestHeaderScore = 0;

      for (let r = 1; r <= Math.min(5, rowCount); r++) {
        const joined = grid[r-1].join(" ").toLowerCase();
        let score = 0;
        
        // Count keyword matches
        headerKeywords.forEach(keyword => {
          if (joined.includes(keyword)) {
            score += keyword === 'name' ? 10 : 5;
          }
        });
        
        // Bonus for having multiple non-empty cells
        const nonEmptyCells = grid[r-1].filter(c => c && c.length > 0).length;
        if (nonEmptyCells >= 3) score += 2;
        
        // Bonus for typical header patterns
        if (joined.includes('student') && joined.includes('name')) score += 5;
        
        console.log(`   Row ${r} header score: ${score} - "${joined.substring(0, 50)}..."`);
        
        if (score > bestHeaderScore) {
          bestHeaderScore = score;
          headerRowIndex = r;
        }
      }
      
      if (headerRowIndex === -1) headerRowIndex = 1;
      console.log(`   Using header row: ${headerRowIndex} (score: ${bestHeaderScore})`);

      // Enhanced name column detection - check first 5 columns
      let nameColumnIndex = -1;
      let bestNameScore = 0;

      // First check headers for name keywords
      const headerRow = grid[headerRowIndex - 1] || [];
      for (let c = 0; c < Math.min(colCount, 5); c++) {
        const h = (headerRow[c] || "").toLowerCase();
        let score = 0;
        
        headerKeywords.forEach(keyword => {
          if (h.includes(keyword)) {
            score += keyword === 'name' ? 15 : 8;
          }
        });
        
        console.log(`   Column ${c + 1} header score: ${score} - "${h}"`);
        
        if (score > bestNameScore) {
          bestNameScore = score;
          nameColumnIndex = c + 1;
        }
      }

      // If no clear header match, analyze content in first 5 columns
      if (nameColumnIndex === -1 || bestNameScore < 8) {
        const contentScores = new Array(Math.min(colCount, 5)).fill(0);
        const maxScanRows = Math.min(50, Math.max(10, rowCount - headerRowIndex));
        
        for (let c = 0; c < contentScores.length; c++) {
          let score = 0;
          let namePatterns = 0;
          
          for (let r = headerRowIndex; r <= Math.min(rowCount, headerRowIndex + maxScanRows); r++) {
            const txt = grid[r-1][c] || "";
            
            if (txt.length >= 3 && /[A-Za-z]/.test(txt)) {
              // Full name patterns (First Last or Last, First)
              if (/\s/.test(txt) || /,/.test(txt)) {
                score += 5;
                namePatterns++;
              }
              // Single name but capitalized properly
              else if (/^[A-Z][a-z]+$/.test(txt)) {
                score += 2;
                namePatterns++;
              }
              // Mixed case names
              else if (/^[A-Z]/.test(txt) && /[a-z]/.test(txt)) {
                score += 1;
              }
              
              // Penalize obvious non-names
              if (/^\d+$/.test(txt) || /^(present|absent|p|a|x|total|sum|date)$/i.test(txt)) {
                score -= 10;
              }
            }
          }
          
          contentScores[c] = score;
          console.log(`   Column ${c + 1} content score: ${score} (${namePatterns} name patterns)`);
        }
        
        const bestContentScore = Math.max(...contentScores);
        if (bestContentScore > bestNameScore) {
          nameColumnIndex = contentScores.indexOf(bestContentScore) + 1;
          bestNameScore = bestContentScore;
        }
      }

      if (nameColumnIndex === -1) nameColumnIndex = 1;
      console.log(`   Using name column: ${nameColumnIndex} (score: ${bestNameScore})`);

      // Extract students with enhanced filtering
      const students = [];
      for (let r = headerRowIndex + 1; r <= rowCount; r++) {
        const cellVal = (grid[r-1][nameColumnIndex - 1] || "").toString().trim();
        
        if (cellVal && 
            cellVal.length >= 2 && 
            /[A-Za-z]/.test(cellVal) &&
            // Enhanced exclusion patterns
            !/^(present|absent|p|a|x|\d+|date|total|sum|average|class|section|teacher|note|remarks?)$/i.test(cellVal) &&
            // Must have reasonable name characteristics
            /[a-zA-Z]{2,}/.test(cellVal) &&
            // Not just numbers or symbols
            !/^[\d\s\-_\.]+$/.test(cellVal)) {
          
          students.push(cellVal);
          allStudents.add(cellVal);
        }
      }

      // Enhanced date column detection - check first 5 rows thoroughly
      const dateColumns = [];
      
      for (let c = 0; c < colCount; c++) {
        let foundDate = false;
        let dateInfo = null;
        
        // Check each of the first 5 rows for date-like content
        for (let r = 1; r <= Math.min(5, rowCount) && !foundDate; r++) {
          const txt = grid[r-1][c] || "";
          const cell = worksheet.getRow(r).getCell(c + 1);
          
          // Check for date patterns in text
          if (looksLikeDateText(txt)) {
            dateInfo = { 
              index: c + 1, 
              header: txt, 
              headerRow: r,
              columnLetter: getColumnLetter(c + 1),
              detectedAs: 'text'
            };
            foundDate = true;
          }
          
          // Check for Excel date formatting
          else if (cell && cell.style && cell.style.numFmt) {
            const numFmt = cell.style.numFmt.toLowerCase();
            if ((numFmt.includes('m') && numFmt.includes('d')) || 
                (numFmt.includes('date')) ||
                (numFmt.includes('mm') && numFmt.includes('dd'))) {
              dateInfo = {
                index: c + 1,
                header: txt || `Date Column ${c + 1}`,
                headerRow: r,
                columnLetter: getColumnLetter(c + 1),
                detectedAs: 'format',
                format: cell.style.numFmt
              };
              foundDate = true;
            }
          }
        }
        
        if (foundDate && dateInfo) {
          dateColumns.push(dateInfo);
          console.log(`   Found date column at ${dateInfo.columnLetter} (${dateInfo.detectedAs}): "${dateInfo.header}"`);
        }
      }

      // Helper function for column letters
      function getColumnLetter(colIndex) {
        let result = '';
        while (colIndex > 0) {
          colIndex--;
          result = String.fromCharCode(65 + (colIndex % 26)) + result;
          colIndex = Math.floor(colIndex / 26);
        }
        return result;
      }

      console.log(`   Final results: ${students.length} students, ${dateColumns.length} date columns`);

      sheets[sheetName] = {
        students,
        studentCount: students.length,
        nameColumnIndex,
        nameColumnLetter: getColumnLetter(nameColumnIndex),
        dateColumns: dateColumns || [], // Ensure it's always an array
        headerRowIndex,
        totalRows: rowCount,
        totalColumns: colCount
      };
    });

    return {
      sheets,
      allStudents: Array.from(allStudents),
      totalStudents: allStudents.size,
      sheetsFound: Object.keys(sheets).length
    };
  } catch (err) {
    console.error('Workbook parsing failed:', err);
    throw new Error('Could not parse workbook: ' + (err?.message || String(err)));
  }
}

// Also fix your multer configuration - replace your existing multer setup with this:
const fixedStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    console.log('MULTER DESTINATION - Field:', file.fieldname, 'Original:', file.originalname);
    
    let dir;
    if (file.fieldname === 'workbook') {
      dir = 'workbooks/';
    } else if (file.fieldname === 'photo') {
      dir = 'uploads/';
    } else {
      // Fallback based on file extension
      const ext = path.extname(file.originalname).toLowerCase();
      if (['.xlsx', '.xls'].includes(ext)) {
        dir = 'workbooks/';
      } else {
        dir = 'uploads/';
      }
    }
    
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    
    console.log('MULTER: Using directory:', dir);
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const userId = req.user ? req.user.userId : Date.now();
    const timestamp = Date.now();
    
    let filename;
    if (file.fieldname === 'workbook') {
      const ext = path.extname(file.originalname);
      filename = `${userId}_workbook_${timestamp}${ext}`;
    } else if (file.fieldname === 'photo') {
      filename = `${userId}_photo_${timestamp}.jpg`;
    } else {
      // Fallback
      const ext = path.extname(file.originalname).toLowerCase();
      if (['.xlsx', '.xls'].includes(ext)) {
        filename = `${userId}_workbook_${timestamp}${ext}`;
      } else {
        filename = `${userId}_photo_${timestamp}.jpg`;
      }
    }
    
    console.log('MULTER: Generated filename:', filename);
    cb(null, filename);
  },
});


// Add a simple test endpoint to check if uploads work
app.post('/api/test-upload', authenticateToken, upload.single('workbook'), (req, res) => {
  console.log('TEST UPLOAD:', req.file);
  
  if (req.file) {
    // Clean up test file
    fs.unlinkSync(req.file.path);
  }
  
  res.json({ 
    success: true, 
    file: req.file ? {
      fieldname: req.file.fieldname,
      originalname: req.file.originalname,
      path: req.file.path,
      destination: req.file.destination
    } : null
  });
});

// 2. Enhanced download endpoint with proper headers
// Replace your download endpoint in server.js with this enhanced version that includes better debugging

app.get('/api/workbook/download', authenticateToken, (req, res) => {
  try {
    console.log(`Download request from user: ${req.user.username}`);
    
    const user = users.find(u => u.userId === req.user.userId);
    if (!user) {
      console.error('User not found for download');
      return res.status(404).json({ error: 'User not found' });
    }

    console.log(`User workbook path: ${user.workbookPath}`);
    console.log(`Path exists: ${user.workbookPath ? fs.existsSync(user.workbookPath) : 'No path set'}`);

    if (!user.workbookPath) {
      console.error('No workbook path set for user');
      return res.status(404).json({ 
        error: 'No workbook found',
        details: 'Please upload a workbook first'
      });
    }

    if (!fs.existsSync(user.workbookPath)) {
      console.error(`Workbook file not found at: ${user.workbookPath}`);
      return res.status(404).json({ 
        error: 'Workbook file not found',
        details: 'The workbook file may have been moved or deleted. Please re-upload your workbook.'
      });
    }

    // Get file stats for debugging
    let fileStats;
    try {
      fileStats = fs.statSync(user.workbookPath);
      console.log(`File size: ${fileStats.size} bytes (${Math.round(fileStats.size/1024)}KB)`);
      console.log(`File modified: ${fileStats.mtime}`);
    } catch (statError) {
      console.error('Error getting file stats:', statError);
      return res.status(500).json({ 
        error: 'File access error',
        details: statError.message 
      });
    }

    // Check if file is readable
    try {
      fs.accessSync(user.workbookPath, fs.constants.R_OK);
      console.log('File is readable');
    } catch (accessError) {
      console.error('File not readable:', accessError);
      return res.status(500).json({ 
        error: 'File access denied',
        details: 'Cannot read the workbook file' 
      });
    }

    const fileName = `${user.username}_attendance_updated_${new Date().toISOString().slice(0, 10)}.xlsx`;
    console.log(`Sending file as: ${fileName}`);

    // Set proper headers for Excel file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Length', fileStats.size.toString());
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');

    // Create read stream and handle errors
    const fileStream = fs.createReadStream(user.workbookPath);
    
    fileStream.on('error', (streamError) => {
      console.error('File stream error:', streamError);
      if (!res.headersSent) {
        res.status(500).json({ 
          error: 'Download failed',
          details: streamError.message 
        });
      }
    });

    fileStream.on('open', () => {
      console.log('File stream opened successfully');
    });

    fileStream.on('end', () => {
      console.log('File download completed successfully');
    });

    res.on('error', (resError) => {
      console.error('Response error:', resError);
    });

    res.on('close', () => {
      console.log('Download response closed');
    });

    // Pipe the file to response
    fileStream.pipe(res);

  } catch (error) {
    console.error('Download endpoint error:', error);
    if (!res.headersSent) {
      res.status(500).json({ 
        error: 'Download failed',
        details: error.message 
      });
    }
  }
});

// Add a debug endpoint to check workbook status
app.get('/api/workbook/debug', authenticateToken, (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }

    const debugInfo = {
      userId: user.userId,
      username: user.username,
      workbookPath: user.workbookPath || null,
      pathExists: user.workbookPath ? fs.existsSync(user.workbookPath) : false,
      studentCount: user.studentDatabase ? user.studentDatabase.length : 0,
      sheetsCount: user.workbookStructure ? Object.keys(user.workbookStructure).length : 0,
      updatedAt: user.updatedAt || null
    };

    if (user.workbookPath && fs.existsSync(user.workbookPath)) {
      try {
        const stats = fs.statSync(user.workbookPath);
        debugInfo.fileSize = stats.size;
        debugInfo.fileModified = stats.mtime;
        debugInfo.isReadable = true;
        
        // Test if it's a valid Excel file
        fs.accessSync(user.workbookPath, fs.constants.R_OK);
      } catch (error) {
        debugInfo.error = error.message;
        debugInfo.isReadable = false;
      }
    }

    // Check workbooks directory
    const workbooksDir = 'workbooks';
    if (fs.existsSync(workbooksDir)) {
      const allFiles = fs.readdirSync(workbooksDir);
      debugInfo.workbooksDirectory = {
        totalFiles: allFiles.length,
        userFiles: allFiles.filter(f => f.includes(user.userId)),
        allFiles: allFiles
      };
    }

    res.json(debugInfo);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 3. Enhanced workbook info endpoint with format details
app.get('/api/workbook/info', authenticateToken, (req, res) => {
  const user = users.find(u => u.userId === req.user.userId);
  if (!user) {
    return res.status(404).json({ error: 'User not found' });
  }

  if (!user.workbookPath || !user.workbookStructure) {
    return res.json({
      hasWorkbook: false,
      message: 'No workbook uploaded yet'
    });
  }

  // Check if workbook file still exists
  if (!fs.existsSync(user.workbookPath)) {
    return res.json({
      hasWorkbook: false,
      error: 'Workbook file not found - please re-upload',
      message: 'Your workbook file appears to have been moved or deleted'
    });
  }

  try {
    const fileStats = fs.statSync(user.workbookPath);
    
    res.json({
      hasWorkbook: true,
      totalStudents: user.studentDatabase.length,
      sheets: Object.keys(user.workbookStructure).map(name => ({
        name: name,
        studentCount: user.workbookStructure[name].studentCount,
        dateColumns: user.workbookStructure[name].dateColumns || [],
        nameColumn: user.workbookStructure[name].nameColumnLetter || 'A',
        sampleStudents: user.workbookStructure[name].students.slice(0, 5),
        hasFormattedDates: (user.workbookStructure[name].dateColumns || []).some(col => col.isFormattedDate)
      })),
      uploadedAt: user.updatedAt,
      fileSize: Math.round(fileStats.size / 1024) + ' KB',
      formatPreservationEnabled: true,
      workbookMetadata: user.workbookMetadata || null
    });
  } catch (error) {
    console.error('Error getting workbook info:', error);
    res.status(500).json({ 
      error: 'Failed to get workbook information',
      hasWorkbook: false 
    });
  }
});

// 4. Helper function to validate Excel file integrity
async function validateExcelFile(filePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    if (!workbook.worksheets || workbook.worksheets.length === 0) {
      throw new Error('No worksheets found in the file');
    }
    
    // Basic validation - check if at least one worksheet has content
    let hasContent = false;
    workbook.eachSheet((worksheet) => {
      if (worksheet.rowCount && worksheet.rowCount > 0) {
        hasContent = true;
      }
    });
    
    if (!hasContent) {
      throw new Error('The Excel file appears to be empty');
    }
    
    return { valid: true };
  } catch (error) {
    return { valid: false, error: error.message };
  }
}

// 5. Enhanced error handling middleware for file uploads
function handleUploadErrors(error, req, res, next) {
  console.error('Upload error:', error);
  
  // Clean up uploaded file if it exists
  if (req.file && fs.existsSync(req.file.path)) {
    try {
      fs.unlinkSync(req.file.path);
    } catch (cleanupError) {
      console.warn('Could not clean up file:', cleanupError.message);
    }
  }
  
  if (error.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({
      error: 'File too large',
      details: 'Please upload a file smaller than 15MB',
      maxSize: '15MB'
    });
  }
  
  if (error.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(400).json({
      error: 'Invalid file field',
      details: 'Please use the correct file upload field'
    });
  }
  
  res.status(500).json({
    error: 'Upload failed',
    details: error.message
  });
}

// Apply error handling middleware
app.use('/api/workbook/upload', handleUploadErrors);
app.use('/api/attendance/process', handleUploadErrors);

app.post('/api/workbook/upload', authenticateToken, upload.single('workbook'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No workbook file uploaded' });
    }

    const originalFileName = req.file.originalname.replace(/\.(xlsx|xls)$/i, '');
    
    const workbookData = await parseWorkbook(req.file.path);
    const user = users.find(u => u.userId === req.user.userId);
    
    if (user) {
      if (user.workbookPath && fs.existsSync(user.workbookPath)) {
        fs.unlinkSync(user.workbookPath);
      }

      user.workbookPath = req.file.path;
      user.studentDatabase = workbookData.allStudents;
      user.workbookStructure = workbookData.sheets;
      user.originalFileName = originalFileName; // Store original name
      user.updatedAt = new Date().toISOString();

      saveUsers();

      res.json({
        message: 'Workbook uploaded successfully',
        totalStudents: workbookData.totalStudents,
        sheetsFound: workbookData.sheetsFound,
        sheets: Object.keys(workbookData.sheets).map(name => ({
          name,
          studentCount: workbookData.sheets[name].studentCount,
          sampleStudents: workbookData.sheets[name].students.slice(0, 3),
        })),
        originalFileName: originalFileName
      });
    }
  } catch (error) {
    console.error('Workbook upload error:', error);
    res.status(500).json({ error: 'Failed to process workbook: ' + error.message });
  }
});

// 6. Backup and restore functionality
function createWorkbookBackup(userId, workbookPath) {
  try {
    if (!fs.existsSync(workbookPath)) return null;
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupDir = path.join('workbooks', 'backups');
    
    if (!fs.existsSync(backupDir)) {
      fs.mkdirSync(backupDir, { recursive: true });
    }
    
    const backupPath = path.join(backupDir, `${userId}_backup_${timestamp}.xlsx`);
    fs.copyFileSync(workbookPath, backupPath);
    
    // Clean up old backups (keep only last 5)
    const backupFiles = fs.readdirSync(backupDir)
      .filter(f => f.startsWith(`${userId}_backup_`))
      .sort()
      .reverse();
    
    if (backupFiles.length > 5) {
      backupFiles.slice(5).forEach(file => {
        try {
          fs.unlinkSync(path.join(backupDir, file));
        } catch (e) {
          console.warn('Could not delete old backup:', e.message);
        }
      });
    }
    
    return backupPath;
  } catch (error) {
    console.error('Backup creation failed:', error);
    return null;
  }
}

// 7. Enhanced file cleanup on server shutdown
process.on('SIGINT', () => {
  console.log('\n🛑 Server shutting down...');
  
  // Clean up temporary files
  ['uploads', 'workbooks/temp'].forEach(dir => {
    if (fs.existsSync(dir)) {
      const files = fs.readdirSync(dir);
      const tempFiles = files.filter(f => 
        f.includes('_temp_') || 
        f.includes('_optimized') || 
        Date.now() - fs.statSync(path.join(dir, f)).mtime.getTime() > 24 * 60 * 60 * 1000
      );
      
      tempFiles.forEach(file => {
        try {
          fs.unlinkSync(path.join(dir, file));
          console.log(`🗑️ Cleaned up: ${file}`);
        } catch (e) {
          console.warn(`Could not clean up ${file}:`, e.message);
        }
      });
    }
  });
  
  process.exit(0);
});

// 8. Health check endpoint for monitoring
app.get('/api/health', (req, res) => {
  const memUsage = process.memoryUsage();
  const uptime = process.uptime();
  
  res.json({
    status: 'healthy',
    uptime: Math.floor(uptime),
    memory: {
      used: Math.round(memUsage.heapUsed / 1024 / 1024) + ' MB',
      total: Math.round(memUsage.heapTotal / 1024 / 1024) + ' MB'
    },
    users: users.length,
    timestamp: new Date().toISOString()
  });
});

app.get('/api/workbook/info', authenticateToken, (req, res) => {
  const user = users.find(u => u.userId === req.user.userId);
  if (!user) {
    return res.status(404).json({ error: 'User not found' });
  }

  if (!user.workbookPath || !user.workbookStructure) {
    return res.json({
      hasWorkbook: false,
      message: 'No workbook uploaded yet'
    });
  }

  res.json({
    hasWorkbook: true,
    totalStudents: user.studentDatabase.length,
    sheets: Object.keys(user.workbookStructure).map(name => ({
      name: name,
      studentCount: user.workbookStructure[name].studentCount,
      dateColumns: user.workbookStructure[name].dateColumns,
      sampleStudents: user.workbookStructure[name].students.slice(0, 5)
    })),
    uploadedAt: user.updatedAt
  });
});

// === ATTENDANCE PROCESSING ===

// OCR Functions (from previous code)
async function optimizeForOCR(inputPath, outputPath) {
  const metadata = await sharp(inputPath).metadata();
  await sharp(inputPath)
    .rotate()
    .resize({ width: 1200, height: 1600, fit: 'inside', withoutEnlargement: true })
    .grayscale()
    .normalize()
    .sharpen({ sigma: 1.0 })
    .jpeg({ quality: 85 })
    .toFile(outputPath);
  
  return { success: true };
}

// normalize text for matching
// Normalizer: remove diacritics, punctuation, common suffixes, collapse spaces, lowercase
function normalizeText(s) {
  if (!s) return '';
  return s.toString()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')        // remove accents
    .replace(/\b(jr|sr|iii|ii|iv|jr\.)\b/ig, '') // remove suffixes
    .replace(/[^\w\s]/g, ' ')               // remove punctuation except spaces
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// Split name into tokens and attempt first/last detection.
// Accepts strings like "Llaneta, Raiceen C." or "Raiceen C. Llaneta"
function splitNameParts(raw) {
  const s = raw || '';
  const hasComma = s.indexOf(',') !== -1;
  if (hasComma) {
    // "Last, First Middle"
    const parts = s.split(',');
    const last = normalizeText(parts[0]);
    const rest = normalizeText(parts.slice(1).join(' '));
    const tokens = rest.split(' ').filter(Boolean);
    return {
      raw: s,
      first: tokens[0] || '',
      last: last.split(' ').pop() || last,
      tokens: (tokens.concat(last.split(' ').filter(Boolean))).filter(Boolean)
    };
  } else {
    // assume "First Middle Last"
    const norm = normalizeText(s);
    const tokens = norm.split(' ').filter(Boolean);
    return {
      raw: s,
      first: tokens[0] || '',
      last: tokens[tokens.length - 1] || '',
      tokens
    };
  }
}

// Replace your current performOCR with this function
// === RESTORE multi-engine performOCROptimized (old working implementation) ===
async function performOCROptimized(imagePath, engine = "2") {
  try {
    console.log(`Performing optimized OCR with engine ${engine}...`);

    // Read file and get stats
    const fileStats = fs.statSync(imagePath);
    const fileBuffer = fs.readFileSync(imagePath);

    console.log(`File size: ${(fileStats.size/1024).toFixed(1)}KB`);

    // Create proper FormData
    const formData = new FormData();
    formData.append("apikey", process.env.OCR_SPACE_API_KEY || "helloworld");
    formData.append("language", "eng");
    formData.append("isOverlayRequired", "false");
    formData.append("detectOrientation", "true");
    formData.append("isTable", "true");
    formData.append("scale", "true");
    formData.append("ocrengine", engine);

    // Create blob with proper MIME type (your environment supported this previously)
    const blob = new Blob([fileBuffer], { type: 'image/jpeg' });
    formData.append("file", blob, "image.jpg");

    console.log(`📡 Calling OCR.Space API with engine ${engine}...`);

    const response = await fetch("https://api.ocr.space/parse/image", {
      method: "POST",
      body: formData,
      timeout: 45000,
      headers: {
        'User-Agent': 'OCR-Client/1.0'
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const result = await response.json();
    console.log(`OCR Response for engine ${engine}:`, result.IsErroredOnProcessing ? 'ERROR' : 'SUCCESS');

    if (result.IsErroredOnProcessing) {
      console.error(`OCR Engine ${engine} Error:`, result.ErrorMessage || result.ErrorDetails);
      return { 
        text: "", 
        success: false, 
        error: Array.isArray(result.ErrorMessage) ? result.ErrorMessage.join(', ') : result.ErrorMessage 
      };
    }

    if (result.ParsedResults && result.ParsedResults.length > 0) {
      const text = result.ParsedResults[0].ParsedText || "";
      console.log(`Engine ${engine} extracted ${text.length} characters`);

      if (text.length > 0) {
        console.log(`Sample: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`);
      }

      return {
        text: text,
        success: true,
        confidence: text.length > 50 ? "high" : "medium"
      };
    }

    return { text: "", success: false, error: "No text found in image" };

  } catch (error) {
    console.error(`❌ OCR Engine ${engine} error:`, error.message);
    return { text: "", success: false, error: error.message };
  }
}

function isExcelFile(filePath) {
  if (!filePath) return false;
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx' || ext === '.xls') {
    // quick check: file exists and can be read by XLSX
    try {
      if (!fs.existsSync(filePath)) return false;
      const wb = XLSX.readFile(filePath, { sheetStubs: true });
      return Array.isArray(wb.SheetNames) && wb.SheetNames.length > 0;
    } catch (e) {
      return false;
    }
  }
  // not xls/xlsx extension -> false
  return false;
}
  

// Name matching functions
function calculateSimilarity(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;
  const matrix = Array(len2 + 1).fill().map(() => Array(len1 + 1).fill(0));

  for (let i = 0; i <= len1; i++) matrix[0][i] = i;
  for (let j = 0; j <= len2; j++) matrix[j][0] = j;

  for (let j = 1; j <= len2; j++) {
    for (let i = 1; i <= len1; i++) {
      const substitutionCost = str1[i - 1] === str2[j - 1] ? 0 : 1;
      matrix[j][i] = Math.min(
        matrix[j][i - 1] + 1,
        matrix[j - 1][i] + 1,
        matrix[j - 1][i - 1] + substitutionCost
      );
    }
  }

  const maxLen = Math.max(len1, len2);
  return maxLen === 0 ? 1 : (maxLen - matrix[len2][len1]) / maxLen;
}

function findTopMatches(ocrText, studentDatabase, topN = 3) {
  const normalizedOcr = normalizeText(ocrText);
  if (!normalizedOcr || normalizedOcr.length < 2) return [];

  const ocrParts = splitNameParts(ocrText);
  const ocrTokens = (ocrParts.tokens || []).filter(Boolean);
  const candidates = [];

  for (const dbName of studentDatabase) {
    const dbParts = splitNameParts(dbName);
    const normalizedDb = normalizeText(dbName);
    if (!normalizedDb) continue;

    // Levenshtein similarity (existing)
    const levSim = calculateSimilarity(normalizedOcr, normalizedDb); // 0..1

    // Jaro (string-similarity)
    const jaro = stringSimilarity.compareTwoStrings(normalizedOcr, normalizedDb); // 0..1

    // Token overlap
    const dbTokens = dbParts.tokens || [];
    let common = 0;
    for (const t of ocrTokens) if (dbTokens.includes(t)) common++;
    const tokenOverlap = dbTokens.length + ocrTokens.length === 0 ? 0 : (2 * common) / (dbTokens.length + ocrTokens.length);

    // Last name match boost
    let lastBoost = 0;
    const ocrLast = ocrParts.last || '';
    const dbLast = dbParts.last || '';
    if (ocrLast && dbLast) {
      if (ocrLast === dbLast) lastBoost = 0.28;
      else {
        const lastSim = calculateSimilarity(ocrLast, dbLast);
        if (lastSim >= 0.85) lastBoost = 0.16;
      }
      const ocrFirstInitial = (ocrParts.first || '').charAt(0);
      const dbFirstInitial = (dbParts.first || '').charAt(0);
      if (lastBoost > 0 && ocrFirstInitial && dbFirstInitial && ocrFirstInitial === dbFirstInitial) {
        lastBoost += 0.12;
      }
    }

    // substring boost
    let substringBoost = 0;
    if (normalizedDb === normalizedOcr) substringBoost = 0.95;
    else if (normalizedDb.includes(normalizedOcr) || normalizedOcr.includes(normalizedDb)) substringBoost = 0.8;

    // composite score: combine levSim, jaro, tokenOverlap and boosts
    const score = Math.min(1,
      (0.35 * levSim) +
      (0.30 * jaro) +
      (0.20 * tokenOverlap) +
      (0.10 * substringBoost) +
      lastBoost
    );

    candidates.push({
      name: dbName,
      normalized: normalizedDb,
      score: Number(score.toFixed(4)),
      details: {
        levSim: Number(levSim.toFixed(3)),
        jaro: Number(jaro.toFixed(3)),
        tokenOverlap: Number(tokenOverlap.toFixed(3)),
        lastBoost: Number(lastBoost.toFixed(3)),
        substringBoost: Number(substringBoost.toFixed(3))
      }
    });
  }

  candidates.sort((a, b) => b.score - a.score);
  return candidates.slice(0, topN);
}


function parseAttendanceFromOCR(ocrText, studentDatabase, selectedSheetStudents) {
  if (!ocrText || ocrText.trim().length === 0) return [];

  // Use only students from the selected sheet
  const targetStudents = selectedSheetStudents || studentDatabase;
  console.log(`Using ${targetStudents.length} students from selected sheet for matching`);

  const lines = ocrText
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  const attendanceList = [];
  const skipPatterns = [/attendance/i, /present/i, /absent/i, /date/i, /class/i, /section/i];

  // REDUCED THRESHOLDS FOR BETTER MATCHING
  const AUTO_MATCH_THRESHOLD = 0.40;  // Lowered from 0.68 to 0.55 (55% similarity)
  const ACCEPT_IF_LASTNAME_MATCH = true;
  const DELTA_THRESHOLD = 0.15;       // Lowered from 0.18 to 0.15

  lines.forEach((line, idx) => {
    if (skipPatterns.some(p => p.test(line))) return;
    if (line.length < 2 || line.length > 120) return;

    const cleanLine = line.replace(/^\d+[\.\)\-]?\s*/, '').replace(/[^\w\s\.\-']/g, ' ').replace(/\s+/g, ' ').trim();
    if (cleanLine.length < 2) return;

    // Match against students from selected sheet
    const top = findTopMatches(cleanLine, targetStudents, 3);
    const best = top[0] || null;
    const second = top[1] || null;

    let accepted = false;
    let reason = '';

    // Rule 1: High confidence match (55%+)
    if (best && best.score >= AUTO_MATCH_THRESHOLD) {
      accepted = true;
      reason = `highConfidence(${(best.score * 100).toFixed(1)}%)`;
    } 
    // Rule 2: Clear winner - best is significantly better than second
    else if (best && second && (best.score - second.score) >= DELTA_THRESHOLD && best.score >= 0.45) {
      accepted = true;
      reason = `clearWinner(${(best.score * 100).toFixed(1)}% vs ${(second.score * 100).toFixed(1)}%)`;
    } 
    // Rule 3: Exact last name match (most reliable)
    else if (best && ACCEPT_IF_LASTNAME_MATCH) {
      const ocrLast = splitNameParts(cleanLine).last || '';
      const dbLast = splitNameParts(best.name).last || '';
      if (ocrLast && dbLast && ocrLast === dbLast) {
        accepted = true;
        reason = `exactLastName("${ocrLast}")`;
      }
    }
    // Rule 4: Very close match even if low score
    else if (best && best.score >= 0.40 && !second) {
      accepted = true;
      reason = `onlyMatch(${(best.score * 100).toFixed(1)}%)`;
    }

    console.log(`Line ${idx + 1}: "${cleanLine}" -> ${best ? best.name : 'NO MATCH'} (${best ? (best.score * 100).toFixed(1) + '%' : 'N/A'}) - ${accepted ? '✓ ACCEPTED' : '✗ REJECTED'} [${reason || 'noMatch'}]`);

    if (best && accepted) {
      attendanceList.push({
        name: best.name,
        originalOCR: cleanLine,
        matchScore: best.score,
        matchMethod: 'auto',
        isMatched: true,
        suggestions: top,
        acceptReason: reason,
        fromSheet: true,
        lineNumber: idx + 1  // Add line number for reference
      });
    } else if (best && best.score >= 0.30) {
      // Keep as suggestion if score is at least 30%
      attendanceList.push({
        name: best.name,
        originalOCR: cleanLine,
        matchScore: best.score,
        matchMethod: 'suggestion',
        isMatched: false,
        suggestions: top,
        acceptReason: reason || `lowConfidence(${(best.score * 100).toFixed(1)}%)`,
        fromSheet: true,
        lineNumber: idx + 1
      });
    } else {
      // No good match found
      attendanceList.push({
        name: cleanLine,
        originalOCR: cleanLine,
        matchScore: 0,
        matchMethod: 'none',
        isMatched: false,
        suggestions: [],
        acceptReason: 'noCandidates',
        fromSheet: false,
        lineNumber: idx + 1
      });
    }
  });

  return attendanceList;
}
// Update Excel workbook with attendance
// Replace your current updateWorkbookWithAttendance with this function
// ---------- Replace your current updateWorkbookWithAttendance with this function ----------
function findHeaderRowIndex(rows) {
  // Try to find the row that contains headings like 'student', 'name', 'student number'
  const lookFor = ['student', 'student number', 'student name', 'name', 'learner'];
  for (let r = 0; r < Math.min(rows.length, 8); r++) {
    const row = rows[r] || [];
    const joined = row.map(c => (c || '').toString().toLowerCase()).join(' ');
    for (const key of lookFor) {
      if (joined.includes(key)) return r;
    }
  }
  // fallback to first non-empty row
  for (let r = 0; r < Math.min(rows.length, 8); r++) {
    const row = rows[r] || [];
    if (row.some(c => (c || '').toString().trim().length > 0)) return r;
  }
  return 0;
}

function detectNameColumnFromFirstCols(rows, headerIdx) {
  // examine columns 0..2 and score how many "name-like" entries they have
  const maxColsToCheck = 3;
  const scores = Array(maxColsToCheck).fill(0);
  const start = headerIdx + 1;
  const end = Math.min(rows.length, headerIdx + 1 + 40); // sample up to 40 rows
  for (let c = 0; c < maxColsToCheck; c++) {
    for (let r = start; r < end; r++) {
      const cell = (rows[r] || [])[c] || '';
      const s = (cell || '').toString().trim();
      // heuristics: has letters, contains space (first + last) or comma (last, first)
      if (s && /[a-zA-Z]/.test(s) && (s.includes(' ') || s.includes(',') ) && s.length > 3) {
        scores[c] += 1;
      }
    }
  }
  // pick highest score column; fallback to 1 then 0
  let best = 0;
  for (let i = 1; i < scores.length; i++) {
    if (scores[i] > scores[best]) best = i;
  }
  return best; // 0-based column index
}

function normalizeTextBasic(s) {
  if (!s) return '';
  return s.toString()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\b(jr|sr|iii|ii|iv|jr\.)\b/ig, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function combinedRowName(row) {
  // combine first three columns to form the likely name string
  const parts = [];
  for (let c = 0; c < 3; c++) {
    const v = (row[c] || '').toString().trim();
    if (v) parts.push(v);
  }
  return normalizeTextBasic(parts.join(' '));
}
// CORRECTED updateWorkbookWithAttendance - Replace in server.js

async function updateWorkbookWithAttendance(user, attendanceList, sheetName, targetDate) {
  try {
    console.log(`Updating workbook: ${sheetName} for ${targetDate}`);

    if (!user?.workbookPath || !fs.existsSync(user.workbookPath)) {
      throw new Error('Workbook file not found');
    }

    const backupPath = user.workbookPath.replace(/\.xlsx?$/i, `_backup_${Date.now()}.xlsx`);
    fs.copyFileSync(user.workbookPath, backupPath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(user.workbookPath);

    const worksheet = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found`);

    // [Keep all your date parsing functions...]
    function tryParseDateFromString(s) {
      if (!s) return null;
      const raw = String(s).trim();
      if (!raw) return null;
      const d1 = new Date(raw);
      if (!isNaN(d1.getTime())) return d1;
      const patterns = [
        /^(\d{1,2})[\/\-](\d{1,2})([\/\-](\d{2,4}))?$/,
        /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/
      ];
      for (const pattern of patterns) {
        const match = raw.match(pattern);
        if (match) {
          if (pattern === patterns[0]) {
            const a = Number(match[1]), b = Number(match[2]), c = match[4] ? Number(match[4]) : null;
            if (c !== null) {
              const fullYear = c < 100 ? 2000 + c : c;
              const dt1 = new Date(fullYear, a - 1, b);
              const dt2 = new Date(fullYear, b - 1, a);
              if (!isNaN(dt1.getTime())) return dt1;
              if (!isNaN(dt2.getTime())) return dt2;
            }
          } else if (pattern === patterns[1]) {
            const dt = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
            if (!isNaN(dt.getTime())) return dt;
          }
        }
      }
      return null;
    }

    function parseCellToDate(cell) {
      if (!cell?.value) return null;
      const v = cell.value;
      if (v instanceof Date) return v;
      if (typeof v === 'number' && !isNaN(v) && v > 1 && v < 100000) {
        try {
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const days = Math.floor(Number(v));
          const ms = Math.round(days * 24 * 60 * 60 * 1000);
          return new Date(excelEpoch.getTime() + ms);
        } catch (e) {
          return null;
        }
      }
      if (typeof v === 'string' || typeof v === 'number') {
        return tryParseDateFromString(String(v));
      }
      if (v && typeof v === 'object' && v.text) {
        return tryParseDateFromString(String(v.text));
      }
      return null;
    }

    function sameYMD(a, b) {
      if (!a || !b) return false;
      return a.getFullYear() === b.getFullYear() && 
             a.getMonth() === b.getMonth() && 
             a.getDate() === b.getDate();
    }

    let parsedTargetDate = tryParseDateFromString(targetDate);
    if (!parsedTargetDate) {
      parsedTargetDate = new Date(String(targetDate));
      if (isNaN(parsedTargetDate.getTime())) parsedTargetDate = null;
    }

    // Find or create date column
    const maxHeaderRow = Math.min(8, worksheet.rowCount || 8);
    let found = { dateCol: -1, headerRowIdx: -1, existingCell: null };

    for (let hr = 1; hr <= maxHeaderRow; hr++) {
      const row = worksheet.getRow(hr);
      const maxCols = Math.min(50, worksheet.columnCount || 50);
      for (let c = 1; c <= maxCols; c++) {
        let cell = row.getCell(c);
        if (cell?.isMerged && cell.master) {
          cell = cell.master;
        }
        const cellDate = parseCellToDate(cell);
        if (cellDate && parsedTargetDate && sameYMD(cellDate, parsedTargetDate)) {
          found = { dateCol: c, headerRowIdx: hr, existingCell: cell };
          break;
        }
        const text = cell?.value ? String(cell.value) : '';
        if (text && typeof targetDate === 'string') {
          const lowerText = text.toLowerCase().replace(/[^\w]/g, '');
          const lowerTarget = String(targetDate).toLowerCase().replace(/[^\w]/g, '');
          if (lowerText === lowerTarget || 
              (lowerText.length > 3 && lowerTarget.includes(lowerText))) {
            found = { dateCol: c, headerRowIdx: hr, existingCell: cell };
            break;
          }
        }
      }
      if (found.dateCol !== -1) break;
    }

    // Create new column if needed
    if (found.dateCol === -1) {
      let bestHeaderRow = 1;
      let maxContent = 0;
      for (let hr = 1; hr <= maxHeaderRow; hr++) {
        const row = worksheet.getRow(hr);
        let content = 0;
        for (let c = 1; c <= 20; c++) {
          const cell = row.getCell(c);
          if (cell?.value) content++;
        }
        if (content > maxContent) {
          maxContent = content;
          bestHeaderRow = hr;
        }
      }
      const headerRow = worksheet.getRow(bestHeaderRow);
      let newCol = (worksheet.columnCount || 0) + 1;
      for (let c = 1; c <= (worksheet.columnCount || 0) + 5; c++) {
        const cell = headerRow.getCell(c);
        if (!cell.value) {
          newCol = c;
          break;
        }
      }
      const newHeaderCell = headerRow.getCell(newCol);
      const prevCell = headerRow.getCell(newCol - 1);
      if (prevCell?.style) {
        newHeaderCell.style = JSON.parse(JSON.stringify(prevCell.style));
      }
      const dateValue = parsedTargetDate;
      if (dateValue && !isNaN(dateValue.getTime())) {
        newHeaderCell.value = dateValue;
        if (!newHeaderCell.style) newHeaderCell.style = {};
        if (!newHeaderCell.style.numFmt) {
          newHeaderCell.style.numFmt = 'mm/dd/yyyy';
        }
      } else {
        newHeaderCell.value = String(targetDate);
      }
      found = { dateCol: newCol, headerRowIdx: bestHeaderRow, existingCell: newHeaderCell };
      console.log(`Created new date column at ${newCol}`);
    }

    // Build student mapping from worksheet
    const rosterStart = found.headerRowIdx + 1;
    const rosterRows = [];
    const lastRow = Math.max(worksheet.rowCount || 0, rosterStart + 200);
    
    for (let r = rosterStart; r <= lastRow; r++) {
      const row = worksheet.getRow(r);
      const parts = [];
      for (let c = 1; c <= 5; c++) {
        const cell = row.getCell(c);
        let cellValue = '';
        if (cell?.value !== null && cell?.value !== undefined) {
          if (typeof cell.value === 'object') {
            if (cell.value.text) {
              cellValue = String(cell.value.text);
            } else if (Array.isArray(cell.value.richText)) {
              cellValue = cell.value.richText.map(rt => rt.text || '').join('');
            } else {
              cellValue = String(cell.value);
            }
          } else {
            cellValue = String(cell.value);
          }
        }
        const cleanValue = cellValue.trim();
        if (cleanValue) parts.push(cleanValue);
      }
      if (parts.length === 0) continue;
      const combined = parts.join(' ')
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^\w\s]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      if (combined.length >= 3 && /[a-zA-Z]/.test(combined) &&
          !/(total|sum|count|average|date|present|absent)$/i.test(combined)) {
        rosterRows.push({ 
          rowNumber: r, 
          combined, 
          rawParts: parts,
          row: row
        });
      }
    }

    // Get students from SELECTED SHEET only
    const selectedSheetData = user.workbookStructure[sheetName];
    const selectedSheetStudents = selectedSheetData.students || [];

    console.log(`Total students in selected sheet "${sheetName}": ${selectedSheetStudents.length}`);

    // Map selected sheet students to worksheet rows
    const nameToRow = {};
    for (const dbName of selectedSheetStudents) {
      const normDb = String(dbName)
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^\w\s]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      let bestMatch = null;
      let bestScore = 0;
      for (const rosterRow of rosterRows) {
        const score = calculateSimilarity(normDb, rosterRow.combined);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = rosterRow;
        }
      }
      if (bestMatch && bestScore >= 0.4) {
        nameToRow[dbName] = bestMatch;
      }
    }

    console.log(`Mapped ${Object.keys(nameToRow).length} sheet students to worksheet rows`);

    // CORRECTED LOGIC: Build set of students who are PRESENT (found in photo)
    const presentStudents = new Set();
    for (const attendance of attendanceList) {
      if (attendance.isMatched) {
        presentStudents.add(attendance.name);
        console.log(`PRESENT in photo: ${attendance.name}`);
      }
    }

    console.log(`Total PRESENT students from photo: ${presentStudents.size}`);

    let updatedCount = 0;
    let absentCount = 0;
    const writeCol = found.dateCol;

    // Find template cell for formatting
    let templateCell = null;
    for (const rosterRow of rosterRows.slice(0, 10)) {
      const cell = rosterRow.row.getCell(writeCol);
      if (cell?.style && cell.value && 
          /^(present|absent|p|a|x)$/i.test(String(cell.value))) {
        templateCell = cell;
        break;
      }
    }

    // Process ALL students in the selected sheet
    for (const dbName of selectedSheetStudents) {
      const targetRow = nameToRow[dbName];
      
      if (!targetRow) {
        console.log(`Student "${dbName}" not found in worksheet rows - skipping`);
        continue;
      }

      const targetCell = targetRow.row.getCell(writeCol);
      const existingStyle = targetCell.style ? 
        JSON.parse(JSON.stringify(targetCell.style)) : {};

      // CORRECTED: Check if student IS in presentStudents set
      if (presentStudents.has(dbName)) {
        // Student WAS found in photo → Mark PRESENT
        targetCell.value = 'Present';
        updatedCount++;
        console.log(`Marked "${dbName}" as PRESENT (found in photo)`);
      } else {
        // Student was NOT found in photo → Mark ABSENT
        targetCell.value = 'Absent';
        absentCount++;
        console.log(`Marked "${dbName}" as ABSENT (not in photo)`);
      }

      // Apply formatting
      if (Object.keys(existingStyle).length > 0) {
        targetCell.style = existingStyle;
      } else if (templateCell?.style) {
        targetCell.style = JSON.parse(JSON.stringify(templateCell.style));
      }
    }

    console.log(`\nFINAL COUNTS:`);
    console.log(`  Present: ${updatedCount}`);
    console.log(`  Absent: ${absentCount}`);
    console.log(`  Total processed: ${updatedCount + absentCount}`);

    // Save workbook
    try {
      await workbook.xlsx.writeFile(user.workbookPath);
      console.log(`Workbook saved successfully`);
      if (fs.existsSync(backupPath)) {
        fs.unlinkSync(backupPath);
      }
    } catch (saveError) {
      console.error('Error saving workbook:', saveError);
      if (fs.existsSync(backupPath)) {
        fs.copyFileSync(backupPath, user.workbookPath);
        fs.unlinkSync(backupPath);
      }
      throw new Error(`Failed to save workbook: ${saveError.message}`);
    }

    return {
      success: true,
      updatedCount: updatedCount,
      absentCount: absentCount,
      totalProcessed: updatedCount + absentCount,
      sheetName: sheetName,
      date: targetDate,
      dateColumnIndex: writeCol,
      preservedFormatting: true
    };

  } catch (error) {
    console.error('updateWorkbookWithAttendance error:', error);
    throw error;
  }
}

app.post('/api/attendance/process', authenticateToken, upload.single('photo'), async (req, res) => {
  let optimizedPath = null;
  
  try {
    console.log(`\nAttendance processing request from ${req.user.username}`);
    console.log(`Request body:`, req.body);

    if (!req.file) {
      return res.status(400).json({ error: 'No photo uploaded' });
    }

    if (req.file.fieldname !== 'photo') {
      return res.status(400).json({ 
        error: 'Invalid file field', 
        details: `Expected 'photo' field, received '${req.file.fieldname}'` 
      });
    }

    const { sheetName, date } = req.body;
    
    if (!sheetName || !date) {
      return res.status(400).json({ 
        error: 'Sheet name and date are required',
        details: 'Please select a sheet and specify the date for attendance'
      });
    }

    const user = users.find(u => u.userId === req.user.userId);
    if (!user) {
      return res.status(400).json({ error: 'User not found' });
    }

    // Validate workbook and sheet
    if (!user.workbookPath || !fs.existsSync(user.workbookPath)) {
      return res.status(400).json({
        error: 'No valid workbook found',
        details: 'Please re-upload your Excel workbook'
      });
    }

    if (!user.workbookStructure || !user.workbookStructure[sheetName]) {
      const availableSheets = user.workbookStructure ? Object.keys(user.workbookStructure) : [];
      return res.status(400).json({ 
        error: 'Sheet not found', 
        details: `Sheet "${sheetName}" not found in your workbook.`,
        availableSheets: availableSheets
      });
    }

    // Get students ONLY from the selected sheet
    const selectedSheetData = user.workbookStructure[sheetName];
    const selectedSheetStudents = selectedSheetData.students || [];

    console.log(`Sheet-limited processing:`);
    console.log(`   Selected sheet: ${sheetName}`);
    console.log(`   Students in selected sheet: ${selectedSheetStudents.length}`);
    console.log(`   Total students in database: ${user.studentDatabase.length}`);
    
    // Optimize image for OCR
    optimizedPath = req.file.path.replace(/\.[^/.]+$/, '_optimized.jpg');
    await optimizeForOCR(req.file.path, optimizedPath);
    
    // Perform OCR
    const engines = ["2", "1"];
    const allResults = [];

    for (const engine of engines) {
      console.log(`\nTrying OCR engine ${engine}...`);
      const ocrResult = await performOCROptimized(optimizedPath, engine);

      if (ocrResult.success && ocrResult.text && ocrResult.text.trim().length > 0) {
        console.log(`Engine ${engine} SUCCESS: ${ocrResult.text.length} characters`);

        // Use ONLY students from selected sheet for matching
        const students = parseAttendanceFromOCR(
          ocrResult.text, 
          user.studentDatabase, // Full database for fallback
          selectedSheetStudents  // Primary matching source
        );

        allResults.push({
          engine,
          text: ocrResult.text,
          students,
          textLength: ocrResult.text.length,
          confidence: ocrResult.confidence || 'unknown'
        });

        const goodMatches = students.filter(s => s.isMatched && (s.matchScore || 0) > 0.7).length;
        if (goodMatches >= 2 || students.length >= 3) {
          console.log(`Engine ${engine}: ${goodMatches} good matches / ${students.length} total - stopping`);
          break;
        }
      } else {
        console.log(`Engine ${engine} FAILED: ${ocrResult.error || 'No text'}`);
      }

      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (allResults.length === 0) {
      return res.json({
        success: false,
        error: "Could not extract text from image",
        suggestions: [
          "Make sure the image is clear and well-lit",
          "Ensure the attendance sheet matches the selected section",
          "Try taking the photo from directly above the paper"
        ],
        attendanceData: [],
        totalStudents: 0
      });
    }

    // Process results and remove duplicates
    let allStudents = [];
    for (const r of allResults) {
      allStudents = allStudents.concat(r.students);
    }

    const uniqueStudents = [];
    const seenNames = new Set();
    allStudents
      .sort((a, b) => {
        if (a.isMatched && !b.isMatched) return -1;
        if (!a.isMatched && b.isMatched) return 1;
        return (b.matchScore || 0) - (a.matchScore || 0);
      })
      .forEach(student => {
        const key = (student.name || '').toLowerCase().trim();
        if (!seenNames.has(key)) {
          seenNames.add(key);
          uniqueStudents.push(student);
        }
      });

    const attendanceList = uniqueStudents;
    const matchedCount = attendanceList.filter(s => s.isMatched).length;
    const sheetMatchedCount = attendanceList.filter(s => s.isMatched && s.fromSheet).length;

    console.log(`\nFINAL RESULTS:`);
    console.log(`   Total found: ${attendanceList.length}`);
    console.log(`   Matched from selected sheet: ${sheetMatchedCount}`);
    console.log(`   Total matched: ${matchedCount}`);
    
    if (attendanceList.length === 0) {
      return res.json({
        success: false,
        error: "No students found in the image",
        extractedText: allResults[0]?.text?.substring(0, 200) || "No text extracted"
      });
    }

    // Update workbook (this already handles sheet-specific updates)
    console.log(`Updating Excel workbook for sheet: ${sheetName}...`);
    const updateResult = await updateWorkbookWithAttendance(user, attendanceList, sheetName, date);
    
    const response = {
      success: true,
      message: `Successfully processed attendance for ${matchedCount} students in ${sheetName}`,
      attendanceData: attendanceList,
      totalStudents: attendanceList.length,
      matchedStudents: matchedCount,
      unmatchedStudents: attendanceList.length - matchedCount,
      sheetMatchedStudents: sheetMatchedCount, // New field
      workbookUpdate: updateResult,
      sheetName: sheetName,
      date: date,
      selectedSheetInfo: {
        name: sheetName,
        studentCount: selectedSheetStudents.length,
        studentsUsedForMatching: selectedSheetStudents.length
      }
    };
    
    res.json(response);
    
  } catch (error) {
    console.error('Attendance processing failed:', error);
    res.status(500).json({
      success: false,
      error: 'Processing failed: ' + error.message
    });
  } finally {
    // Cleanup temporary files
    try {
      if (req.file && fs.existsSync(req.file.path)) {
        fs.unlinkSync(req.file.path);
      }
      if (optimizedPath && fs.existsSync(optimizedPath)) {
        fs.unlinkSync(optimizedPath);
      }
    } catch (cleanupError) {
      console.warn(`Cleanup warning:`, cleanupError.message);
    }
  }
});

app.get('/api/workbook/download', authenticateToken, (req, res) => {
  try {
    const user = users.find(u => u.userId === req.user.userId);
    
    if (!user?.workbookPath || !fs.existsSync(user.workbookPath)) {
      return res.status(404).json({ 
        error: 'Workbook not found',
        details: 'Please re-upload your workbook file'
      });
    }

    // Get original filename or use default
    let originalName = 'attendance';
    
    if (user.driveFileName) {
      // From Google Drive
      originalName = user.driveFileName.replace(/\.(xlsx|xls)$/i, '');
    } else if (user.workbookPath) {
      // From local upload - extract original name if stored
      const pathParts = user.workbookPath.split(/[\/\\]/);
      const filename = pathParts[pathParts.length - 1];
      // Remove timestamp and _workbook_ prefix
      const match = filename.match(/_workbook_\d+(.+)$/);
      if (match) {
        originalName = 'attendance';
      }
    }
    
    // Create filename: originalname_updated_YYYY-MM-DD.xlsx
    const today = new Date().toISOString().slice(0, 10);
    const fileName = `${originalName}_updated_${today}.xlsx`;
    
    const stat = fs.statSync(user.workbookPath);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Length', stat.size.toString());
    res.setHeader('Cache-Control', 'no-cache');
    
    const fileStream = fs.createReadStream(user.workbookPath);
    fileStream.pipe(res);
    
    fileStream.on('error', (error) => {
      console.error('Download stream error:', error);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Download failed' });
      }
    });
    
    console.log(`Download initiated: ${fileName}`);
  } catch (error) {
    console.error('Download error:', error);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Download failed: ' + error.message });
    }
  }
});

// Initialize
loadUsers();

// Create directories
['uploads', 'workbooks'].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

app.listen(port, () => {
  console.log(`Server: http://localhost:${port}`);
});

// To this:
app.listen(port, '0.0.0.0', () => {
});

// Add this enhanced cleanup function to your server.js to fix the corruption issue

function fixCorruptedWorkbookPaths() {
  console.log('Fixing corrupted workbook paths...');
  
  let fixedCount = 0;
  
  users.forEach(user => {
    let userModified = false;
    
    // Check if workbook path is corrupted (pointing to photo file)
    if (user.workbookPath && (
        user.workbookPath.includes('_photo_') || 
        user.workbookPath.endsWith('.jpg') || 
        user.workbookPath.endsWith('.jpeg') || 
        user.workbookPath.endsWith('.png') ||
        !user.workbookPath.includes('workbook')
      )) {
      
      console.log(`Found corrupted workbook path for ${user.username}: ${user.workbookPath}`);
      
      // Try to find the actual workbook file
      const workbooksDir = 'workbooks';
      if (fs.existsSync(workbooksDir)) {
        const files = fs.readdirSync(workbooksDir);
        const userWorkbooks = files.filter(file => 
          file.startsWith(`${user.userId}_workbook_`) && 
          (file.endsWith('.xlsx') || file.endsWith('.xls'))
        );
        
        if (userWorkbooks.length > 0) {
          // Use the most recent workbook
          const latestWorkbook = userWorkbooks.sort().reverse()[0];
          const correctPath = path.join(workbooksDir, latestWorkbook);
          
          if (fs.existsSync(correctPath)) {
            user.workbookPath = correctPath;
            console.log(`Fixed workbook path: ${correctPath}`);
            userModified = true;
            
            // Re-parse workbook if student database is missing
            if (!user.studentDatabase || user.studentDatabase.length === 0) {
              console.log(`Re-parsing workbook for ${user.username}...`);
              parseWorkbook(correctPath).then(workbookData => {
                user.studentDatabase = workbookData.allStudents;
                user.workbookStructure = workbookData.sheets;
                user.updatedAt = new Date().toISOString();
                saveUsers();
                console.log(`Restored data for ${user.username}: ${workbookData.allStudents.length} students`);
              }).catch(err => {
                console.error(`Failed to re-parse workbook for ${user.username}:`, err);
              });
            }
          }
        } else {
          // No valid workbook found, reset user data
          console.log(`No valid workbook found for ${user.username}, resetting data`);
          user.workbookPath = null;
          user.studentDatabase = [];
          user.workbookStructure = null;
          userModified = true;
        }
      }
    }
    
    // Also check if workbook path exists but file is missing
    if (user.workbookPath && !fs.existsSync(user.workbookPath)) {
      console.log(`Workbook file missing for ${user.username}: ${user.workbookPath}`);
      user.workbookPath = null;
      user.studentDatabase = [];
      user.workbookStructure = null;
      userModified = true;
    }
    
    if (userModified) {
      fixedCount++;
    }
  });
  
  if (fixedCount > 0) {
    saveUsers();
    console.log(`Fixed ${fixedCount} corrupted user records`);
  } else {
    console.log('No corruption found');
  }
  
  return fixedCount;
}

app.post('/api/admin/fix-corruption', authenticateToken, (req, res) => {
  try {
    const fixedCount = fixCorruptedWorkbookPaths();
    res.json({ 
      success: true, 
      message: `Fixed ${fixedCount} corrupted records`,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(500).json({ 
      success: false, 
      error: 'Fix failed: ' + error.message 
    });
  }
});

// Run the fix on server startup
console.log('Running corruption fix on startup...');
fixCorruptedWorkbookPaths();