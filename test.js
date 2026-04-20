const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require('express-fileupload');
const https = require('https');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const cors = require('cors');
const { promisify } = require('util');
const request = promisify(require('request'));

const app = express();

// Middleware setup
app.use(express.json());
app.use(cors());
app.use(fileUpload({
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB limit
  useTempFiles: true,
  tempFileDir: '/tmp/'
}));
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Environment variables
const PORT = process.env.PORT || 5000;
const LF_API_USERNAME = process.env.LF_API_USERNAME || "lfapi";
const LF_API_PASSWORD = process.env.LF_API_PASSWORD || "lPJ!@875fcnbVJFRdoyUwbNJ"; // Should be in env vars
const LF_REPOSITORY_ID = process.env.LF_REPOSITORY_ID || "2705591";
const ZOHO_BASE_URL = "download-accl.zoho.com";
const LF_BASE_URL = "https://dms.tbank.com/LFRepositoryAPI/v1";

// Start server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

// Root endpoint
app.get("/", (req, res) => {
  res.json({ message: "API is running. Use /utl endpoint for file uploads." });
});

// File upload and transfer endpoint
app.post("/utl", async (req, res) => {
  try {
    // Validate request
    const { auth, fileId, fileName } = req.body;
    
    if (!auth || !fileId || !fileName) {
      return res.status(400).json({ error: "Missing required parameters" });
    }
    
    // Sanitize filename to prevent path traversal
    const sanitizedFilename = path.basename(fileName);
    const tempFilePath = path.join(__dirname, sanitizedFilename);
    
    console.log(`Processing file: ${sanitizedFilename}`);

    // Download file from Zoho
    await downloadFileFromZoho(fileId, auth, tempFilePath);
    
    // Get LF access token
    const lfToken = await getLFAccessToken();
    
    // Upload file to LF repository
    const uploadResult = await uploadFileToLF(tempFilePath, sanitizedFilename, lfToken);
    
    console.log("Upload result:", uploadResult);
    
    // Schedule file cleanup
    setTimeout(async () => {
      try {
        await fs.unlink(tempFilePath);
        console.log(`Temporary file ${sanitizedFilename} deleted`);
      } catch (err) {
        console.error(`Error deleting temporary file: ${err.message}`);
      }
    }, 5000);

    res.json({ status: "Success", message: "File transferred to LF repository" });
  } catch (error) {
    console.error("Error processing request:", error);
    res.status(500).json({ error: "Failed to process file", details: error.message });
  }
});

// Helper function to download file from Zoho
function downloadFileFromZoho(fileId, authToken, filePath) {
  return new Promise((resolve, reject) => {
    const file = fsSync.createWriteStream(filePath);
    
    const options = {
      hostname: ZOHO_BASE_URL,
      path: `/v1/workdrive/download/${fileId}`,
      headers: {
        'Authorization': `Zoho-oauthtoken ${authToken}`
      }
    };

    https.get(options, (response) => {
      if (response.statusCode !== 200) {
        reject(new Error(`Failed to download file: ${response.statusCode}`));
        return;
      }

      response.pipe(file);
      
      file.on("finish", () => {
        file.close();
        resolve();
      });
      
      file.on("error", (err) => {
        fsSync.unlink(filePath, () => {});
        reject(err);
      });
    }).on("error", (err) => {
      fsSync.unlink(filePath, () => {});
      reject(err);
    });
  });
}

// Helper function to get LF access token
async function getLFAccessToken() {
  try {
    const response = await fetch(`${LF_BASE_URL}/Repositories/DMS/Token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        grant_type: "password",
        username: LF_API_USERNAME,
        password: LF_API_PASSWORD
      })
    });

    if (!response.ok) {
      throw new Error(`Failed to get token: ${response.status}`);
    }

    const result = await response.json();
    return result.access_token;
  } catch (error) {
    console.error("Error getting LF token:", error);
    throw error;
  }
}

// Helper function to upload file to LF repository
async function uploadFileToLF(filePath, fileName, accessToken) {
  try {
    const options = {
      method: 'POST',
      url: `${LF_BASE_URL}/Repositories/DMS/Entries/${LF_REPOSITORY_ID}/${fileName}?autoRename=false`,
      headers: {
        'Authorization': `Bearer ${accessToken}`
      },
      formData: {
        'electronicDocument': {
          'value': fsSync.createReadStream(filePath),
          'options': {
            'filename': fileName,
            'contentType': null
          }
        }
      }
    };

    const response = await request(options);
    return response.body;
  } catch (error) {
    console.error("Error uploading to LF:", error);
    throw error;
  }
}

// Get folder listing endpoint
app.post("/gfl", async (req, res) => {
  try {
    const { folder } = req.body;
    console.log(req.body);

    if (!folder) {
      return res.status(400).json({ error: "Missing required parameter: folder" });
    }

    const lfToken = await getLFAccessToken();
    const folderListing = await getFolderChildren(folder, lfToken);

    res.json(folderListing);
  } catch (error) {
    console.error("Error getting folder listing:", error);
    res.status(500).json({ error: "Failed to get folder listing", details: error.message });
  }
});

// Helper function to get children of a folder from LF repository
async function getFolderChildren(entryId, accessToken) {
  try {
    const response = await fetch(
      `${LF_BASE_URL}/Repositories/DMS/Entries/${encodeURIComponent(entryId)}/Laserfiche.Repository.Folder/children`,
      {
        method: "GET",
        headers: {
          "Authorization": `Bearer ${accessToken}`
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to get folder listing: ${response.status}`);
    }

    return await response.json();
  } catch (error) {
    console.error("Error fetching folder children:", error);
    throw error;
  }
}

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ error: "Something went wrong", details: err.message });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: "Endpoint not found" });
});