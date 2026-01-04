// Final Instructions:
// 1. Install dependencies: npm install
// 2. Run the script: node upload.js
// On the first run, you will be prompted to authorize the application.
// Follow the URL in the console, grant permission, and paste the resulting code back into the console.

import { google } from 'googleapis';
import { promises as fs } from 'fs';
import path from 'path';
import process from 'process';
import axios from 'axios';
import xlsx from 'xlsx';
import { createInterface } from 'readline';
import cron from "node-cron";
import ffmpegPath from 'ffmpeg-static';
import ffmpeg from 'fluent-ffmpeg';

if (ffmpegPath) {
  ffmpeg.setFfmpegPath(ffmpegPath);
}
// const fs = require('fs');

const youtube = google.youtube('v3');
const OAuth2 = google.auth.OAuth2;

// --- CONFIGURATION ---
const CLIENT_SECRETS_PATH = 'client_secrets.json';
const TOKEN_PATH = 'token.json';
const XLSX_URL = 'https://raw.githubusercontent.com/swipeviewer99-ai/blogs-video-project/fix/intermittent-video-stopping/assets/SEOVideos_updated.xlsx';
const DOWNLOAD_DIR = `C:\\Users\\deept\\blogs-project-branches\\blogs-video-project-refactored\\output1`;

// Scopes required for the YouTube API.
const SCOPES = ['https://www.googleapis.com/auth/youtube.upload'];

/**
 * Reads the local client secrets file.
 */
async function getClientSecrets() {
    const content = await fs.readFile(CLIENT_SECRETS_PATH);
    return JSON.parse(content);
}

/**
 * Authenticate with Google's OAuth2 service.
 */
async function authenticate() {
    console.log('Authenticating...');
    const secrets = await getClientSecrets();
    const client = new OAuth2(
        secrets.installed.client_id,
        secrets.installed.client_secret,
        secrets.installed.redirect_uris[0]
    );

    // Check for a previously saved token.
    try {
        const token = await fs.readFile(TOKEN_PATH);
        client.setCredentials(JSON.parse(token));
        console.log('Authentication successful from saved token.');
        return client;
    } catch (err) {
        return getNewToken(client);
    }
}

/**
 * Get and store a new token after prompting for user authorization.
 */
async function getNewToken(client) {
    const authUrl = client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });

    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    return new Promise((resolve, reject) => {
        rl.question('Enter the code from that page here: ', async (code) => {
            rl.close();
            try {
                const { tokens } = await client.getToken(code);
                client.setCredentials(tokens);
                await fs.writeFile(TOKEN_PATH, JSON.stringify(tokens));
                console.log('Token stored to', TOKEN_PATH);
                console.log('Authentication successful.');
                resolve(client);
            } catch (err) {
                console.error('Error while trying to retrieve access token', err);
                reject(err);
            }
        });
    });
}

/**
 * Downloads the XLSX file from the given URL.
 */
// const axios = require('axios');
// const xlsx = require('xlsx');

async function readXLSXFromUrl(source) {
    console.log('Reading XLSX...');

    try {
        let fileBuffer;

        if (source.startsWith('http://') || source.startsWith('https://')) {
            // Remote URL
            console.log('Fetching from URL...');
            const response = await axios.get(source, { responseType: 'arraybuffer' });
            fileBuffer = response.data;
        } else {
            // Local file path
            console.log('Reading from local file...');
            fileBuffer = await fs.readFile(source);
        }

        const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
        console.log(data[0]);

        console.log('XLSX processed successfully.');
        return data;
    } catch (error) {
        console.error('Error reading XLSX:', error.message);
        throw error;
    }
}


/**
 * Downloads a video from a URL to a local path.
 */
async function downloadVideo(videoUrl, title) {
    // Create a safe filename
    const safeTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const filePath = path.join(DOWNLOAD_DIR, `${safeTitle}.mp4`);

    console.log(`Downloading video: ${title}`);

    // Ensure download directory exists
    await fs.mkdir(DOWNLOAD_DIR, { recursive: true });

    const writer = (await import('fs')).createWriteStream(filePath);
    const response = await axios({
        url: videoUrl,
        method: 'GET',
        responseType: 'stream',
    });

    response.data.pipe(writer);

    return new Promise((resolve, reject) => {
        writer.on('finish', () => {
            console.log(`Finished downloading: ${title}`);
            resolve(filePath);
        });
        writer.on('error', (err) => {
            console.error(`Error downloading video: ${title}`, err);
            reject(err);
        });
    });
}

/**
 * Prepare a short by trimming to `maxSeconds` and converting to vertical 9:16 (720x1280).
 * Returns path to the processed file.
 */
async function prepareShort(originalPath, maxSeconds = 59) {
    const parsed = path.parse(originalPath);

    // Construct the new processed file path
    const outFile = path.join(parsed.dir, `${parsed.name}_short${parsed.ext}`);

    return new Promise((resolve, reject) => {
        ffmpeg(originalPath)
            .videoFilters([
                { filter: 'scale', options: '720:1280:force_original_aspect_ratio=decrease' },
                { filter: 'pad',   options: '720:1280:(ow-iw)/2:(oh-ih)/2' }
            ])
            .format('mp4')
            .outputOptions([
                '-c:v libx264',
                '-profile:v high',
                '-level 4.0',
                '-pix_fmt yuv420p',
                '-c:a aac',
                '-b:a 128k',
                `-t ${maxSeconds}`
            ])
            .on('start', cmd => console.log('FFmpeg CMD:', cmd))
            .on('error', err => {
                console.error('FFmpeg error:', err.message);
                reject(err);
            })
            .on('end', () => {
                console.log('Prepared short saved to:', outFile);
                resolve(outFile);
            })
            .save(outFile);
    });
}



/**
 * Uploads a single video to YouTube.
 */
async function uploadVideo(auth, videoData, localVideoPath) {
    const {
        'YouTube Title': title,
        'YouTube Description': description,
        'YouTube Tags': tags,
        categoryId,
        privacyStatus,
        selfDeclaredMadeForKids,
    } = videoData;

    console.log(`Uploading video: ${title}...`);

    try {
        const res = await youtube.videos.insert({
            auth: auth,
            part: 'snippet,status',
            requestBody: {
                snippet: {
                    title,
                    description,
                    tags: tags ? tags.split(',').map(tag => tag.trim()) : [],
                    categoryId: categoryId || '22', // Default to 'People & Blogs' if not specified
                },
                status: {
                    privacyStatus: privacyStatus || 'private', // Default to 'private'
                    selfDeclaredMadeForKids: selfDeclaredMadeForKids === 'true' || selfDeclaredMadeForKids === true,
                },
            },
            media: {
                body: (await import('fs')).createReadStream(localVideoPath),
            },
        });
        console.log(`Successfully uploaded "${title}" with Video ID: ${res.data.id}`);
        return res.data;
    } catch (error) {
        console.error(`Error uploading video: ${title}`, error.message);
        throw error; // Re-throw to be caught by the main loop
    }
}

/**
 * Uploads the provided video as a YouTube Short.
 * Ensures #shorts tag in title/tags and prepares the file (trim + vertical) before uploading.
 */
async function uploadShort(auth, videoData, localVideoPath) {
    const {
        'YouTube Title': rawTitle,
        'YouTube Description': description,
        'YouTube Tags': tags,
        categoryId,
        privacyStatus,
        selfDeclaredMadeForKids,
    } = videoData;

    // Ensure we have a title and add #shorts if not present
    let title = rawTitle || 'Untitled Short';
    if (!/#[sS]horts/.test(title)) {
        title = `${title} #shorts`;
    }

    // Ensure tags include '#shorts'
    const tagArray = tags ? tags.split(',').map(t => t.trim()).filter(Boolean) : [];
    if (!tagArray.some(t => /#[sS]horts/.test(t))) tagArray.unshift('#shorts');

    console.log(`Preparing and uploading Short: ${title}`);

    // 1) Prepare the short (trim + vertical) - this produces a new file
    let preparedPath;
    try {
        preparedPath = await prepareShort(localVideoPath, 59); // 59s to be safe
    } catch (err) {
        console.error('Error preparing short:', err.message);
        throw err;
        // fallback: attempt to upload original file (still may be rejected by YouTube if >60s or wrong aspect)
        preparedPath = localVideoPath;
    }

    // 2) Upload using same youtube.videos.insert but with #shorts in metadata
    try {
        const res = await youtube.videos.insert({
            auth: auth,
            part: 'snippet,status',
            requestBody: {
                snippet: {
                    title,
                    description: description || '',
                    tags: tagArray,
                    categoryId: categoryId || '22',
                },
                status: {
                    privacyStatus: privacyStatus || 'private',
                    selfDeclaredMadeForKids: selfDeclaredMadeForKids === 'true' || selfDeclaredMadeForKids === true,
                },
            },
            media: {
                body: (await import('fs')).createReadStream(preparedPath),
            },
        });

        console.log(`Successfully uploaded Short "${title}" with Video ID: ${res.data.id}`);
        return res.data;
    } catch (error) {
        console.error(`Error uploading short "${title}":`, error.message);
        throw error;
    } finally {
        // Clean up the prepared file if different from original
        try {
            if (preparedPath && preparedPath !== localVideoPath) {
                await fs.unlink(preparedPath).catch(()=>{});
            }
        } catch (cleanupErr) {
            console.warn('Failed to cleanup prepared short file:', cleanupErr.message);
        }
    }
}

async function fileExists(filePath) {
    try {
        await fs.access(filePath);
        return true;  // File exists
    } catch (err) {
        return false; // File does NOT exist
    }
}



async function updateExcelStatus(filePath, updatedRow) {
    try{
  console.log('for updating excel status!************');

  if (!updatedRow || !updatedRow["Title"]) {
    console.log('No Title provided in updatedRow');
    return;
  }

  // Read workbook and first sheet
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert sheet to array-of-arrays (preserves exact header text)
  const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  if (!sheetData || sheetData.length === 0) {
    throw new Error("Sheet is empty or unreadable");
  }

  // header row (as-is)
  const headers = sheetData[0].map(h => (h === undefined || h === null) ? "" : String(h));

  // find Title column index (case-insensitive, trimmed)
  const titleCol = headers.findIndex(h => String(h).trim().toLowerCase() === "title");
  if (titleCol === -1) {
    throw new Error("Could not find a 'Title' column in the sheet headers");
  }

  // find Uploaded column index (case-insensitive, trimmed)
  let uploadedCol = headers.findIndex(h => String(h).trim().toLowerCase() === "uploaded");
  if (uploadedCol === -1) {
    // add an 'Uploaded' header at the end (preserve header exact text as 'Uploaded')
    headers.push("Uploaded");
    uploadedCol = headers.length - 1;
    sheetData[0] = headers;
  }

  // normalize title to find
  const titleToFind = String(updatedRow["Title"]).trim();

  // Flag to check if we updated anything
  let updatedAnything = false;

  // Iterate data rows and update matching row
  for (let r = 1; r < sheetData.length; r++) {
    const row = sheetData[r] || [];

    const cellTitle = row[titleCol] !== undefined && row[titleCol] !== null ? String(row[titleCol]).trim() : "";

    if (cellTitle === titleToFind) {
      // ensure row array has enough columns
      while (row.length <= uploadedCol) row.push("");
      row[uploadedCol] = "Yes"; // write capitalized "Yes"
      sheetData[r] = row;
      updatedAnything = true;
      break; // remove break to update all matching rows
    }
  }

  if (!updatedAnything) {
    console.warn(`Title "${titleToFind}" not found in sheet. No changes made.`);
    return;
  }

  // Convert the modified array-of-arrays back to a sheet
  const newSheet = xlsx.utils.aoa_to_sheet(sheetData);

  // Update workbook and write file
  workbook.Sheets[sheetName] = newSheet;
  xlsx.writeFile(workbook, filePath);

  console.log(`✔ Excel updated: ${updatedRow["Title"]} marked as Uploaded = Yes`);
}
catch(err)
{
    console.log(err);
}

}


/**
 * Main function to run the script.
 */
async function main() {
    try {
        const auth = await authenticate();
        const videoMetadata = await readXLSXFromUrl('./SEOVideos_updated.xlsx');

        const filteredVideos = videoMetadata.filter(row =>
            !(row.Uploaded && row.Uploaded.toString().toLowerCase() === 'yes')
        );
        for (const videoData of filteredVideos) {
            const videoUrl = videoData.BlobUrl;
            const title = videoData['Title'];

            if (!videoUrl || !title) {
                console.warn('Skipping row due to missing bloburl or resume title:', videoData);
                continue;
            }

            let localVideoPath;
            try {
                // 1. Download video

                localVideoPath = 'C:\\Users\\deept\\blogs-project-branches\\intermittent-video-stopping\\blogs-video-project-refactored\\output1';

                const safeTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
                let filePath = path.join(localVideoPath, `${safeTitle}.mp4`);
                const exists = await fileExists("C:\\path\\to\\file.mp4");
                if (!exists) {
                    console.log(`File doesn't exist hence going for downloading`);
                    localVideoPath = await downloadVideo(videoUrl, title);
                    filePath = localVideoPath;
                }
                // 2. Upload video

                try {
                   await uploadVideo(auth, videoData, filePath);
                    //await uploadShort(auth, videoData, filePath);
                    await updateExcelStatus('./SEOVideos_updated.xlsx', videoData);

                } catch (error) {
                    if (error.message.includes("exceeded the number of videos")) {
                        console.error("STOPPING: YouTube upload limit reached.");
                        break; // exit loop safely
                    }
                    console.log(error);
                }

            } catch (error) {
                console.error(`Failed to process video "${title}". Continuing to next video.`);
                // The specific error is already logged in uploadVideo or downloadVideo
            } finally {
                // 3. Clean up downloaded file
                if (localVideoPath) {
                    try {
                        await fs.unlink(localVideoPath);
                        console.log(`Cleaned up temporary file: ${localVideoPath}`);
                    } catch (cleanupError) {
                        console.error(`Error cleaning up file ${localVideoPath}:`, cleanupError.message);
                    }
                }
            }
        }
        console.log('All videos processed.');

    } catch (error) {
        console.error('A critical error occurred:', error.message);
        process.exit(1);
    }
}
cron.schedule(
  '0 0 20 * * *',
  () => {
    console.log("Cron fired at:", new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }));
    main();
  },
  {
    timezone: "Asia/Kolkata",
    scheduled: true
  }
);

// main();

// cron.schedule("*/10 * * * * *", () => {
 
// });
 main();




