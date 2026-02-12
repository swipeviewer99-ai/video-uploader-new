
import { google } from 'googleapis';
import fs from 'fs';
import { promises as fsPromises } from 'fs';
import path from 'path';
import process from 'process';
import axios from 'axios';
import xlsx from 'xlsx';
import { createInterface } from 'readline';

// --- CONFIGURATION ---
const CLIENT_SECRETS_PATH = 'client_secrets.json';
const TOKEN_PATH = 'token.json';
const REMOTE_XLSX_URL = 'https://raw.githubusercontent.com/swipeviewer99-ai/video-uploader-new/main/SEOVideos_updated.xlsx';
const LOCAL_XLSX_PATH = 'SEOVideos_updated.xlsx';

const SCOPES = ['https://www.googleapis.com/auth/youtube.force-ssl'];

const youtube = google.youtube('v3');
const OAuth2 = google.auth.OAuth2;

/**
 * Reads the local client secrets file.
 */
async function getClientSecrets() {
    try {
        const content = await fsPromises.readFile(CLIENT_SECRETS_PATH);
        return JSON.parse(content);
    } catch (err) {
        console.warn('Warning: client_secrets.json not found. Skipping authentication for dry run.');
        return null;
    }
}

/**
 * Authenticate with Google's OAuth2 service.
 */
async function authenticate() {
    console.log('Authenticating...');
    const secrets = await getClientSecrets();
    if (!secrets) return null; // Return null if secrets are missing (dry run mode)

    const client = new OAuth2(
        secrets.installed.client_id,
        secrets.installed.client_secret,
        secrets.installed.redirect_uris[0]
    );

    // Check for a previously saved token.
    try {
        const token = await fsPromises.readFile(TOKEN_PATH);
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
                await fsPromises.writeFile(TOKEN_PATH, JSON.stringify(tokens));
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
 * Downloads the XLSX file from the remote URL and returns JSON data.
 */
async function getRemoteExcelData() {
    console.log(`Downloading Excel from: ${REMOTE_XLSX_URL}`);
    try {
        const response = await axios.get(REMOTE_XLSX_URL, { responseType: 'arraybuffer' });
        const workbook = xlsx.read(response.data, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert to JSON (array of arrays to handle potentially missing headers)
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Extract headers from the first row
        const headers = data[0];
        console.log('Headers found:', headers);

        // Map remaining rows to objects based on headers
        const rows = data.slice(1).map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            // Handle cases where the YouTube URL might be in a column without a header or specific index
            // Based on previous inspection, it seems to be the last column or index 12 (M)
            if (!obj['YouTube URL'] && row.length > headers.length) {
                 // Fallback: Check if the last element looks like a YouTube URL
                 const lastElement = row[row.length - 1];
                 if (typeof lastElement === 'string' && lastElement.includes('youtube.com')) {
                     obj['YouTube URL'] = lastElement;
                 }
            }
            return obj;
        });

        return { rows, workbook, sheetName, headers };
    } catch (error) {
        console.error('Error downloading or parsing Excel:', error.message);
        throw error;
    }
}

/**
 * Extracts Video ID from a YouTube URL.
 */
function extractVideoId(url) {
    if (!url) return null;
    const match = url.match(/(?:v=|\/)([0-9A-Za-z_-]{11}).*/);
    return match ? match[1] : null;
}

/**
 * Updates the local Excel file to mark a video as updated.
 */
async function markVideoAsUpdated(videoData, index) {
    try {
        // Assume local file exists and is initialized
        if (!fs.existsSync(LOCAL_XLSX_PATH)) {
             console.error("Local file not initialized. Cannot update status.");
             return;
        }

        const workbook = xlsx.readFile(LOCAL_XLSX_PATH);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert to JSON (array of arrays)
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const headers = data[0];

        // Find or add 'updated_description' column
        let statusColIndex = headers.indexOf('updated_description');
        if (statusColIndex === -1) {
            headers.push('updated_description');
            statusColIndex = headers.length - 1;
            data[0] = headers; // Update header row
        }

        // Update the specific row
        // note: 'index' passed from main loop is 0-based index of data rows (excluding header)
        // so actual row in 'data' array is index + 1
        const rowIndex = index + 1;

        if (data[rowIndex]) {
            // Ensure row has enough columns
            while (data[rowIndex].length <= statusColIndex) {
                data[rowIndex].push('');
            }
            data[rowIndex][statusColIndex] = 'yes';
        } else {
             console.error(`Row ${rowIndex} not found in local Excel data.`);
        }

        // Write back to file
        const newSheet = xlsx.utils.aoa_to_sheet(data);
        workbook.Sheets[sheetName] = newSheet;
        xlsx.writeFile(workbook, LOCAL_XLSX_PATH);
        console.log(`Updated Excel status for row ${rowIndex} to 'yes'.`);

    } catch (err) {
        console.error(`Error updating Excel status for row ${index}:`, err.message);
    }
}

/**
 * Initialize local Excel file from remote data
 */
function initializeLocalExcel(workbook) {
    xlsx.writeFile(workbook, LOCAL_XLSX_PATH);
    console.log(`Initialized local Excel file at ${LOCAL_XLSX_PATH}`);
}


/**
 * Updates the video description on YouTube.
 */
async function updateVideoDescription(auth, videoId, newDescription) {
    if (!auth) {
        console.log(`[DRY RUN] Would update video ID: ${videoId}`);
        return true; // Simulate success
    }

    // Truncate description if it's too long (safeguard)
    if (newDescription.length > 4900) {
        console.warn(`WARNING: Description for video ${videoId} is ${newDescription.length} chars. Truncating to 4900.`);
        newDescription = newDescription.substring(0, 4900);
    }

    try {
        // 1. Get current video details (snippet) to preserve title, tags, etc.
        const response = await youtube.videos.list({
            auth: auth,
            part: 'snippet',
            id: videoId
        });

        if (response.data.items.length === 0) {
            console.warn(`Video not found for ID: ${videoId}`);
            return false;
        }

        const videoSnippet = response.data.items[0].snippet;
        const currentDescription = videoSnippet.description;

        if (currentDescription === newDescription) {
            console.log(`Description for video ${videoId} is already up to date.`);
            return true; // Considered success as state is correct
        }

        // 2. Update the description
        videoSnippet.description = newDescription;

        await youtube.videos.update({
            auth: auth,
            part: 'snippet',
            requestBody: {
                id: videoId,
                snippet: videoSnippet
            }
        });

        console.log(`Successfully updated description for video ID: ${videoId}`);
        return true;

    } catch (error) {
        console.error(`Failed to update video ${videoId}:`, error.message);
        if (error.response && error.response.data && error.response.data.error) {
            console.error('API Error Details:', JSON.stringify(error.response.data.error, null, 2));
        }
        return false;
    }
}

/**
 * Main execution function.
 */
async function main() {
    try {
        const auth = await authenticate();

        // 1. Check if local Excel file exists and use it if so (to preserve progress)
        let rows, workbook;

        if (fs.existsSync(LOCAL_XLSX_PATH)) {
            console.log(`Using existing local Excel file: ${LOCAL_XLSX_PATH}`);
            workbook = xlsx.readFile(LOCAL_XLSX_PATH);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            const headers = data[0];
            console.log('Headers found:', headers);

            rows = data.slice(1).map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = row[index];
                });
                // Fallback for missing header match on YouTube URL if needed (less likely on local re-read but good safety)
                if (!obj['YouTube URL'] && row.length > headers.length) {
                     const lastElement = row[row.length - 1];
                     if (typeof lastElement === 'string' && lastElement.includes('youtube.com')) {
                         obj['YouTube URL'] = lastElement;
                     }
                }
                return obj;
            });

        } else {
            // 2. Otherwise download from remote
            const result = await getRemoteExcelData();
            rows = result.rows;
            workbook = result.workbook;
            // Initialize local file immediately so we can start tracking updates
            initializeLocalExcel(workbook);
        }


        console.log(`Found ${rows.length} rows to process.`);

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            // Check if already updated
            if (row['updated_description'] && row['updated_description'].toString().toLowerCase() === 'yes') {
                // console.log(`Skipping row ${i+1}: Already updated.`);
                continue;
            }

            const youtubeUrl = row['YouTube URL'];
            const newDescription = row['YouTube Description'];

            if (!youtubeUrl) continue;

            const videoId = extractVideoId(youtubeUrl);
            if (!videoId) {
                console.warn(`Invalid YouTube URL: ${youtubeUrl}`);
                continue;
            }

            if (!newDescription) {
                console.warn(`No description found for video ID: ${videoId}`);
                continue;
            }

            console.log(`Processing Video ID: ${videoId}...`);
            const success = await updateVideoDescription(auth, videoId, newDescription);

            if (success) {
                await markVideoAsUpdated(row, i);
            }
        }

        console.log('All updates completed.');

    } catch (error) {
        console.error('Critical error:', error);
    }
}

main();
