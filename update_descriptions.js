
import { google } from 'googleapis';
import { promises as fs } from 'fs';
import path from 'path';
import process from 'process';
import axios from 'axios';
import xlsx from 'xlsx';
import { createInterface } from 'readline';

// --- CONFIGURATION ---
const CLIENT_SECRETS_PATH = 'client_secrets.json';
const TOKEN_PATH = 'token.json';
const REMOTE_XLSX_URL = 'https://raw.githubusercontent.com/swipeviewer99-ai/video-uploader-new/main/SEOVideos_updated.xlsx';

const SCOPES = ['https://www.googleapis.com/auth/youtube.force-ssl'];

const youtube = google.youtube('v3');
const OAuth2 = google.auth.OAuth2;

/**
 * Reads the local client secrets file.
 */
async function getClientSecrets() {
    try {
        const content = await fs.readFile(CLIENT_SECRETS_PATH);
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

        return rows;
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
 * Updates the video description on YouTube.
 */
async function updateVideoDescription(auth, videoId, newDescription) {
    if (!auth) {
        console.log(`[DRY RUN] Would update video ID: ${videoId}`);
        // console.log(`[DRY RUN] New Description Preview: ${newDescription.substring(0, 50)}...`);
        return;
    }

    // Truncate description if it's too long (safeguard)
    if (newDescription.length > 4900) {
        console.warn(`WARNING: Description for video ${videoId} is ${newDescription.length} chars. Truncating to 4900.`);
        newDescription = newDescription.substring(0, 4900);
    }

    // Check for angle brackets (basic HTML check)
    if (newDescription.includes('<') || newDescription.includes('>')) {
        console.warn(`WARNING: Description for video ${videoId} contains potentially invalid characters (< or >).`);
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
            return;
        }

        const videoSnippet = response.data.items[0].snippet;
        const currentDescription = videoSnippet.description;

        if (currentDescription === newDescription) {
            console.log(`Description for video ${videoId} is already up to date.`);
            return;
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

    } catch (error) {
        console.error(`Failed to update video ${videoId}:`, error.message);
        // Log more specific error details if available
        if (error.response && error.response.data && error.response.data.error) {
            console.error('API Error Details:', JSON.stringify(error.response.data.error, null, 2));
        } else if (error.errors) {
             console.error('Error Object:', JSON.stringify(error.errors, null, 2));
        }
    }
}

/**
 * Main execution function.
 */
async function main() {
    try {
        const auth = await authenticate();
        const videoData = await getRemoteExcelData();

        console.log(`Found ${videoData.length} rows in Excel.`);

        // Log first few rows description length for debugging
        if (videoData.length > 0) {
             console.log('--- DEBUG: First few rows description check ---');
             for(let i=0; i<Math.min(3, videoData.length); i++) {
                 const desc = videoData[i]['YouTube Description'] || '';
                 const id = extractVideoId(videoData[i]['YouTube URL']);
                 console.log(`Row ${i+1} (ID: ${id}): Length=${desc.length} chars. Contains '<'=${desc.includes('<')}, Contains '>'=${desc.includes('>')}`);
             }
             console.log('--- END DEBUG ---');
        }


        for (const row of videoData) {
            const youtubeUrl = row['YouTube URL'];
            const newDescription = row['YouTube Description'];

            if (!youtubeUrl) {
                // console.log('Skipping row: No YouTube URL found.', row);
                continue;
            }

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
            await updateVideoDescription(auth, videoId, newDescription);
        }

        console.log('All updates completed.');

    } catch (error) {
        console.error('Critical error:', error);
    }
}

main();
