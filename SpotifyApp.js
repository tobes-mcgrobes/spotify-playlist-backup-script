const express = require('express');
const querystring = require('querystring');
const axios = require('axios');
const bodyParser = require('body-parser');
const session = require('express-session');
const fs = require('fs');
const xlsx = require('xlsx');

const client_id = 'YOUR_CLIENT_ID'; // Replace with your client ID
const client_secret = 'YOUR_CLIENT_SECRET'; // Replace with your client secret
const redirect_uri = 'http://localhost:8888/callback'; // Ensure this matches the registered redirect URI

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

const sessionSecret = 'your_super_secret_key'; // Replace this with your own secret key

app.use(session({
  secret: sessionSecret,
  resave: false,
  saveUninitialized: true,
  cookie: { secure: false } // Set to true if using https
}));

let playlistIDs = []; // Global variable to store playlist IDs

function generateRandomString(length) {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  
  for (let i = 0; i < length; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

// Sanitize sheet name function
function sanitizeSheetName(name) {
    return name.replace(/[:\\\/?*[\]]/g, '_'); // Replace disallowed characters with underscores
}

app.get('/login', (req, res) => {
  const state = generateRandomString(16);
  const scope = 'user-read-private user-read-email playlist-read-private';

  res.redirect('https://accounts.spotify.com/authorize?' +
    querystring.stringify({
      response_type: 'code',
      client_id: client_id,
      scope: scope,
      redirect_uri: redirect_uri,
      state: state
    }));
});

app.get('/callback', async (req, res) => {
  const code = req.query.code || null;

  if (!code) {
    res.send('No code found in query parameters');
    return;
  }

  try {
    const response = await axios.post('https://accounts.spotify.com/api/token', querystring.stringify({
      code: code,
      redirect_uri: redirect_uri,
      grant_type: 'authorization_code'
    }), {
      headers: {
        'Authorization': 'Basic ' + Buffer.from(client_id + ':' + client_secret).toString('base64'),
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });

    req.session.access_token = response.data.access_token;
    req.session.refresh_token = response.data.refresh_token;

    res.send('Authorization successful. You can now access the /playlists endpoint.');
  } catch (error) {
    console.error('Error getting token:', error);
    res.send(`Error: ${error.response ? error.response.data.error_description : error.message}`);
  }
});

app.get('/playlists', async (req, res) => {
    const access_token = req.session.access_token;
  
    if (!access_token) {
      res.send('No access token available');
      return;
    }
  
    try {
      const response = await axios.get('https://api.spotify.com/v1/me/playlists', {
        headers: {
          'Authorization': `Bearer ${access_token}`
        }
      });
  
      // Extract playlist IDs and details
      playlistIDs = response.data.items.map(playlist => ({
        id: playlist.id,
        name: playlist.name,
        description: playlist.description,
        imageUrl: (playlist.images && playlist.images.length > 0) ? playlist.images[0].url : ''
      }));
  
      // Create a new workbook
      const workbook = xlsx.utils.book_new();
  
      // Add summary sheet
      const summarySheetData = playlistIDs.map(playlist => ({
        Name: playlist.name,
        Description: playlist.description,
        'Image URL': playlist.imageUrl
      }));
      const summarySheet = xlsx.utils.json_to_sheet(summarySheetData);
      xlsx.utils.book_append_sheet(workbook, summarySheet, 'Summary');
  
      for (const playlist of playlistIDs) {
        try {
          const tracksResponse = await axios.get(`https://api.spotify.com/v1/playlists/${playlist.id}/tracks`, {
            headers: {
              'Authorization': `Bearer ${access_token}`
            }
          });
  
          const trackDetails = tracksResponse.data.items.map(item => {
            const track = item.track;
            if (track) {
              return {
                'Playlist Name': playlist.name,
                'Track Name': track.name,
                'Album Name': track.album.name,
                'Artist Name': track.artists.map(artist => artist.name).join(', ')
              };
            } else {
              return {
                'Playlist Name': playlist.name,
                'Track Name': 'Unknown',
                'Album Name': 'Unknown',
                'Artist Name': 'Unknown'
              };
            }
          });
  
          // Truncate the playlist name if it exceeds 31 characters
          const sheetName = sanitizeSheetName(playlist.name).substring(0, 31); // Truncate to 31 characters after sanitization
  
          // Convert track details to a worksheet
          const trackSheet = xlsx.utils.json_to_sheet(trackDetails);
          // Add the worksheet to the workbook with the truncated playlist name as the sheet name
          xlsx.utils.book_append_sheet(workbook, trackSheet, sheetName);
        } catch (trackError) {
          console.error(`Error fetching tracks for playlist ${playlist.id}:`, trackError);
        }
      }
  
      // Write the workbook to a file
      xlsx.writeFile(workbook, 'tracks.xlsx');
  
      res.send('Tracks from all playlists have been saved to tracks.xlsx.');
    } catch (error) {
      console.error('Error fetching playlists:', error);
      res.send(`Error: ${error.response ? error.response.data.error.message : error.message}`);
    }
  });
  

const PORT = process.env.PORT || 8888;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
