const express = require('express');
const axios = require("axios");
const qs = require('qs');
const multer = require("multer");
const jwt = require("jsonwebtoken");
const crypto = require("crypto");

const router = express.Router();
const upload = multer();
const graphAPI = "https://graph.microsoft.com/v1.0/sites";
require('dotenv').config();

const generateAccessToken = () => {
  let payload = {
    "aud": process.env.AUTH_URL,
    "exp": Math.floor(Date.now() / 1000) + 120,
    "iss": process.env.APP_ID,
    "jti": crypto.randomUUID(), // add to .env
    "nbf": Math.floor(Date.now() / 1000),
    "sub": process.env.APP_ID
  };
  let header = {
    "alg": "PS256",
    "typ": "JWT",
    "x5t": process.env.THUMBPRINT
  };

  return jwt.sign(
    payload,
    process.env.PRIVATE_KEY,
    { header: header }
  );
}

// Get token from Graph API
const getToken = async(token) => {
  let data;

  await axios.post("https://login.microsoftonline.com/mcsurfacesinc.onmicrosoft.com/oauth2/v2.0/token", qs.stringify({
    client_id: process.env.APP_ID,
    scope: "https://graph.microsoft.com/.default",
    grant_type: 'client_credentials',
    client_assertion: token,
    client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
  }))
    .then(res => {
      data = res.data;
    })
    .catch(err => {
      console.log(err);
    });

  return data;
}

// Get folder by ID
router.get('/folder', async (req, res) => {
  let token = generateAccessToken();
  let tokenRes = await getToken(token);

  let axiosRes = await axios.get(`${graphAPI}/${process.env.SITE_ID}/drive/items/${req.query.id}/children`, {
    headers: {
      'Authorization': `Bearer ${tokenRes.access_token}`,
    }
  });

  res.send(axiosRes.data);
})

// Get top level folder in 'Sage' SharePoint Site
router.get('/top-level', async function(req, res, next) {
  let token = generateAccessToken();
  let tokenRes = await getToken(token);

  let axiosRes = await axios.get(`${graphAPI}/${process.env.SITE_ID}/drive`, {
    headers: {
      'Authorization': `Bearer ${tokenRes.access_token}`,
    }
  });

  res.send(axiosRes.data);
});

// Get list of subfolders
// query: parentId
router.get('/folder-children', async function(req, res, next) {
  let token = generateAccessToken();
  let tokenRes = await getToken(token);

  let axiosRes = await axios.get(`${graphAPI}/${process.env.SITE_ID}/drive/items/${req.query.parentId}/children`, {
    headers: {
      'Authorization': `Bearer ${tokenRes.access_token}`,
    }
  });

  res.send(axiosRes.data);
});

// Create file under specified parent
// body : formData - file
// query: parentId
router.post('/file', upload.single("file"), async function(req, res, next) {
  let token = generateAccessToken();
  let tokenRes = await getToken(token);
  let file = req.file;
  let parentId = req.query.parentId;

  if (file.size >= 250000000) {
    res.send({ message: 'File size is too large.' });
  }

  let axiosRes = await axios.put(`${graphAPI}/${process.env.SITE_ID}/drive/items/${parentId}:/${file.originalname}:/content`, file, {
    headers: {
      "Authorization": `Bearer ${tokenRes.access_token}`,
      "Content-Type": "multipart/form-data",
    }
  });

  res.send(axiosRes.data);
});

// Create folder under specified parent
// query: parentId
// query: folder
router.post('/folder', async function(req, res, next) {
  let token = generateAccessToken();
  let tokenRes = await getToken(token);
  // let parentId = req.query.parentId;
  let body = {
    "name": req.query.folder,
    "folder": {},
    "@microsoft.graph.conflictBehavior": "rename",
  }

  console.log(req.query)

  let axiosRes = await axios.post(`${graphAPI}/${process.env.SITE_ID}/drive/items/${req.query.parentId}/children`, body, {
    headers: {
      "Authorization": `Bearer ${tokenRes.access_token}`,
      "Content-Type": "application/json",
    }
  });

  res.send(axiosRes.data);
});

module.exports = router;
