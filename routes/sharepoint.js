const express = require('express');
const axios = require("axios");
const qs = require('qs');
const multer = require("multer");

const router = express.Router();
const upload = multer();
const graphAPI = "https://graph.microsoft.com/v1.0/sites";
require('dotenv').config();

const getToken = async() => {
  let data;

  await axios.post("https://login.microsoftonline.com/mcsurfacesinc.onmicrosoft.com/oauth2/v2.0/token", qs.stringify({
    grant_type: "client_credentials",
    client_id: process.env.CLIENT_ID,
    scope: "https://graph.microsoft.com/.default",
    client_secret: process.env.CLIENT_SECRET,
  }))
    .then(res => {
      data = res.data;
    })
    .catch(err => {
      console.log(err);
    });

  return data;
}

const getTerritoryId = (territory) => {
  switch (territory.toUpperCase()) {
    case "AUSTIN":
      return process.env.SP_AUSTIN_ID;
    case "DALLAS":
      return process.env.SP_DALLAS_ID;
    case "HOUSTON":
      return process.env.SP_HOUSTON_ID;
    case "SAN ANTONIO":
      return process.env.SP_SANANT_ID;
    default:
      return;
  }
}

// Get top level folder in 'Sage' SharePoint Site
router.get('/top-level', async function(req, res, next) {
  let tokenRes = await getToken();

  let axiosRes = await axios.get(`${graphAPI}/${process.env.SITE_ID}/drive`, {
    headers: {
      'Authorization': `Bearer ${tokenRes.access_token}`,
    }
  });

  res.send(axiosRes.data);
});

router.get('/client-folder', async function(req, res, next) {
  let tokenRes = await getToken();
  let territory_id;

  if (req.query.territory) {
    territory_id = getTerritoryId(req.query.territory.toUpperCase());
  } else {
    res.send({ message: 'No territory found.' });
  }

  let axiosRes = await axios.get(`${graphAPI}/${process.env.SITE_ID}/drive/items/${territory_id}/children`, {
    headers: {
      'Authorization': `Bearer ${tokenRes.access_token}`,
    }
  });

  res.send(axiosRes.data);
});

// Create file under specified parent
router.post('/file', upload.single("file"), async function(req, res, next) {
  let tokenRes = await getToken();
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
// body: name
router.post('/folder', async function(req, res, next) {
  let tokenRes = await getToken();
  let parentId = req.query.parentId;
  let body = {
    "name": req.body.name,
    "folder": {},
    "@microsoft.graph.conflictBehavior": "rename",
  }

  console.log(body);

  let axiosRes = await axios.post(`${graphAPI}/${process.env.SITE_ID}/drive/items/${parentId}/children`, body, {
    headers: {
      "Authorization": `Bearer ${tokenRes.access_token}`,
      "Content-Type": "application/json",
    }
  });

  res.send(axiosRes.data);
});

module.exports = router;
