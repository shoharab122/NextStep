require('dotenv').config();
const express = require('express');
const path = require('path');
const cors = require('cors');
const apiRoutes = require('./routes/api');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static site (index.html, about.html, contact.html, admin.html, css/js)
app.use(express.static(path.join(__dirname, 'routes', 'public')));

// API routes (contact form, student applications, admin auth & dashboard, export)
app.use('/api', apiRoutes);

// Fallback routes for direct page loads
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'routes', 'public', 'index.html'));
});
app.get('/about', (req, res) => {
  res.sendFile(path.join(__dirname, 'routes', 'public', 'about.html'));
});
app.get('/contact', (req, res) => {
  res.sendFile(path.join(__dirname, 'routes', 'public', 'contact.html'));
});
app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'routes', 'public', 'admin.html'));
});

app.listen(PORT, () => {
  console.log(`NextStep Immigration server running at http://localhost:${PORT}`);
});