const fs = require('fs');
const path = require('path');

// Read the bundled JavaScript
const jsPath = path.join(__dirname, '..', 'dist', 'ui.js');
const js = fs.readFileSync(jsPath, 'utf8');

// Read the HTML template
const htmlTemplatePath = path.join(__dirname, '..', 'src', 'ui.html');
const htmlTemplate = fs.readFileSync(htmlTemplatePath, 'utf8');

// Inject the bundled JS into the HTML
const finalHtml = htmlTemplate.replace(
  '<!-- BUNDLED_JS -->',
  `<script>${js}</script>`
);

// Write the final HTML
const outputPath = path.join(__dirname, '..', 'dist', 'ui.html');
fs.writeFileSync(outputPath, finalHtml);

console.log('UI HTML built successfully!');

