const http = require('http');
const fs = require('fs');

const indexHtml = fs.readFileSync('Index.html', 'utf8');
const stylesHtml = fs.readFileSync('Styles.html', 'utf8');
const componentsHtml = fs.readFileSync('Components.html', 'utf8');
const jsHtml = fs.readFileSync('JavaScript.html', 'utf8');
const codeGs = fs.readFileSync('Code.gs', 'utf8');

// Extremely basic template tag replacement for local dev
let finalHtml = indexHtml
  .replace('<?!= include(\'Styles\'); ?>', stylesHtml)
  .replace('<?!= include(\'Components\'); ?>', componentsHtml)
  .replace('<?!= include(\'JavaScript\'); ?>', jsHtml);

// Inject mock theme overrides css block that doGet would generate
const mockOverrides = `
html[data-theme="boring"] {
  --accent-color: #FF0000;
}
html[data-theme="anesthesia"] {
  --accent-color: #00FF00;
  --bg-color: #000000;
}
html[data-theme="ketamine"] {
  --accent-color: #0000FF;
}
`;
finalHtml = finalHtml.replace('<?!= themeOverridesCss ?>', mockOverrides);

const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/html' });
  res.end(finalHtml);
});

server.listen(3000, () => {
  console.log('Server running at http://localhost:3000/');
});
