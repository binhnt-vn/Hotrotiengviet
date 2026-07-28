const fs = require('fs');
const path = require('path');

const filesToCopy = [
  'index.html',
  'app.js',
  'style.css',
  'app_standalone.html',
  'commands.html',
  'manifest.xml',
  'tu_dien_viet_tat.json'
];

const rootDir = path.join(__dirname, '..');
const destDir = path.join(rootDir, 'public');

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

filesToCopy.forEach(file => {
  const srcFile = path.join(rootDir, file);
  const destFile = path.join(destDir, file);
  if (fs.existsSync(srcFile)) {
    fs.copyFileSync(srcFile, destFile);
    console.log(`Copied ${file} -> public/${file}`);
  }
});

const assetsSrc = path.join(rootDir, 'assets');
const assetsDest = path.join(destDir, 'assets');
if (fs.existsSync(assetsSrc)) {
  fs.cpSync(assetsSrc, assetsDest, { recursive: true });
  console.log(`Copied assets/ -> public/assets/`);
}

console.log('Build completed successfully!');
