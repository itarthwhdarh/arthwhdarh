const fs = require('fs');
const path = require('path');

const emojiRegex = /[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F600}-\u{1F64F}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]/gu;

const files = [
  'cloud/index.html',
  'cloud/cloud.js',
  'cloud/cloud.css'
];

let totalEmojis = 0;

files.forEach(file => {
  const filePath = path.join(__dirname, '..', file);
  const content = fs.readFileSync(filePath, 'utf8');
  const matches = content.match(emojiRegex);
  if (matches) {
    console.log(`Found ${matches.length} emojis in ${file}:`, matches);
    totalEmojis += matches.length;
  } else {
    console.log(`No emojis found in ${file}`);
  }
});

console.log(`Emoji audit completed. Total emojis: ${totalEmojis}`);
