const path = require('path');
const file_path = path.dirname(__filename);

const fs = require('fs');

const Prism = require(path.join(file_path, 'prism.js'));

const code = Buffer.from(process.argv[2], 'base64').toString('utf-8');
const language = process.argv[3];

const highlightedCode = Prism.highlight(code, Prism.languages[language], language);

console.log(Buffer.from(highlightedCode, 'utf-8').toString('base64'));
