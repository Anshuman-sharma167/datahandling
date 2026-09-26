'use strict';
const fs = require('node:fs');
const path = require('node:path');
const root = path.join(__dirname, '..');
const output = path.join(root, 'public');
fs.mkdirSync(output, { recursive: true });
for (const file of ['index.html', 'data-engine.js']) {
  fs.copyFileSync(path.join(root, file), path.join(output, file));
}
