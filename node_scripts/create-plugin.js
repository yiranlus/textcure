const fs = require('fs');
const archiver = require('archiver');
const path = require('path');

const buildDir = path.resolve(__dirname, '../textcure');
const output = fs.createWriteStream(path.resolve(__dirname, '../onlyoffice-textcure.plugin'));

const archive = archiver('zip', { zlib: { level: 9 } });

output.on('close', () => {
  console.log(`${archive.pointer()} total bytes`);
  console.log(`Plugin zip ready: onlyoffice-textcure.plugin`);
});

archive.pipe(output);
const resourcesDir = path.resolve(__dirname, '../resources');
archive.directory(buildDir, false);  // Zip entire build/
archive.directory(resourcesDir, 'resources');
archive.finalize();
