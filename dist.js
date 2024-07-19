/*
 * @Author: Marlon
 * @Date: 2024-07-12 17:02:17
 * @Description: 
 */
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const colors = require('colors')

const packageJsonPath = path.resolve(__dirname, 'package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
packageJson.main = 'dist/index.min.js';
fs.writeFileSync(packageJsonPath, JSON.stringify(packageJson, null, 2));
console.log(colors.green('Successfully updated "main" field to dist/index.min.js.'));

console.log(colors.blue('Executing npm run build...'));
const buildProcess = exec('npm run build');

buildProcess.stdout.on('data', (data) => {
  process.stdout.write(data);
});

buildProcess.stderr.on('data', (data) => {
  process.stderr.write(data);
});

buildProcess.on('close', (code) => {
  if (code === 0) {
    console.log('npm run build completed successfully.');
  } else {
    console.error(`npm run build failed with code ${code}.`);
  }
});

