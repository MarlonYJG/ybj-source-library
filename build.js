/*
 * @Author: Marlon
 * @Date: 2024-07-12 17:02:17
 * @Description: 
 */
const fs = require('fs');
const shell = require('shelljs');
const colors = require('colors')
const path = require('path');

if (shell.test('-d', 'dist')) {
  shell.rm('-rf', 'dist/*');
} else {
  shell.mkdir('-p', 'dist');
}

console.log(colors.yellow('Copying files from src to dist...'));
shell.cp('-R', 'src/*', 'dist/');
console.log(colors.green('Files copied successfully.'));

console.log(colors.yellow('Modifying import paths in .js files...'));
shell.find('dist').forEach(file => {
  if (file.endsWith('.js')) {
    let content = fs.readFileSync(file, 'utf8');

    content = content.replace(/import\s+(\w+|\{[^{}]*\})\s+from\s+(['"])(\.{1,2}\/[^'"]+)\2/g, (match, importName, quote, path) => {
      if (!path.endsWith('.min.js')) {
        return match.replace(path, `${path}.min.js`);
      } else {
        return match;
      }
    });
    fs.writeFileSync(file, content, 'utf8');
  }
});
console.log(colors.green('Import paths modified successfully.'));

console.log(colors.yellow('Minifying .js files...'));
shell.ls('dist/**/*.js').forEach(file => {
  if (!file.endsWith('.min.js')) {
    const dir = path.dirname(file);
    const baseName = path.basename(file, '.js');
    const minFile = path.join(dir, `${baseName}.min.js`);
    shell.mkdir('-p', path.dirname(minFile));

    const result = shell.exec(`terser ${file} -o ${minFile} --compress --mangle`);

    if (result.code !== 0) {
      console.log(colors.red(`Failed to minify ${file}: ${result.stderr}`));
    } else {
      console.log(colors.green(`Successfully minified ${file} to ${minFile}`));
    }
  }
});
console.log(colors.green('Minification completed.'));

console.log(colors.yellow('Deleting non-minified .js files...'));
shell.ls('dist/**/*').forEach(file => {
  if (file.endsWith('.js') && !file.endsWith('.min.js')) {
    shell.rm(file);
  }
});
console.log(colors.green('Non-minified .js files deleted.'));

console.log(colors.green.bold('All tasks completed successfully!'));
