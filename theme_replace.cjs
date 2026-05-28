const fs = require('fs');
const path = require('path');

function walk(dir) {
    let results = [];
    const list = fs.readdirSync(dir);
    list.forEach(function(file) {
        file = path.join(dir, file);
        const stat = fs.statSync(file);
        if (stat && stat.isDirectory()) { 
            results = results.concat(walk(file));
        } else { 
            if (file.endsWith('.tsx') || file.endsWith('.ts') || file.endsWith('.css') || file.endsWith('.html')) results.push(file);
        }
    });
    return results;
}

const files = walk('./src');

files.forEach(file => {
    let content = fs.readFileSync(file, 'utf8');
    
    // Replacements
    content = content.replace(/#FAF9F6/gi, '#ffffff'); // pure white for main backgrounds
    content = content.replace(/#E5E5DF/gi, '#e2e8f0'); // slate-200 for borders/subtle bg
    content = content.replace(/#5A5A40/gi, '#2563eb'); // blue-600 for primary elements
    content = content.replace(/#8C8C70/gi, '#64748b'); // slate-500 for secondary text
    content = content.replace(/#4A4A3C/gi, '#0f172a'); // slate-900 for primary text
    
    fs.writeFileSync(file, content, 'utf8');
});
console.log("Colors replaced successfully.");
