const fs = require('fs');

function mount() {
    const indexHtml = fs.readFileSync('Index.html', 'utf8');
    const stylesHtml = fs.readFileSync('Styles.html', 'utf8');
    const compHtml = fs.readFileSync('Components.html', 'utf8');
    const jsHtml = fs.readFileSync('JavaScript.html', 'utf8');

    let finalHtml = indexHtml;
    finalHtml = finalHtml.replace(/<\?!= include\('Styles'\); \?>/, stylesHtml.replace(/<style>/, '').replace(/<\/style>/, ''));
    finalHtml = finalHtml.replace(/<\?!= include\('Components'\); \?>/, compHtml);
    finalHtml = finalHtml.replace(/<\?!= include\('JavaScript'\); \?>/, jsHtml);

    fs.writeFileSync('tests/merged.html', finalHtml);
}

module.exports = { mount };
