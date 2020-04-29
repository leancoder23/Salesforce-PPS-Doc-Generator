const xml2js = require('xml2js');
const fs = require('fs').promises;

module.exports = async(filePath)=>{
    const parser = new xml2js.Parser();
    const content = fs.readFile(filePath);
    return parser.parseString(content);
    
}
