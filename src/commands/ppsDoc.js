const program = require('commander');
const path = require('path');
const fs = require('fs').promises;
const xml2js = require('xml2js');
const PermissionXlsxExportHandler = require('./pps2Xlsx');

const subDirs = ['profiles','permissionsets'];
const xml2jsParser = new xml2js.Parser();
generateDoc = async function(root,outDir){

    let exportHandler  = new PermissionXlsxExportHandler();
    await exportHandler.initialize('ppsExport.xlsx');

    for(let subDir of subDirs){
        let files = await fs.readdir(path.join(root,subDir));
        for(let fileName of files){
            if(fileName.indexOf('.profile')>-1 || fileName.indexOf('.permissionset')>-1){
                const content = await fs.readFile(path.join(root,subDir,fileName));
                //console.log(content);
                const jsContent = await xml2jsParser.parseStringPromise(content);
                //export it to excel or any other format
                await exportHandler.add(fileName,jsContent);
            }
        }
    }
    await exportHandler.save(outDir);
}



program
.command('pps-doc <source>')
.description('Genarate document for profile and permission sets')
.option('-o, --output <string>', 'output directory')
.action((source,options) => {
    console.log('Generating profile and permissionset document from '+ source)
    generateDoc(source,options.output ||'./').catch(err =>console.log(err));
});

module.export = program