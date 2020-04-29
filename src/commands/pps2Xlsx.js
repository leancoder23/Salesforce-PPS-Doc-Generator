const xlsxPopulate = require('xlsx-populate');
const path = require('path');

class PermissionXlsxExportHandler {

    async initialize(fileName) {
        this.xlsxName = fileName;
        this.workbook = await xlsxPopulate.fromBlankAsync();
    }

    async  add(name, ppsJs) {
        if (!this.workbook)
            throw new Error('Export handler is not initialized');

        let pps = ppsJs.Profile || ppsJs.PermissionSet;


        if (!pps)
            throw new Error(`${name} not a valid salesforce profile or permission set`);

        let ppsName = name.substring(0, name.indexOf('.'));

        const sheet = this.workbook.addSheet(ppsName);

        /* WE will arrange the data in two logical column section
            First it logical column will display following attributes
                Object Permission  -  Column A to G
                    Col A - Object name
                    Col B - Allow read
                    Col C - Allow create
                    Col E - Allow edit
                    Col F - Allow read
                    Col F - View all records
                    Col G - Modify all records
                Record Type visibility - Column A to E
                    Col A - Object name
                    Col B - Record type name
                    Col C - Default flag
                    Col D - visible flag
                Layout assignment - Column A to C
                    Col A - Object name
                    Col B - record type
                    Col C - layout name
                Tab visibility - Column A to B
                    Col A - tab name
                    Col B - visibility    
                Page accesses  - Column A  (display only enabled) 
                    Col A - page name
                Class accesses - Column A  (display only enabled) 
                    Col A - Apex class name
                User permissions - Column A  (display only enabled)
                    Col A - permission name
            Second logical column will display field permission
                    Col I - Object name
                    Col J - Field name
                    Col K - Readable
                    Col L - Editable 
        */

        const getPropValue = function (prop) {
            if (Array.isArray(prop))
                return prop[0];
            else
                return prop;
        }

        let headerStyle = {
            bold: true,
            border: true,
            fill: {
                type: 'solid',
                color: {
                    rgb: '2c94ab'
                }
            },
            fontColor: 'ffffff',
            horizontalAlignment:'center',
            verticalAlignment:'center',
            fontSize:16    
        }
        let sectionHeaderStyle = {
            bold: true,
            border: true,
            fill: {
                type: 'solid',
                color: {
                    rgb: '2c94ab'
                }
            },
            fontColor: 'ffffff'
        }
        let sectionStyle = {
            border: true
        }

        const exportObjectPermission = function (rowIndex) {
            //ObjectAssignment
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Object Permissions');
            sheet.range(`A${rowIndex}:G${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('Name');
            sheet.cell(`B${rowIndex}`).value('Allow Read');
            sheet.cell(`C${rowIndex}`).value('Allow Create');
            sheet.cell(`D${rowIndex}`).value('Allow Edit');
            sheet.cell(`E${rowIndex}`).value('Allow Delete');
            sheet.cell(`F${rowIndex}`).value('View All');
            sheet.cell(`G${rowIndex}`).value('Modify All');
            sheet.range(`A${rowIndex}:G${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.objectPermissions || []) {
               
                let allowRead = getPropValue(ppsItem.allowRead)=='true';
                let allowCreate = getPropValue(ppsItem.allowCreate) =='true';
                let allowEdit = getPropValue(ppsItem.allowEdit)=='true';
                let allowDelete = getPropValue(ppsItem.allowDelete)=='true';
                let viewAllRecords = getPropValue(ppsItem.viewAllRecords)=='true';
                let modifyAllRecords = getPropValue(ppsItem.modifyAllRecords)=='true';

                if(allowRead||allowCreate||allowEdit||allowDelete||viewAllRecords||modifyAllRecords){
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.object));
                    sheet.cell(`B${rowIndex}`).value(allowRead);
                    sheet.cell(`C${rowIndex}`).value(allowCreate);
                    sheet.cell(`D${rowIndex}`).value(allowEdit);
                    sheet.cell(`E${rowIndex}`).value(allowDelete);
                    sheet.cell(`F${rowIndex}`).value(viewAllRecords);
                    sheet.cell(`G${rowIndex}`).value(modifyAllRecords);
                }
               
            }
            sheet.range(`A${sectionStartRow}:G${rowIndex}`).style(sectionStyle);

            return rowIndex;

        }

        const exportRecordTypeVisibility = function (rowIndex) {

            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Record type  visibility');
            sheet.range(`A${rowIndex}:D${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('Tab name');
            sheet.cell(`B${rowIndex}`).value('Record type');
            sheet.cell(`C${rowIndex}`).value('Default');
            sheet.cell(`D${rowIndex}`).value('Visible');
            sheet.range(`A${rowIndex}:D${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.recordTypeVisibilities || []) {
                rowIndex++;
                let recordTypeParts = getPropValue(ppsItem.recordType).split('.');
                sheet.cell(`A${rowIndex}`).value(recordTypeParts[0]);
                sheet.cell(`B${rowIndex}`).value(recordTypeParts[1]);
                sheet.cell(`C${rowIndex}`).value(getPropValue(ppsItem.default));
                sheet.cell(`D${rowIndex}`).value(getPropValue(ppsItem.visible));

            }
            sheet.range(`A${sectionStartRow}:D${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportLayoutVisibility = function (rowIndex) {
            //Layout assignment
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Layout  Assignment');
            sheet.range(`A${rowIndex}:C${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('Object name');
            sheet.cell(`B${rowIndex}`).value('Record type');
            sheet.cell(`C${rowIndex}`).value('Layout name');
            sheet.range(`A${rowIndex}:C${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.layoutAssignments || []) {
                rowIndex++;
                let recordType = getPropValue(ppsItem.recordType);
                if (recordType) {
                    let recordTypeParts = recordType.split('.');
                    sheet.cell(`A${rowIndex}`).value(recordTypeParts[0]);
                    sheet.cell(`B${rowIndex}`).value(recordTypeParts[1]);
                }
                sheet.cell(`C${rowIndex}`).value(getPropValue(ppsItem.layout));

            }
            sheet.range(`A${sectionStartRow}:C${rowIndex}`).style(sectionStyle);

            return rowIndex;

        }

        const exportTabVisibility = function (rowIndex) {
            //tab visiblity 
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Tab  Visibility');
            sheet.range(`A${rowIndex}:B${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('Tab name');
            sheet.cell(`B${rowIndex}`).value('Visible');
            sheet.range(`A${rowIndex}:B${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.tabVisibilities || pps.tabSettings || []) {
                rowIndex++;
                sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.tab));
                sheet.cell(`B${rowIndex}`).value(getPropValue(ppsItem.visibility));

            }
            sheet.range(`A${sectionStartRow}:B${rowIndex}`).style(sectionStyle);

            return rowIndex;

        }

        const exportApplicationVisiblity = function (rowIndex) {

            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Application visibilities');
            sheet.range(`A${rowIndex}:B${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('Application');
            sheet.cell(`B${rowIndex}`).value('Default');
            sheet.range(`A${rowIndex}:B${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.applicationVisibilities || []) {
                if (getPropValue(ppsItem.visible) == 'true') {
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.application));
                    sheet.cell(`B${rowIndex}`).value(getPropValue(ppsItem.default));
                }
            }

            sheet.range(`A${sectionStartRow}:B${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportPageAccess = function (rowIndex) {
            //Page Access 
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Page Access');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('name');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);



            for (let ppsItem of pps.pageAccesses || []) {
                if (getPropValue(ppsItem.enabled) == 'true') {
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.apexPage));
                }
            }
            sheet.range(`A${sectionStartRow}:A${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportClassAccess = function (rowIndex) {
            //Class Access 
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Apex Class Access');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('name');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.classAccesses || []) {
                if (getPropValue(ppsItem.enabled) == 'true') {
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.apexClass));
                }
            }
            sheet.range(`A${sectionStartRow}:A${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportCustomSettingAccesses = function (rowIndex) {
            //Page Access 
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('Custom Setting Accesses');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('name');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);



            for (let ppsItem of pps.customSettingAccesses || []) {
                if (getPropValue(ppsItem.enabled) == 'true') {
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.name));
                }
            }
            sheet.range(`A${sectionStartRow}:A${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportUserPermissions = function (rowIndex) {
            //Page Access 
            let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`A${rowIndex}`).value('User Permissions');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`A${rowIndex}`).value('name');
            sheet.cell(`A${rowIndex}`).style(sectionHeaderStyle);



            for (let ppsItem of pps.userPermissions || []) {
                if (getPropValue(ppsItem.enabled) == 'true') {
                    rowIndex++;
                    sheet.cell(`A${rowIndex}`).value(getPropValue(ppsItem.name));
                }
            }
            sheet.range(`A${sectionStartRow}:A${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        const exportFieldAccess = function (rowIndex) {
            //Now add field permission
           let sectionStartRow = rowIndex;
            //Header
            sheet.cell(`I${rowIndex}`).value('Field  Permissions');
            sheet.range(`I${rowIndex}:L${rowIndex}`).merged(true).style(sectionHeaderStyle);

            rowIndex++; //new line 
            sheet.cell(`I${rowIndex}`).value('Object name');
            sheet.cell(`J${rowIndex}`).value('Field Name');
            sheet.cell(`K${rowIndex}`).value('Readable');
            sheet.cell(`L${rowIndex}`).value('Editable');
            sheet.range(`I${rowIndex}:L${rowIndex}`).style(sectionHeaderStyle);

            for (let ppsItem of pps.fieldPermissions||[]) {
                
                let fieldParts = getPropValue(ppsItem.field).split('.');
                let readable = getPropValue(ppsItem.readable)=='true';
                let editable = getPropValue(ppsItem.editable)=='true';
                if(readable||editable){
                    rowIndex++;
                    sheet.cell(`I${rowIndex}`).value(fieldParts[0]);
                    sheet.cell(`J${rowIndex}`).value(fieldParts[1]);
                    sheet.cell(`K${rowIndex}`).value(readable);
                    sheet.cell(`L${rowIndex}`).value(editable);
                }
                

            }
            sheet.range(`I${sectionStartRow}:L${rowIndex}`).style(sectionStyle);

            return rowIndex;
        }

        //header row
        sheet.cell('A1').value(`${ppsName} - ${getPropValue(pps.description||'')}`);

        sheet.range('A1:L2').merged(true).style(headerStyle);

        let rowIndex = 3; //starting row
        rowIndex = exportObjectPermission(rowIndex);
        rowIndex = rowIndex + 2; //add padding
        rowIndex=exportRecordTypeVisibility(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportLayoutVisibility(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportTabVisibility(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportPageAccess(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportClassAccess(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportApplicationVisiblity(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportCustomSettingAccesses(rowIndex);
        rowIndex = rowIndex + 2; 
        rowIndex=exportUserPermissions(rowIndex);
        rowIndex = rowIndex + 2; 

        let fieldRowStartIndex = 3;
        exportFieldAccess(fieldRowStartIndex);

        //set column width
        sheet.column("A").width(40);
        sheet.column("B").width(40);
        sheet.column("C").width(40);
        sheet.column("D").width(10);
        sheet.column("E").width(10);
        sheet.column("F").width(10);
        sheet.column("G").width(10);
        sheet.column("H").width(10);
        sheet.column("I").width(40);
        sheet.column("J").width(40);
        sheet.column("K").width(10);
        sheet.column("L").width(10);

        
    }

    async save(outDir) {
        if (!this.workbook)
            throw new Error('Export handler is not initialized');

        //remove sheet 1 
        this.workbook.deleteSheet('Sheet1');

        await this.workbook.toFileAsync(path.join(outDir, this.xlsxName));
    }
}

module.exports = PermissionXlsxExportHandler;