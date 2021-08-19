//const xlsx = require('xlsx')
const Excel = require('exceljs');
//const iconv = require('iconv-lite'); //한글깨짐방지

module.exports = async function (context, req) {
    
    // @files 엑셀 파일을 가져온다.

    context.log('JavaScript HTTP trigger function processed a request.');
    // read from a file
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("HANDY_HSO_Approval_ini_20210531.xlsx");
    let raw = [];
        workbook.eachSheet((sheet, sid)=> {
            //console.log('sheet name=' + sheet.name);
            if(sheet.name != 'globals.properties') {
                return;
            }
            sheet.eachRow((row, rid)=> {
                if(rid < 4) {
                    return;
                }
                
                //4행 옵션명, 설정값, 설명, Default, Version, 비고
                let optionname = '';
                let optionvalue = '';
                let optiondescription = '';
                let optiondefault = '';
                let optionversion = '';
                let optionetc = '';
                row.eachCell((cell, cid)=> {
                    if(cid == 1) {
                        optionname = cell.value
                    } else if(cid == 2) {
                        optionvalue = cell.value
                    } else if(cid == 3) {
                        optiondescription = cell.value
                    } else if(cid == 4) {
                        optiondefault = cell.value
                    } else if(cid == 5) {
                        optionversion = cell.value
                    } else if(cid == 6) {
                        optionetc = cell.value
                    }
                    //console.log('cid='+cid+' cell.value=' + cell.value)
                })
                raw.push({
                    'optionname':optionname, 'optionvalue':optionvalue, 'optiondescription':optiondescription,
                    'optiondefault':optiondefault, 'optionversion':optionversion, 'optionetc':optionetc
                });
            })
         })
      
    const name = (req.query.name || (req.body && req.body.name));

    context.res = {
        // status: 200, /* Defaults to 200 */
        headers: {"Access-Control-Allow-Origin": "*"},
        body: {data: raw}
        //res ini
    };
}