// Require library
const fs = require('fs')
const fss = require('fs').promises
const readline = require('readline')
//const AWS = require('aws-sdk')
const events = require('events')
const xl = require('excel4node')
const sql = require('mssql')
const path = require('path')
const moment = require('moment')
const ejs = require('exceljs');
//AWS.config.update({ region: 'us-east-1' })
let dbServer
let dbName
let fsBaseFilePath
let dbSecret
let dataArr = []
let pathToHere = __dirname
let config = {}
/*const dbConfig = async () => {
    let envFileDirs = fs.readdirSync(pathToHere)
    envFileDirs = envFileDirs.filter(file => file === '.paymentSummRprtEnv')
    envFileDirs = envFileDirs[0]
    const readable = fs.createReadStream(`${pathToHere}/${envFileDirs}`);
    const reader = readline.createInterface({ input: readable });
    reader.on('line', (line) => {
        dataArr.push(line)
    });
    await events.once(reader, 'close');
    for (line of dataArr) {
        line = line.split('=')
        let firstVal = line[0]
        let secVal = line[1]?.replace(/['"]+/g, '')
        if (firstVal === 'DB_NAME') {
            dbName = secVal
        }
        if (firstVal === 'FACETS_DB_HOST') {
            dbServer = secVal
        }
        if (firstVal === 'DB_SECRET') {
            dbSecret = secVal
        }
        if (firstVal === 'FSX_PATH') {
            fsBaseFilePath = secVal
        }
    }
    config.user = undefined
    config.password = undefined
    config.server = dbServer
    config.database = dbName
    config.options = {
        trustedConnection: false,
        trustedDomain: true,
        trustServerCertificate: true
    }
    return config
}*/
 
/**
     * Method to get username password from secrets manager
     * @returns 
     */
/*const getDatabaseCredentials = async () => {
    try {
        let config = await dbConfig()
        const secretsManager = new AWS.SecretsManager();
        const secretData = await secretsManager.getSecretValue({ SecretId: dbSecret }).promise();
        const { username, password } = JSON.parse(secretData.SecretString);
        config.user = username
        config.password = password
        return config
    } catch (err) {
        console.error(`INVALID database credentials: ${err}`);
        console.log("EXECUTION UNSUCCESSFULL");
        fs.appendFileSync(`${fsBaseFilePath}error.log`, '\r\nINVALID database credentials')
        return 8
    }
}*/
 
 
const createTempExcel = async () => {
    try {
        /*console.log("***************Node JS Execution Started***************")
        let config = await getDatabaseCredentials()
        if (!config?.user || !config?.password) {
            console.log("Error while connecting to Database")
            return 8
        }
        let pool = await sql.connect(config)
        new sql.Request()
        console.log("***************Database Connected Successfully***************")
        /********** Database Connected successfully **********/
        
 
        /********** Converting data to formatted excel file **********/
        
        ///test only
       
        const wb = new xl.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        
        // Headers
        const headers = ['Receipt_Date', 'Paid_Date', 'Trading_Partner', 'Claim_Number'];
        
        // Hardcoded data
        const claimNumbers = ['24A000035800', '24A000035801', '24A000035802','','66','ff4fed'];
        
        // Add headers to the sheet
        headers.forEach((header, index) => {
          ws.cell(1, index + 1).string(header);
        });
        
        // Add hardcoded data under the Claim_Number column
        claimNumbers.forEach((claimNumber, index) => {
          ws.cell(index + 2, headers.indexOf('Claim_Number') + 1).string(claimNumber);
        });

       // let filePathOut =  `C:/Users/Toto/Desktop/releaseclaims/`
        ///testing
        //let fileName = (data.recordset[0].FILE_TITLE).toString()
       // const outO = path.join(filePathOut, `${fileName}.xlsx`)
       // wb.write(outO)
        
        // Save the workbook to a file
        wb.write('C:/Users/Toto/Desktop/releaseclaims/tempClaims.xlsx', (err, stats) => {
          if (err) {
            console.error(err);
          } else {
            console.log('Excel file created successfully: tempClaims.xlsx');
          }

        
        });    
        //////////new

 
        
        return 0
    } catch (err) {
        console.log(err)
        return 8
    }
 
 
}
const releaseClaims = async () => {
    try {
        /*console.log("***************Node JS Execution Started***************")
        let config = await getDatabaseCredentials()
        if (!config?.user || !config?.password) {
            console.log("Error while connecting to Database")
            return 8
        }
        let pool = await sql.connect(config)
        new sql.Request()
        console.log("***************Database Connected Successfully***************")
        /********** Database Connected successfully **********/
        
 
        /********** Converting data to formatted excel file **********/
        
        ///test only
       
        const workbook = new ejs.Workbook();
workbook.xlsx.readFile('C:/Users/Toto/Desktop/releaseclaims/tempClaims.xlsx')
  .then(() => {
    // Assuming 'Id' is in column A
    const worksheet = workbook.getWorksheet('Sheet1');
    const columnA = worksheet.getColumn('D');

    // Get the values from each cell in the 'Id' column
    const idList = [];
    const rejectedID = [];
    columnA.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
        // Skip the header row (assuming it's the first row)
        if (rowNumber > 1) {
          const id = cell.value;
  
          // Check the length of the ID
          if (id && id.length === 12) {
            idList.push(id);
          } else {
            // Add to rejectedID list if length is not 12
            rejectedID.push(id);
          }
        }
      });
  
      // Log the list of rejected IDs and count
      console.log('List of rejected IDs: ', rejectedID);
      console.log('Count of rejected IDs: ', rejectedID.length);
  
      // Now idList contains all the valid IDs
      console.log('List of valid IDs: ', idList);
  })
  .catch(error => {
    console.error(error.message);
  });
        //////////new

 
        
        return 0
    } catch (err) {
        console.log(err)
        return 8
    }
 
 
}
 
(async () => {
    const tempCreated = await createTempExcel()
   // if (tempCreated === 0) {
        const releaseSuccess = await releaseClaims()
   // }
   // console.log('Payment Summary Report>>>>>>>> ', outputRpt)
  /*  if (tempCreated === 0 && releaseSuccess === 0) {
        console.log('Return 0 to abf runbook')
        return 0
    } else {
        console.log('Return 8 to abf runbook')
        return 8
    }*/
})()