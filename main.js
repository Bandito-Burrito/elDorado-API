const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron')
const sqlite3 = require('sqlite3').verbose();
require('dotenv').config();
const XLSX = require('xlsx');
const path = require('path');
var parseFullName = require('parse-full-name').parseFullName;
var stringSimilarity = require("string-similarity");


var data = [];
var expandedData = [];
var ocAPIcount = 0;
var endAPIcount = 0;

function createWindow () {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  })

  win.loadFile('index.html')

  const originalConsoleLog = console.log;
  console.log = (...args) => {
    originalConsoleLog(...args);
    if (win && win.webContents) {
      win.webContents.send('log', args.join(' '));
    }
  };

}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

ipcMain.handle('open-file-dialog', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile']
  })
  
  if (result.canceled) {
    return null
  } else {
    return result.filePaths[0]
  }
})

ipcMain.on('selected-file', (event, filePath) => {

  if (!filePath) {
    return
  }

  const workbook = XLSX.readFile(filePath);

  if (workbook.Sheets["Sheet1"]) {
    var worksheet = workbook.Sheets["Sheet1"];

    console.log('\nChecking for Names in column NAME');
    data = XLSX.utils.sheet_to_json(worksheet);

    var namesPresent = false;
    var firmsPresent = false;

    var minRows = Math.min(data.length, 5);

    for (var i = 0; i < minRows; i++) {
      if (data[i]['NAME']) {
        namesPresent = true;
        break;
      }
    }

    for (var i = 0; i < minRows; i++) {
      if (data[i]['FIRM']) {
        firmsPresent = true;
        break;
      }
    }

    if (namesPresent) {
      console.log('Names found in column NAME. Proceeding with People API Call.');

      let businessCall = false;

      if(firmsPresent){
        businessCall = true;
      }

      main(filePath, businessCall)
      .then(() => {
        console.log("\n");
        console.log('👍👍👍Execution completed successfully with no errors.👍👍👍');
        console.log('❗❗❗ Remember to wait one minute between large calls to avoid rate limiting ❗❗❗');
      })
      .catch(err => {
        if (err) {
          console.log("\n");
          console.log('❗❗❗ An error occurred:', err);
          console.log('If a large amount of records was sent to the api you may want to wait a minute to avoid rate limiting');
        }
      });
    }
    else {
      (async () => {
        console.log('No Names found in first 5 rows of column NAME. Proceeding with Business API Call.');
    
        try {
          await main2(filePath);
          console.log("\n");
          console.log('👍👍👍 Business API call completed successfully with no errors. 👍👍👍');
          console.log('❗❗❗ Remember to wait one minute between large calls to avoid rate limiting ❗❗❗');
    
          let businessCall = true;
          console.log('\nNow proceeding to Customer API');
    
          await main(filePath, businessCall);
          console.log("\n");
          console.log('👍👍👍 Customer API call completed successfully with no errors. 👍👍👍');
          console.log('❗❗❗ Remember to wait one minute between large calls to avoid rate limiting ❗❗❗');
        } catch (err) {
          console.log("\n");
          console.log('❗❗❗ An error occurred:', err);
          console.log('If a large amount of records was sent to the API, you may want to wait a minute to avoid rate limiting');
        }
      })();
    }
  }
  else
  {
    console.log('No Sheet1 found in the Excel file. Ending execution.');
    return
  }


  })


async function makeApiCall(FirstName, LastName, addressLine1, addressLine2, rowOfExcel,db, businessCall) {

  var firstMobile = "";
  var firstLandline = "";
  var secondMobile = "";
  var matchValue = "";

  if (endAPIcount % 100 === 0 && endAPIcount !== 0) {
    console.log(`Pausing for 1 minute at index ${rowOfExcel} to respect rate limits.`);
    await new Promise(resolve => setTimeout(resolve, 60000)); // 1 minute delay
  }



  const options = {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'galaxy-ap-name': process.env.AP_NAME,
      'galaxy-ap-password': process.env.AP_PASS,
      'galaxy-search-type': 'DevAPIContactEnrich',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      "FirstName": FirstName,
      "LastName": LastName,
      "Address": {
        "addressLine1": addressLine1,
        "addressLine2": addressLine2
      }
    })
  };

  try {

    const response = await fetch('https://devapi.endato.com/Contact/Enrich', options);

    const jsonResponse = await response.json();
    endAPIcount++;

    if (response.status !== 200) {
      console.log('API Error:', response.status);
      console.error('API Error:', response.status);
      return response.status;
    }

    if (jsonResponse && jsonResponse.person) {
      data[rowOfExcel] ['MATCH'] = 'Yes';
      matchValue = 'Yes';

      var firstMobileMatch = false;
      var firstLandlineMatch = false;
      var secondMobileMatch = false;
    

      jsonResponse.person.phones.forEach((phone,index) => {

        if (firstMobileMatch && firstLandlineMatch) {
          return false; // break
        }

        if (phone.type === 'mobile' && !firstMobileMatch) {
          data[rowOfExcel] ['First Mobile'] = phone.number;
          firstMobile = phone.number;
          firstMobileMatch = true;
          console.log('-   First Mobile:', phone.number);
        }
        if (phone.type === 'landline' && !firstLandlineMatch) {
          data[rowOfExcel] ['First Landline'] = phone.number;
          firstLandline = phone.number;
          firstLandlineMatch = true;
          console.log('-   First Landline:', phone.number);
        }
        if (phone.type === 'mobile' && firstMobileMatch && index == 1) {
          data[rowOfExcel] ['Second Mobile (if applicable)'] = phone.number;
          secondMobile = phone.number;
          secondMobileMatch = true;
          console.log('-   Second Mobile:', phone.number);
        }
      });
      
      if (!firstMobileMatch && !firstLandlineMatch && !secondMobileMatch) {
        data[rowOfExcel] ['MATCH'] = 'Match No Phone';
        matchValue = 'Match No Phone';
      }

      }
    else {
      data[rowOfExcel] ['MATCH'] = jsonResponse.message
      console.log("-   "+jsonResponse.message);
      matchValue = jsonResponse.message;
    }
    
    if(!businessCall){
    await insertRecord(FirstName, LastName, addressLine1, addressLine2, firstMobile, firstLandline, secondMobile, matchValue, db);
    }else{
    await insertRecord3(data[rowOfExcel]['FIRM'], FirstName, LastName, addressLine2, firstMobile, firstLandline, secondMobile, matchValue, db);
    }  
    return jsonResponse;
  } catch (err) {
    // Handle any errors that occur during the fetch
    console.error('Error during API call:', err);
    console.log('Error: ', err.message);
  }
}


async function processRowsSequentially(data,db,businessCall) {
  for (const [index , row] of data.entries()) {

    const FirstName = row['FIRST'];
    const LastName = row['LAST'];

    if (businessCall){
    var addressLine1 = null;
    var addressLine2 = `${row['ZIP5']}`;

    }else{
    var addressLine1 = row['MAILING'];
    var addressLine2 = `${row['CITY']}, ${row['STATE']}`;
    }

    console.log('\nProcessing:', index+1, FirstName, LastName, addressLine1, addressLine2);

    if (!businessCall){
    if (!FirstName || !LastName || !addressLine1 || !addressLine2) {
      console.log('Skipping row due to missing data:');
      data[index]['MATCH'] = 'Incomplete Data';
      continue;
    }
    }
    else{
      if (!FirstName || !LastName || !addressLine2) {
        console.log('Skipping row due to missing data:');
        data[index]['MATCH'] = 'Incomplete Data';
        continue;
      }
    }

    if(!businessCall){
    const isDuplicate = await checkForDuplicate(FirstName, LastName, addressLine1, addressLine2,db);
    if (isDuplicate) {
      console.log('Duplicate record found. Skipping API call');

      const duplicateFields = await getAdditionalFields(FirstName, LastName, addressLine1, addressLine2,  db);

      data[index]['MATCH'] = 'Duplicate';
      data[index]['First Mobile'] = duplicateFields.FirstMobile || '';
      data[index]['First Landline'] = duplicateFields.FirstLandline || '';
      data[index]['Second Mobile (if applicable)'] = duplicateFields.SecondMobile || '';
      data[index]['Initial Match'] = duplicateFields.InitialMatch || '';

      continue;
    }
    }else{
    const isDuplicate = await checkForDuplicate3(row['FIRM'], FirstName, LastName, addressLine2, db);
    if (isDuplicate) {
      console.log('Duplicate record found. Skipping API call');
      
      const duplicateFields = await getAdditionalFields3(row['FIRM'], FirstName, LastName, addressLine2, db);

      data[index]['MATCH'] = 'Duplicate';
      data[index]['First Mobile'] = duplicateFields.FirstMobile || '';
      data[index]['First Landline'] = duplicateFields.FirstLandline || '';
      data[index]['Second Mobile (if applicable)'] = duplicateFields.SecondMobile || '';
      data[index]['Initial Match'] = duplicateFields.InitialMatch || '';

      continue;
    }
    }

    const status =  await makeApiCall(FirstName, LastName, addressLine1, addressLine2, index, db, businessCall);
    if (status == 403){
      console.log('Rate Limit Exceeded. Ending Execution');
      break;
    }


    
  }
}

function checkForDuplicate(FirstName, LastName, AddressLine1, AddressLine2,db) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT COUNT(*) as count FROM records
      WHERE FirstName = ? AND LastName = ? AND AddressLine1 = ? AND AddressLine2 = ?
    `;
    db.get(query, [FirstName, LastName, AddressLine1, AddressLine2], (err, row) => {
      if (err) {
        console.error('Error checking for duplicate:', err.message);
        reject(err);
      } else {
        resolve(row.count > 0);
      }
    });
  });
}

function checkForDuplicate3(Firm, FirstName, LastName, AddressLine2, db) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT COUNT(*) as count FROM records
      WHERE Firm = ? AND FirstName = ? AND LastName = ? AND AddressLine2 = ?
    `;
    db.get(query, [Firm, FirstName, LastName, AddressLine2], (err, row) => {
      if (err) {
        console.error('Error checking for duplicate:', err.message);
        reject(err);
      } else {
        resolve(row.count > 0);
      }
    });
  });
}


function getAdditionalFields(FirstName, LastName, AddressLine1, AddressLine2, db) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT FirstMobile, FirstLandline, SecondMobile, InitialMatch FROM records
      WHERE FirstName = ? AND LastName = ? AND AddressLine1 = ? AND AddressLine2 = ?
    `;
    db.get(query, [FirstName, LastName, AddressLine1, AddressLine2], (err, row) => {
      if (err) {
        console.error('Error fetching additional fields:', err.message);
        reject(err);
      } else {
        resolve(row);
      }
    });
  });
}

function getAdditionalFields3(Firm, FirstName, LastName, AddressLine2, db) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT FirstMobile, FirstLandline, SecondMobile, InitialMatch FROM records
      WHERE Firm = ? AND FirstName = ? AND LastName = ? AND AddressLine2 = ?
    `;
    db.get(query, [Firm, FirstName, LastName, AddressLine2], (err, row) => {
      if (err) {
        console.error('Error fetching additional fields:', err.message);
        reject(err);
      } else {
        resolve(row);
      }
    });
  });
}


function insertRecord(FirstName, LastName, AddressLine1, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch, db) {
  return new Promise((resolve, reject) => {
    const query = `
      INSERT INTO records (FirstName, LastName, AddressLine1, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `;
    db.run(query, [FirstName, LastName, AddressLine1, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch], function(err) {
      if (err) {
        // Handle unique constraint violation (duplicate record)
        if (err.message.includes('UNIQUE constraint failed')) {
          console.log('Record already exists in the database.');
          resolve(false);
        } else {
          console.error('Error inserting record:', err.message);
          reject(err);
        }
      } else {
        console.log('Record inserted into database, index:', this.lastID);
        resolve(true);
      }
    });
  });
}

function insertRecord3(Firm, FirstName, LastName, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch, db) {
  return new Promise((resolve, reject) => {
    const query = `
      INSERT INTO records (Firm, FirstName, LastName, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `;
    db.run(query, [Firm, FirstName, LastName, AddressLine2, FirstMobile, FirstLandline, SecondMobile, InitialMatch], function(err) {
      if (err) {
        // Handle unique constraint violation (duplicate record)
        if (err.message.includes('UNIQUE constraint failed')) {
          console.log('Record already exists in the database.');
          resolve(false);
        } else {
          console.error('Error inserting record:', err.message);
          reject(err);
        }
      } else {
        console.log('Record inserted into database, index:', this.lastID);
        resolve(true);
      }
    });
  });
}

async function setupDatabase() {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database('./records.db', (err) => {
      if (err) {
        console.error('Error opening database:', err.message);
        reject(err); // Reject promise if there's an error
      } else {
        console.log('\nConnected to the SQLite people database.');

        db.serialize(() => {
          db.run(`
            CREATE TABLE IF NOT EXISTS records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              FirstName TEXT NOT NULL,
              LastName TEXT NOT NULL,
              AddressLine1 TEXT NOT NULL,
              AddressLine2 TEXT NOT NULL,
              FirstMobile TEXT,   
              FirstLandline TEXT, 
              SecondMobile TEXT, 
              InitialMatch TEXT, 
              UNIQUE(FirstName, LastName, AddressLine1, AddressLine2)
            )
          `, (err) => {
            if (err) {
              console.error('Error creating table:', err.message);
              reject(err); // Reject if table creation fails
            } else {
              resolve(db); // Resolve the promise when successful
            }
          });
        });
      }
    });
  });
}

async function setupDatabase3() {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database('./businessPersonal.db', (err) => {
      if (err) {
        console.error('Error opening database:', err.message);
        reject(err); // Reject promise if there's an error
      } else {
        console.log('\nConnected to the SQLite business-people database.');

        db.serialize(() => {
          db.run(`
            CREATE TABLE IF NOT EXISTS records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              Firm TEXT NOT NULL,
              FirstName TEXT NOT NULL,
              LastName TEXT NOT NULL,
              AddressLine2 TEXT NOT NULL,
              FirstMobile TEXT,   
              FirstLandline TEXT, 
              SecondMobile TEXT, 
              InitialMatch TEXT, 
              UNIQUE(Firm, FirstName, LastName, AddressLine2)
            )
          `, (err) => {
            if (err) {
              console.error('Error creating table:', err.message);
              reject(err); // Reject if table creation fails
            } else {
              resolve(db); // Resolve the promise when successful
            }
          });
        });
      }
    });
  });
}

ipcMain.on('open-instructions', () => {

  const instructionsPath = path.join(__dirname, 'instructions.txt'); 

  shell.openPath(instructionsPath)
      .then(response => {
          if (response) {
              console.error('Error opening file:', response);
          }
      });

    });

async function main(filePath, businessCall) {

  if(businessCall){
  var db = await setupDatabase3(); //unique records for business calls are firm fname lname and zip only
  }else{
  var db = await setupDatabase(); //unique records for personal calls are fname lname and full address
  }

  endAPIcount = 0;
  ocAPIcount = 0;

  const workbook = XLSX.readFile(filePath);



  if (businessCall){
    var worksheet = workbook.Sheets["Business Match"];
    console.log('\nPulling data from "Business Match" of the Excel file...');
  }
  else
  {
    var worksheet = workbook.Sheets["Sheet1"];
    console.log('\nPulling data from "Sheet1" of the Excel file...');
  }


  data = XLSX.utils.sheet_to_json(worksheet);
  

  await processRowsSequentially(data,db,businessCall);

  if (businessCall){

  let previousCompanyName ="";
  let previousCompanyOfficerName = "";
  let previousCompanyOfficerPosition = "";
  let previousCompanyFM = "";
  let previousCompanyFL = "";
  let previousCompanySM = "";


  var newData = [];

  data.forEach((firm) => {

    if (previousCompanyName === firm.FIRM && previousCompanyName !== "") {
      const newDataLength = newData.length;
      newData[newDataLength - 1]['Officer Name 2'] = firm['Officer Name'];
      newData[newDataLength - 1]['Officer Position 2'] = firm['Officer Position'];
      newData[newDataLength - 1]['First Mobile 2'] = firm['First Mobile'];
      newData[newDataLength - 1]['First Landline 2'] = firm['First Landline'];
      newData[newDataLength - 1]['Second Mobile 2'] = firm['Second Mobile (if applicable)'];
      if(firm['MATCH'] === 'Yes'){
        newData[newDataLength - 1]['MATCH'] = 'Yes';
      }
    }
    else{
      newData.push({ ...firm });
    }
    previousCompanyName = firm['FIRM'];
  });

  }

  var matchData=[];
  var noMatchData=[];
  var matchNoNumber=[];
  var duplicateData=[];
  var incompleteData=[];
  var notProcessed=[];


if (!businessCall){
data.forEach((row) => {
  switch (row.MATCH) {
    case "Yes":
      matchData.push(row);
      break;
      
    case "Match No Phone":
      matchNoNumber.push(row);
      break;
      
    case "Duplicate":
      duplicateData.push(row);
      break;

    case "Incomplete Data":
      incompleteData.push(row);
      break;
    
    case "":
      notProcessed.push(row);
      break;
      
    default:
      noMatchData.push(row);
      break;
  }
});
}else{
  newData.forEach((row) => {
    switch (row.MATCH) {
      case "Yes":
        matchData.push(row);
        break;
        
      case "Match No Phone":
        matchNoNumber.push(row);
        break;
        
      case "Duplicate":
        duplicateData.push(row);
        break;
  
      case "Incomplete Data":
        incompleteData.push(row);
        break;
      
      case "":
        notProcessed.push(row);
        break;
        
      default:
        noMatchData.push(row);
        break;
    }
  });
}


const matchWorksheet = XLSX.utils.json_to_sheet(matchData);
const noMatchWorksheet = XLSX.utils.json_to_sheet(noMatchData);
const matchNoNumberWorksheet = XLSX.utils.json_to_sheet(matchNoNumber);
const duplicateWorksheet = XLSX.utils.json_to_sheet(duplicateData);
const incompleteDataWorksheet = XLSX.utils.json_to_sheet(incompleteData);
const notProcessedWorksheet = XLSX.utils.json_to_sheet(notProcessed);
  

  console.log("\n");
  if (workbook.Sheets["Match"]) {
    // Sheet exists, overwrite it
    workbook.Sheets["Match"] = matchWorksheet;
    console.log(`Overwrote existing Excel sheet: Match`);
  } else {
    // Sheet doesn't exist, append new sheet
    XLSX.utils.book_append_sheet(workbook, matchWorksheet, "Match");
    console.log(`Created new Excel sheet: Match`);
  }

  if (workbook.Sheets["No Match"]) {

    workbook.Sheets["No Match"] = noMatchWorksheet;
    console.log(`Overwrote existing Excel sheet: No Match`);
  } else {

    XLSX.utils.book_append_sheet(workbook, noMatchWorksheet, "No Match");
    console.log(`Created new Excel sheet: No Match`);
  }

  if (workbook.Sheets["Match No Phone"]) {

    workbook.Sheets["Match No Phone"] = matchNoNumberWorksheet;
    console.log(`Overwrote existing Excel sheet: Match No Phone`);
  } else {

    XLSX.utils.book_append_sheet(workbook, matchNoNumberWorksheet, "Match No Phone");
    console.log(`Created new Excel sheet: Match No Phone`);
  }

  if (workbook.Sheets["Duplicates"]) {

    workbook.Sheets["Duplicates"] = duplicateWorksheet;
    console.log(`Overwrote existing Excel sheet: Duplicates`);
  } else {

    XLSX.utils.book_append_sheet(workbook, duplicateWorksheet, "Duplicates");
    console.log(`Created new Excel sheet: Duplicates`);
  }

  if (workbook.Sheets["Incomplete Data"]) {

    workbook.Sheets["Incomplete Data"] = incompleteDataWorksheet;
    console.log(`Overwrote existing Excel sheet: Incomplete Data`);
  } else {

    XLSX.utils.book_append_sheet(workbook, incompleteDataWorksheet, "Incomplete Data");
    console.log(`Created new Excel sheet: Incomplete Data`);
  }

  if (workbook.Sheets["Not Processed"]) {

    workbook.Sheets["Not Processed"] = notProcessedWorksheet;
    console.log(`Overwrote existing Excel sheet: Not Processed`);
  } else {

    XLSX.utils.book_append_sheet(workbook, notProcessedWorksheet, "Not Processed");
    console.log(`Created new Excel sheet: Not Processed`);
  }


  XLSX.writeFile(workbook, filePath);


  db.close((err) => {
    if (err) {
      console.error('\nError closing the database:', err.message);
    } else {
      console.log('\nDatabase connection closed.');
    }
  });

}

//second api call








///////////////////////////////////////////////////////////////////////////////////////////////











//second api call

function insertOfficer(recordId, OfficerName, Title, dbBusiness) {
  return new Promise((resolve, reject) => {
    const query = `
      INSERT INTO officers (record_id, OfficerName, Title)
      VALUES (?, ?, ?)
    `;
    dbBusiness.run(query, [recordId, OfficerName, Title], function (err) {
      if (err) {
        console.error('Error inserting into officers:', err.message);
        reject(err);
      } else {
        console.log('Officer inserted with ID:', this.lastID);
        resolve(this.lastID);
      }
    });
  });
}


function insertRecord2(Firm, State, Mailing, InitialMatch, dbBusiness) {
  return new Promise((resolve, reject) => {
    const query = `
      INSERT INTO records (Firm, State, Mailing, InitialMatch)
      VALUES (?, ?, ?, ?)
    `;
    dbBusiness.run(query, [Firm, State, Mailing, InitialMatch], function (err) {
      if (err) {
        console.error('Error inserting into records:', err.message);
        reject(err);
      } else {
        console.log('\nRecord inserted with ID:', this.lastID);
        resolve(this.lastID); // Return the ID of the newly inserted record
      }
    });
  });
}


async function makeApiCall2(searchName, jurisdictionCode, searchMailing, index, dbBusiness) {
  try {

        var companyID = "";
        var InitialMatch = "";
        var officerArray = [];

        officerArray.length = 0;

        if (ocAPIcount % 100 === 0 && ocAPIcount !== 0) {
          console.log(`Pausing for 1 minute at index ${index} to respect rate limits.`);
          await new Promise(resolve => setTimeout(resolve, 60000)); // 1 minute delay
        }

        const url = `https://api.opencorporates.com/v0.4/companies/search?q=${searchName}&jurisdiction_code=${jurisdictionCode}&order=score&normalise_company_name=true&api_token=${process.env.OC_KEY}`;

        console.log("Making First call to Business API")
 

        const response = await fetch(url);
        const apiData = await response.json();

        ocAPIcount++;

        if (response.status !== 200) {
          console.log('API Error:', response.status);
          console.error('API Error:', response.status);
          return response.status;
        }

        const companies = apiData.results.companies;

        

        if (companies !== undefined && companies.length > 0) {

          console.log(companies.length,'Companies found');
          if (companies.length > 1) {
            console.log('\nComparing Addresses for best match');
            const threshold = 0.1; // threshold for a strong match
          
            // Filter out null or undefined addresses
            const addresses = companies
              .map(company => company.company.registered_address_in_full)
              .filter(address => address !== null && address !== undefined);
          
            // Check if we have any valid addresses left to compare
            if (addresses.length === 0) {
              console.log('No valid addresses to compare. Taking First Company');
              companyID = companies[0].company.company_number; // Default to the first company
            } else {
              const { ratings, bestMatch, bestMatchIndex } = stringSimilarity.findBestMatch(searchMailing, addresses);
          
              if (bestMatch.rating >= threshold) {
                console.log('Best Match Found:', bestMatch.target);
                console.log('Company Address:', searchMailing);
                companyID = companies[bestMatchIndex].company.company_number;
              } else {
                // No strong match, default to the first company
                console.log('No strong match found, defaulting to first company');
                companyID = companies[0].company.company_number;
              }
            }
          }
          else
          {
            companyID = companies[0].company.company_number;
          }

        }
        else
        {
          console.log('No Companies found');
          InitialMatch = "No Companies Found";
          let newRow = { ...data[index] };
          newRow['Officer List'] = "No Companies Found";
          expandedData.push(newRow);
        }


        if (companyID !== "") {

            
            console.log("\nMaking Second call to Business API")

            console.log('Company ID:',companyID);
            console.log('\n');

            const detailUrl = `https://api.opencorporates.com/v0.4/companies/${jurisdictionCode}/${companyID}?api_token=${process.env.OC_KEY}`;
            const detailResponse = await fetch(detailUrl);
            ocAPIcount++;
            const detailData = await detailResponse.json();
            
            if (detailResponse.status !== 200) {
              console.log('API Error:', detailResponse.status);
              console.error('API Error:', detailResponse.status);
              return detailResponse.status;
            }

            // Extract specific ownership information
            const companyDetails = detailData.results.company;

            const processedNames = new Set();
            const processedLastNames = new Set();
            let takenNames = 0;
            
            if (companyDetails.officers && companyDetails.officers.length > 0) {
              let officersFound = false;

              companyDetails.officers.forEach(officerData => {
                
                const nameObject = parseFullName(officerData.officer.name);
                const fullNameKey = (nameObject.first + nameObject.last).toLowerCase();

                if (processedNames.has(fullNameKey)) {
                  return; // Skip duplicate first + last names
                }

                if (takenNames >= 2) {
                  return; // Skip more than 2 officers
                }

                const lastNameDifferentFirst = processedLastNames.has(nameObject.last.toLocaleLowerCase()) && !processedNames.has(fullNameKey);

                if(lastNameDifferentFirst||
                  (officerData.officer.position!=="treasurer" && 
                    officerData.officer.position!=="secretary" && 
                    officerData.officer.position!=="agent")){
                let newRow = { ...data[index] };

                // Add officer data to the new row
                newRow['Officer List'] = "True";
                newRow['Officer Name'] = officerData.officer.name;
                newRow['Officer Position'] = officerData.officer.position;
                newRow['NAME'] = officerData.officer.name;
                newRow['FIRST'] = nameObject.first;
                newRow['LAST'] = nameObject.last;
                
                console.log('-    Officer Position:', officerData.officer.position);
                console.log('-    Officer Name:', officerData.officer.name);
              

                officerArray.push({OfficerName: officerData.officer.name, Title: officerData.officer.position});
                takenNames++;

                // Push the complete row to expandedData
                expandedData.push(newRow);
                officersFound = true;

                processedNames.add(fullNameKey);
                processedLastNames.add(nameObject.last.toLowerCase());

                }
              });

              if (officersFound === true)
              {
                InitialMatch = "True";
              }
              else
              {
                InitialMatch = "Company found but no officers"
                console.log('Company found but no officers');
                let newRow = { ...data[index] };
                newRow['Officer List'] = "Company found but no officers";
                expandedData.push(newRow);
              }

            }
            else
            {
              let newRow = { ...data[index] };
    
              // Add officer data to the new row
              newRow['Officer List'] = "Company found but no officers";
              InitialMatch = "Company found but no officers";
              console.log('Company found but no officers');
              // Push the complete row to expandedData
              expandedData.push(newRow);
            }

          }
      
      let newRow = { ...data[index] };
      const recordId = await insertRecord2(newRow['FIRM'], newRow['STATE'], newRow['MAILING']+", "+newRow['ZIP5'], InitialMatch, dbBusiness);
      for (let i = 0; i < officerArray.length; i++) {
        let OfficerName = officerArray[i].OfficerName;
        let Title = officerArray[i].Title;
        await insertOfficer(recordId, OfficerName, Title, dbBusiness);
      }

      } catch (error) {
        console.error('Error:', error);
        console.log('Error: ', error.message);
      }
    }


function getAdditionalFields2(Firm, State, Mailing, dbBusiness) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT records.InitialMatch, officers.OfficerName, officers.Title
      FROM records
      LEFT JOIN officers ON records.id = officers.record_id
      WHERE records.Firm = ? AND records.State = ? AND records.Mailing = ?
    `;
    
    dbBusiness.all(query, [Firm, State, Mailing], (err, rows) => {
      if (err) {
        console.error('Error fetching additional fields:', err.message);
        reject(err);
      } else {
        resolve(rows); // 'rows' will contain all officers and their titles for the given firm, state, and mailing
      }
    });
  });
}



function checkForDuplicate2(Firm, State, Mailing, dbBusiness) {
  return new Promise((resolve, reject) => {
    const query = `
      SELECT COUNT(*) as count FROM records
      WHERE Firm = ? AND State = ? AND Mailing = ? 
    `;
    dbBusiness.get(query, [Firm, State, Mailing], (err, row) => {
      if (err) {
        console.error('Error checking for duplicate:', err.message);
        reject(err);
      } else {
        resolve(row.count > 0);
      }
    });
  });
}

async function processRowsSequentially2(data,dbBusiness) {
  for (const [index , row] of data.entries()) {


    const companyName = row['FIRM'];
    const companyState = row['STATE'];
    const searchMailing = row['MAILING']+", "+row['ZIP5'];

    console.log('\nProcessing:', index+1, companyName, companyState);

    if (!companyName || !companyState) {
      console.log('Skipping row due to missing data:', index);
      
      let newRow = { ...data[index] };
      newRow['Officer List'] = "Incomplete Data";
      expandedData.push(newRow);
      continue;
    }

    const isDuplicate = await checkForDuplicate2(companyName, companyState, searchMailing, dbBusiness);
    console.log(companyName, companyState, searchMailing);

    if (isDuplicate) {
      console.log('Duplicate record found. Skipping API call');

      const duplicateFields = await getAdditionalFields2(companyName, companyState, searchMailing, dbBusiness);
      

      if(duplicateFields.length > 0){
        duplicateFields.forEach((record) => {
          let newRow = { ...data[index] };
          newRow['Officer List'] = "Duplicate";
          newRow['Initial Match'] = record.InitialMatch;
          newRow['Officer Name'] = record.OfficerName;
          newRow['Officer Position'] = record.Title;

          const nameObject = parseFullName(record.OfficerName);

          newRow['NAME'] = record.OfficerName;
          newRow['FIRST'] = nameObject.first;
          newRow['LAST'] = nameObject.last;
          expandedData.push(newRow);

        });
      }

      continue;
    }
      const jurisdictionCode = 'us_' + companyState.toLowerCase();
      const searchName = encodeURIComponent(companyName);


     const status = await makeApiCall2(searchName, jurisdictionCode, searchMailing, index, dbBusiness);

     if (status == 403){
      console.log('Rate Limit Exceeded. Ending Execution');

      for (let remainingIndex = index; remainingIndex < data.length; remainingIndex++) {
        let newRow = { ...data[remainingIndex] };
        expandedData.push(newRow);
      }      

      break;
     }
    
  }
}


async function setupDatabase2() {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database('./business.db', (err) => {
      if (err) {
        console.error('Error opening database:', err.message);
        reject(err); // Reject promise if there's an error
      } else {
        console.log('\nConnected to the SQLite business database.');

        db.serialize(() => {
          // Create the records table (same as before)
          db.run(`
            CREATE TABLE IF NOT EXISTS records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              Firm TEXT NOT NULL,
              State TEXT NOT NULL,
              Mailing TEXT NOT NULL,
              InitialMatch TEXT,
              UNIQUE(Firm, State, Mailing)
            )
          `, (err) => {
            if (err) {
              console.error('Error creating records table:', err.message);
              reject(err); // Reject if table creation fails
            }
          });

          // Create the officers table
          db.run(`
            CREATE TABLE IF NOT EXISTS officers (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              record_id INTEGER NOT NULL,
              OfficerName TEXT,
              Title TEXT,
              FOREIGN KEY(record_id) REFERENCES records(id) ON DELETE CASCADE
            )
          `, (err) => {
            if (err) {
              console.error('Error creating officers table:', err.message);
              reject(err); // Reject if table creation fails
            } else {
              resolve(db); // Resolve the promise when both tables are created successfully
            }
          });
        });
      }
    });
  });
}



async function main2(filePath) {

  expandedData.length = 0;
  const dbBusiness = await setupDatabase2();

  const workbook = XLSX.readFile(filePath);
  var worksheet = workbook.Sheets["Sheet1"];

  console.log('\nPulling data from "Sheet1" of the Excel file...');
  data = XLSX.utils.sheet_to_json(worksheet);



  await processRowsSequentially2(data,dbBusiness);



  var matchData=[];
  var noMatchData=[];
  var notProcessed=[]



expandedData.forEach((row) => {
  if (row['Officer List'] === "True" || row['Initial Match'] === "True") {
    matchData.push(row);
  } else if (row['Officer List'] === undefined) {
    notProcessed.push(row);
  }
  else {
    noMatchData.push(row);
  }
});


const matchWorksheet = XLSX.utils.json_to_sheet(matchData);
const noMatchWorksheet = XLSX.utils.json_to_sheet(noMatchData);
const notProcessedWorksheet = XLSX.utils.json_to_sheet(notProcessed);
  

  console.log("\n");
  if (workbook.Sheets["Business Match"]) {
    // Sheet exists, overwrite it
    workbook.Sheets["Business Match"] = matchWorksheet;
    console.log(`Overwrote existing Excel sheet: Business Match`);
  } else {
    // Sheet doesn't exist, append new sheet
    XLSX.utils.book_append_sheet(workbook, matchWorksheet, "Business Match");
    console.log(`Created new Excel sheet: Business Match`);
  }

  if (workbook.Sheets["Business No Match"]) {

    workbook.Sheets["Business No Match"] = noMatchWorksheet;
    console.log(`Overwrote existing Excel sheet: Business No Match`);
  } else {

    XLSX.utils.book_append_sheet(workbook, noMatchWorksheet, "Business No Match");
    console.log(`Created new Excel sheet: Business No Match`);
  }

  if (workbook.Sheets["Business Not Processed"]) {
    // Sheet exists, overwrite it
    workbook.Sheets["Business Not Processed"] = notProcessedWorksheet;
    console.log(`Overwrote existing Excel sheet: Business Not Processed`);
  } else {
    // Sheet doesn't exist, append new sheet
    XLSX.utils.book_append_sheet(workbook, notProcessedWorksheet, "Business Not Processed");
    console.log(`Created new Excel sheet: Business Not Processed`);
  }

  XLSX.writeFile(workbook, filePath);


  dbBusiness.close((err) => {
    if (err) {
      console.error('\nError closing the database:', err.message);
    } else {
      console.log('\nBusiness Database connection closed.');
    }
  });

}