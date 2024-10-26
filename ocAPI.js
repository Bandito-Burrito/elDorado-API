require('dotenv').config();
const XLSX = require('xlsx');
var data = [];
var expandedData = [];

const filePath = 'C:/Users/ihn19/OneDrive/Desktop/HTML-CSS-Projects/EnDato API Call/Data/ocStubborn.xlsx';


async function main(filePath) {
const workbook = XLSX.readFile(filePath);
var worksheet = workbook.Sheets["Sheet1"];
data = XLSX.utils.sheet_to_json(worksheet);
await processRowsSequentially(data);

const newWorksheet = XLSX.utils.json_to_sheet(expandedData);
    
// Add it as a new sheet to the existing workbook
XLSX.utils.book_append_sheet(workbook, newWorksheet, "API Return");

// Write back to the same file
XLSX.writeFile(workbook, filePath);
console.log('New sheet added to workbook');
}

async function processRowsSequentially(data) {
    for (const [index , row] of data.entries()) {
      // Extract variables from the row
      const companyName = row['FIRM'];
      const companyState = row['STATE'];

      // Log the variables if needed
      console.log('Processing:', companyName, companyState);
  
      if (!companyName || !companyState) {
        console.log('Skipping row due to missing data:', index);
        continue;
      }

      const jurisdictionCode = 'us_' + companyState.toLowerCase();
      const searchName = encodeURIComponent(companyName).replace(/%20/g, '+');
      
      // Call the API function with the variables
     await makeApiCall(searchName, jurisdictionCode, index);
    }
  }


async function makeApiCall(companyName, jurisdictionCode, index) {
  try {

        if (index % 50 === 0 && index !== 0) {
          console.log(`Pausing for 1 minute at index ${index} to respect rate limits.`);
          await new Promise(resolve => setTimeout(resolve, 60000)); // 1 minute delay
        }

        const url = `https://api.opencorporates.com/v0.4/companies/search?q=${companyName}&jurisdiction_code=${jurisdictionCode}&order=score&normalise_company_name=true&api_token=${process.env.OC_KEY}`;

        console.log('API URL:', url);

        const response = await fetch(url);
        const apiData = await response.json();

        console.log('Full API Response:', apiData);
        console.log('\n');

        const companies = apiData.results.companies;
        console.log('Companies:', companies);
        console.log('\n');

        if (companies !== undefined && companies.length > 0) {

            const companyID = companies[0].company.company_number;
            console.log('Company ID:', companyID);
            const detailUrl = `https://api.opencorporates.com/v0.4/companies/${jurisdictionCode}/${companyID}?api_token=${process.env.OC_KEY}`;
            const detailResponse = await fetch(detailUrl);
            const detailData = await detailResponse.json();
            
            // Log the detailed company information
            console.log('\n');
            console.log('Second Endpoint:', detailData);
            console.log('\n');
      
            // Extract specific ownership information
            const companyDetails = detailData.results.company;
            
            


            // Log specific ownership-related fields
            console.log('\n');
            console.log('Officers:', companyDetails.officers);
            console.log('Owners:', companyDetails.ultimate_beneficial_owners);
            console.log('\n');

            if (companyDetails.officers && companyDetails.officers.length > 0) {
              companyDetails.officers.forEach(officerData => {
                let newRow = { ...data[index] };
        
                // Add officer data to the new row
                newRow['Officer List'] = "True";
                newRow['Officer Name'] = officerData.officer.name;
                newRow['Officer Position'] = officerData.officer.position;
                
                // Push the complete row to expandedData
                expandedData.push(newRow);
              });

            }
            else
            {
              let newRow = { ...data[index] };
    
              // Add officer data to the new row
              newRow['Officer List'] = "False";
              
              // Push the complete row to expandedData
              expandedData.push(newRow);
            }

          }
        
      } catch (error) {
        console.error('Error:', error);
      }
    }

    main(filePath);