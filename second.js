const {shell, ipcRenderer} = require('electron')
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs').promises;
const OpenAI = require("openai");
const dotenv = require("dotenv");

dotenv.config();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY, dangerouslyAllowBrowser: true
});


document.getElementById('open-instructions').addEventListener('click', () => {
          const instructionsPath = path.join(__dirname, 'lead-instructions.txt'); 
        
          shell.openPath(instructionsPath)
              .then(response => {
                  if (response) {
                      console.error('Error opening file:', response);
                  }
              });
  });

  document.getElementById('sunBizScript').addEventListener('click', async () => {
        const filePath = await ipcRenderer.invoke('open-file-dialog')
    if (filePath) {
      console.log('Selected file path:', filePath)
      
      try {
        // Read the file
        const data = await fs.readFile(filePath, 'utf8');
        const lines = data.split('\n');
        const records = [];
  
        // Process each line
        for (const line of lines) {
          if (line.trim() === '') continue;

          console.log('Processing line:', line);
  
          // Extract fields based on fixed positions
          const record = {
            FIRM: line.substring(13, 60).trim(),
            MAILING: line.substring(69, 155).trim(),
            CITY: line.substring(156, 181).trim(),
            STATE: line.substring(185, 187).trim(),
            ZIP5: line.substring(187, 192).trim()
          };
  
          records.push(record);
        }
  
        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(records);
  
        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  
        // Save the Excel file in the same directory as the input file
        const outputPath = filePath.replace('.txt', '_ExtractedFirms.xlsx');
        XLSX.writeFile(wb, outputPath);
        
        console.log('File processed successfully! Saved to:', outputPath);
        
        // Show alert with file location
        alert(`File processed successfully!\nSaved to: ${outputPath}`);
  
      } catch (error) {
        console.error('Error processing file:', error);
        alert('Error processing file: ' + error.message);
      }
      
    } else {
      console.log('No file selected');
    }
  });


  document.getElementById('mdjScript').addEventListener('click', async () => {
    try {
      // a) Let user pick a PDF
      const filePath = await ipcRenderer.invoke('open-file-dialog');
      if (!filePath) {
        console.log('No file selected');
        return;
      }
      
      document.documentElement.classList.add('waiting');
      const outputPath = filePath.replace('.pdf', '_ExtractedData.xlsx');

      // b) Parse the PDF in main
      const pdfText = await ipcRenderer.invoke('parse-pdf', filePath);
      console.log('Raw PDF text:', pdfText);
  
      // c) Filter the lines
      const lines = pdfText.split('\n');
      const federalLiens = filterFederalLiens(lines);
      console.log('Filtered federal liens:', federalLiens);

        // **Check if we have data before sending to OpenAI**
        if (!federalLiens || federalLiens.length === 0) {
            alert("No valid federal liens found. There's nothing to process.");
            return;
            }

      // d) Send to OpenAI and save final to Excel
    await sendFederalLiensToOpenAIAndSaveToExcel(federalLiens, outputPath);

    alert(`File processed successfully!\nSaved to: ${outputPath}`);
  
    } catch (error) {
      console.error('Error in renderer flow:', error);
      alert('Error processing file: ' + error.message);
    }
    finally {
    // Revert cursor back to default whether success or error
    document.documentElement.classList.remove('waiting');
    }
  });

// Same filter function as before
function filterFederalLiens(lines) {
  const results = [];
  for (let i = 0; i < lines.length; i++) {
    const lowerLine = lines[i].toLowerCase();
    if (lowerLine.includes('federal lien')) {
      if (i + 1 >= lines.length || !lines[i + 1].startsWith('Plaintiff Address:')) {
        continue;
      }
      let addressLineIndex = -1;
      const limit = Math.min(i + 10, lines.length - 1);
      for (let j = i + 1; j <= limit; j++) {
        if (lines[j].startsWith('C-')) {
          break;
        }
        if (lines[j].startsWith('Defendant Address:')) {
          addressLineIndex = j;
          break;
        }
      }
      if (addressLineIndex === -1) {
        continue;
      }
      const block = [];
      let count = 0;
      for (let k = addressLineIndex; k < lines.length; k++) {
        if (k > addressLineIndex && lines[k].startsWith('C-')) {
          break;
        }
        block.push(lines[k]);
        count++;
        if (count === 6) {
          break;
        }
      }
      results.push(block);
    }
  }
  return results;
}

async function sendFederalLiensToOpenAIAndSaveToExcel(federalLiens, outputPath) {
    try {
      // Convert your array to a JSON string to send as the "user" message
      const federalLiensJson = JSON.stringify(federalLiens, null, 2);
  
      const instruction = `
        Extract fields FIRM, MAILING, CITY, STATE, FIRST, LAST, MI, ZIP5 where applicable from Defendant Addresses.
        Convert 'L L C' to 'LLC' where applicable.
        Return the result only as a JSON array of objects. Each object must contain exactly these fields (some fields may be empty if unknown).
      `;
  
      // For OpenAI v4 library:
      const response = await openai.chat.completions.create({
        model: "gpt-4o-2024-08-06",  
        messages: [
          { role: "system", content: instruction },
          { role: "user", content: federalLiensJson },
        ],
      });
  
      const rawContent = response.choices[0].message.content;
      console.log("OpenAI raw response:", rawContent);
  
      // Clean out ```json ``` if the model includes those
      const cleanResponse = rawContent.replace(/```json|```/g, '').trim();
  
      // Attempt to parse as JSON
      const extractedData = JSON.parse(cleanResponse);
  
      if (Array.isArray(extractedData)) {
        // Convert JSON array to Excel
        saveToExcel(extractedData, outputPath);
        console.log('Data has been saved to federal_liens.xlsx');
      } else {
        console.error("Response data is not in the expected array format.");
      }
  
    } catch (error) {
      console.error("Error sending to OpenAI or saving to Excel:", error);
      throw error;
    }
  }

  function saveToExcel(data, fileName) {

    console.log(data);

    const firmsData = data.filter(row => row.FIRM && row.FIRM.trim() !== '');
    const personalData = data.filter(row => !row.FIRM || row.FIRM.trim() === '');

    if (firmsData.length > 0) {
        const firmsWorkbook = XLSX.utils.book_new();
        const firmsWorksheet = XLSX.utils.json_to_sheet(firmsData);
        XLSX.utils.book_append_sheet(firmsWorkbook, firmsWorksheet, 'Sheet1');
        const outputPath1 = fileName.replace('_ExtractedData.xlsx', '_ExtractedFirms.xlsx');
        XLSX.writeFile(firmsWorkbook, outputPath1);
    }

    if (personalData.length > 0) {
        const personalWorkbook = XLSX.utils.book_new();
        const personalWorksheet = XLSX.utils.json_to_sheet(personalData);
        XLSX.utils.book_append_sheet(personalWorkbook, personalWorksheet, 'Sheet1');
        const outputPath2 = fileName.replace('_ExtractedData.xlsx', '_ExtractedPersonal.xlsx');
        XLSX.writeFile(personalWorkbook, outputPath2);
    }
}