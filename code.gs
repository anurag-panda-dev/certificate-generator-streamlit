// Google Apps Script - Deploy as Web App
// Execute as: Me
// Access: Anyone

// Replace with your actual Google Spreadsheet ID
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// Certificate type to sheet name mapping
const SHEET_MAPPING = {
  'Python for Beginners': 'Python',
  'Web Development for Beginners': 'WebDev',
  'Tree Plantation': 'TreePlantation',
  'Debate Competition': 'Debate',
  'Yoga Camp': 'Yoga',
  'Blood Donation Camp': 'BloodDonation',
  'Arduino for Beginners': 'Arduino'
};

/**
 * Main function to handle incoming requests
 */
function doPost(e) {
  try {
    // Enable CORS for cross-origin requests
    const response = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    };
    
    // Parse the request data
    const data = JSON.parse(e.postData.contents);
    
    // Log the received data for debugging
    console.log('Received data:', data);
    
    // Validate required fields
    if (!data.type || !data.name || !data.email || !data.date) {
      return createResponse(400, 'Missing required fields', response);
    }
    
    // Get the appropriate sheet name
    const sheetName = SHEET_MAPPING[data.type];
    if (!sheetName) {
      return createResponse(400, 'Invalid certificate type', response);
    }
    
    // Process the form submission
    const result = processFormSubmission(data, sheetName);
    
    if (result.success) {
      return createResponse(200, result.message, response);
    } else {
      return createResponse(500, result.message, response);
    }
    
  } catch (error) {
    console.error('Error processing request:', error);
    return createResponse(500, 'Internal server error: ' + error.message, {
      'Access-Control-Allow-Origin': '*'
    });
  }
}

/**
 * Handle OPTIONS requests for CORS preflight
 */
function doOptions() {
  return HtmlService.createHtmlOutput()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('Access-Control-Allow-Origin', '*')
    .addMetaTag('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .addMetaTag('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Process form submission and add to appropriate sheet
 */
function processFormSubmission(data, sheetName) {
  try {
    // Open the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get or create the sheet
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = createSheet(spreadsheet, sheetName);
    }
    
    // Prepare the row data
    const timestamp = new Date();
    const rowData = [
      timestamp,
      data.name,
      data.email,
      data.date,
      data.blood_group || ''
    ];
    
    // Add the data to the sheet
    sheet.appendRow(rowData);
    
    // Generate and send certificate automatically
    const certificateResult = generateCertificate(data);
    
    if (certificateResult.success) {
      console.log(`Certificate generated and sent to: ${data.email}`);
      return {
        success: true,
        message: 'Certificate generated and sent to your email successfully!'
      };
    } else {
      console.error(`Certificate generation failed: ${certificateResult.message}`);
      // Still return success for form submission, but note certificate issue
      return {
        success: true,
        message: 'Form submitted successfully! Certificate generation is in progress.'
      };
    }
    
  } catch (error) {
    console.error('Error processing form submission:', error);
    return {
      success: false,
      message: 'Error saving data: ' + error.message
    };
  }
}

/**
 * Create a new sheet with proper headers
 */
function createSheet(spreadsheet, sheetName) {
  const sheet = spreadsheet.insertSheet(sheetName);
  
  // Add headers
  const headers = ['Timestamp', 'Name', 'Email', 'Date', 'Blood Group'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format the header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  return sheet;
}

/**
 * Create a properly formatted response
 */
function createResponse(status, message, headers = {}) {
  const response = {
    message: message,
    status: status,
    timestamp: new Date().toISOString()
  };
  
  const output = HtmlService.createHtmlOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Add CORS headers
  if (headers['Access-Control-Allow-Origin']) {
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return output;
}

/**
 * Optional: Send email notification to the user
 */
function sendEmailNotification(data) {
  try {
    const subject = `Certificate Request Confirmed: ${data.type}`;
    const body = `
Dear ${data.name},

Thank you for your certificate request! We have received your submission for the following certificate:

Certificate Type: ${data.type}
Name: ${data.name}
Email: ${data.email}
Date of Issue: ${data.date}
${data.blood_group ? `Blood Group: ${data.blood_group}` : ''}

Your certificate will be processed and sent to your email address shortly.

Best regards,
Certificate Generation Team
    `;
    
    GmailApp.sendEmail(data.email, subject, body);
    
  } catch (error) {
    console.error('Error sending email notification:', error);
  }
}

/**
 * Generate PDF certificate using Google Slides template
 */
function generateCertificate(data) {
  try {
    // Certificate template mapping - Replace with your actual Google Slides template IDs
    const CERTIFICATE_TEMPLATES = {
      'Python for Beginners': 'YOUR_PYTHON_TEMPLATE_ID',
      'Web Development for Beginners': 'YOUR_WEBDEV_TEMPLATE_ID',
      'Tree Plantation': 'YOUR_TREE_TEMPLATE_ID',
      'Debate Competition': 'YOUR_DEBATE_TEMPLATE_ID',
      'Yoga Camp': 'YOUR_YOGA_TEMPLATE_ID',
      'Blood Donation Camp': 'YOUR_BLOOD_TEMPLATE_ID',
      'Arduino for Beginners': 'YOUR_ARDUINO_TEMPLATE_ID'
    };
    
    // Get the appropriate template ID
    const templateId = CERTIFICATE_TEMPLATES[data.type];
    if (!templateId) {
      throw new Error(`No template found for certificate type: ${data.type}`);
    }
    
    // Make a copy of the template
    const templateFile = DriveApp.getFileById(templateId);
    const newFileName = `Certificate_${data.name.replace(/\s+/g, '_')}_${Date.now()}`;
    const newFile = templateFile.makeCopy(newFileName);
    
    // Open the presentation
    const presentation = SlidesApp.openById(newFile.getId());
    const slides = presentation.getSlides();
    
    // Process each slide in the presentation
    slides.forEach((slide, index) => {
      console.log(`Processing slide ${index + 1}`);
      
      // Get all text elements on the slide
      const pageElements = slide.getPageElements();
      
      pageElements.forEach(element => {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          if (shape.getText()) {
            let textContent = shape.getText().asString();
            
            // Replace placeholders with actual data
            textContent = replacePlaceholders(textContent, data);
            
            // Update the text if changes were made
            if (textContent !== shape.getText().asString()) {
              shape.getText().setText(textContent);
              
              // Optional: Style the text
              styleText(shape.getText(), data);
            }
          }
        }
        
        // Handle text boxes
        if (element.getPageElementType() === SlidesApp.PageElementType.TEXT_BOX) {
          const textBox = element.asTextBox();
          if (textBox.getText()) {
            let textContent = textBox.getText().asString();
            
            // Replace placeholders
            textContent = replacePlaceholders(textContent, data);
            
            // Update the text
            if (textContent !== textBox.getText().asString()) {
              textBox.getText().setText(textContent);
              
              // Optional: Style the text
              styleText(textBox.getText(), data);
            }
          }
        }
      });
    });
    
    // Save the presentation
    presentation.saveAndClose();
    
    // Convert to PDF
    const pdf = DriveApp.getFileById(newFile.getId()).getAs('application/pdf');
    pdf.setName(`${data.name}_Certificate_${data.type.replace(/\s+/g, '_')}.pdf`);
    
    // Send via email
    sendCertificateEmail(data, pdf);
    
    // Optional: Save PDF to a specific folder
    savePdfToFolder(pdf, data);
    
    // Clean up temporary presentation file
    DriveApp.getFileById(newFile.getId()).setTrashed(true);
    
    return {
      success: true,
      message: 'Certificate generated and sent successfully!'
    };
    
  } catch (error) {
    console.error('Error generating certificate:', error);
    return {
      success: false,
      message: 'Error generating certificate: ' + error.message
    };
  }
}

/**
 * Replace placeholders in text content
 */
function replacePlaceholders(text, data) {
  // Format date properly
  const formattedDate = formatDate(data.date);
  
  // Replace all placeholders
  return text
    .replace(/\{\{NAME\}\}/g, data.name.toUpperCase())
    .replace(/\{\{FULL_NAME\}\}/g, data.name)
    .replace(/\{\{CERTIFICATE_TYPE\}\}/g, data.type)
    .replace(/\{\{COURSE_NAME\}\}/g, data.type)
    .replace(/\{\{DATE\}\}/g, formattedDate)
    .replace(/\{\{ISSUE_DATE\}\}/g, formattedDate)
    .replace(/\{\{BLOOD_GROUP\}\}/g, data.blood_group || '')
    .replace(/\{\{EMAIL\}\}/g, data.email)
    .replace(/\{\{YEAR\}\}/g, new Date().getFullYear().toString())
    .replace(/\{\{MONTH\}\}/g, new Date().toLocaleString('default', { month: 'long' }))
    .replace(/\{\{CURRENT_DATE\}\}/g, new Date().toLocaleDateString());
}

/**
 * Style text elements (optional customization)
 */
function styleText(textRange, data) {
  try {
    // Apply different styles based on content
    const text = textRange.asString();
    
    // Style name fields
    if (text.includes(data.name)) {
      textRange.getTextStyle()
        .setFontSize(24)
        .setBold(true)
        .setForegroundColor('#1a73e8');
    }
    
    // Style certificate type
    if (text.includes(data.type)) {
      textRange.getTextStyle()
        .setFontSize(18)
        .setBold(true)
        .setForegroundColor('#34a853');
    }
    
    // Style date
    if (text.includes(formatDate(data.date))) {
      textRange.getTextStyle()
        .setFontSize(12)
        .setItalic(true)
        .setForegroundColor('#666666');
    }
    
  } catch (error) {
    console.log('Error styling text:', error);
  }
}

/**
 * Format date for display
 */
function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  } catch (error) {
    return dateStr; // Return original if parsing fails
  }
}

/**
 * Send certificate via email
 */
function sendCertificateEmail(data, pdf) {
  try {
    const subject = `🏆 Your Certificate: ${data.type}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #1a73e8;">Congratulations ${data.name}!</h2>
        
        <p>We are pleased to present you with your certificate for <strong>${data.type}</strong>.</p>
        
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
          <h3 style="color: #34a853; margin-top: 0;">Certificate Details:</h3>
          <ul style="list-style: none; padding: 0;">
            <li><strong>Name:</strong> ${data.name}</li>
            <li><strong>Certificate Type:</strong> ${data.type}</li>
            <li><strong>Issue Date:</strong> ${formatDate(data.date)}</li>
            ${data.blood_group ? `<li><strong>Blood Group:</strong> ${data.blood_group}</li>` : ''}
          </ul>
        </div>
        
        <p>Your certificate is attached to this email as a PDF file. Please save it for your records.</p>
        
        <div style="background-color: #e8f0fe; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <p style="margin: 0; color: #1a73e8;"><strong>💡 Tip:</strong> You can print this certificate on high-quality paper for the best results.</p>
        </div>
        
        <p>Thank you for your participation!</p>
        
        <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 30px 0;">
        
        <p style="color: #666; font-size: 12px;">
          Best regards,<br>
          Certificate Generation Team<br>
          <em>This is an automated email. Please do not reply.</em>
        </p>
      </div>
    `;
    
    const textBody = `
Dear ${data.name},

Congratulations! We are pleased to present you with your certificate for ${data.type}.

Certificate Details:
- Name: ${data.name}
- Certificate Type: ${data.type}
- Issue Date: ${formatDate(data.date)}
${data.blood_group ? `- Blood Group: ${data.blood_group}` : ''}

Your certificate is attached to this email as a PDF file. Please save it for your records.

Thank you for your participation!

Best regards,
Certificate Generation Team
    `;
    
    GmailApp.sendEmail(
      data.email,
      subject,
      textBody,
      {
        htmlBody: htmlBody,
        attachments: [pdf]
      }
    );
    
    console.log(`Certificate emailed to: ${data.email}`);
    
  } catch (error) {
    console.error('Error sending certificate email:', error);
  }
}

/**
 * Save PDF to a specific folder (optional)
 */
function savePdfToFolder(pdf, data) {
  try {
    // Replace with your actual folder ID where you want to save certificates
    const CERTIFICATE_FOLDER_ID = 'YOUR_CERTIFICATE_FOLDER_ID';
    
    if (CERTIFICATE_FOLDER_ID && CERTIFICATE_FOLDER_ID !== 'YOUR_CERTIFICATE_FOLDER_ID') {
      const folder = DriveApp.getFolderById(CERTIFICATE_FOLDER_ID);
      
      // Create a subfolder for the certificate type if it doesn't exist
      const certificateTypeFolder = getOrCreateSubfolder(folder, data.type);
      
      // Save the PDF to the folder
      const savedFile = DriveApp.createFile(pdf);
      certificateTypeFolder.addFile(savedFile);
      
      // Remove from root folder
      DriveApp.getRootFolder().removeFile(savedFile);
      
      console.log(`Certificate saved to folder: ${data.type}`);
    }
    
  } catch (error) {
    console.error('Error saving PDF to folder:', error);
  }
}

/**
 * Get or create a subfolder
 */
function getOrCreateSubfolder(parentFolder, folderName) {
  try {
    const subfolders = parentFolder.getFoldersByName(folderName);
    if (subfolders.hasNext()) {
      return subfolders.next();
    } else {
      return parentFolder.createFolder(folderName);
    }
  } catch (error) {
    console.error('Error creating subfolder:', error);
    return parentFolder; // Return parent folder if subfolder creation fails
  }
}

/**
 * Bulk generate certificates from spreadsheet data
 */
function bulkGenerateCertificates(sheetName, startRow = 2) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`Sheet ${sheetName} not found`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const nameCol = headers.indexOf('Name');
    const emailCol = headers.indexOf('Email');
    const dateCol = headers.indexOf('Date');
    const bloodGroupCol = headers.indexOf('Blood Group');
    
    // Process each row
    for (let i = startRow - 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[nameCol] && row[emailCol]) {
        const certificateData = {
          type: Object.keys(SHEET_MAPPING).find(key => SHEET_MAPPING[key] === sheetName),
          name: row[nameCol],
          email: row[emailCol],
          date: row[dateCol],
          blood_group: row[bloodGroupCol] || ''
        };
        
        console.log(`Generating certificate for: ${certificateData.name}`);
        const result = generateCertificate(certificateData);
        
        if (result.success) {
          console.log(`✓ Certificate generated for ${certificateData.name}`);
        } else {
          console.error(`✗ Failed to generate certificate for ${certificateData.name}: ${result.message}`);
        }
        
        // Add a small delay to avoid hitting API limits
        Utilities.sleep(1000);
      }
    }
    
  } catch (error) {
    console.error('Error in bulk certificate generation:', error);
  }
}

/**
 * Test function to verify the script works
 */
function testScript() {
  const testData = {
    type: 'Python for Beginners',
    name: 'Test User',
    email: 'test@example.com',
    date: '2025-01-15',
    blood_group: ''
  };
  
  const result = processFormSubmission(testData, 'Python');
  console.log('Test result:', result);
}

/**
 * Initialize the spreadsheet with all required sheets
 */
function initializeSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Create all required sheets
    Object.values(SHEET_MAPPING).forEach(sheetName => {
      if (!spreadsheet.getSheetByName(sheetName)) {
        createSheet(spreadsheet, sheetName);
        console.log(`Created sheet: ${sheetName}`);
      }
    });
    
    console.log('Spreadsheet initialization complete');
    
  } catch (error) {
    console.error('Error initializing spreadsheet:', error);
  }
}