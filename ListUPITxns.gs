function processUPIEmails() {
  const query = 'subject:(UPI) OR "You have done a UPI txn" OR "You sent money using UPI" OR "Received via UPI" OR "View: Account update for your HDFC Bank A/c"';
  const threads = GmailApp.search(query);
  const label = getOrCreateLabel_("Deleted");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UPI_Transactions");

  // Add header if empty
  if (sheet.getLastRow() === 0) {
    //sheet.appendRow(["Date", "Amount", "Type", "VPA/Account", "From VPA", "Reference No", "Email Subject", "Email Link"]);
    sheet.appendRow(["Date", "Time", "Amount", "Type", "VPA/Account", "From VPA", "Reference No", "Email Subject", "Email Link"]);

  }

  // Load last processed message ID from the tracker sheet
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ProcessingTracker");
  let lastProcessedId = trackerSheet ? trackerSheet.getRange("A2").getValue() : null;

  let processedCount = 0; // Counter for processed emails in this run

  for (let i = threads.length - 1; i >= 0; i--) { // Oldest to newest
    const messages = threads[i].getMessages();
    for (const message of messages) {
      const messageId = message.getId();

      // Skip already processed messages
      if (lastProcessedId && messageId <= lastProcessedId) {
        continue;
      }

      try {
        // Get plain text content directly from GmailApp
        const body = message.getPlainBody();
        const emailLink = "https://mail.google.com/mail/u/0/#inbox/" + messageId;

        // Skip non-transaction emails
        if (body.includes('OTP') || body.includes('password') || body.includes('IPIN')) {
          Logger.log("Skipping non-transaction email: " + message.getSubject());
          continue;
        }

        Logger.log("Processing email: " + message.getSubject());
        Logger.log("Email Link: " + emailLink);


        const details = parseTransactionDetails(body);

        // Check if it's a valid transaction (either credit or debit)
        const isValidCredit = details.type === 'credited' && details.amount;
        const isValidDebit = details.type === 'debited' && details.amount && (details.vpa || details.toAccount);

        if (isValidCredit || isValidDebit) {
          const emailDateObj = message.getDate();
          const emailTimeStr = Utilities.formatDate(emailDateObj, Session.getScriptTimeZone(), "HH:mm");

          sheet.appendRow([
            details.date,
            details.time || emailTimeStr,
            details.amount,
            details.type,
            details.vpa || (details.toAccount ? "Account: " + details.toAccount : ""),
            details.fromVPA || "",
            details.reference,
            message.getSubject(),
            emailLink
          ]);
          //message.moveToTrash();
          threads[i].addLabel(label);
          processedCount++; // Increment the counter

        } else {
          Logger.log("No valid transaction details found in email: " + message.getSubject());
          Logger.log("Transaction details: " + JSON.stringify(details));
          continue;
        }

        // Save the last processed message ID *only* if we processed at least one email
        if (trackerSheet && processedCount > 0) {
          trackerSheet.getRange("A2").setValue(messageId);
        }

      } catch (e) {
        Logger.log("Error processing message: " + e.message);
        Logger.log("Message subject: " + message.getSubject());
        continue;
      }
    }
    // After processing all messages in a thread, if we processed any emails, update the tracker
    if (trackerSheet && processedCount > 0) {
      trackerSheet.getRange("A2").setValue(messages[messages.length - 1].getId()); // Use the last message ID from the thread
      processedCount = 0; // Reset counter for the next thread
    }
  }
}

function parseTransactionDetails(body) {
  if (!body) {
    Logger.log("Empty body received");
    return {
      amount: "",
      time:"",
      type: "",
      date: "",
      vpa: "",
      reference: "",
      toAccount: "",
      fromVPA: ""
    };
  }

  // Normalize the body (works for plain text too)
  const cleaned = body.replace(/\s+/g, ' ').trim();
  //Logger.log("Processing email content: " + cleaned.substring(0, 100) + "..."); // Only log first 100 chars

  // Enhanced amount matching - handles HDFC format with or without space after Rs.
  const amountMatch = cleaned.match(/Rs\.?\s*([0-9,]+\.\d{2})/i);

  // Enhanced type matching - handles both debit and credit formats
  const typeMatch = cleaned.match(/\b(debited|credited|sent|received)\b/i);

  // Enhanced VPA matching - handles both debit and credit formats
  const vpaMatch = cleaned.match(/to\s+VPA\s+([a-zA-Z0-9.\-_@]+)(?:\s+[A-Z\s]+)?/i);
  const fromVpaMatch = cleaned.match(/(?:by|from)\s+VPA\s+([a-zA-Z0-9.\-_@]+)(?:\s+[A-Z\s]+)?/i);

  // Account transfer matching
  const accountMatch = cleaned.match(/to\s+account\s+\*\*(\d{4})/i);

  // Enhanced date matching - handles HDFC format
  const dateMatch = cleaned.match(/on\s+(\d{2}-\d{2}-\d{2})/i);

  const timeMatch = cleaned.match(/(?:on\s+\d{2}-\d{2}-\d{2}[^\d]*)?(\d{1,2}:\d{2}(?:\s?[APMapm]{2})?)/i);

  // Enhanced reference number matching - handles HDFC format
  const refMatch = cleaned.match(/reference number is\s+(\d{12})/i);

  // Log matches in a more concise way
  const matches = {
    amount: amountMatch ? amountMatch[1] : "not found",
    type: typeMatch ? typeMatch[1] : "not found",
    vpa: vpaMatch ? vpaMatch[1] : "not found",
    fromVPA: fromVpaMatch ? fromVpaMatch[1] : "not found",
    toAccount: accountMatch ? accountMatch[1] : "not found",
    date: dateMatch ? dateMatch[1] : "not found",
    time: timeMatch ? timeMatch[1] : "not found",
    reference: refMatch ? refMatch[1] : "not found"
  };
  //Logger.log("Matches found: " + JSON.stringify(matches));

  // Extract VPA without any trailing name
  let vpa = "";
  if (vpaMatch) {
    vpa = vpaMatch[1].trim();
  }

  // Determine transaction type
  let transactionType = "";
  const lowerBody = cleaned.toLowerCase();

  // First check for explicit credit indicators
  if (lowerBody.includes('successfully credited') ||
    lowerBody.includes('is credited') ||
    lowerBody.includes('has been credited')) {
    transactionType = 'credited';
  }
  // Then check for explicit debit indicators
  else if (lowerBody.includes('debited') ||
    lowerBody.includes('has been debited') ||
    lowerBody.includes('is debited')) {
    transactionType = 'debited';
  }
  // Finally check for simple credit/debit words
  else if (typeMatch) {
    const matchedType = typeMatch[1].toLowerCase();
    if (['debited', 'sent'].includes(matchedType)) {
      transactionType = 'debited';
    } else if (['credited', 'received'].includes(matchedType)) {
      transactionType = 'credited';
    }
  }

  Logger.log("Transaction type detection: " + JSON.stringify({
    matchedType: typeMatch ? typeMatch[1] : "none",
    finalType: transactionType,
    hasCredited: lowerBody.includes('credited'),
    hasDebited: lowerBody.includes('debited'),
    hasSuccessfullyCredited: lowerBody.includes('successfully credited'),
    hasIsCredited: lowerBody.includes('is credited'),
    hasBeenCredited: lowerBody.includes('has been credited')
  }));

  return {
    amount: amountMatch ? amountMatch[1] : "",
    type: transactionType,
    date: dateMatch ? dateMatch[1] : "",
    time: timeMatch ? timeMatch[1] : "",
    vpa: vpa,
    reference: refMatch ? refMatch[1] : "",
    toAccount: accountMatch ? accountMatch[1] : "",
    fromVPA: fromVpaMatch ? fromVpaMatch[1] : ""
  };
}


function extractPlainTextFromPayload(payload) {
  try {
    if (payload.parts) {
      for (const part of payload.parts) {
        const result = extractPlainTextFromPayload(part);
        if (result) return result;
      }
    }

    if (!payload.body || !payload.body.data) {
      return "";
    }

    // Ensure data is a string
    const base64Data = typeof payload.body.data === 'string'
      ? payload.body.data.replace(/-/g, '+').replace(/_/g, '/')
      : payload.body.data;

    if (payload.mimeType === "text/plain") {
      try {
        const decodedText = Utilities.newBlob(Utilities.base64Decode(base64Data)).getDataAsString();
        //Logger.log("Successfully decoded plain text email");
        return decodedText;
      } catch (e) {
        Logger.log("Failed to decode plain text: " + e.message);
        // Try alternative decoding method
        try {
          const decodedText = Utilities.newBlob(Utilities.base64DecodeWebSafe(payload.body.data)).getDataAsString();
         // Logger.log("Successfully decoded plain text email using WebSafe method");
          return decodedText;
        } catch (e2) {
          Logger.log("Failed to decode plain text with WebSafe method: " + e2.message);
          return "";
        }
      }
    }

    if (payload.mimeType === "text/html") {
      try {
        const decodedHtml = Utilities.newBlob(Utilities.base64Decode(base64Data)).getDataAsString();
        //Logger.log("Successfully decoded HTML email");
        // Remove HTML tags and normalize whitespace
        return decodedHtml.replace(/<[^>]+>/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
      } catch (e) {
        Logger.log("Failed to decode HTML: " + e.message);
        // Try alternative decoding method
        try {
          const decodedHtml = Utilities.newBlob(Utilities.base64DecodeWebSafe(payload.body.data)).getDataAsString();
         // Logger.log("Successfully decoded HTML email using WebSafe method");
          return decodedHtml.replace(/<[^>]+>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        } catch (e2) {
          Logger.log("Failed to decode HTML with WebSafe method: " + e2.message);
          return "";
        }
      }
    }

    return "";
  } catch (e) {
    Logger.log("Error in extractPlainTextFromPayload: " + e.message);
    // Log the payload structure for debugging
    Logger.log("Payload structure: " + JSON.stringify({
      mimeType: payload.mimeType,
      hasBody: !!payload.body,
      dataType: payload.body ? typeof payload.body.data : 'undefined'
    }));
    return "";
  }
}


function getOrCreateLabel_(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  return label ? label : GmailApp.createLabel(labelName);
}

// Function to manually start processing from a specific date
function processFromDate(startDate) {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ProcessingTracker");
  if (trackerSheet) {
    trackerSheet.getRange("A2").clearContent(); // Clear last processed email ID
    trackerSheet.getRange("B2").setValue(startDate.toISOString()); // Set start date
  }
  processUPIEmails();
}

// Function to process all emails from the beginning
function processAllEmails() {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ProcessingTracker");
  if (trackerSheet) {
    trackerSheet.getRange("A2:B2").clearContent(); // Clear last processed email ID and date
  }
  processUPIEmails();
}
