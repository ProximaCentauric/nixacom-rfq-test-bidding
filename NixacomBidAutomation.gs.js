/**
 * ════════════════════════════════════════════════════════════════════
 * NIXACOM BID AUTOMATION — Merged Final Script
 * ════════════════════════════════════════════════════════════════════
 * Base: your working Phase 3 script (triggers, menus, assignment)
 * Added: full Phase 2 auto-submission (proposal doc, PDF, email, portal)
 *
 * SETUP (do once):
 *   1. Uncomment setSecrets(), fill in values, click Run, re-comment it
 *   2. Register triggers in Triggers panel (clock icon):
 *      - onStatusChange  → On change
 *      - pollApproved    → Every 30 minutes
 *      - dailyDiscovery  → Day timer, 7-8 AM
 *   3. Reload sheet — "Nixacom Bids" menu appears
 * ════════════════════════════════════════════════════════════════════
 */


// ── Column indexes (1-based) ──────────────────────────────────────────────────
const TASKS_SHEET  = "Tasks";
const LOGS_SHEET   = "Logs";
const STATUS_COL   = 5;   // E — Status (main trigger column)
const DECISION_COL = 6;   // F — GPT Decision   ← ADD THIS LINE
const CAUTION_COL  = 8;   // H — Caution Override ("Confirmed")
const CONFIRM_COL  = 9;   // I — Submission Confirmation Number
const FORMAT_COL   = 10;  // J — Proposal Format Used
const PHASE2_COL   = 11;  // K — Phase 2 Status detail
const OWNER_COL    = 3;   // C — Assigned Owner
const SOLICIT_COL  = 2;   // B — Solicitation Number
const URL_COL      = 1;   // A — Portal URL

// ── Team email map — update when team changes ─────────────────────────────────
const TEAM_EMAILS = {
  "CEO_name":   "CEO_name@nixacom.com",
  "name_1":  "name_1@nixacom.com",
  "name_2":  "name_2@nixacom.com",
};


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 1 — SETUP
// Fill in values below, run once, then re-comment.
// ════════════════════════════════════════════════════════════════════════════════

// function setSecrets() {
//   PropertiesService.getScriptProperties().setProperties({

//     // ── Mode ───────────────────────────────────────────────────────────────
//     // "true"  = test mode: emails go to TEST_EMAIL, portal = mock portal
//     // "false" = production: emails go to real bid contacts
//     "TEST_MODE":                    "true",
//     "TEST_EMAIL":                   "[PRIVATED]",

//     // ── Nixacom emails ────────────────────────────────────────────────────
//     "[REDACTED]_EMAIL":                 "[PRIVATED]",
//     "CONTRACTS_EMAIL":              "[PRIVATED]",
//     "FROM_NAME":                    "Company Name — Contracts",

//     // ── Spreadsheet ID (from Sheet URL) ───────────────────────────────────
//     "SPREADSHEET_ID":               "[PRIVATED]",

//     // ── Mock portal (test mode only) ──────────────────────────────────────
//     // Run: ngrok http 5000 → paste the https URL here
//     // Leave blank to skip portal upload and do email-only
//     "MOCK_PORTAL_URL":               "https://your-ngrok-url.ngrok-free.dev",
//     "PORTAL_USERNAME":              "[PRIVATED]",
//     "PORTAL_PASSWORD":              "[PRIVATED]",

//     // ── Cloud Function URLs (leave blank until deployed to GCP) ───────────
//     "PHASE1_CLOUD_FUNCTION_URL":    "",
//     "PHASE2_CLOUD_FUNCTION_URL":    "",
//     "DISCOVERY_CLOUD_FUNCTION_URL": "",
//     "OPENAI_API_KEY":               "[PRIVATED]",
//     "SAM_GOV_API_KEY": "[PRIVATED]"
//   });
//   Logger.log("Secrets saved.");
// }


function checkSecrets() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log(JSON.stringify(props, null, 2));
}

// ════════════════════════════════════════════════════════════════════════════════
// SECTION 2 — TRIGGERS
// ════════════════════════════════════════════════════════════════════════════════

/**
 * INSTALLABLE TRIGGER — On change.
 * Fires when any cell in the spreadsheet is edited.
 * Routes based on which column changed.
 */
function onStatusChange(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName(TASKS_SHEET);
    if (!sheet) return;

    const activeSheet = e.source.getActiveSheet();
    if (activeSheet.getName() !== TASKS_SHEET) return;

    const range = e.source.getActiveRange();
    if (!range) return;

    const col = range.getColumn();
    const row = range.getRow();
    if (row <= 1) return; // skip header

    const task = getTaskRow(sheet, row);

    // Status column changed → route to correct phase
    if (col === STATUS_COL) {
      Logger.log(`Status changed → row ${row}: "${task.status}"`);
      switch (task.status) {
        case "Not Started":
          runPhase1(task, row, sheet); // Remove direct analyzeBidWithGPT call
          break;
        case "Approved":           runPhase2(task, row, sheet);     break;
        case "Approved - CAUTION": handleCaution(task, row, sheet); break;
        case "Rejected":           logRejection(task, row, sheet);  break;
        default: Logger.log(`No handler for: "${task.status}"`);
      }
    }

    // Owner column changed → send assignment email
    if (col === OWNER_COL && task.owner) {
      handleOwnerAssignment(task, row, sheet);
    }

    // Caution column set to Confirmed → proceed with Phase 2
    if (col === CAUTION_COL &&
        task.caution === "Confirmed" &&
        task.status === "Approved - CAUTION") {
      Logger.log(`CAUTION confirmed row ${row} — running Phase 2`);
      runPhase2(task, row, sheet);
    }

  } catch (err) {
    Logger.log("onStatusChange error: " + err);
  }
}


/**
 * INSTALLABLE TIME-BASED TRIGGER — every 30 minutes.
 * Safety net: catches Approved rows missed by onStatusChange
 * (e.g. when a Cloud Function writes status programmatically).
 */
function pollApproved() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(TASKS_SHEET);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let processed = 0;

  for (let i = 1; i < data.length; i++) {
    const status = (data[i][STATUS_COL - 1] || "").toString().trim();
    const row    = i + 1;
    const task   = getTaskRow(sheet, row);

    if (status === "Approved") {
      Logger.log(`pollApproved: row ${row} (${task.solicitation})`);
      runPhase2(task, row, sheet);
      processed++;
      Utilities.sleep(3000);
    } else if (status === "Approved - CAUTION" && task.caution === "Confirmed") {
      Logger.log(`pollApproved: CAUTION confirmed row ${row}`);
      runPhase2(task, row, sheet);
      processed++;
      Utilities.sleep(3000);
    }
  }

  if (processed > 0) {
    logToSheet("Poll", "Processed", `${processed} row(s)`, "pollApproved");
  }
}


// /**
//  * INSTALLABLE TIME-BASED TRIGGER — daily at 7 AM ET.
//  * Calls discovery Cloud Function to find new opportunities.
//  */
// function dailyDiscovery() {
//   const props    = PropertiesService.getScriptProperties();
//   const endpoint = props.getProperty("DISCOVERY_CLOUD_FUNCTION_URL");

//   Logger.log("dailyDiscovery: starting...");

//   if (!endpoint) {
//     Logger.log("No DISCOVERY_CLOUD_FUNCTION_URL set — skipping.");
//     return;
//   }

//   try {
//     const resp = UrlFetchApp.fetch(endpoint, {
//       method: "post",
//       contentType: "application/json",
//       payload: JSON.stringify({
//         spreadsheet_id: props.getProperty("SPREADSHEET_ID"),
//         tasks_sheet: TASKS_SHEET
//       }),
//       muteHttpExceptions: true
//     });

//     const code = resp.getResponseCode();
//     if (code !== 200 && code !== 202) {
//       sendAlert("Discovery Failed", `HTTP ${code}: ${resp.getContentText()}`, null);
//     } else {
//       Logger.log("dailyDiscovery: triggered successfully.");
//     }
//   } catch (err) {
//     sendAlert("Discovery Exception", err.toString(), null);
//   }
// }


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 3 — PHASE 1
// Calls Cloud Function for GPT analysis + draft proposal generation.
// Falls back to "Pending Analysis" if no URL is configured yet.
// ════════════════════════════════════════════════════════════════════════════════

function runPhase1(task, row, sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET);
  }

  sheet.getRange(row, STATUS_COL).setValue("Processing — Phase 1");
  logToSheet("Phase 1", "Started", task.solicitation, task.url);

  try {
    // Run the native Apps Script OpenAI analysis directly.
    const result = analyzeBidWithGPT(task, row, sheet);

    if (result && result.ok) {
      // DO NOT overwrite STATUS_COL here — analyzeBidWithGPT already set it to 
      // "Under Review (Phase 1 Complete)" or "Rejected".
      logToSheet("Phase 1", "Completed (Native)", task.solicitation, "Ran analyzeBidWithGPT natively");
    } else {
      const reason = (result && result.reason) || "Unknown failure in analyzeBidWithGPT";
      sheet.getRange(row, STATUS_COL).setValue("Phase 1 Failed");
      logToSheet("Phase 1", "Failed", task.solicitation, reason);
      sendAlert("Phase 1 Failed", reason, task);
    }
  } catch (err) {
    sheet.getRange(row, STATUS_COL).setValue("Phase 1 Failed");
    logToSheet("Phase 1", "Exception", task.solicitation, err.toString().substring(0, 200));
    sendAlert("Phase 1 Exception", err.toString(), task);
  }
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 4 — PHASE 2 (Fully automatic — runs entirely inside Apps Script)
//
// Flow:
//   1. Validate (due date, decision)
//   2. Generate proposal as Google Doc → saved to "Nixacom Proposals" Drive folder
//   3. Export Doc as PDF blob
//   4. Send PDF to bid contact email (or TEST_EMAIL in test mode)
//   5. Submit to mock portal via API (if MOCK_PORTAL_URL is set)
//   6. Write confirmation number + status back to Sheet
//   7. Send confirmation email to CEO_name + owner
// ════════════════════════════════════════════════════════════════════════════════

async function runPhase2(task, row, sheet) {
  const props    = PropertiesService.getScriptProperties();
  const testMode = props.getProperty("TEST_MODE") === "true";
  const SLIDE_TEMPLATE_ID = "1SXQDaJJgrwXXxDA8AY1uHeoEUu7Rmd1DtIiYcTWT-ls";

  // Guard: due date passed
  if (task.dueDate && isPastDue(task.dueDate)) {
    sheet.getRange(row, STATUS_COL).setValue("Expired");
    sendAlert("Bid Expired", `Row ${row} (${task.solicitation}) due date has passed.`, task);
    logToSheet("Phase 2", "Expired", task.solicitation, `Due: ${task.dueDate}`);
    return;
  }

  // Guard: NO GO decision
  if (task.decision && task.decision.toString().includes("NO GO")) {
    sheet.getRange(row, STATUS_COL).setValue("Rejected — No Go");
    logToSheet("Phase 2", "Skipped", task.solicitation, "Decision = NO GO");
    return;
  }

  sheet.getRange(row, STATUS_COL).setValue("Processing — Phase 2");
  logToSheet("Phase 2", "Started", task.solicitation, task.url);

  try {
  // Step 1 & 2 — Generate merged PDF proposal (Cover + Body with replaced placeholders)
    const pdfUrl = await generateMergedProposal(task, testMode);
    logToSheet("Phase 2", "Merged Proposal Created", task.solicitation, pdfUrl);

    // Fetch the final merged PDF blob for email attachment & portal submission
    const pdfId = pdfUrl.match(/\/d\/([^\/]+)/)[1];
    const pdfFile = DriveApp.getFileById(pdfId);
    const pdfBlob = pdfFile.getBlob().setName(`Proposal - ${task.solicitation}.pdf`);

    // Step 3 — Determine submission email
    const submitEmail = testMode
      ? props.getProperty("TEST_EMAIL") || "contracts@nixacom.com"
      : props.getProperty("CONTRACTS_EMAIL");

    // Step 4 — Send proposal email with PDF attached
    sendProposalEmail(task, pdfBlob, submitEmail, testMode, props);
    logToSheet("Phase 2", "Email Sent", task.solicitation, submitEmail);

    // Step 5 — Portal submission (if mock portal URL is configured)
    const portalResult = submitToPortal(task, pdfBlob, props, testMode);
    logToSheet("Phase 2", "Portal " + portalResult.method,
               task.solicitation, portalResult.confirmation);

    // Step 6 — Write results back to Sheet
    const conf = portalResult.confirmation || generateConfirmation();
    const ts   = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm");
    sheet.getRange(row, STATUS_COL).setValue("Submitted");
    sheet.getRange(row, CONFIRM_COL).setValue(conf);
    sheet.getRange(row, FORMAT_COL).setValue("AUTO");
    sheet.getRange(row, PHASE2_COL).setValue(
      `Submitted via ${portalResult.method} | ${ts}`
    );

    // Step 7 — Send confirmation to CEO_name + owner
    sendConfirmationEmail(task, conf, submitEmail, portalResult, pdfUrl, props);
    logToSheet("Phase 2", "Complete", task.solicitation, `Conf: ${conf}`);

    Logger.log(`✅ Phase 2 complete: ${task.solicitation} | ${conf}`);

  } catch (err) {
    Logger.log("Phase 2 error: " + err);
    sheet.getRange(row, STATUS_COL).setValue("Submission Failed");
    sheet.getRange(row, PHASE2_COL).setValue("Error: " + err.toString());
    sendAlert("Phase 2 Failed", err.toString(), task);
    logToSheet("Phase 2", "Failed", task.solicitation, err.toString());
  }
}

async function testRunPhase2() {
  const spreadsheet = SpreadsheetApp.openById("1b_E0zRaOnrhpDXMs3o8ZmFxyaWlnzjCYzYXvAzODAH8"); 
  const sheet = spreadsheet.getActiveSheet();
  
  const mockTask = {
    agencyName: "Nixacom",
    projectTitle: "Website Redesign",
    dueDate: "2026-12-31",
    solicitation: "RFP 2026-TEST-001",
    url: "https://sam.gov",
    decision: "GO"
  };
  
  await runPhase2(mockTask, 2, sheet);
}

// ════════════════════════════════════════════════════════════════════════════════
// SECTION 5 — PROPOSAL DOCUMENT GENERATION
// Creates a Google Doc proposal and saves it to "Nixacom Proposals" in Drive.
// ════════════════════════════════════════════════════════════════════════════════


async function testPipelinePhase2() {
  const sampleTask = {
    agencyName: "Nixacom",
    projectTitle: "Website Redesign",
    solicitation: "RFP 2026-TEST-001",
    url: "https://sam.gov"
  };

  // Await the async PDF generation
  const pdfUrl = await generateMergedProposal(sampleTask, true);
  Logger.log("Generated Merged Proposal PDF URL: " + pdfUrl);
}

async function generateMergedProposal(task, testMode) {
  const SLIDE_TEMPLATE_ID = "1SXQDaJJgrwXXxDA8AY1uHeoEUu7Rmd1DtIiYcTWT-ls";

  const folder       = getOrCreateFolder("Nixacom Proposals");
  const agency       = task.agencyName   || "Procurement Office";
  const projectTitle = task.projectTitle  || "Digital Services";
  const solNum       = task.solicitation || "TBD";
  const dateStr      = Utilities.formatDate(new Date(), "America/New_York", "MMMM d, yyyy");
  const baseName     = `Nixacom Proposal — ${solNum} — ${Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd")}`;

  // ════════════════════════════════════════════════════════════════════════════
  // 1. POPULATE & EXPORT SLIDES COVER PAGE
  // ════════════════════════════════════════════════════════════════════════════
  const templateFile = DriveApp.getFileById(SLIDE_TEMPLATE_ID);
  const tempSlideFile = templateFile.makeCopy(`Temp Cover — ${baseName}`, folder);

  const tempPresentation = SlidesApp.openById(tempSlideFile.getId());
  const slide            = tempPresentation.getSlides()[0];

  slide.replaceAllText("{{PROJECT_TITLE}}", projectTitle);
  slide.replaceAllText("{{SOL_NUM}}",       solNum);
  slide.replaceAllText("{{AGENCY}}",        agency);
  slide.replaceAllText("{{DATE}}",          dateStr);

  tempPresentation.saveAndClose();

  const coverPdfBlob = tempSlideFile.getAs("application/pdf");

  // ════════════════════════════════════════════════════════════════════════════
  // 2. GENERATE DOC BODY PDF
  // ════════════════════════════════════════════════════════════════════════════
  const docId = generateProposalBodyDoc(task, testMode, baseName, folder);
  Logger.log("docId value: " + docId + " | type: " + typeof docId);
  const docFile = DriveApp.getFileById(docId);
  const bodyPdfBlob = docFile.getAs("application/pdf");

  // ════════════════════════════════════════════════════════════════════════════
  // 3. MERGE PDFs & CLEAN UP
  // ════════════════════════════════════════════════════════════════════════════
  const mergedPdfBlob = await mergePdfBlobs([coverPdfBlob, bodyPdfBlob], `Proposal - ${task.solicitation}.pdf`);
  const finalPdfFile  = folder.createFile(mergedPdfBlob);

  // Send temporary working files to trash
  tempSlideFile.setTrashed(true);
  docFile.setTrashed(true);

  return finalPdfFile.getUrl();
}

/**
 * Merges multiple PDF blobs into a single PDF blob using PDF-lib.
 * @param {Blob[]} blobs - Array of Apps Script PDF Blobs.
 * @param {string} fileName - Output file name.
 * @return {Promise<Blob>} Merged PDF Blob.
 */
async function mergePdfBlobs(blobs, fileName) {
  // Polyfill browser timer APIs for Apps Script environment
  if (typeof setTimeout === "undefined") {
    this.setTimeout = function (callback, delay) {
      Utilities.sleep(delay || 0);
      callback();
    };
    this.clearTimeout = function () {};
  }

  // Load pdf-lib CDN into global execution scope
  const cdnUrl = "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js";
  const cdnContent = UrlFetchApp.fetch(cdnUrl).getContentText();
  (0, eval)(cdnContent);

  const mergedPdf = await PDFLib.PDFDocument.create();

  for (const blob of blobs) {
    const rawBytes = blob.getBytes();
    const uint8Array = new Uint8Array(
      rawBytes.map((byte) => (byte < 0 ? byte + 256 : byte))
    );

    const pdf = await PDFLib.PDFDocument.load(uint8Array);
    const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    copiedPages.forEach((page) => mergedPdf.addPage(page));
  }

  const mergedPdfBytes = await mergedPdf.save();
  return Utilities.newBlob(mergedPdfBytes, "application/pdf", fileName);
}



/**
 * Generates a Nixacom proposal Google Doc Body.
 * Cover page: solid navy background, white/blue text.
 *
 * Requires: Advanced Google Docs Service enabled
 *   Apps Script editor → Services (+) → "Google Docs API" → Add
 * (Only used to enforce Letter page dimensions; the rest of the script
 * uses the standard DocumentApp service as before.)
 */
function generateProposalBodyDoc(task, testMode, baseTitle) {
  const folder       = getOrCreateFolder("Nixacom Proposals"); // <--- ADD THIS LINE
  const agency       = task.agencyName   || "Procurement Office";
  const projectTitle = task.projectTitle  || "Digital Services";
  const solNum       = task.solicitation || "TBD";
  const portalUrl    = task.url          || "";
  const dateStr      = Utilities.formatDate(new Date(), "America/New_York", "MMMM d, yyyy");

  let doc = DocumentApp.create(baseTitle + " (Body)");
  const docId = doc.getId();
  let body = doc.getBody();

  // Standard 1-inch margins (72pt) for the entire body document
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // ════════════════════════════════════════════════════════════════════════════
  // TABLE OF CONTENTS (Starts on Page 1)
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "Table of Contents");

  const toc = [
    ["Letter of Transmittal",              "3"],
    ["1. Executive Summary",               "4"],
    ["2. Company Profile",                 "5"],
    ["3. Understanding of the Project",    "6"],
    ["4. Technical Approach",              "7"],
    ["5. Project Management",              "9"],
    ["6. Quality Assurance",               "11"],
    ["7. Pricing Proposal",                "12"],
    ["8. Conclusion",                      "13"],
  ];

  const tocTable = body.appendTable();
  toc.forEach(([section, page]) => {
    const row  = tocTable.appendTableRow();
    const cell = row.appendTableCell(section);
    cell.editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
    const pageCell = row.appendTableCell(page);
    pageCell.setWidth(50);
    pageCell.editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });
  tocTable.setBorderWidth(0);

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // LETTER OF TRANSMITTAL
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "Letter of Transmittal");

  body.appendParagraph(dateStr).setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph("Procurement Evaluation Committee").setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  body.appendParagraph(agency).setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph(portalUrl).setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(`Subject: Proposal for ${projectTitle} – Solicitation No. ${solNum}`).setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph("Dear Members of the Evaluation Committee,").setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    `Nixacom Corporation is pleased to submit this proposal in response to Solicitation No. ${solNum} ` +
    `for ${projectTitle}.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    `We appreciate the opportunity to partner with ${agency} in delivering a modern, secure, ` +
    `accessible, and results-driven digital solution that supports the Agency's mission and ` +
    `long-term strategic objectives.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph("Nixacom proposes a comprehensive end-to-end solution including:").setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  [
    "Strategic planning and discovery",
    "User experience design and information architecture",
    "Custom software and web development",
    "ADA and WCAG 2.1 AA accessibility compliance",
    "Secure cloud hosting and cybersecurity hardening",
    "Quality assurance and testing",
    "Staff training and knowledge transfer",
    "Post-launch maintenance and support",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL)
        .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    "Our multidisciplinary team combines expertise in project management, software engineering, " +
    "UI/UX design, accessibility, cybersecurity, cloud infrastructure, and quality assurance. " +
    "We apply industry best practices throughout the project lifecycle to ensure that every " +
    "deliverable meets the highest standards of quality, performance, security, and maintainability."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    "We certify that all information contained in this proposal is accurate and complete to the " +
    "best of our knowledge. We appreciate your consideration and welcome the opportunity to " +
    "discuss our proposal in greater detail."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  body.appendParagraph("Respectfully submitted,").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("Nixacom Corporation").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  body.appendParagraph("Authorized Representative").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("North Carolina, USA").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 1 — EXECUTIVE SUMMARY
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "1. Executive Summary");

  addSubHeader(body, "Introduction");
  body.appendParagraph(
    "Nixacom Corporation is a technology consulting and software development company " +
    "specializing in digital transformation for government agencies, public organizations, " +
    "and regulated industries. Our mission is to help public-sector organizations modernize " +
    "their digital services through secure, accessible, scalable, and user-centered technology solutions."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, `Understanding of ${agency}'s Objectives`);
  body.appendParagraph(
    `Based on our review of Solicitation No. ${solNum}, ${agency} seeks a qualified technology ` +
    `partner capable of delivering ${projectTitle}. Our proposal addresses each stated requirement ` +
    `through a structured, low-risk implementation methodology.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Our Solution");
  body.appendParagraph(
    `Nixacom will deliver a comprehensive digital solution based on four core principles:`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  [
    "Citizen-Centered Design — Every feature designed to improve usability and public access.",
    "Accessibility by Design — WCAG 2.1 Level AA compliance integrated throughout, not treated as an afterthought.",
    "Security First — Controls embedded throughout the development lifecycle.",
    "Future Scalability — Architecture supports future expansion and integrations.",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Expected Outcomes");
  [
    "A fully delivered solution meeting all solicitation requirements.",
    "Improved organizational efficiency and public engagement.",
    "ADA-compliant and accessible user experience.",
    "Enterprise-grade security and high-performance hosting.",
    "Comprehensive training and long-term operational independence.",
    "A scalable platform supporting future digital initiatives.",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 2 — COMPANY PROFILE
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "2. Company Profile");

  addSubHeader(body, "Corporate Overview");
  body.appendParagraph(
    "Nixacom Corporation is a minority-owned North Carolina C-Corporation specializing in " +
    "digital transformation, custom software development, enterprise web solutions, and " +
    "government technology services. Our multidisciplinary team delivers innovative digital " +
    "solutions that help public agencies modernize operations, improve citizen engagement, " +
    "and increase organizational efficiency through secure, scalable, and user-centered technology."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Core Competencies");
  [
    "Enterprise Web Application Development",
    "Government Website Design and Redesign",
    "Content Management Systems (WordPress, Drupal, Headless CMS)",
    "Cloud Infrastructure (Microsoft Azure, AWS, Cloudflare)",
    "Cybersecurity Implementation and SSDLC",
    "ADA and WCAG 2.1 AA Accessibility Compliance",
    "UI/UX Design and User Research",
    "API Design and Systems Integration",
    "Workflow Automation and Business Process Modernization",
    "DevSecOps and Continuous Integration/Delivery",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 3 — UNDERSTANDING OF THE PROJECT
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "3. Understanding of the Project");

  body.appendParagraph(
    `${agency} requires more than a standard vendor — the Agency seeks a strategic technology ` +
    `partner capable of delivering ${projectTitle} that reflects its mission, supports operational ` +
    `objectives, and enables staff to efficiently manage day-to-day functions.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    `Based on our review of Solicitation No. ${solNum}, we understand that the successful ` +
    `contractor will be expected to:`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  [
    `Deliver ${projectTitle} on time and within budget.`,
    "Provide a secure, scalable, and maintainable solution.",
    "Ensure full ADA and WCAG 2.1 AA accessibility compliance.",
    "Deliver comprehensive staff training and knowledge transfer.",
    "Provide ongoing maintenance and technical support.",
    "Maintain transparent communication and project governance throughout.",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 4 — TECHNICAL APPROACH
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "4. Technical Approach");

  body.appendParagraph(
    "Nixacom will execute this project using a structured Agile framework combined with " +
    "industry-standard software engineering practices. Each project phase concludes with a " +
    "formal review and approval before progressing to the next stage."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  const phases = [
    {
      name: "Phase 1 – Project Initiation and Discovery",
      duration: "Weeks 1–2",
      activities: [
        "Kickoff meeting and stakeholder interviews",
        "Business requirements gathering and validation",
        "Technical assessment and risk identification",
        "Project schedule and communication plan development",
      ],
      deliverables: [
        "Project Charter",
        "Requirements Specification",
        "Initial Risk Register",
        "Work Breakdown Structure (WBS)",
      ],
    },
    {
      name: "Phase 2 – Solution Design",
      duration: "Weeks 3–5",
      activities: [
        "User experience research and persona development",
        "Information architecture and sitemap",
        "Wireframes and high-fidelity mockups",
        "Technical architecture and database design",
        "Security architecture planning",
      ],
      deliverables: [
        "Functional Design Document",
        "Technical Design Document",
        "UI/UX Mockups and Style Guide",
        "System Architecture Diagrams",
      ],
    },
    {
      name: "Phase 3 – Development",
      duration: "Weeks 6–12",
      activities: [
        "Frontend and backend development",
        "CMS configuration and custom templates",
        "Authentication and security implementation",
        "Third-party integrations and API development",
        "Accessibility implementation throughout",
      ],
      deliverables: [
        "Sprint Demonstrations",
        "Source Code Repository",
        "API Documentation",
      ],
    },
    {
      name: "Phase 4 – Testing and Quality Assurance",
      duration: "Weeks 13–14",
      activities: [
        "Functional, integration, and regression testing",
        "Accessibility validation (WCAG 2.1 AA)",
        "Security and performance testing",
        "User Acceptance Testing (UAT)",
      ],
      deliverables: [
        "Test Plans and Test Results",
        "Defect Log",
        "Accessibility Report",
        "UAT Sign-Off",
      ],
    },
    {
      name: "Phase 5 – Deployment",
      duration: "Week 15",
      activities: [
        "Production environment preparation",
        "DNS migration and SSL activation",
        "Performance monitoring and go-live verification",
        "Hypercare support",
      ],
      deliverables: [
        "Production Environment",
        "Deployment Report",
        "Go-Live Approval",
      ],
    },
    {
      name: "Phase 6 – Training and Knowledge Transfer",
      duration: "Week 16",
      activities: [
        "Administrator and end-user training",
        "CMS and system administration training",
        "Technical documentation handover",
        "Knowledge transfer workshops",
      ],
      deliverables: [
        "Administrator Guide",
        "User Manual",
        "Training Videos",
        "Maintenance Procedures",
      ],
    },
  ];

  phases.forEach(phase => {
    addSubHeader(body, `${phase.name} — ${phase.duration}`);

    body.appendParagraph("Activities:").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
    phase.activities.forEach(a => {
      body.appendParagraph(`● ${a}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
    });

    body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
    body.appendParagraph("Deliverables:").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
    phase.deliverables.forEach(d => {
      body.appendParagraph(`● ${d}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
    });
    body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 5 — PROJECT MANAGEMENT
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "5. Project Management");

  addSubHeader(body, "Project Management Approach");
  body.appendParagraph(
    "Nixacom Corporation will execute this project using a hybrid Agile/Scrum methodology " +
    "aligned with PMI (Project Management Institute) best practices. A dedicated Project Manager " +
    "will serve as the primary point of contact throughout the engagement. Weekly status meetings, " +
    "sprint reviews, and executive progress reports will provide complete visibility into project performance."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Project Schedule");

  const schedTable = body.appendTable();
  const schedHeader = schedTable.appendTableRow();
  schedHeader.appendTableCell("Phase").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  schedHeader.appendTableCell("Duration").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  schedHeader.getCell(0).setBackgroundColor("#D6E4F0");
  schedHeader.getCell(1).setBackgroundColor("#D6E4F0");

  [
    ["Project Initiation", "Weeks 1–2"],
    ["Solution Design", "Weeks 3–5"],
    ["Development", "Weeks 6–12"],
    ["Testing & QA", "Weeks 13–14"],
    ["Deployment", "Week 15"],
    ["Training & Knowledge Transfer", "Week 16"],
    ["Total Estimated Duration", "16 Weeks"],
  ].forEach(([phase, dur], i) => {
    const row = schedTable.appendTableRow();
    row.appendTableCell(phase).editAsText().setFontSize(11).setBold(phase === "Total Estimated Duration").setForegroundColor("#000000");
    row.appendTableCell(dur).editAsText().setFontSize(11).setBold(phase === "Total Estimated Duration").setForegroundColor("#000000");
    if (i % 2 === 0) {
      row.getCell(0).setBackgroundColor("#D6E4F0");
      row.getCell(1).setBackgroundColor("#D6E4F0");
    }
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Risk Management");
  [
    ["Scope Changes",           "Formal Change Control Process"],
    ["Schedule Delays",         "Weekly Progress Monitoring and Agile Sprint Planning"],
    ["Integration Challenges",  "Early Technical Validation and Proof-of-Concept"],
    ["User Adoption",           "Comprehensive Training and Change Management"],
    ["Security Threats",        "Secure SDLC, Penetration Testing, Continuous Monitoring"],
    ["Data Loss",               "Automated Backups and Disaster Recovery Procedures"],
  ].forEach(([risk, mitigation]) => {
    body.appendParagraph(`● ${risk}: ${mitigation}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 6 — QUALITY ASSURANCE
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "6. Quality Assurance");

  body.appendParagraph(
    "Quality assurance activities will be integrated throughout the project lifecycle. " +
    "Every software release will undergo the following testing protocols before production deployment:"
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  [
    "Functional Testing — Verification that every feature operates according to approved specifications.",
    "Responsive Testing — Validated on desktop, tablet, and mobile across Windows, macOS, Android, and iOS.",
    "Accessibility Testing — WCAG 2.1 AA review, screen-reader compatibility, keyboard navigation, and contrast validation.",
    "Performance Testing — Page speed, load testing, database optimization, and server response time.",
    "Security Testing — Vulnerability scanning, authentication testing, permission validation, and dependency analysis.",
    "User Acceptance Testing (UAT) — Client personnel validate all business processes before go-live approval.",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  body.appendParagraph(
    "No production deployment will occur until all critical issues have been resolved and " +
    "written acceptance has been received."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 7 — PRICING PROPOSAL
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "7. Pricing Proposal");

  addSubHeader(body, "One-Time Implementation Services");

  const priceTable = body.appendTable();
  const priceHeader = priceTable.appendTableRow();
  priceHeader.appendTableCell("Service").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  priceHeader.appendTableCell("Cost (USD)").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  priceHeader.getCell(0).setBackgroundColor("#D6E4F0");
  priceHeader.getCell(1).setBackgroundColor("#D6E4F0");

  [
    ["Project Management",          "$18,000"],
    ["Business Analysis",           "$12,000"],
    ["UI/UX Design",                "$15,000"],
    ["Frontend Development",        "$28,000"],
    ["Backend Development",         "$32,000"],
    ["API Integration",             "$12,000"],
    ["Cybersecurity Configuration", "$9,000"],
    ["Quality Assurance & Testing", "$15,000"],
    ["Training & Documentation",    "$10,000"],
    ["Deployment & Go-Live Support","$10,000"],
    ["Total One-Time Cost",         "$161,000"],
  ].forEach(([service, cost], i) => {
    const row = priceTable.appendTableRow();
    row.appendTableCell(service).editAsText().setFontSize(11).setBold(service.startsWith("Total")).setForegroundColor("#000000");
    row.appendTableCell(cost).editAsText().setFontSize(11).setBold(service.startsWith("Total")).setForegroundColor("#000000");
    if (i % 2 === 0 && !service.startsWith("Total")) {
      row.getCell(0).setBackgroundColor("#D6E4F0");
      row.getCell(1).setBackgroundColor("#D6E4F0");
    }
    if (service.startsWith("Total")) {
      row.getCell(0).setBackgroundColor("#D6E4F0");
      row.getCell(1).setBackgroundColor("#D6E4F0");
    }
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  addSubHeader(body, "Annual Support & Maintenance");

  const supportTable = body.appendTable();
  const supportHeader = supportTable.appendTableRow();
  supportHeader.appendTableCell("Service").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  supportHeader.appendTableCell("Annual Cost").editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  supportHeader.getCell(0).setBackgroundColor("#D6E4F0");
  supportHeader.getCell(1).setBackgroundColor("#D6E4F0");

  [
    ["Application Support",    "$24,000"],
    ["Security Updates",       "$8,000"],
    ["System Monitoring",      "$6,000"],
    ["Preventive Maintenance", "$5,000"],
    ["Help Desk Services",     "$7,000"],
    ["Annual Total",           "$50,000"],
  ].forEach(([service, cost], i) => {
    const row = supportTable.appendTableRow();
    row.appendTableCell(service).editAsText().setFontSize(11).setBold(service === "Annual Total").setForegroundColor("#000000");
    row.appendTableCell(cost).editAsText().setFontSize(11).setBold(service === "Annual Total").setForegroundColor("#000000");
    if (i % 2 === 0 && service !== "Annual Total") {
      row.getCell(0).setBackgroundColor("#D6E4F0");
      row.getCell(1).setBackgroundColor("#D6E4F0");
    }
    if (service === "Annual Total") {
      row.getCell(0).setBackgroundColor("#D6E4F0");
      row.getCell(1).setBackgroundColor("#D6E4F0");
    }
  });

  body.appendPageBreak();

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 8 — CONCLUSION
  // ════════════════════════════════════════════════════════════════════════════
  addSectionHeader(body, "8. Conclusion");

  body.appendParagraph(
    `Nixacom Corporation appreciates the opportunity to submit this proposal in response to ` +
    `Solicitation No. ${solNum} — ${projectTitle}.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    `We understand that ${agency} requires a trusted technology partner capable of delivering ` +
    `a secure, accessible, scalable, and results-driven solution that strengthens operations ` +
    `and supports the Agency's long-term strategic objectives.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph("Nixacom is committed to:").setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  [
    "Delivering the project on schedule and within budget.",
    "Maintaining transparent communication throughout the engagement.",
    "Ensuring full compliance with accessibility and security requirements.",
    "Providing a high-performance, maintainable, and scalable solution.",
    "Equipping Agency personnel with the knowledge and tools for long-term success.",
  ].forEach(item => {
    body.appendParagraph(`● ${item}`).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  });

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);

  body.appendParagraph(
    "We believe our experience, technical expertise, disciplined delivery methodology, and " +
    "commitment to client success make Nixacom the ideal partner for this initiative. " +
    "We respectfully request the opportunity to discuss our proposal further."
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  body.appendParagraph(
    `Thank you for your consideration. We look forward to the opportunity to partner with ${agency}.`
  ).setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  body.appendParagraph("Respectfully Submitted,").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("Nixacom Corporation").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(true).setForegroundColor("#000000");
  body.appendParagraph("Authorized Representative").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");
  body.appendParagraph("North Carolina, USA").setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setFontSize(11).setBold(false).setForegroundColor("#000000");

  // ── Save and move to folder ──────────────────────────────────────────────────
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return doc.getId();
}

function testGenerateProposalDoc() {
  const dummyTask = {
    agencyName:   "Nixacom",
    projectTitle: "Website Redesign",
    solicitation: "RFP 2026-TEST-001",
    url:          "https://sam.gov",
  };
  const url = generateProposalBodyDoc(dummyTask, true); // true = test mode
  Logger.log("Test doc created: " + url);
}

// ── Helper: section header (navy background style) ───────────────────────────
function addSectionHeader(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.editAsText()
   .setFontSize(13)
   .setForegroundColor("#1F4E79")
   .setBold(true);
  p.setSpacingBefore(12);
  p.setSpacingAfter(8);
  return p;
}

// ── Helper: sub-header ────────────────────────────────────────────────────────
function addSubHeader(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p.editAsText()
   .setFontSize(11)
   .setForegroundColor("#245D91")
   .setBold(true);
  p.setSpacingBefore(8);
  p.setSpacingAfter(4);
  return p;
}

function getPdfBlob(docId, solicitation) {
  // Wait a moment for the Doc to be fully available
  Utilities.sleep(3000);

  const url  = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    // Fallback — export via Drive API instead
    const file = DriveApp.getFileById(docId);
    return file.getAs("application/pdf").setName(
      `Nixacom_Proposal_${solicitation.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`
    );
  }

  return resp.getBlob().setName(
    `Nixacom_Proposal_${solicitation.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`
  );
}

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 6 — EMAIL SUBMISSION
// Sends proposal PDF via GmailApp — no SendGrid needed, 100% Google.
// ════════════════════════════════════════════════════════════════════════════════

function sendProposalEmail(task, pdfBlob, submitEmail, testMode, props) {
  const fromName = props.getProperty("FROM_NAME") || "Nixacom Corp";
  const subject  = `Proposal Submission — ${task.solicitation} — Nixacom Corp` +
                   (testMode ? " [TEST]" : "");

  const body =
    `Dear Procurement Office,\n\n` +
    `Please find attached Nixacom Corp's proposal in response to ` +
    `Solicitation ${task.solicitation}.\n\n` +
    (testMode
      ? `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
        `⚠️  TEST MODE — Pipeline simulation\n` +
        `Solicitation: ${task.solicitation}\n` +
        `Portal: ${task.url || "N/A"}\n` +
        `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`
      : "") +
    `Submitted:  ${Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm")} ET\n` +
    `Vendor:     Nixacom Corp\n` +
    `Contact:    contracts@nixacom.com\n\n` +
    `Respectfully,\n${fromName}\ncontracts@nixacom.com`;

  GmailApp.sendEmail(submitEmail, subject, body, {
    name:        fromName,
    attachments: [pdfBlob],
    cc:          props.getProperty("CONTRACTS_EMAIL") || "",
  });
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 7 — PORTAL SUBMISSION
// Calls mock portal API if MOCK_PORTAL_URL is set.
// Falls back gracefully to email-only if not configured.
// ════════════════════════════════════════════════════════════════════════════════

function submitToPortal(task, pdfBlob, props, testMode) {
  const portalUrl = props.getProperty("MOCK_PORTAL_URL");

  // No portal URL set — email only
  if (!portalUrl) {
    return {
      method:       "Email Only",
      confirmation: generateConfirmation(),
      success:      true,
    };
  }

  try {
    // Login to mock portal
    const loginResp = UrlFetchApp.fetch(portalUrl + "/api/login", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        username: props.getProperty("PORTAL_USERNAME") || "nixacom",
        password: props.getProperty("PORTAL_PASSWORD") || "nixacom2026",
      }),
      muteHttpExceptions: true,
    });

    if (loginResp.getResponseCode() !== 200) {
      throw new Error("Portal login failed: HTTP " + loginResp.getResponseCode());
    }

    // Get session cookie
    const rawCookies = loginResp.getAllHeaders()["Set-Cookie"];
    const cookieStr  = Array.isArray(rawCookies)
      ? rawCookies.join("; ")
      : (rawCookies || "");

    // Build multipart upload
    const boundary = "NixacomBoundary" + Utilities.getUuid().replace(/-/g, "");
    const pdfB64   = Utilities.base64Encode(pdfBlob.getBytes());
    const filename  = pdfBlob.getName();
    const solId     = task.solicitation || "SOL-TEST-001";

    let mp = `--${boundary}\r\n`;
    mp += `Content-Disposition: form-data; name="solicitation_id"\r\n\r\n${solId}\r\n`;
    mp += `--${boundary}\r\n`;
    mp += `Content-Disposition: form-data; name="vendor_name"\r\n\r\nNixacom Corp\r\n`;
    mp += `--${boundary}\r\n`;
    mp += `Content-Disposition: form-data; name="contact_email"\r\n\r\ncontracts@nixacom.com\r\n`;
    mp += `--${boundary}\r\n`;
    mp += `Content-Disposition: form-data; name="proposal"; filename="${filename}"\r\n`;
    mp += `Content-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n`;
    mp += pdfB64 + "\r\n";
    mp += `--${boundary}--\r\n`;

    const uploadResp = UrlFetchApp.fetch(portalUrl + "/api/submit", {
      method:      "post",
      contentType: `multipart/form-data; boundary=${boundary}`,
      payload:     mp,
      headers:     { "Cookie": cookieStr },
      muteHttpExceptions: true,
    });

    const code = uploadResp.getResponseCode();
    const data = JSON.parse(uploadResp.getContentText());

    if (code === 200 && data.success) {
      return {
        method:       "Portal Upload",
        confirmation: data.confirmation,
        success:      true,
      };
    }
    throw new Error("Upload failed: " + uploadResp.getContentText());

  } catch (err) {
    Logger.log("Portal error (falling back to email-only): " + err);
    return {
      method:       "Email Only (portal error)",
      confirmation: generateConfirmation(),
      success:      false,
      error:        err.toString(),
    };
  }
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 8 — CONFIRMATION EMAIL
// Sent to CEO_name + owner after submission completes successfully.
// ════════════════════════════════════════════════════════════════════════════════

function sendConfirmationEmail(task, conf, submitEmail, portalResult, docUrl, props) {
  const CEO_nameEmail = props.getProperty("CEO_name_EMAIL");
  const ownerEmail  = TEAM_EMAILS[task.owner] || "";
  const toList      = [CEO_nameEmail, ownerEmail].filter(Boolean).join(",");
  const ts          = Utilities.formatDate(
    new Date(), "America/New_York", "yyyy-MM-dd HH:mm"
  );

  GmailApp.sendEmail(
    toList,
    `✅ Submitted — ${task.solicitation} — Conf: ${conf}`,
    `Hi team,\n\n` +
    `The proposal for ${task.solicitation} has been submitted automatically.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `  SUBMISSION CONFIRMATION\n` +
    `  Confirmation #: ${conf}\n` +
    `  Solicitation:   ${task.solicitation}\n` +
    `  Submitted To:   ${submitEmail}\n` +
    `  Method:         ${portalResult.method}\n` +
    `  Timestamp:      ${ts} ET\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `Proposal Doc: ${docUrl}\n` +
    `Portal: ${task.url || "N/A"}\n` +
    `Owner: ${task.owner || "N/A"}\n\n` +
    `No further action needed.\n\n— Nixacom Automated Bid Pipeline`,
    { name: "Nixacom Bid Pipeline" }
  );
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 9 — CAUTION + REJECTION HANDLERS
// ════════════════════════════════════════════════════════════════════════════════

function handleCaution(task, row, sheet) {
  if (task.caution && task.caution.toString().trim() === "Confirmed") {
    Logger.log(`CAUTION confirmed row ${row} — proceeding to Phase 2`);
    runPhase2(task, row, sheet);
  } else {
    const props = PropertiesService.getScriptProperties();
    MailApp.sendEmail({
      to:      props.getProperty("CEO_name_EMAIL"),
      subject: `⚠️ Action Required — CAUTION Bid: ${task.solicitation}`,
      body:
        `This bid was flagged CAUTION and needs explicit confirmation.\n\n` +
        `Solicitation: ${task.solicitation}\n` +
        `Portal: ${task.url}\nRow: ${task.row}\n\n` +
        `To submit:\n` +
        `1. Set Status = "Approved - CAUTION"\n` +
        `2. Add "Confirmed" to Column H of row ${task.row}\n\n` +
        `The pipeline will then submit automatically.\n\n— Nixacom Bid Pipeline`
    });
    logToSheet("Caution", "Alert Sent", task.solicitation, "Awaiting Col H confirmation");
  }
}

function logRejection(task, row, sheet) {
  const ts = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm");
  sheet.getRange(row, CONFIRM_COL).setValue(`Rejected: ${ts}`);
  logToSheet("Rejection", "Logged", task.solicitation, `Row ${row}`);
  Logger.log(`Row ${row} (${task.solicitation}) rejected.`);
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 10 — OWNER ASSIGNMENT
// Fires when Column C (Owner) is changed.
// ════════════════════════════════════════════════════════════════════════════════

function handleOwnerAssignment(task, row, sheet) {
  const email = TEAM_EMAILS[task.owner];
  if (!email) {
    Logger.log(`No email for owner: ${task.owner}`);
    return;
  }

  const props    = PropertiesService.getScriptProperties();
  const sheetId  = props.getProperty("SPREADSHEET_ID");
  const sheetUrl = sheetId
    ? `https://docs.google.com/spreadsheets/d/${sheetId}`
    : SpreadsheetApp.getActiveSpreadsheet().getUrl();

  MailApp.sendEmail({
    to:      email,
    cc:      props.getProperty("CEO_name_EMAIL"),
    subject: `[Nixacom] Bid Assigned to You — ${task.solicitation}`,
    body:
      `Hi ${task.owner},\n\n` +
      `You have been assigned to the following bid opportunity:\n\n` +
      `Solicitation: ${task.solicitation}\n` +
      `Portal: ${task.url}\n` +
      `Due Date: ${task.dueDate}\n` +
      `Decision: ${task.decision}\n\n` +
      `Please review the draft proposal and update the status in the sheet.\n\n` +
      `Bid Support Sheet: ${sheetUrl}\n\n` +
      `Questions? Contact contracts@nixacom.com\n\n— Nixacom Bid Pipeline`
  });

  logToSheet("Assignment", "Email Sent", task.solicitation, `To: ${email}`);
  Logger.log(`Assignment email sent to ${email} for row ${row}.`);
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 11 — CUSTOM MENU
// ════════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Nixacom Bids")
    .addItem("▶ Run Phase 2 on Selected Row",  "manualPhase2")
    .addItem("✅ Mark Selected as Approved",    "bulkApprove")
    .addItem("❌ Mark Selected as Rejected",    "bulkReject")
    .addItem("👤 Assign Selected Rows to...",   "assignSelectedRows")
    .addSeparator()
    .addItem("🔍 Check Approved Rows Now",      "pollApproved")
    .addItem("🌐 Run Discovery Now",            "runDiscoveryNow")
    .addSeparator()
    .addItem("📊 View Pipeline Status",         "showPipelineStatus")
    .addToUi();
}

function manualPhase2() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(TASKS_SHEET);
  const row   = sheet.getActiveRange().getRow();
  if (row <= 1) { ui.alert("Select a data row first."); return; }

  const task = getTaskRow(sheet, row);
  const res  = ui.alert(
    `Run Phase 2 for row ${row}?`,
    `Solicitation: ${task.solicitation}\n` +
    `This will generate and submit the proposal automatically.`,
    ui.ButtonSet.YES_NO
  );
  if (res === ui.Button.YES) runPhase2(task, row, sheet);
}

function bulkApprove() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(TASKS_SHEET);
  const sel   = sheet.getActiveRange();
  const start = sel.getRow();
  const count = sel.getNumRows();

  if (ui.alert(
    `Approve ${count} bid(s)?`,
    "This triggers Phase 2 for each selected row.",
    ui.ButtonSet.YES_NO
  ) !== ui.Button.YES) return;

  for (let i = start; i < start + count; i++) {
    if (i <= 1) continue;
    sheet.getRange(i, STATUS_COL).setValue("Approved");
    Utilities.sleep(500);
  }
  ui.alert(`${count} bid(s) approved. Phase 2 running automatically.`);
}

function bulkReject() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(TASKS_SHEET);
  const sel   = sheet.getActiveRange();
  const start = sel.getRow();
  const count = sel.getNumRows();

  if (ui.alert(
    `Reject ${count} bid(s)?`,
    "This cannot be undone automatically.",
    ui.ButtonSet.YES_NO
  ) !== ui.Button.YES) return;

  for (let i = start; i < start + count; i++) {
    if (i <= 1) continue;
    sheet.getRange(i, STATUS_COL).setValue("Rejected");
    Utilities.sleep(300);
  }
  ui.alert(`${count} bid(s) marked Rejected.`);
}

function assignSelectedRows() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(TASKS_SHEET);
  const names = Object.keys(TEAM_EMAILS).join(", ");
  const res   = ui.prompt(
    "Assign Bids",
    `Enter team member name (${names}):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const owner = res.getResponseText().trim();
  if (!TEAM_EMAILS[owner]) {
    ui.alert(`Unknown: "${owner}". Check the TEAM_EMAILS list.`);
    return;
  }

  const sel   = sheet.getActiveRange();
  const start = sel.getRow();
  const count = sel.getNumRows();
  let n = 0;
  for (let i = start; i < start + count; i++) {
    if (i <= 1) continue;
    sheet.getRange(i, OWNER_COL).setValue(owner);
    n++;
    Utilities.sleep(300);
  }
  ui.alert(`Assigned ${n} bid(s) to ${owner}. Assignment emails sent.`);
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 12 — PIPELINE STATUS SIDEBAR
// ════════════════════════════════════════════════════════════════════════════════

function showPipelineStatus() {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet()
                                 .getSheetByName(TASKS_SHEET);
  const data     = sheet.getDataRange().getValues().slice(1);
  const props    = PropertiesService.getScriptProperties();
  const testMode = props.getProperty("TEST_MODE") === "true";

  const counts = {};
  data.forEach(row => {
    const s = (row[STATUS_COL - 1] || "Empty").toString().trim();
    counts[s] = (counts[s] || 0) + 1;
  });

  const rows = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .map(([s, c]) => `<tr><td>${s}</td><td><b>${c}</b></td></tr>`)
    .join("");

  const html = HtmlService.createHtmlOutput(`
    <style>
      body  { font-family: Arial, sans-serif; font-size:13px; padding:12px; }
      h3    { color:#1F4E79; margin-bottom:6px; }
      .mode { background:${testMode ? "#fef08a" : "#dcfce7"};
              color:${testMode ? "#713f12" : "#166534"};
              padding:4px 10px; border-radius:3px; font-size:12px;
              font-weight:700; margin-bottom:10px; display:inline-block; }
      table { width:100%; border-collapse:collapse; margin-top:8px; }
      th    { background:#1F4E79; color:#fff; padding:6px 8px; font-size:12px; }
      td    { padding:6px 8px; border-bottom:1px solid #eee; }
      .btn  { background:#2E75B6; color:#fff; border:none; padding:8px 14px;
              border-radius:4px; cursor:pointer; margin-top:8px;
              width:100%; font-size:13px; }
      .btn2 { background:#1E6B3C; margin-top:4px; }
    </style>
    <h3>📊 Bid Pipeline</h3>
    <div class="mode">${testMode ? "⚠️ TEST MODE" : "✅ PRODUCTION"}</div>
    <table>
      <tr><th>Status</th><th>Count</th></tr>
      ${rows}
    </table>
    <button class="btn"
      onclick="google.script.run.pollApproved(); this.innerText='Running...'">
      🔍 Check Approved Now
    </button>
    <button class="btn btn2"
      onclick="google.script.run.showPipelineStatus()">
      🔄 Refresh
    </button>
  `).setTitle("Bid Pipeline").setWidth(270);

  SpreadsheetApp.getUi().showSidebar(html);
}


// ════════════════════════════════════════════════════════════════════════════════
// SECTION 13 — SHARED UTILITIES
// ════════════════════════════════════════════════════════════════════════════════

function getTaskRow(sheet, row) {
  if (!sheet) {
    Logger.log("getTaskRow: sheet is undefined for row " + row);
    return {
      row, url: "", solicitation: "", owner: "", dueDate: "",
      status: "", decision: "", docUrls: "", caution: "",
      confirm: "", format: "", phase2: "", agencyName: "", projectTitle: "",
      description: ""
    };
  }
  
  // Grab 14 columns (A through N)
  const vals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  return {
    row,
    url:          vals[0]  || "",
    solicitation: vals[1]  || "",
    owner:        vals[2]  || "",
    dueDate:      vals[3]  || "",
    status:       vals[4]  || "",
    decision:     vals[5]  || "",
    docUrls:      vals[6]  || "",
    caution:      vals[7]  || "",
    confirm:      vals[8]  || "",
    format:       vals[9]  || "",
    phase2:       vals[10] || "",
    agencyName:   vals[11] || "",
    projectTitle: vals[12] || "",
    description:  vals[13] || "", // Column N
  };
}

function sendAlert(subject, detail, task) {
  const props = PropertiesService.getScriptProperties();
  const to    = props.getProperty("CEO_name_EMAIL");
  const cc    = task && task.owner && TEAM_EMAILS[task.owner]
               ? TEAM_EMAILS[task.owner]
               : (props.getProperty("CONTRACTS_EMAIL") || "");

  MailApp.sendEmail({
    to, cc,
    subject: `[Nixacom Alert] ${subject}`,
    body:
      `Nixacom Pipeline Alert\n\n` +
      `Issue: ${subject}\n\n` +
      (task
        ? `Solicitation: ${task.solicitation}\nRow: ${task.row}\nPortal: ${task.url}\n\n`
        : "") +
      `Detail: ${detail}\n\n` +
      `Time: ${Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm")} ET`
  });
}

function logToSheet(phase, action, solicitation, detail) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   log = ss.getSheetByName(LOGS_SHEET);

  if (!log) {
    log = ss.insertSheet(LOGS_SHEET);
    const hdr = log.getRange(1, 1, 1, 5);
    hdr.setValues([["Timestamp (ET)", "Phase", "Action", "Solicitation", "Detail"]]);
    hdr.setFontWeight("bold")
       .setBackground("#1F4E79")
       .setFontColor("#FFFFFF");
    log.setFrozenRows(1);
  }

  const ts = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss");
  log.appendRow([ts, phase, action, solicitation, detail]);
}

function generateConfirmation() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let result  = "";
  for (let i = 0; i < 8; i++) {
    result += chars[Math.floor(Math.random() * chars.length)];
  }
  return "NX-" + result;
}

function todayET() {
  return Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
}

function isPastDue(dueDateStr) {
  if (!dueDateStr) return false;
  try { return new Date(dueDateStr) < new Date(); }
  catch (e) { return false; }
}

function analyzeBidWithGPT(task, row, sheet) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("OPENAI_API_KEY") || "";

  if (!apiKey || apiKey === "sk-dummy" || !apiKey.startsWith("sk-")) {
    Logger.log("BidGPT: No real API key — skipping.");
    logToSheet("BidGPT", "Skipped", task.solicitation, "No real OpenAI API key set");
    return { ok: false, reason: "No real OpenAI API key set" };
  }

  const existing = sheet.getRange(row, DECISION_COL).getValue().toString().trim();
  if (existing) return { ok: true, reason: "Decision already set — not overwritten" };

  logToSheet("BidGPT", "Analyzing", task.solicitation, task.url);

  // Construct structured system and user prompts safely
  const systemPrompt = "You are the Nixacom Bid Analyzer. Nixacom specializes in web design, CMS platforms, digital transformation, and WCAG accessibility compliance for government and institutional clients.";
  
  const userContent = `Analyze this bid and return ONLY valid JSON:
Solicitation ID: ${task.solicitation || "Unknown"}
Portal URL: ${task.url || "Unknown"}
Due Date: ${task.dueDate || "Unknown"}
Scope / Description: ${task.description || task.projectTitle || "No description provided."}

JSON Schema Output Required:
{
  "decision": "PURSUE" | "NO GO",
  "confidence": "HIGH" | "MEDIUM" | "LOW",
  "reasoning": "One paragraph explanation.",
  "estimated_value": "Dollar range or Unknown",
  "key_risks": ["risk1", "risk2"],
  "recommended_approach": "How Nixacom should position this bid."
}`;

  const requestPayload = {
    model: "gpt-4o-mini",
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent }
    ],
    max_tokens: 500,
    temperature: 0.3
  };

  const fetchOptions = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true,
    escaping: false // Prevents Apps Script from corrupting the URL endpoint
  };

  try {
    let response;
    let code;
    const maxAttempts = 3;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", fetchOptions);
      code = response.getResponseCode();

      if (code === 429 && attempt < maxAttempts) {
        const backoffMs = attempt === 1 ? 2000 : 5000;
        Logger.log(`BidGPT: HTTP 429 for ${task.solicitation}, retrying in ${backoffMs}ms...`);
        logToSheet("BidGPT", "Retrying", task.solicitation, `HTTP 429 — attempt ${attempt}/${maxAttempts}`);
        Utilities.sleep(backoffMs);
        continue;
      }
      break;
    }

    if (code === 401) {
      logToSheet("BidGPT", "Skipped", task.solicitation, "Invalid API key");
      return { ok: false, reason: "Invalid API key" };
    }
    if (code !== 200) {
      logToSheet("BidGPT", "Failed", task.solicitation, `HTTP ${code}`);
      return { ok: false, reason: `HTTP ${code}` };
    }

    const data = JSON.parse(response.getContentText());
    const result = JSON.parse(data.choices[0].message.content);
    const decision = result.decision || "PURSUE";
    const confidence = (result.confidence || "UNKNOWN").toUpperCase();

    // Write decision back to sheet (Column F)
    const decisionCell = sheet.getRange(row, DECISION_COL);
    decisionCell.setValue(decision);

    // Update Status back to sheet (Column E)
    const statusCell = sheet.getRange(row, STATUS_COL);

    if (decision === "PURSUE") {
      decisionCell.setBackground("#D6F0E0").setFontColor("#1E6B3C");
      statusCell.setValue("Under Review (Phase 1 Complete)");
    } else if (decision === "NO GO") {
      decisionCell.setBackground("#FEE2E2").setFontColor("#B91C1C");
      statusCell.setValue("Rejected");
      const timeZone = Session.getScriptTimeZone();
      const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
      sheet.getRange(row, 9).setValue(`Rejected: ${timestamp}`);
    }

    // Write Confidence Level to Column H (Column 8)
    const confidenceCell = sheet.getRange(row, 8);
    confidenceCell.setValue(confidence);

    if (confidence === "HIGH") {
      confidenceCell.setBackground("#D6F0E0").setFontColor("#1E6B3C");
    } else if (confidence === "MEDIUM") {
      confidenceCell.setBackground("#FEF3C7").setFontColor("#92400E");
    } else if (confidence === "LOW") {
      confidenceCell.setBackground("#FEE2E2").setFontColor("#B91C1C");
    }

    PropertiesService.getScriptProperties().setProperty(
      `GPT_REASONING_ROW_${row}`,
      JSON.stringify(result)
    );

    logToSheet("BidGPT", `Decision: ${decision}`, task.solicitation, (result.reasoning || "").substring(0, 200));
    sendGPTAnalysisEmail(task, result, props);

    return { ok: true, reason: `Decision: ${decision}` };

  } catch (err) {
    logToSheet("BidGPT", "Error", task.solicitation, err.toString().substring(0, 200));
    return { ok: false, reason: err.toString().substring(0, 200) };
  }
}

function sendGPTAnalysisEmail(task, result, props) {
  const testMode = props.getProperty("TEST_MODE") === "true";
  const to       = testMode
    ? props.getProperty("TEST_EMAIL")
    : [props.getProperty("CEO_name_EMAIL"), TEAM_EMAILS[task.owner]].filter(Boolean).join(",");

  const decision = result.decision || "UNKNOWN";
  const emoji    = decision === "PURSUE" ? "✅" : "❌";

  GmailApp.sendEmail(to,
    `${emoji} BidGPT — ${task.solicitation} — ${decision}`,
    `BidGPT has analyzed a new bid.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `  ${emoji} DECISION: ${decision}\n` +
    `  Confidence:      ${result.confidence || "MEDIUM"}\n` +
    `  Solicitation:    ${task.solicitation}\n` +
    `  Portal:          ${task.url}\n` +
    `  Due Date:        ${task.dueDate || "Not specified"}\n` +
    `  Estimated Value: ${result.estimated_value || "Unknown"}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `REASONING:\n${result.reasoning || "None."}\n\n` +
    `KEY RISKS:\n${(result.key_risks || []).map(r => `• ${r}`).join("\n") || "None."}\n\n` +
    `RECOMMENDED APPROACH:\n${result.recommended_approach || "Standard proposal."}\n\n` +
    `— Nixacom BidGPT Analyzer`,
    { name: "Nixacom BidGPT" }
  );
}

function getGPTReasoning(row) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(`GPT_REASONING_ROW_${row}`);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}



// ════════════════════════════════════════════════════════════════════════════════
// PHASE 1 — SAM.GOV DISCOVERY
// ════════════════════════════════════════════════════════════════════════════════
//
// ADD these functions to your NixacomBidAutomation.gs script.
// They REPLACE the existing dailyDiscovery() function.
//
// SETUP:
//   1. Get a free SAM.gov API key:
//      → Go to https://sam.gov
//      → Sign in → Click your name → API Keys → Generate Key
//      → Copy the key
//
//   2. Add to setSecrets():
//      "SAM_GOV_API_KEY": "your-key-here",
//
//   3. Run setSecrets() once, re-comment it.
//
//   4. The dailyDiscovery trigger (7 AM daily) already exists —
//      it will now call this script automatically.
//
// WHAT IT DOES:
//   - Searches SAM.gov for opportunities matching Nixacom's service areas
//   - Filters out duplicates already in the Sheet
//   - Writes new opportunities as "Not Started" rows to the Tasks tab
//   - Sends a discovery summary email to CEO_name
//   - BidGPT then auto-analyzes each new row (once real OpenAI key is set)
// ════════════════════════════════════════════════════════════════════════════════


// ── Keywords that match Nixacom's service areas ───────────────────────────────
// SAM.gov searches title field only — each keyword runs as a separate query
const DISCOVERY_KEYWORDS = [
  "website redesign",
  "web design",
  "CMS",
  "content management",
  "digital transformation",
  "web development",
  "accessibility",
  "WCAG",
  "portal development",
  "mobile application",
  "intranet",
  "digital services",
];

// ── NAICS codes for Nixacom's core services ───────────────────────────────────
// 541511 = Custom Computer Programming Services
// 541512 = Computer Systems Design Services
// 541519 = Other Computer Related Services
// 541430 = Graphic Design Services
const NAICS_CODES = ["541511", "541512", "541519", "541430"];

// ── How many days back to search ─────────────────────────────────────────────
const DAYS_BACK = 30; // daily run checks last 24 hours

// ── Max results per keyword search ───────────────────────────────────────────
const MAX_PER_KEYWORD = 10;


// ════════════════════════════════════════════════════════════════════════════════
// REPLACE your existing dailyDiscovery() with this version
// ════════════════════════════════════════════════════════════════════════════════

function dailyDiscovery() {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("SAM_GOV_API_KEY") || "";

  Logger.log("dailyDiscovery: Starting multi-source discovery (BidNet + SAM.gov)...");

  // ── 1. RUN BIDNET DIRECT DISCOVERY (Gmail Parsing) ─────────────────────────
  let bidNetAdded = 0;
  try {
    logToSheet("Discovery", "Started", "BidNet Direct", "Checking Gmail label: BidNet-Alerts");
    bidNetAdded = parseBidNetEmails(); // Returns count of newly added bids
    logToSheet("Discovery", "Complete", "BidNet Direct", `New: ${bidNetAdded}`);
  } catch (err) {
    Logger.log(`BidNet discovery error: ${err}`);
    logToSheet("Discovery", "Failed", "BidNet Direct", err.toString());
  }

  // ── 2. RUN SAM.GOV DISCOVERY (API Search) ──────────────────────────────────
  logToSheet("Discovery", "Started", "SAM.gov", `Keywords: ${DISCOVERY_KEYWORDS.length}`);

  if (!apiKey || apiKey === "your-key-here") {
    Logger.log("No SAM_GOV_API_KEY set — skipping SAM.gov discovery.");
    logToSheet("Discovery", "Skipped", "SAM.gov", "No SAM_GOV_API_KEY in Script Properties");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET);
  if (!sheet) {
    Logger.log("Tasks sheet not found.");
    return;
  }

  // Get existing solicitation numbers and URLs to avoid duplicates
  const existing = getExistingSolicitations(sheet);
  Logger.log(`Existing solicitations in sheet: ${existing.size}`);

  // Fetch opportunities from SAM.gov
  const opportunities = discoverSAMGov(apiKey);
  Logger.log(`SAM.gov returned ${opportunities.length} raw opportunities`);

  // Filter duplicates AND check if opportunity is relevant IT/Web work
  const newOpps = opportunities.filter(opp => {
    const isNotDuplicate = !existing.has(opp.solicitationNumber) && !existing.has(opp.url);
    const isRelevant = isITWebBid(opp.title || opp.solicitationNumber, opp.description || opp.title);
    
    if (!isRelevant && isNotDuplicate) {
      Logger.log(`Skipped non-IT bid: ${opp.title || opp.solicitationNumber}`);
    }
    
    return isNotDuplicate && isRelevant;
  });
  
  Logger.log(`New SAM.gov opportunities after dedup & IT filtering: ${newOpps.length}`);

  // Write new SAM.gov opportunities to sheet & immediately execute Phase 1
  let samAdded = 0;
  newOpps.forEach(opp => {
    try {
      // 1. Add row to sheet via helper function
      addOpportunityToSheet(sheet, opp);
      samAdded++;
      
      const newRowIndex = sheet.getLastRow();

      // 2. Set Status to "Processing" while GPT analyzes
      sheet.getRange(newRowIndex, STATUS_COL).setValue("Processing — Phase 1");
      SpreadsheetApp.flush();

      // 3. Immediately trigger Phase 1 AI analysis
      analyzeBidWithGPT(opp, newRowIndex, sheet);

      Utilities.sleep(200); // Small delay between writes/API calls
    } catch (err) {
      Logger.log(`Failed to process opportunity ${opp.solicitationNumber}: ${err}`);
    }
  });

  // Send combined summary email
  sendDiscoverySummaryEmail(opportunities.length, samAdded + bidNetAdded, newOpps, props);

  logToSheet("Discovery", "Complete", "SAM.gov",
    `Found: ${opportunities.length} | New SAM: ${samAdded} | New BidNet: ${bidNetAdded}`);
  Logger.log(`Discovery complete. Added and processed ${samAdded} SAM.gov and ${bidNetAdded} BidNet bids.`);
}







// ════════════════════════════════════════════════════════════════════════════════
// SAM.GOV SEARCH
// ════════════════════════════════════════════════════════════════════════════════

// REPLACE the DISCOVERY_KEYWORDS.forEach loop with a single combined query:
function discoverSAMGov(apiKey) {
  const BASE_URL = "https://api.sam.gov/opportunities/v2/search";
  const allOpps  = [];
  const seen     = new Set();

  // Look back 30 days dynamically
  const today = new Date();
  const from  = new Date(today);
  from.setDate(from.getDate() - 30);

  const postedFrom = Utilities.formatDate(from,  "America/New_York", "MM/dd/yyyy");
  const postedTo   = Utilities.formatDate(today, "America/New_York", "MM/dd/yyyy");

  // Loop through target keywords (e.g., "website", "web", "CMS")
  const keywords = ["website", "web", "CMS"];

  keywords.forEach(kw => {
    try {
      const params = {
        api_key:    apiKey,
        limit:      100,
        postedFrom: postedFrom,
        postedTo:   postedTo,
        ptype:      "o,p,k,r",
        title:      kw
      };

      const url      = BASE_URL + "?" + buildQueryString(params);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const opps = data.opportunitiesData || [];

        opps.forEach(opp => {
          const solNum = opp.solicitationNumber || opp.noticeId || "";
          if (solNum && !seen.has(solNum)) {
            seen.add(solNum);
            const parsed = parseSAMOpportunity(opp);
            if (parsed) allOpps.push(parsed);
          }
        });
      }
      Utilities.sleep(300); // Small pause between requests
    } catch (err) {
      Logger.log(`Error searching keyword '${kw}': ${err}`);
    }
  });

  Logger.log(`SAM.gov multi-keyword search returned ${allOpps.length} unique opportunities.`);
  return allOpps;
}


/**
 * Phase 1 Extension: Parses BidNet Direct email alerts into the Tasks sheet.
 */
function parseBidNetEmails() {
  const labelName = "BidNet-Alerts";
  const label = GmailApp.getUserLabelByName(labelName);
  
  if (!label) {
    Logger.log(`Gmail label '${labelName}' not found. Skipping BidNet parsing.`);
    return 0;
  }

  const threads = label.getThreads(0, 5); // Fetch top 5 unhandled email threads
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET);
  if (!sheet) return 0;

  const existing = getExistingSolicitations(sheet);
  let newBidsCount = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(msg => {
      const body = msg.getBody(); // Read HTML body
      
      // Match BidNet solicitation URLs and link title text
      const linkRegex = /<a\s+href="(https:\/\/www\.bidnetdirect\.com\/[^\"]+)"[^>]*>(.*?)<\/a>/gi;
      let match;

      while ((match = linkRegex.exec(body)) !== null) {
        const rawUrl = match[1];
        const titleText = match[2].replace(/<[^>]+>/g, '').trim();

        if (rawUrl.includes("solicitations") || rawUrl.includes("open-bids")) {
          const solNum = "BIDNET-" + Math.floor(Math.random() * 8999 + 1000);

          // Check against existing entries
          if (!existing.has(rawUrl) && !existing.has(solNum)) {

            // Check if title/URL is relevant to IT, Software, or Web Development
            if (!isITWebBid(titleText, rawUrl)) {
              Logger.log(`Skipped non-IT BidNet link: ${titleText}`);
              continue;
            }
            
            // Construct opportunity object strictly mapped to addOpportunityToSheet schema
            const bidNetOpp = {
              url:          rawUrl,
              solicitation: solNum,
              owner:        "",
              dueDate:      "",
              status:       "Not Started",
              decision:     "",
              docUrls:      "",
              agencyName:   "BidNet Direct Agency",
              projectTitle: titleText || "BidNet Opportunity",
              description:  titleText || "BidNet Opportunity"
            };

            try {
              // 1. Add row to sheet
              addOpportunityToSheet(sheet, bidNetOpp);
              existing.add(rawUrl); // Add to local deduplication set
              newBidsCount++;

              const newRowIndex = sheet.getLastRow();

              // 2. Set Status to "Processing" while GPT analyzes
              sheet.getRange(newRowIndex, STATUS_COL).setValue("Processing — Phase 1");
              SpreadsheetApp.flush();

              // 3. Immediately trigger Phase 1 AI analysis
              analyzeBidWithGPT(bidNetOpp, newRowIndex, sheet);

              Utilities.sleep(200); // Small delay between writes/API calls
            } catch (err) {
              Logger.log(`Failed to process BidNet opportunity ${solNum}: ${err}`);
            }
          }
        }
      }
    });

    // Remove label after parsing so threads are processed only once
    thread.removeLabel(label);
  });

  return newBidsCount;
}

function getGPTReasoning(row) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(`GPT_REASONING_ROW_${row}`);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

// ════════════════════════════════════════════════════════════════════════════════
// PARSE SAM.GOV OPPORTUNITY INTO NIXACOM SHEET FORMAT
// ════════════════════════════════════════════════════════════════════════════════

function parseSAMOpportunity(opp) {
  try {
    const solNum    = opp.solicitationNumber || opp.noticeId || "UNKNOWN";
    const title     = opp.title             || "Untitled";
    const agency    = opp.fullParentPathName || opp.organizationHierarchy ||
                      opp.department         || "Federal Agency";
    const dueDate   = parseDate(opp.responseDeadLine || opp.archiveDate);
    const noticeUrl = opp.uiLink            ||
                      `https://sam.gov/opp/${opp.noticeId}/view`;

    // Skip if due date has already passed
    if (dueDate && new Date(dueDate) < new Date()) {
      return null;
    }

    // Get the notice type label
    const typeMap = {
      "o": "Solicitation",
      "p": "Pre-Solicitation",
      "k": "Combined Synopsis",
      "r": "Sources Sought",
      "g": "Sale of Surplus",
      "s": "Special Notice",
      "i": "Intent to Bundle",
    };
    const noticeType = typeMap[opp.type] || opp.type || "RFP";

    return {
      url:          noticeUrl,
      solicitation: solNum,
      owner:        "",           // to be assigned manually or by CEO_name
      dueDate:      dueDate,
      status:       "Not Started",
      decision:     "",           // BidGPT fills this
      docUrls:      "",
      agencyName:   cleanAgencyName(agency),
      projectTitle: title,
      noticeType:   noticeType,
      naics:        opp.naicsCode || "",
      postedDate:   parseDate(opp.postedDate),
      description:  (opp.description || "").substring(0, 500),
    };

  } catch (err) {
    Logger.log("parseSAMOpportunity error: " + err);
    return null;
  }
}

function cleanAgencyName(raw) {
  if (!raw) return "Federal Agency";
  // SAM.gov returns hierarchy like "DEPT OF DEFENSE.ARMY.CORPS OF ENGINEERS"
  // Take the last segment which is the most specific agency
  const parts = raw.split(".");
  return parts[parts.length - 1].trim()
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

function parseDate(raw) {
  if (!raw) return "";
  try {
    // SAM.gov dates come in various formats
    // "2026-07-15T00:00:00-04:00" or "07/15/2026"
    const d = new Date(raw);
    if (isNaN(d.getTime())) return "";
    return Utilities.formatDate(d, "America/New_York", "yyyy-MM-dd");
  } catch (e) {
    return "";
  }
}


// ════════════════════════════════════════════════════════════════════════════════
// WRITE NEW OPPORTUNITY TO TASKS SHEET
// ════════════════════════════════════════════════════════════════════════════════

function addOpportunityToSheet(sheet, opp) {
  // Find the next empty row
  const lastRow = sheet.getLastRow();
  const newRow  = lastRow + 1;

  // Write all 13 columns (A through M)
  sheet.getRange(newRow, 1, 1, 13).setValues([[
    opp.url,          // A — Portal URL
    opp.solicitation, // B — Solicitation Number
    opp.owner,        // C — Owner (blank — to be assigned)
    opp.dueDate,      // D — Due Date
    opp.status,       // E — Status = "Not Started" (triggers BidGPT + Phase 1)
    opp.decision,     // F — Decision (blank — BidGPT fills this)
    opp.docUrls,      // G — Document URLs (blank for now)
    "",               // H — Caution Override
    "",               // I — Submission Confirmation
    "",               // J — Proposal Format
    "",               // K — Phase 2 Status
    opp.agencyName,   // L — Agency Name
    opp.projectTitle, // M — Project Title
  ]]);

  // Color the new row light green so it's visible as new
  sheet.getRange(newRow, 1, 1, 13)
       .setBackground("#F0FFF0")
       .setBorder(true, true, true, true, false, false,
                  "#C8E6C9", SpreadsheetApp.BorderStyle.SOLID);

  Logger.log(`Added: ${opp.solicitation} — ${opp.projectTitle}`);
}


// ════════════════════════════════════════════════════════════════════════════════
// DEDUPLICATION — GET EXISTING SOLICITATION NUMBERS
// ════════════════════════════════════════════════════════════════════════════════

function getExistingSolicitations(sheet) {
  const existing = new Set();
  const lastRow  = sheet.getLastRow();
  if (lastRow <= 1) return existing;

  // Read columns A (URL) and B (Solicitation Number)
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach(row => {
    if (row[0]) existing.add(row[0].toString().trim()); // URL
    if (row[1]) existing.add(row[1].toString().trim()); // Solicitation number
  });

  return existing;
}


// ════════════════════════════════════════════════════════════════════════════════
// DISCOVERY SUMMARY EMAIL
// ════════════════════════════════════════════════════════════════════════════════

function sendDiscoverySummaryEmail(total, added, newOpps, props) {
  const testMode    = props.getProperty("TEST_MODE") === "true";
  const CEO_nameEmail = props.getProperty("CEO_name_EMAIL");
  const to          = testMode
    ? props.getProperty("TEST_EMAIL")
    : CEO_nameEmail;

  if (!to) return;

  const today   = Utilities.formatDate(new Date(), "America/New_York", "MMMM d, yyyy");
  const subject = `📋 Nixacom Discovery — ${added} New Bid${added !== 1 ? "s" : ""} Found — ${today}`;

  // Build opportunity list
  const oppList = newOpps.length > 0
    ? newOpps.map((opp, i) =>
        `${i + 1}. ${opp.projectTitle}\n` +
        `   Agency: ${opp.agencyName}\n` +
        `   Solicitation: ${opp.solicitation}\n` +
        `   Due: ${opp.dueDate || "Not specified"}\n` +
        `   URL: ${opp.url}\n`
      ).join("\n")
    : "No new opportunities found today.";

  const body =
    `Nixacom Daily Bid Discovery Report\n` +
    `${today}\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `  SOURCE:    SAM.gov (Federal)\n` +
    `  SCANNED:   ${total} opportunities\n` +
    `  NEW TODAY: ${added} added to Bid Support Sheet\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    (added > 0
      ? `NEW OPPORTUNITIES:\n\n${oppList}\n\n` +
        `These rows have been added to the Bid Support Sheet as "Not Started".\n` +
        (props.getProperty("OPENAI_API_KEY") &&
         props.getProperty("OPENAI_API_KEY") !== "sk-dummy"
          ? `BidGPT will analyze each one automatically.\n`
          : `BidGPT analysis will run once a real OpenAI API key is configured.\n`)
      : `No new matching opportunities were found today.\n`) +
    `\n— Nixacom Discovery Pipeline\n` +
    (testMode ? "\n⚠️ TEST MODE" : "");

  GmailApp.sendEmail(to, subject, body, { name: "Nixacom Discovery" });
}


// ════════════════════════════════════════════════════════════════════════════════
// MANUAL TRIGGER — run discovery right now from the Nixacom Bids menu
// ════════════════════════════════════════════════════════════════════════════════

function testSAMGov() {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("SAM_GOV_API_KEY");
  
  const today = new Date();
  const from  = new Date(today);
  from.setDate(from.getDate() - 30); // past 30 days

  const postedFrom = Utilities.formatDate(from,  "America/New_York", "MM/dd/yyyy");
  const postedTo   = Utilities.formatDate(today, "America/New_York", "MM/dd/yyyy");

  const url = "https://api.sam.gov/opportunities/v2/search" +
    "?api_key=" + apiKey +
    "&limit=100" +
    "&postedFrom=" + postedFrom +
    "&postedTo=" + postedTo +
    "&ptype=o" +
    "&title=website";

  Logger.log("Calling: " + url.replace(apiKey, "HIDDEN"));
  
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  
  Logger.log("HTTP Code: " + code);
  Logger.log("Response snippet: " + body.substring(0, 500));
}


function runDiscoveryNow() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "Run SAM.gov Discovery Now?",
    "This will search SAM.gov for new opportunities and add them to the Tasks tab.\n\n" +
    "This may take 1-2 minutes.",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  try {
    dailyDiscovery();
    ui.alert("Discovery complete! Check the Tasks tab for new rows and your email for a summary.");
  } catch (err) {
    ui.alert("Discovery failed: " + err.toString());
  }
}


function isITWebBid(title, description) {
  const text = (title + " " + (description || "")).toLowerCase();

  // 1. Mandatory Exclusions (Drop false matches immediately)
  const excludePatterns = [
    /\bwebbing\b/i,
    /\bwebster\b/i,
    /\bwebbers\b/i,
    /\bstrap\b/i,
    /\bstraps\b/i,
    /\bbelt\b/i,
    /\bconstruction management\b/i,
    /\bmass spectrometer\b/i,
    /\bfuel\b/i,
    /\bcompressor\b/i,
    /\bcompressors\b/i,
    /\bcabinets\b/i,
    /\bcisco webex\b/i,
    /\bwebex\b/i,
    /\bwebinar\b/i,
    /\brefuse\b/i,
    /\brecycling\b/i,
    /\bjanitorial\b/i,
    /\bhvac\b/i
  ];

  for (let pattern of excludePatterns) {
    if (pattern.test(text)) return false;
  }

  // 2. Strict Inclusion Keywords (Must match targeted Nixacom web/IT offerings)
  const includePatterns = [
    /\bwebsite\b/i,
    /\bweb design\b/i,
    /\bweb development\b/i,
    /\bweb portal\b/i,
    /\bweb hosting\b/i,
    /\bapplication development\b/i,
    /\bapp development\b/i,
    /\bdigital transformation\b/i,
    /\bwcag\b/i,
    /\baccessibility compliance\b/i,
    /\bcms redesign\b/i,
    /\bcms migration\b/i,
    /\bcontent management system\b/i,
    /\bwordpress\b/i,
    /\bdrupal\b/i
  ];

  return includePatterns.some(pattern => pattern.test(text));
}

// ════════════════════════════════════════════════════════════════════════════════
// UTILITY
// ════════════════════════════════════════════════════════════════════════════════

function buildQueryString(params) {
  return Object.entries(params)
    .filter(([, v]) => v !== null && v !== undefined && v !== "")
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");
}

function processAllPendingBids() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  
  // Loop through all rows skipping the header (Row 1)
  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const currentStatus = data[i][STATUS_COL - 1]; // STATUS_COL is 5 (Column E)

    if (currentStatus === "Not Started") {
      Logger.log(`Processing Row ${rowNum}...`);

      // 1. Reconstruct opportunity object from sheet data
      const opp = {
        title: data[i][0],        // Column A: Title / Agency
        url: data[i][1],          // Column B: URL
        solicitation: data[i][2], // Column C: Solicitation
        description: data[i][3],  // Column D: Description
        dueDate: data[i][6]       // Column G: Due Date
      };

      // 2. Set Status to "Processing"
      sheet.getRange(rowNum, STATUS_COL).setValue("Processing — Phase 1");
      SpreadsheetApp.flush();

      // 3. Run Phase 1 GPT Analysis
      try {
        analyzeBidWithGPT(opp, rowNum, sheet);
      } catch (err) {
        Logger.log(`Failed Row ${rowNum}: ${err}`);
        sheet.getRange(rowNum, STATUS_COL).setValue("Not Started"); // Reset if failed
      }

      Utilities.sleep(300); // Pace API requests
    }
  }

  Logger.log("All pending 'Not Started' bids have been processed!");
}
