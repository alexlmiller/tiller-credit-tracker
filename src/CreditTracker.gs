/**
 * Credit Tracker - Dashboard Setup
 *
 * Creates Credits Config and Credits Tracker sheets with Tiller-matching design.
 * Tracks recurring statement credits by matching transaction descriptions.
 * Designed for credit card credits but works with any recurring credit/reimbursement.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // Using public Gist for now; switch to repo URL when public:
  // templateUrl: 'https://raw.githubusercontent.com/alexlmiller/credit-card-tracker/main/templates/cards.json',
  templateUrl: 'https://gist.githubusercontent.com/alexlmiller/2dec55a2f8dca6abaeee1a2ab545a5cc/raw/cards.json',
  // Default Tiller column mappings (can be overridden in Credits Config Settings)
  tiller: {
    dateColumn: 'B',
    amountColumn: 'E',
    accountColumn: 'G',
    descriptionColumn: 'N'
  }
};

// ============================================================================
// DESIGN CONSTANTS
// ============================================================================

const COLORS = {
  headerBlue: '#4285f4',
  headerText: '#ffffff',
  lightGray: '#f8f9fa',
  mediumGray: '#e8eaed',
  white: '#ffffff',
  success: '#137333',
  successBg: '#e6f4ea',
  warning: '#b06000',
  warningBg: '#fef7e0',
  danger: '#c5221f',
  dangerBg: '#fce8e6',
  neutral: '#5f6368',
  border: '#dadce0',
  progressGreen: '#34a853',
  progressYellow: '#fbbc04',
  progressBg: '#e8eaed',
};

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

function setupCreditTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = createConfigSheet(ss);
  const trackerSheet = createTrackerSheet(ss);
  ss.setActiveSheet(configSheet);

  try {
    SpreadsheetApp.getUi().alert(
      'Setup Complete!',
      'Credits Config and Credits Tracker sheets have been created.\n\n' +
      'Next steps:\n' +
      '1. Use "Credit Tracker > Add Card from Template" to add cards\n' +
      '2. Fill in your Tiller Account names and Anniversary dates\n' +
      '3. Run "Credit Tracker > Rebuild Dashboard" to update Credits Tracker',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    console.log('Setup complete! Check Credits Config and Credits Tracker sheets.');
  }
}

// ============================================================================
// CONFIG SHEET
// ============================================================================

function createConfigSheet(ss) {
  let sheet = ss.getSheetByName('Credits Config');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('Credits Config');

  sheet.getRange('A:N').setFontFamily('Arial').setFontSize(10);
  sheet.setHiddenGridlines(true);

  // Column widths - A is spacer like tracker sheet
  //        A    B    C    D    E    F    G
  const configColWidths = [20, 180, 180, 90, 100, 60, 20];
  configColWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ─────────────────────────────────────────────────────────────────────────
  // TITLE BAR (matches tracker sheet style)
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(1, 8);
  sheet.setRowHeight(2, 36);
  sheet.setRowHeight(3, 4);

  sheet.getRange('B2').setValue('Credit Tracker — Configuration')
    .setFontSize(18).setFontWeight('bold').setFontColor(COLORS.headerBlue);

  sheet.getRange('B3:M3').setBackground(COLORS.headerBlue);

  // ─────────────────────────────────────────────────────────────────────────
  // ACCOUNTS SECTION
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(4, 15); // spacer row

  sheet.getRange('B5:G5').merge()
    .setValue('ACCOUNTS')
    .setBackground(COLORS.headerBlue)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(11);
  sheet.setRowHeight(5, 28);

  sheet.getRange('B6:G6')
    .setValues([['Account Name', 'Tiller Account', 'Annual Fee', 'Anniversary', 'Active', '']])
    .setFontWeight('bold')
    .setBackground(COLORS.mediumGray);

  // Empty rows for accounts (users add via menu)
  const emptyCards = Array(8).fill(['', '', '', '', true, '']);
  sheet.getRange(7, 2, emptyCards.length, 6).setValues(emptyCards);

  sheet.getRange('D7:D14').setNumberFormat('$#,##0');
  sheet.getRange('E7:E14').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F7:F14').insertCheckboxes();
  sheet.getRange('B6:G14').setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Green background for user-editable cells (Tiller convention)
  sheet.getRange('B7:E14').setBackground('#E2FFE2');

  // ─────────────────────────────────────────────────────────────────────────
  // CREDITS SECTION (rows 16+)
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(15, 15); // spacer row

  sheet.getRange('B16:I16').merge()
    .setValue('CREDIT BENEFITS')
    .setBackground(COLORS.headerBlue)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(11);
  sheet.setRowHeight(16, 28);

  sheet.getRange('B17:I17')
    .setValues([['Credit Name', 'Account', 'Amount', 'Frequency', 'Reset', 'Detection Pattern', 'Track', '']])
    .setFontWeight('bold')
    .setBackground(COLORS.mediumGray);

  // Empty rows for credits (populated when cards are added)
  const emptyCredits = Array(30).fill(['', '', '', '', '', '', true, '']);
  sheet.getRange(18, 2, emptyCredits.length, 8).setValues(emptyCredits);

  sheet.getRange('D18:D47').setNumberFormat('$#,##0.00');
  sheet.getRange('H18:H47').insertCheckboxes();

  // Account dropdown - references the Accounts section
  const cardListValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('$B$7:$B$14'), true)
    .build();
  sheet.getRange('C18:C47').setDataValidation(cardListValidation);

  const freqList = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monthly', 'Quarterly', 'SemiAnnual', 'Annual'], true)
    .build();
  sheet.getRange('E18:E47').setDataValidation(freqList);

  const resetList = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Calendar', 'Anniversary'], true)
    .build();
  sheet.getRange('F18:F47').setDataValidation(resetList);

  sheet.getRange('B17:I47').setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Green background for user-editable credit cells
  sheet.getRange('B18:G47').setBackground('#E2FFE2');

  // Additional column widths for credits section
  sheet.setColumnWidth(7, 250); // G - Detection Pattern
  sheet.setColumnWidth(8, 70);  // H - Track
  sheet.setColumnWidth(9, 20);  // I - Spacer

  // ─────────────────────────────────────────────────────────────────────────
  // SETTINGS SECTION (columns K-M, next to Accounts)
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setColumnWidth(10, 20);  // J - Spacer
  sheet.setColumnWidth(11, 150); // K - Setting name
  sheet.setColumnWidth(12, 120); // L - Value
  sheet.setColumnWidth(13, 20);  // M - Spacer

  sheet.getRange('K5:M5').merge()
    .setValue('TILLER SETTINGS')
    .setBackground(COLORS.headerBlue)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(11);

  sheet.getRange('K6:M12')
    .setValues([
      ['Setting', 'Value', ''],
      ['Transactions Sheet', 'Transactions', ''],
      ['Date Column', 'B', ''],
      ['Amount Column', 'E', ''],
      ['Account Column', 'G', ''],
      ['Description Column', 'N', ''],
      ['', '', '']
    ]);

  sheet.getRange('K6:L6')
    .setFontWeight('bold')
    .setBackground(COLORS.mediumGray);

  sheet.getRange('K6:M12').setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Alternating row colors for settings
  sheet.getRange('K8:M8').setBackground(COLORS.lightGray);
  sheet.getRange('K10:M10').setBackground(COLORS.lightGray);

  // Green for editable settings values
  sheet.getRange('L7:L11').setBackground('#E2FFE2');

  // ─────────────────────────────────────────────────────────────────────────
  // INSTRUCTIONS PANEL (below Settings, next to Credits)
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(14, 15); // spacer row

  sheet.getRange('K16:M16').merge()
    .setValue('QUICK START GUIDE')
    .setBackground(COLORS.headerBlue)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(11);

  // Instruction rows - merge K:M for each row to give text more room
  const instructionRows = [
    { row: 17, text: '1. ADD CARDS', isHeader: true },
    { row: 18, text: 'Use menu: Credit Tracker → Add Card from Template' },
    { row: 19, text: 'Or manually fill in the Accounts section above' },
    { row: 20, text: '' },
    { row: 21, text: '2. CONFIGURE TILLER', isHeader: true },
    { row: 22, text: 'Set your Tiller Account name for each account' },
    { row: 23, text: 'Verify the column mappings in Settings above' },
    { row: 24, text: '' },
    { row: 25, text: '3. SET ANNIVERSARY DATES', isHeader: true },
    { row: 26, text: 'Enter account open date for anniversary-based credits' },
    { row: 27, text: '' },
    { row: 28, text: '4. CUSTOMIZE DETECTION', isHeader: true },
    { row: 29, text: 'The "Detection Pattern" matches transaction text' },
    { row: 30, text: 'Adjust patterns if credits aren\'t being detected' },
    { row: 31, text: '' },
    { row: 32, text: '5. REFRESH DASHBOARD', isHeader: true },
    { row: 33, text: 'Run: Credit Tracker → Refresh Dashboard' },
    { row: 34, text: 'Do this after adding/removing accounts' }
  ];

  instructionRows.forEach(({ row, text, isHeader }) => {
    const range = sheet.getRange(`K${row}:M${row}`);
    range.merge().setValue(text).setFontSize(9).setWrap(true);
    if (isHeader) {
      range.setFontWeight('bold').setFontColor(COLORS.headerBlue);
    }
  });

  sheet.getRange('K16:M34').setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('K17:M34').setBackground(COLORS.lightGray);

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER: Tiller Account List (hidden column N)
  // Dynamically extracts unique account names from Transactions sheet
  // ─────────────────────────────────────────────────────────────────────────

  sheet.getRange('N1').setValue('TillerAccounts').setFontWeight('bold').setFontSize(8);

  // UNIQUE formula using INDIRECT to reference Settings values (L7=sheet, L10=column)
  const accountListFormula = '=IFERROR(SORT(UNIQUE(FILTER(INDIRECT($L$7&"!"&$L$10&":"&$L$10),INDIRECT($L$7&"!"&$L$10&":"&$L$10)<>""))),"")';
  sheet.getRange('N2').setFormula(accountListFormula);

  // Data validation for Tiller Account column (C7:C14) using the dynamic list
  const tillerAccountValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('$N$2:$N$50'), true)
    .setAllowInvalid(true)  // Allow manual entry if account not in list yet
    .build();
  sheet.getRange('C7:C14').setDataValidation(tillerAccountValidation);

  sheet.hideColumns(14); // Hide column N

  sheet.setFrozenRows(3);

  return sheet;
}

// ============================================================================
// TRACKER SHEET (Dashboard)
// ============================================================================

function createTrackerSheet(ss) {
  let sheet = ss.getSheetByName('Credits Tracker');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('Credits Tracker');

  sheet.getRange('A:AH').setFontFamily('Arial').setFontSize(10);
  sheet.setHiddenGridlines(true);

  // Column widths for side-by-side layout
  // Left panel (Card sections): A=spacer, B-J=card data (9 cols)
  // Right panel (Expiring Soon): K=spacer, L-O=expiring data, P=spacer
  // Calculation area: Q onwards (hidden)
  //        A    B    C    D    E    F    G    H    I    J    K    L    M    N    O    P
  //             Credit Val Period Rcvd Prog  Reset  Ann YTD YTD%      Cred  Acct  Left Resets
  const colWidths = [20, 145, 55, 80, 60, 80, 85, 55, 55, 80, 20, 130, 125, 60, 85, 20];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ─────────────────────────────────────────────────────────────────────────
  // TITLE BAR
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(1, 8);
  sheet.setRowHeight(2, 36);
  sheet.setRowHeight(3, 4);

  sheet.getRange('B2').setValue('Credit Tracker')
    .setFontSize(18).setFontWeight('bold').setFontColor(COLORS.headerBlue);

  sheet.getRange('O2').setFormula('="as of "&TEXT(NOW(),"mmm d")')
    .setFontSize(9).setFontColor(COLORS.neutral).setHorizontalAlignment('right');

  sheet.getRange('B3:O3').setBackground(COLORS.headerBlue);

  // ─────────────────────────────────────────────────────────────────────────
  // SUMMARY METRICS
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(4, 15);
  sheet.setRowHeight(5, 18);
  sheet.setRowHeight(6, 40);

  // Merge cells for labels (row 5) and values (row 6) to prevent cutoff
  sheet.getRange('B5:C5').merge().setValue('CAPTURED YTD').setFontSize(9).setFontWeight('bold').setFontColor(COLORS.neutral);
  sheet.getRange('D5:F5').merge().setValue('AVAILABLE').setFontSize(9).setFontWeight('bold').setFontColor(COLORS.neutral);
  sheet.getRange('G5:H5').merge().setValue('REMAINING').setFontSize(9).setFontWeight('bold').setFontColor(COLORS.neutral);
  sheet.getRange('I5:J5').merge().setValue('CAPTURE %').setFontSize(9).setFontWeight('bold').setFontColor(COLORS.neutral);

  // Formulas reference calculation area (filtered by Track=TRUE) - columns R-AG
  // AF = Track, AD = YTD, AE = AnnualMax
  // B6: Total credits captured YTD
  sheet.getRange('B6:C6').merge().setFormula('=SUMIF(AF:AF,TRUE,AD:AD)')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLORS.success).setNumberFormat('$#,##0');

  // D6: Total annual credit potential
  sheet.getRange('D6:F6').merge().setFormula('=SUMIF(AF:AF,TRUE,AE:AE)')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLORS.neutral).setNumberFormat('$#,##0');

  // G6: Remaining credits available this year
  sheet.getRange('G6:H6').merge().setFormula('=D6-B6')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLORS.warning).setNumberFormat('$#,##0');

  // I6: Capture rate percentage
  sheet.getRange('I6:J6').merge().setFormula('=IFERROR(B6/D6,0)')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLORS.headerBlue).setNumberFormat('0%');

  const rules = [];

  // ─────────────────────────────────────────────────────────────────────────
  // RIGHT PANEL: EXPIRING SOON (can expand vertically)
  // ─────────────────────────────────────────────────────────────────────────

  sheet.setRowHeight(7, 15);
  sheet.setRowHeight(8, 28);

  sheet.getRange('L8:O8').merge()
    .setValue('⚠️  EXPIRING SOON')
    .setBackground(COLORS.headerBlue)
    .setFontColor(COLORS.headerText)
    .setFontWeight('bold')
    .setFontSize(11);

  sheet.getRange('L9:O9')
    .setValues([['Credit', 'Account', 'Left', 'Resets']])
    .setFontWeight('bold')
    .setBackground(COLORS.mediumGray)
    .setFontSize(9);

  // Expiring Soon: 75% through period AND not fully used AND Track=TRUE
  // Sorted by days left ascending - references calculation area in columns R-AG
  // R=CreditName, S=CardName, Z=Remaining, AB=DaysLeft, AC=PctThruPeriod, AA=PctUsed, AF=Track
  const expiringFormula = `=IFERROR(SORT(FILTER({R2:R17,S2:S17,Z2:Z17,IF(AB2:AB17<=1,"tomorrow","in "&AB2:AB17&" days")},(AC2:AC17>=0.75)*(AA2:AA17<1)*(R2:R17<>"")*(AF2:AF17=TRUE)),3,TRUE),"No credits expiring soon")`;
  sheet.getRange('L10').setFormula(expiringFormula);

  sheet.getRange('N10:N30').setNumberFormat('$#,##0.00');
  sheet.getRange('L9:O9').setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

  // Highlight urgent items
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('tomorrow')
    .setBackground(COLORS.dangerBg).setFontColor(COLORS.danger).setBold(true)
    .setRanges([sheet.getRange('O10:O30')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('in 1 day').setBackground(COLORS.dangerBg).setFontColor(COLORS.danger)
    .setRanges([sheet.getRange('O10:O30')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('in 2 day').setBackground(COLORS.warningBg).setFontColor(COLORS.warning)
    .setRanges([sheet.getRange('O10:O30')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('in 3 day').setBackground(COLORS.warningBg).setFontColor(COLORS.warning)
    .setRanges([sheet.getRange('O10:O30')]).build());

  // ─────────────────────────────────────────────────────────────────────────
  // LEFT PANEL: PER-ACCOUNT SECTIONS (stacked vertically, dynamic spacing)
  // Using INDEX/MATCH to allow sparklines in Progress column
  // ─────────────────────────────────────────────────────────────────────────

  // Read account configs dynamically from Credits Config
  const cardConfigs = getCardConfigsFromSheet();

  if (cardConfigs.length === 0) {
    // No accounts configured yet - show placeholder message
    sheet.getRange('B8').setValue('No accounts configured. Use Credit Tracker menu to add accounts.');
  }

  let sectionStartRow = 8;

  cardConfigs.forEach((card, cardIndex) => {
    const cardName = card.name;

    // Section header with account-level summary stats
    // Card name in B-E, renewal warning in F-G
    sheet.getRange(`B${sectionStartRow}:E${sectionStartRow}`).merge()
      .setValue(`💳  ${cardName.toUpperCase()}`)
      .setBackground(COLORS.headerBlue)
      .setFontColor(COLORS.headerText)
      .setFontWeight('bold')
      .setFontSize(11);
    sheet.setRowHeight(sectionStartRow, 28);

    // F-G: Renewal warning (shows if anniversary is within 30 days)
    // Looks up anniversary date from Credits Config and calculates days until next renewal
    const renewalFormula = `=IFERROR(LET(
      anniv,VLOOKUP("${cardName}",'Credits Config'!$B$7:$E$14,4,FALSE),
      anniv_this_year,DATE(YEAR(TODAY()),MONTH(anniv),DAY(anniv)),
      next_anniv,IF(anniv_this_year<TODAY(),DATE(YEAR(TODAY())+1,MONTH(anniv),DAY(anniv)),anniv_this_year),
      days_until,next_anniv-TODAY(),
      IF(AND(anniv<>"",days_until<=30),"⚠️ Renews in "&days_until&"d","")
    ),"")`;
    sheet.getRange(`F${sectionStartRow}:G${sectionStartRow}`).merge().setFormula(renewalFormula)
      .setFontColor('#fef7e0').setBackground(COLORS.headerBlue)
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left');

    // Account summary stats in header row (right side)
    // H: YTD captured for this account
    sheet.getRange(`H${sectionStartRow}`).setFormula(
      `=SUMIF($S$2:$S$21,"${cardName}",$AD$2:$AD$21)`
    ).setNumberFormat('$#,##0').setFontColor(COLORS.headerText).setFontWeight('bold')
      .setBackground(COLORS.headerBlue).setHorizontalAlignment('right');

    // I: "of" label
    sheet.getRange(`I${sectionStartRow}`).setValue('of')
      .setFontColor(COLORS.headerText).setBackground(COLORS.headerBlue)
      .setHorizontalAlignment('center').setFontSize(9);

    // J: Annual potential for this account
    sheet.getRange(`J${sectionStartRow}`).setFormula(
      `=SUMIF($S$2:$S$21,"${cardName}",$AE$2:$AE$21)`
    ).setNumberFormat('$#,##0').setFontColor(COLORS.headerText).setFontWeight('bold')
      .setBackground(COLORS.headerBlue).setHorizontalAlignment('left');

    // Table header - extended to include YTD columns
    const headerRow = sectionStartRow + 1;
    sheet.getRange(`B${headerRow}:J${headerRow}`)
      .setValues([['Credit', 'Value', 'Period', 'Received', 'Progress', 'Next Reset', 'Annual', 'YTD', 'YTD %']])
      .setFontWeight('bold')
      .setBackground(COLORS.mediumGray)
      .setFontSize(9);

    // Data rows using INDEX/MATCH with RankInCard for sorting
    // Column mapping: S=CardName, AG=RankInCard, R=CreditName, T=MaxAmt, U=Frequency, V=Received, W=PeriodEnd, AD=YTD, AE=AnnualMax
    const dataRow = headerRow + 1;
    const dataRows = card.maxCredits;

    for (let rank = 1; rank <= dataRows; rank++) {
      const displayRow = dataRow + rank - 1;

      // Helper: MATCH formula to find the row with this card and rank
      // Returns row index within R2:R21 range
      const matchFormula = `MATCH(1,($S$2:$S$21="${cardName}")*($AG$2:$AG$21=${rank}),0)`;

      // B: Credit Name
      sheet.getRange(`B${displayRow}`).setFormula(
        `=IFERROR(INDEX($R$2:$R$21,${matchFormula}),"")`
      );

      // C: Value (MaxAmt)
      sheet.getRange(`C${displayRow}`).setFormula(
        `=IFERROR(INDEX($T$2:$T$21,${matchFormula}),"")`
      ).setNumberFormat('$#,##0.00');

      // D: Period (Frequency)
      sheet.getRange(`D${displayRow}`).setFormula(
        `=IFERROR(INDEX($U$2:$U$21,${matchFormula}),"")`
      );

      // E: Received
      sheet.getRange(`E${displayRow}`).setFormula(
        `=IFERROR(INDEX($V$2:$V$21,${matchFormula}),"")`
      ).setNumberFormat('$#,##0.00');

      // F: Progress (Sparkline!) - bar chart showing received vs remaining
      sheet.getRange(`F${displayRow}`).setFormula(
        `=IF(C${displayRow}="","",SPARKLINE({E${displayRow},C${displayRow}-E${displayRow}},{"charttype","bar";"max",C${displayRow};"color1",IF(E${displayRow}>=C${displayRow},"${COLORS.progressGreen}","${COLORS.progressYellow}");"color2","${COLORS.progressBg}"}))`
      );

      // G: Next Reset (PeriodEnd)
      sheet.getRange(`G${displayRow}`).setFormula(
        `=IFERROR(INDEX($W$2:$W$21,${matchFormula}),"")`
      ).setNumberFormat('mm/dd/yyyy');

      // H: Annual Value (AnnualMax)
      sheet.getRange(`H${displayRow}`).setFormula(
        `=IFERROR(INDEX($AE$2:$AE$21,${matchFormula}),"")`
      ).setNumberFormat('$#,##0');

      // I: YTD Received
      sheet.getRange(`I${displayRow}`).setFormula(
        `=IFERROR(INDEX($AD$2:$AD$21,${matchFormula}),"")`
      ).setNumberFormat('$#,##0');

      // J: YTD Progress (sparkline)
      sheet.getRange(`J${displayRow}`).setFormula(
        `=IF(H${displayRow}="","",SPARKLINE({I${displayRow},H${displayRow}-I${displayRow}},{"charttype","bar";"max",H${displayRow};"color1",IF(I${displayRow}>=H${displayRow},"${COLORS.progressGreen}","${COLORS.progressYellow}");"color2","${COLORS.progressBg}"}))`
      );

      // Alternating row colors (extended to column J)
      if (rank % 2 === 0) {
        sheet.getRange(`B${displayRow}:J${displayRow}`).setBackground(COLORS.lightGray);
      }
    }

    // Add borders around the table (extended to column J)
    const tableEndRow = dataRow + dataRows - 1;
    sheet.getRange(`B${headerRow}:J${tableEndRow}`).setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);

    // Conditional formatting for Next Reset dates (only if credit not fully used)
    // E = Received, C = Value - only highlight if E < C (still has value to capture)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(G${dataRow}<>"",$G${dataRow}<=TODAY()+7,$G${dataRow}>TODAY()+1,$E${dataRow}<$C${dataRow})`)
      .setBackground(COLORS.warningBg).setFontColor(COLORS.warning)
      .setRanges([sheet.getRange(`G${dataRow}:G${tableEndRow}`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(G${dataRow}<>"",$G${dataRow}<=TODAY()+1,$E${dataRow}<$C${dataRow})`)
      .setBackground(COLORS.dangerBg).setFontColor(COLORS.danger).setBold(true)
      .setRanges([sheet.getRange(`G${dataRow}:G${tableEndRow}`)]).build());

    // Move to next section: header(1) + subheader(1) + data rows + spacer(1)
    sectionStartRow = dataRow + dataRows + 1;
  });

  // Apply all conditional formatting rules
  sheet.setConditionalFormatRules(rules);

  // ─────────────────────────────────────────────────────────────────────────
  // CALCULATION AREA (Hidden columns R-AG)
  // ─────────────────────────────────────────────────────────────────────────

  setupCalculations(sheet);
  sheet.hideColumns(18, 16); // R through AG (columns 18-33)

  sheet.setFrozenRows(3);

  return sheet;
}

// ============================================================================
// CALCULATION AREA
// ============================================================================

function setupCalculations(sheet) {
  // Read Tiller settings from Credits Config
  const tiller = getTillerSettings();
  const txSheet = tiller.sheetName;
  const dateCol = tiller.dateCol;
  const amtCol = tiller.amountCol;
  const acctCol = tiller.accountCol;
  const descCol = tiller.descriptionCol;

  // Headers - Calculation area now starts at column R (moved from O to make room for YTD columns)
  const headers = [
    'CreditName',   // R
    'AccountName',  // S
    'MaxAmt',       // T
    'Frequency',    // U
    'Received',     // V
    'PeriodEnd',    // W
    'PeriodStart',  // X (helper)
    'TillerAcct',   // Y (helper)
    'Remaining',    // Z
    'PctUsed',      // AA
    'DaysLeft',     // AB
    'PctThruPeriod',// AC
    'YTD',          // AD
    'AnnualMax',    // AE
    'Track',        // AF
    'RankInAcct'    // AG - rank within account by PctUsed ascending
  ];
  sheet.getRange('R1:AG1').setValues([headers]).setFontWeight('bold').setFontSize(8);

  const numCredits = 20;

  for (let i = 0; i < numCredits; i++) {
    const row = 2 + i;
    const configRow = 18 + i;  // Credits section starts at row 18 in new layout

    // R: Credit Name (column B in config)
    sheet.getRange(`R${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",'Credits Config'!B${configRow})`);

    // S: Account Name (column C in config)
    sheet.getRange(`S${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",'Credits Config'!C${configRow})`);

    // T: Max Amount (column D in config)
    sheet.getRange(`T${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",'Credits Config'!D${configRow})`);

    // U: Frequency (column E in config)
    sheet.getRange(`U${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",'Credits Config'!E${configRow})`);

    // Y: Tiller Account (lookup from accounts section B7:C14)
    sheet.getRange(`Y${row}`).setFormula(
      `=IF('Credits Config'!B${configRow}="","",VLOOKUP('Credits Config'!C${configRow},'Credits Config'!$B$7:$C$14,2,FALSE))`
    );

    // X: Period Start (Reset is column F in config, Anniversary lookup from B7:E14)
    const periodStart = `=IF('Credits Config'!B${configRow}="","",
      IF('Credits Config'!E${configRow}="Monthly",DATE(YEAR(TODAY()),MONTH(TODAY()),1),
      IF('Credits Config'!E${configRow}="Quarterly",DATE(YEAR(TODAY()),(ROUNDUP(MONTH(TODAY())/3,0)-1)*3+1,1),
      IF('Credits Config'!E${configRow}="SemiAnnual",IF(MONTH(TODAY())<=6,DATE(YEAR(TODAY()),1,1),DATE(YEAR(TODAY()),7,1)),
      IF('Credits Config'!E${configRow}="Annual",
        IF('Credits Config'!F${configRow}="Calendar",DATE(YEAR(TODAY()),1,1),
          DATE(YEAR(TODAY()),MONTH(VLOOKUP('Credits Config'!C${configRow},'Credits Config'!B7:E14,4,FALSE)),
          DAY(VLOOKUP('Credits Config'!C${configRow},'Credits Config'!B7:E14,4,FALSE)))-
          IF(TODAY()<DATE(YEAR(TODAY()),MONTH(VLOOKUP('Credits Config'!C${configRow},'Credits Config'!B7:E14,4,FALSE)),
          DAY(VLOOKUP('Credits Config'!C${configRow},'Credits Config'!B7:E14,4,FALSE))),365,0)),
      "")))))`;
    sheet.getRange(`X${row}`).setFormula(periodStart);

    // W: Period End
    const periodEnd = `=IF('Credits Config'!B${configRow}="","",
      IF('Credits Config'!E${configRow}="Monthly",EOMONTH(X${row},0),
      IF('Credits Config'!E${configRow}="Quarterly",EOMONTH(X${row},2),
      IF('Credits Config'!E${configRow}="SemiAnnual",EOMONTH(X${row},5),
      IF('Credits Config'!E${configRow}="Annual",DATE(YEAR(X${row})+1,MONTH(X${row}),DAY(X${row}))-1,"")))))`;
    sheet.getRange(`W${row}`).setFormula(periodEnd);

    // V: Amount Received (current period) - Detection Pattern is column G
    const received = `=IF('Credits Config'!B${configRow}="",0,
      IFERROR(SUMIFS('${txSheet}'!$${amtCol}:$${amtCol},
        '${txSheet}'!$${acctCol}:$${acctCol},Y${row},
        '${txSheet}'!$${dateCol}:$${dateCol},">="&X${row},
        '${txSheet}'!$${dateCol}:$${dateCol},"<="&W${row},
        '${txSheet}'!$${amtCol}:$${amtCol},">0",
        '${txSheet}'!$${descCol}:$${descCol},"*"&'Credits Config'!G${configRow}&"*"),0))`;
    sheet.getRange(`V${row}`).setFormula(received);

    // Z: Remaining
    sheet.getRange(`Z${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",MAX(0,T${row}-V${row}))`);

    // AA: Percent Used (0 to 1)
    sheet.getRange(`AA${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",IF(T${row}=0,0,V${row}/T${row}))`);

    // AB: Days Left
    sheet.getRange(`AB${row}`).setFormula(`=IF('Credits Config'!B${configRow}="","",MAX(0,W${row}-TODAY()))`);

    // AC: Percent Through Period (0 to 1)
    sheet.getRange(`AC${row}`).setFormula(
      `=IF('Credits Config'!B${configRow}="","",IF(W${row}=X${row},1,(TODAY()-X${row})/(W${row}-X${row})))`
    );

    // AD: YTD Received - Detection Pattern is column G
    const ytd = `=IF('Credits Config'!B${configRow}="",0,
      IFERROR(SUMIFS('${txSheet}'!$${amtCol}:$${amtCol},
        '${txSheet}'!$${acctCol}:$${acctCol},Y${row},
        '${txSheet}'!$${dateCol}:$${dateCol},">="&DATE(YEAR(TODAY()),1,1),
        '${txSheet}'!$${dateCol}:$${dateCol},"<="&TODAY(),
        '${txSheet}'!$${amtCol}:$${amtCol},">0",
        '${txSheet}'!$${descCol}:$${descCol},"*"&'Credits Config'!G${configRow}&"*"),0))`;
    sheet.getRange(`AD${row}`).setFormula(ytd);

    // AE: Annual Max (Amount is column D, Frequency is column E)
    const annualMax = `=IF('Credits Config'!B${configRow}="",0,
      IF('Credits Config'!E${configRow}="Monthly",'Credits Config'!D${configRow}*12,
      IF('Credits Config'!E${configRow}="Quarterly",'Credits Config'!D${configRow}*4,
      IF('Credits Config'!E${configRow}="SemiAnnual",'Credits Config'!D${configRow}*2,
      'Credits Config'!D${configRow}))))`;
    sheet.getRange(`AE${row}`).setFormula(annualMax);

    // AF: Track (column H in config)
    sheet.getRange(`AF${row}`).setFormula(`=IF('Credits Config'!B${configRow}="",FALSE,'Credits Config'!H${configRow})`);

    // AG: Rank within card by PctUsed ascending (for sorting in display)
    // Only ranks tracked credits, untracked get blank
    // Uses credit name as tiebreaker when PctUsed is equal (e.g., multiple 0% credits)
    sheet.getRange(`AG${row}`).setFormula(
      `=IF(OR(R${row}="",AF${row}=FALSE),"",SUMPRODUCT(($S$2:$S$21=S${row})*($AF$2:$AF$21=TRUE)*($R$2:$R$21<>"")*((AA$2:AA$21<AA${row})+((AA$2:AA$21=AA${row})*(R$2:R$21<R${row}))))+1)`
    );
  }

  // Number formats for calculation area (R-AG)
  sheet.getRange('T:T').setNumberFormat('$#,##0.00');   // MaxAmt
  sheet.getRange('V:V').setNumberFormat('$#,##0.00');   // Received
  sheet.getRange('W:W').setNumberFormat('mm/dd/yyyy');  // PeriodEnd
  sheet.getRange('X:X').setNumberFormat('mm/dd/yyyy');  // PeriodStart
  sheet.getRange('Z:Z').setNumberFormat('$#,##0.00');   // Remaining
  sheet.getRange('AA:AA').setNumberFormat('0%');        // PctUsed
  sheet.getRange('AC:AC').setNumberFormat('0%');        // PctThruPeriod
  sheet.getRange('AD:AD').setNumberFormat('$#,##0.00'); // YTD
  sheet.getRange('AE:AE').setNumberFormat('$#,##0.00'); // AnnualMax
}

// ============================================================================
// MENU
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Credit Tracker')
    .addItem('Add Card from Template...', 'showAddCardDialog')
    .addItem('Add Custom Account', 'addCustomCard')
    .addSeparator()
    .addItem('Refresh Dashboard', 'rebuildDashboard')
    .addSeparator()
    .addItem('Full Setup / Reset', 'setupCreditTracker')
    .addToUi();
}

// ============================================================================
// TEMPLATE FUNCTIONS
// ============================================================================

/**
 * Fetches card templates from GitHub
 */
function fetchCardTemplates() {
  try {
    const response = UrlFetchApp.fetch(CONFIG.templateUrl);
    const data = JSON.parse(response.getContentText());
    return data;
  } catch (e) {
    throw new Error('Failed to fetch templates: ' + e.message);
  }
}

/**
 * Shows dialog to select a card template to add
 */
function showAddCardDialog() {
  const ui = SpreadsheetApp.getUi();

  let templates;
  try {
    templates = fetchCardTemplates();
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
    return;
  }

  // Build list of card names
  const cardNames = templates.cards.map(c => c.name);

  // Show prompt with numbered list
  let message = 'Available card templates:\n\n';
  cardNames.forEach((name, i) => {
    message += `${i + 1}. ${name}\n`;
  });
  message += '\nEnter the number of the card to add:';

  const result = ui.prompt('Add Card from Template', message, ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText().trim(), 10);
    if (selection >= 1 && selection <= cardNames.length) {
      const selectedCard = templates.cards[selection - 1];
      addCardFromTemplate(selectedCard);
      ui.alert('Success', `Added "${selectedCard.name}" with ${selectedCard.credits.length} credits.\n\nRun "Rebuild Dashboard" to update Credits Tracker.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Invalid selection. Please enter a number from the list.');
    }
  }
}

/**
 * Adds a card and its credits from a template to Credits Config
 */
function addCardFromTemplate(cardTemplate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Credits Config');

  if (!configSheet) {
    throw new Error('Credits Config sheet not found. Run "Full Setup / Reset" first.');
  }

  // Find next empty row in Accounts section (rows 7-14, column B)
  const cardsRange = configSheet.getRange('B7:B14');
  const cardsValues = cardsRange.getValues();
  let cardRow = -1;
  for (let i = 0; i < cardsValues.length; i++) {
    if (!cardsValues[i][0]) {
      cardRow = 7 + i;
      break;
    }
  }

  if (cardRow === -1) {
    throw new Error('No empty account slots available. Maximum 8 accounts supported.');
  }

  // Add the account (columns B-F)
  configSheet.getRange(cardRow, 2, 1, 5).setValues([[
    cardTemplate.name,
    '', // Tiller Account - user needs to fill in
    cardTemplate.annualFee,
    '', // Anniversary date - user needs to fill in
    true // Active
  ]]);

  // Find next empty row in Credits section (rows 18-47, column B)
  const creditsRange = configSheet.getRange('B18:B47');
  const creditsValues = creditsRange.getValues();
  let creditRow = -1;
  for (let i = 0; i < creditsValues.length; i++) {
    if (!creditsValues[i][0]) {
      creditRow = 18 + i;
      break;
    }
  }

  if (creditRow === -1) {
    throw new Error('No empty credit slots available.');
  }

  // Add all credits for this account (columns B-H)
  cardTemplate.credits.forEach((credit, i) => {
    configSheet.getRange(creditRow + i, 2, 1, 7).setValues([[
      credit.name,
      cardTemplate.name,
      credit.amount,
      credit.frequency,
      credit.reset,
      credit.pattern,
      true // Track
    ]]);
  });
}

/**
 * Adds a blank custom account entry
 */
function addCustomCard() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Add Custom Account', 'Enter the account name:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() === ui.Button.OK) {
    const cardName = result.getResponseText().trim();
    if (cardName) {
      addCardFromTemplate({
        name: cardName,
        annualFee: 0,
        credits: []
      });
      ui.alert('Success', `Added "${cardName}".\n\nNow add credits for this account in Credits Config, then run "Rebuild Dashboard".`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Rebuilds just the Credits Tracker dashboard based on current Credits Config
 */
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Credits Config');

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Credits Config sheet not found. Run "Full Setup / Reset" first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  createTrackerSheet(ss);

  SpreadsheetApp.getUi().alert('Dashboard Rebuilt', 'Credits Tracker has been updated based on your current Credits Config.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Reads Tiller settings from Credits Config sheet
 * Returns object with sheet name and column mappings
 */
function getTillerSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Credits Config');

  if (!configSheet) {
    // Return defaults if no config sheet
    return {
      sheetName: 'Transactions',
      dateCol: 'B',
      amountCol: 'E',
      accountCol: 'G',
      descriptionCol: 'N'
    };
  }

  // Read settings from L7:L11 (Value column in Settings section)
  const settingsRange = configSheet.getRange('L7:L11');
  const values = settingsRange.getValues();

  return {
    sheetName: values[0][0] || 'Transactions',
    dateCol: values[1][0] || 'B',
    amountCol: values[2][0] || 'E',
    accountCol: values[3][0] || 'G',
    descriptionCol: values[4][0] || 'N'
  };
}

/**
 * Reads account configuration from Credits Config sheet
 * Returns array of { name, maxCredits } for each unique account with credits
 */
function getCardConfigsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Credits Config');

  if (!configSheet) {
    return [];
  }

  // Read credits section (B18:C47) to get credit names and account names
  const creditsRange = configSheet.getRange('B18:C47');
  const creditsValues = creditsRange.getValues();

  // Count credits per account
  const cardCounts = {};
  creditsValues.forEach(row => {
    const creditName = row[0];
    const cardName = row[1];
    if (creditName && cardName) {
      cardCounts[cardName] = (cardCounts[cardName] || 0) + 1;
    }
  });

  // Convert to array format, preserving order of first appearance
  const cardConfigs = [];
  const seenCards = new Set();
  creditsValues.forEach(row => {
    const cardName = row[1];
    if (cardName && !seenCards.has(cardName)) {
      seenCards.add(cardName);
      cardConfigs.push({
        name: cardName,
        maxCredits: cardCounts[cardName]
      });
    }
  });

  return cardConfigs;
}
