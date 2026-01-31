const ExcelJS = require('exceljs');

async function createStartingStrengthTracker() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Barbell Strength Tracker';
    workbook.created = new Date();

    // Color scheme
    const colors = {
        headerBg: 'FF2F5496',
        headerText: 'FFFFFFFF',
        subHeaderBg: 'FFD6DCE4',
        lightBg: 'FFF2F2F2',
        success: 'FF92D050',
        inputCell: 'FFFFFFCC'
    };

    // Reusable style helpers
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
    };
    const inputFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.inputCell } };
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };

    // ==================== README SHEET ====================
    const readmeSheet = workbook.addWorksheet('README', {
        properties: { tabColor: { argb: 'FF0070C0' } }
    });

    readmeSheet.columns = [
        { key: 'col1', width: 5 },
        { key: 'col2', width: 50 },
        { key: 'col3', width: 50 },
        { key: 'col4', width: 5 }
    ];

    // ===== HEADER BANNER =====
    const bannerRow1 = readmeSheet.addRow(['', '', '', '']);
    readmeSheet.mergeCells(1, 1, 1, 4);
    bannerRow1.height = 15;
    bannerRow1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };

    const titleRow = readmeSheet.addRow(['', 'BARBELL STRENGTH TRACKER', '', '']);
    readmeSheet.mergeCells(2, 1, 2, 4);
    titleRow.height = 45;
    titleRow.getCell(1).value = '🏋️  BARBELL STRENGTH TRACKER  🏋️';
    titleRow.getCell(1).font = { bold: true, size: 28, color: { argb: 'FFFFFFFF' } };
    titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    const subtitleRow = readmeSheet.addRow(['', '', '', '']);
    readmeSheet.mergeCells(3, 1, 3, 4);
    subtitleRow.height = 25;
    subtitleRow.getCell(1).value = 'Novice Linear Progression Program by Mark Rippetoe';
    subtitleRow.getCell(1).font = { italic: true, size: 14, color: { argb: 'FFFFFFFF' } };
    subtitleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    subtitleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    const versionRow = readmeSheet.addRow(['', '', '', '']);
    readmeSheet.mergeCells(4, 1, 4, 4);
    versionRow.height = 20;
    versionRow.getCell(1).value = 'Version 2.5 | Auto-Detection, Progress Tracking & Smart Exercise Introduction';
    versionRow.getCell(1).font = { size: 10, color: { argb: 'FFD6DCE4' } };
    versionRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    versionRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    readmeSheet.addRow(['']); // Spacer

    // ===== QUICK START BOX =====
    const quickStartHeader = readmeSheet.addRow(['', '⚡ QUICK START', '', '']);
    readmeSheet.mergeCells(6, 2, 6, 3);
    quickStartHeader.height = 28;
    quickStartHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    quickStartHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
    quickStartHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const quickSteps = [
        '1️⃣  Go to Settings → Enter your starting weights',
        '2️⃣  Go to Workout Log → Select Date, Workout (A/B), Exercise',
        '3️⃣  Enter Actual Weight lifted and Reps for each set',
        '4️⃣  Check Status column → "OK" = success, "STALL" = missed reps',
        '5️⃣  Review Progress Chart → Watch your gains grow!'
    ];

    quickSteps.forEach(step => {
        const row = readmeSheet.addRow(['', step, '', '']);
        readmeSheet.mergeCells(row.number, 2, row.number, 3);
        row.getCell(2).font = { size: 11 };
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
        row.getCell(2).alignment = { vertical: 'middle', indent: 1 };
        row.height = 22;
    });

    readmeSheet.addRow(['']); // Spacer

    // ===== PROGRAM OVERVIEW BOX =====
    const programHeader = readmeSheet.addRow(['', '📋 THE PROGRAM', '', '']);
    readmeSheet.mergeCells(programHeader.number, 2, programHeader.number, 3);
    programHeader.height = 28;
    programHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    programHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } };
    programHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    // Workout A
    const workoutARow = readmeSheet.addRow(['', '  WORKOUT A', 'Squat 3×5  |  Bench Press 3×5  |  Deadlift 1×5', '']);
    workoutARow.height = 26;
    workoutARow.getCell(2).font = { bold: true, size: 12, color: { argb: 'FF1F4E79' } };
    workoutARow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
    workoutARow.getCell(3).font = { size: 11 };
    workoutARow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
    workoutARow.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };

    // Workout B
    const workoutBRow = readmeSheet.addRow(['', '  WORKOUT B', 'Squat 3×5  |  Overhead Press 3×5  |  Deadlift 1×5', '']);
    workoutBRow.height = 26;
    workoutBRow.getCell(2).font = { bold: true, size: 12, color: { argb: 'FF375623' } };
    workoutBRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    workoutBRow.getCell(3).font = { size: 11 };
    workoutBRow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    workoutBRow.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };

    // Schedule
    const scheduleRow = readmeSheet.addRow(['', '  📅 SCHEDULE', '3 days/week: Mon-Wed-Fri alternating A-B-A, B-A-B', '']);
    scheduleRow.height = 26;
    scheduleRow.getCell(2).font = { bold: true, size: 11 };
    scheduleRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    scheduleRow.getCell(3).font = { size: 11 };
    scheduleRow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    scheduleRow.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };

    readmeSheet.addRow(['']); // Spacer

    // ===== SHEETS GUIDE =====
    const sheetsHeader = readmeSheet.addRow(['', '📑 SHEETS GUIDE', '', '']);
    readmeSheet.mergeCells(sheetsHeader.number, 2, sheetsHeader.number, 3);
    sheetsHeader.height = 28;
    sheetsHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    sheetsHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFED7D31' } };
    sheetsHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const sheets = [
        { name: '⚙️ Settings', desc: 'Configure starting weights, increments, bar weight' },
        { name: '💪 Assistance Exercises', desc: 'Add/remove optional exercises (don\'t affect progression)' },
        { name: '📝 Workout Log', desc: 'Main tracking - auto-calculates targets & detects stalls' },
        { name: '🔥 Warm-Up Calculator', desc: 'Enter working weight, see warm-up progression' },
        { name: '⚖️ Body Weight Log', desc: 'Optional weekly weigh-ins with statistics' },
        { name: '📊 Progress Summary', desc: 'Auto-updating weekly best lifts table' },
        { name: '📈 Progress Chart', desc: 'Visual chart of your strength progression' },
        { name: '🎯 Program Phase', desc: 'Auto-detects stalls & transition timing' }
    ];

    sheets.forEach((sheet, index) => {
        const row = readmeSheet.addRow(['', sheet.name, sheet.desc, '']);
        row.height = 24;
        row.getCell(2).font = { bold: true, size: 11 };
        row.getCell(3).font = { size: 11, color: { argb: 'FF404040' } };
        const bgColor = index % 2 === 0 ? 'FFFEF4EC' : 'FFFFFFFF';
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    });

    readmeSheet.addRow(['']); // Spacer

    // ===== TIPS BOX =====
    const tipsHeader = readmeSheet.addRow(['', '💡 PRO TIPS', '', '']);
    readmeSheet.mergeCells(tipsHeader.number, 2, tipsHeader.number, 3);
    tipsHeader.height = 28;
    tipsHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FF7030A0' } };
    tipsHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4DFEC' } };
    tipsHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const tips = [
        '✓ Target Weight auto-fills based on your last successful lift + increment',
        '✓ Just enter Actual Weight and Reps - everything else calculates automatically',
        '✓ STALL means you missed target reps - try same weight next session',
        '✓ After 3 stalls on an exercise, consider moving to Intermediate program',
        '✓ Deload 10% after stalls, work back up before changing programs'
    ];

    tips.forEach(tip => {
        const row = readmeSheet.addRow(['', tip, '', '']);
        readmeSheet.mergeCells(row.number, 2, row.number, 3);
        row.getCell(2).font = { size: 11 };
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F0F7' } };
        row.getCell(2).alignment = { vertical: 'middle', indent: 1 };
        row.height = 22;
    });

    readmeSheet.addRow(['']); // Spacer

    // ===== WARNING BOX =====
    const warningHeader = readmeSheet.addRow(['', '⚠️ WHEN TO TRANSITION TO INTERMEDIATE', '', '']);
    readmeSheet.mergeCells(warningHeader.number, 2, warningHeader.number, 3);
    warningHeader.height = 28;
    warningHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FF9C5700' } };
    warningHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    warningHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const warnings = [
        '→ Program Phase sheet shows "⚠️ TRANSITION NEEDED" for an exercise',
        '→ You\'ve accumulated 3+ stalls at the same weight',
        '→ Deloading and working back up no longer helps',
        '→ Recovery between sessions feels insufficient'
    ];

    warnings.forEach(warning => {
        const row = readmeSheet.addRow(['', warning, '', '']);
        readmeSheet.mergeCells(row.number, 2, row.number, 3);
        row.getCell(2).font = { size: 11 };
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
        row.getCell(2).alignment = { vertical: 'middle', indent: 1 };
        row.height = 22;
    });

    readmeSheet.addRow(['']); // Spacer

    // ===== EXERCISE INTRODUCTION BOX =====
    const introHeader = readmeSheet.addRow(['', '🎯 EXERCISE INTRODUCTION TIMING', '', '']);
    readmeSheet.mergeCells(introHeader.number, 2, introHeader.number, 3);
    introHeader.height = 28;
    introHeader.getCell(2).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    introHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
    introHeader.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const introInfo = [
        '🏋️ Light Squat: Appears when any lift reaches INTERMEDIATE (stall threshold)',
        '   → Uses 80% of your Squat working weight (configurable in Settings)',
        '   → Helps maintain squat frequency when transitioning programs',
        '💪 Chin Ups: Appears after 2 weeks of training (configurable in Settings)',
        '   → Gives time to establish baseline before adding pulling work',
        '⚙️ Customize timing in Settings under "EXERCISE INTRODUCTION"'
    ];

    introInfo.forEach(info => {
        const row = readmeSheet.addRow(['', info, '', '']);
        readmeSheet.mergeCells(row.number, 2, row.number, 3);
        row.getCell(2).font = { size: 11 };
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F0F7' } };
        row.getCell(2).alignment = { vertical: 'middle', indent: 1 };
        row.height = 22;
    });

    readmeSheet.addRow(['']); // Spacer

    // ===== FOOTER =====
    const footerRow = readmeSheet.addRow(['', 'Based on "Starting Strength: Basic Barbell Training" by Mark Rippetoe', '', '']);
    readmeSheet.mergeCells(footerRow.number, 2, footerRow.number, 3);
    footerRow.getCell(2).font = { italic: true, size: 10, color: { argb: 'FF808080' } };
    footerRow.getCell(2).alignment = { horizontal: 'center' };

    const footer2 = readmeSheet.addRow(['', 'Get stronger. Add weight. Repeat. 💪', '', '']);
    readmeSheet.mergeCells(footer2.number, 2, footer2.number, 3);
    footer2.getCell(2).font = { bold: true, size: 12, color: { argb: 'FF2F5496' } };
    footer2.getCell(2).alignment = { horizontal: 'center' };

    // Disclaimer
    readmeSheet.addRow(['']);
    const disclaimerRow = readmeSheet.addRow(['', '⚠️ DISCLAIMER: This is an unofficial community tool, not affiliated with or endorsed by Starting Strength, Inc. or Aasgaard Company.', '', '']);
    readmeSheet.mergeCells(disclaimerRow.number, 2, disclaimerRow.number, 3);
    disclaimerRow.getCell(2).font = { italic: true, size: 9, color: { argb: 'FF808080' } };
    disclaimerRow.getCell(2).alignment = { horizontal: 'center' };

    // Add bottom banner
    readmeSheet.addRow(['']);
    const bottomBanner = readmeSheet.addRow(['', '', '', '']);
    readmeSheet.mergeCells(bottomBanner.number, 1, bottomBanner.number, 4);
    bottomBanner.height = 15;
    bottomBanner.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };

    // ==================== SETTINGS SHEET ====================
    const settingsSheet = workbook.addWorksheet('Settings', {
        properties: { tabColor: { argb: 'FF00B050' } }
    });

    settingsSheet.columns = [
        { key: 'setting', width: 30 },
        { key: 'value', width: 15 },
        { key: 'unit', width: 10 },
        { key: 'notes', width: 50 }
    ];

    const settingsHeader = settingsSheet.addRow(['Setting', 'Value', 'Unit', 'Notes']);
    settingsHeader.font = { bold: true, color: { argb: colors.headerText } };
    settingsHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };

    const settingsData = [
        ['Bar Weight', 45, 'lbs', 'Standard Olympic barbell'],
        ['', '', '', ''],
        ['STARTING WEIGHTS', '', '', 'Enter your tested 5RM or conservative starting weight'],
        ['Squat Starting Weight', 135, 'lbs', ''],
        ['Bench Press Starting Weight', 95, 'lbs', ''],
        ['Deadlift Starting Weight', 155, 'lbs', ''],
        ['Overhead Press Starting Weight', 65, 'lbs', ''],
        ['Power Clean Starting Weight', 95, 'lbs', '(Optional - can add later)'],
        ['', '', '', ''],
        ['WEIGHT INCREMENTS', '', '', 'Amount to add each successful session (min 5 lbs with 2.5 lb plates)'],
        ['Squat Increment', 5, 'lbs', 'Can use 10 lbs early on'],
        ['Bench Press Increment', 5, 'lbs', 'Minimum with 2.5 lb plates per side'],
        ['Deadlift Increment', 10, 'lbs', 'Can use 5 lbs when progress slows'],
        ['Overhead Press Increment', 5, 'lbs', 'Slowest progressing lift - stay patient'],
        ['Power Clean Increment', 5, 'lbs', ''],
        ['', '', '', ''],
        ['AVAILABLE PLATES', '', '', '2.5, 5, 10, 25, 45 lbs'],
        ['', '', '', ''],
        ['PROGRAM SETTINGS', '', '', ''],
        ['Stall Threshold', 3, 'sessions', 'Consecutive stalls before suggesting program change'],
        ['Deload Percentage', 10, '%', 'Amount to reduce weight after stall'],
        ['', '', '', ''],
        ['TARGET REPS', '', '', 'Expected reps per set for each exercise'],
        ['Squat Target Reps', 5, 'reps', '3 sets of 5'],
        ['Bench Press Target Reps', 5, 'reps', '3 sets of 5'],
        ['Deadlift Target Reps', 5, 'reps', '1 set of 5'],
        ['Overhead Press Target Reps', 5, 'reps', '3 sets of 5'],
        ['Power Clean Target Reps', 3, 'reps', '5 sets of 3'],
        ['', '', '', ''],
        ['SESSION TRACKING', '', '', ''],
        ['Total Sessions', '', '', 'Auto-calculated from Workout Log (unique dates)'],
        ['Current Week', '', '', 'CEILING(TotalSessions/3, 1) - assumes 3 sessions/week'],
        ['', '', '', ''],
        ['EXERCISE INTRODUCTION', '', '', 'When exercises become available'],
        ['Chin Ups Introduction Week', 2, 'weeks', 'Chin Ups appear in dropdown after this many weeks'],
        ['Light Squat Percentage', 80, '%', 'Light Squat weight = this % of Squat working weight'],
        ['Light Squat Increment', 5, 'lbs', 'Weight increase per session for Light Squat']
    ];

    settingsData.forEach((row, index) => {
        const newRow = settingsSheet.addRow(row);
        if (row[0] && (row[0].includes('WEIGHTS') || row[0].includes('INCREMENTS') ||
            row[0].includes('SETTINGS') || row[0].includes('PLATES') || row[0].includes('TARGET REPS') ||
            row[0].includes('SESSION TRACKING') || row[0].includes('EXERCISE INTRODUCTION'))) {
            newRow.font = { bold: true };
            newRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
        }
        if (row[1] !== '' && typeof row[1] === 'number') {
            newRow.getCell(2).fill = inputFill;
            newRow.getCell(2).border = thinBorder;
        }
    });

    // Named cells for settings (using cell.name property)
    settingsSheet.getCell('B2').name = 'BarWeight';
    settingsSheet.getCell('B5').name = 'SquatStart';
    settingsSheet.getCell('B6').name = 'BenchStart';
    settingsSheet.getCell('B7').name = 'DeadliftStart';
    settingsSheet.getCell('B8').name = 'PressStart';
    settingsSheet.getCell('B9').name = 'CleanStart';
    settingsSheet.getCell('B12').name = 'SquatInc';
    settingsSheet.getCell('B13').name = 'BenchInc';
    settingsSheet.getCell('B14').name = 'DeadliftInc';
    settingsSheet.getCell('B15').name = 'PressInc';
    settingsSheet.getCell('B16').name = 'CleanInc';

    // Session tracking formulas
    // Total Sessions: Count unique dates in Workout Log
    settingsSheet.getCell('B32').value = { formula: 'SUMPRODUCT(--(FREQUENCY(IF(\'Workout Log\'!A10:A209<>"",\'Workout Log\'!A10:A209),IF(\'Workout Log\'!A10:A209<>"",\'Workout Log\'!A10:A209))>0))' };
    settingsSheet.getCell('B32').name = 'TotalSessions';
    // Current Week: Ceiling of sessions / 3
    settingsSheet.getCell('B33').value = { formula: 'MAX(1,CEILING(B32/3,1))' };
    settingsSheet.getCell('B33').name = 'CurrentWeek';

    // Exercise introduction settings
    settingsSheet.getCell('B36').name = 'ChinUpsIntroWeek';
    settingsSheet.getCell('B37').name = 'LightSquatPct';
    settingsSheet.getCell('B38').name = 'LightSquatInc';

    // ==================== ASSISTANCE EXERCISES SHEET ====================
    const assistanceSheet = workbook.addWorksheet('Assistance Exercises', {
        properties: { tabColor: { argb: 'FF9933FF' } }
    });

    assistanceSheet.columns = [
        { key: 'exercise', width: 25 },
        { key: 'scheme', width: 15 },
        { key: 'notes', width: 40 },
        { key: 'spacer', width: 3 },
        { key: 'dropdown', width: 18 }
    ];

    // Title
    const assistTitle = assistanceSheet.addRow(['ASSISTANCE EXERCISES', '', '', '', 'EXERCISE LIST']);
    assistanceSheet.mergeCells('A1:C1');
    assistTitle.height = 30;
    assistTitle.getCell(1).font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
    assistTitle.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9933FF' } };
    assistTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    // Column E header for dropdown list
    assistTitle.getCell(5).font = { bold: true, size: 11, color: { argb: colors.headerText } };
    assistTitle.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    assistTitle.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    // Instructions
    const assistInstr = assistanceSheet.addRow(['Add or remove exercises below. These appear in the Workout Log dropdown but do NOT affect progression tracking.', '', '', '', 'Squat']);
    assistanceSheet.mergeCells('A2:C2');
    assistInstr.getCell(1).font = { italic: true, size: 10, color: { argb: 'FF666666' } };
    assistInstr.height = 20;
    // E2: Squat (main lift)
    assistInstr.getCell(5).font = { bold: true };

    // Row 3: Spacer + Bench Press
    const row3 = assistanceSheet.addRow(['', '', '', '', 'Bench Press']);
    row3.getCell(5).font = { bold: true };

    // Header row + Deadlift
    const assistHeader = assistanceSheet.addRow(['Exercise Name', 'Suggested Scheme', 'Notes', '', 'Deadlift']);
    [1, 2, 3].forEach(col => {
        assistHeader.getCell(col).font = { bold: true, color: { argb: colors.headerText } };
        assistHeader.getCell(col).fill = headerFill;
    });
    assistHeader.getCell(5).font = { bold: true }; // E4: Deadlift

    // Default assistance exercises with corresponding dropdown entries
    // Column E continues with: E5=OHP, E6=Power Clean, E7=Light Squat (conditional), E8=Chin Ups (conditional), E9+=Assistance
    const defaultAssistance = [
        ['Barbell Curls', '3x10', 'Bicep isolation', '', 'Overhead Press'],           // Row 5, E5=OHP
        ['Back Extension', '3x10-15', 'Lower back and glutes', '', 'Power Clean'],    // Row 6, E6=Power Clean
        ['Skull Crushers', '3x10', 'Tricep isolation - use EZ bar', '', ''],          // Row 7, E7=Light Squat (formula)
        ['Tricep Pushdown', '3x10-12', 'Cable machine', '', ''],                      // Row 8, E8=Chin Ups (formula)
        ['Dips', '3x8-10', 'Can add weight when bodyweight becomes easy', '', ''],    // Row 9, E9=Barbell Curls (formula)
        ['', '', '', '', ''],  // Row 10
        ['', '', '', '', ''],  // Row 11
        ['', '', '', '', ''],  // Row 12
        ['', '', '', '', ''],  // Row 13
        ['', '', '', '', '']   // Row 14
    ];

    defaultAssistance.forEach((row, index) => {
        const newRow = assistanceSheet.addRow(row);
        const rowNum = newRow.number; // Actual row number (5, 6, 7, ...)

        // Highlight input cells for assistance exercises (columns A-C)
        for (let col = 1; col <= 3; col++) {
            newRow.getCell(col).fill = inputFill;
            newRow.getCell(col).border = thinBorder;
        }

        // Style column E (dropdown list)
        if (index < 2) {
            // E5=Overhead Press, E6=Power Clean (already set as static values)
            newRow.getCell(5).font = { bold: true };
        }
    });

    // Set formulas for conditional and assistance exercises in column E
    // E7: Light Squat (conditional - appears when stall threshold reached)
    assistanceSheet.getCell('E7').value = { formula: 'IF(MAX(\'Program Phase\'!B9:B13)>=Settings!$B$21,"Light Squat","")' };
    assistanceSheet.getCell('E7').font = { italic: true, color: { argb: 'FF7030A0' } };

    // E8: Chin Ups (conditional - appears after configured weeks)
    assistanceSheet.getCell('E8').value = { formula: 'IF(Settings!$B$33>=Settings!$B$36,"Chin Ups","")' };
    assistanceSheet.getCell('E8').font = { italic: true, color: { argb: 'FF7030A0' } };

    // E9-E18: Reference assistance exercises from column A (A5-A14)
    for (let i = 9; i <= 18; i++) {
        const assistRow = i - 4; // E9->A5, E10->A6, etc.
        assistanceSheet.getCell(`E${i}`).value = { formula: `IF(A${assistRow}="","",A${assistRow})` };
        assistanceSheet.getCell(`E${i}`).font = { size: 10 };
    }

    // Add note at bottom
    assistanceSheet.addRow(['']);
    const noteRow = assistanceSheet.addRow(['💡 Tip: Add exercises in column A. They will automatically appear in the Workout Log dropdown.']);
    assistanceSheet.mergeCells(`A${noteRow.number}:C${noteRow.number}`);
    noteRow.getCell(1).font = { italic: true, size: 10, color: { argb: 'FF7030A0' } };

    const noteRow2 = assistanceSheet.addRow(['Assistance exercises show "—" for Target Weight and Status (no auto-progression).']);
    assistanceSheet.mergeCells(`A${noteRow2.number}:C${noteRow2.number}`);
    noteRow2.getCell(1).font = { italic: true, size: 10, color: { argb: 'FF7030A0' } };

    assistanceSheet.views = [{ state: 'frozen', ySplit: 4 }];

    // Named cells for stall settings (after assistance sheet creation)
    settingsSheet.getCell('B21').name = 'StallThreshold';
    settingsSheet.getCell('B22').name = 'DeloadPct';

    // ==================== WORKOUT LOG SHEET ====================
    const workoutSheet = workbook.addWorksheet('Workout Log', {
        properties: { tabColor: { argb: 'FFFF6600' } }
    });

    workoutSheet.columns = [
        { key: 'date', width: 12 },
        { key: 'workout', width: 8 },
        { key: 'exercise', width: 16 },
        { key: 'scheme', width: 8 },
        { key: 'targetWeight', width: 11 },
        { key: 'actualWeight', width: 11 },
        { key: 'set1', width: 6 },
        { key: 'set2', width: 6 },
        { key: 'set3', width: 6 },
        { key: 'set4', width: 6 },
        { key: 'set5', width: 6 },
        { key: 'totalReps', width: 8 },
        { key: 'targetReps', width: 8 },
        { key: 'status', width: 8 },
        { key: 'notes', width: 20 },
        // Helper columns for Progress Summary (hidden)
        { key: 'squatWt', width: 10 },
        { key: 'benchWt', width: 10 },
        { key: 'deadliftWt', width: 10 },
        { key: 'ohpWt', width: 10 },
        { key: 'cleanWt', width: 10 }
    ];

    // Add exercise reference legend stacked in first rows
    const setsGuideData = [
        'SETS GUIDE:',
        'Squat: 3 sets x 5 reps (3x5)',
        'Bench Press: 3 sets x 5 reps (3x5)',
        'Deadlift: 1 set x 5 reps (1x5)',
        'Overhead Press: 3 sets x 5 reps (3x5)',
        'Power Clean: 5 sets x 3 reps (5x3)',
        'ASSISTANCE: 3x10 (configure in "Assistance Exercises" sheet — does NOT affect progression)'
    ];

    setsGuideData.forEach((text, index) => {
        const row = workoutSheet.addRow([text]);
        row.font = { bold: index === 0, size: 10, color: { argb: 'FF0070C0' } };
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
        // Merge cells A through O for the legend rows
        workoutSheet.mergeCells(index + 1, 1, index + 1, 15);
    });

    const workoutHeader = workoutSheet.addRow([
        'Date', 'Type', 'Exercise', 'Scheme', 'Target Weight', 'Actual Weight',
        'Set 1', 'Set 2', 'Set 3', 'Set 4', 'Set 5', 'Total Reps', 'Target Reps', 'Status', 'Notes',
        'SquatWt', 'BenchWt', 'DeadliftWt', 'OHPWt', 'CleanWt'
    ]);
    workoutHeader.font = { bold: true, color: { argb: colors.headerText } };
    workoutHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };

    const subHeader = workoutSheet.addRow([
        '', 'A/B', 'Select ↓', 'Auto', 'Auto', 'Enter',
        'Reps', 'Reps', 'Reps', '-', '-', 'Auto', 'Auto', 'Auto', '',
        'Helper', 'Helper', 'Helper', 'Helper', 'Helper'
    ]);
    subHeader.font = { italic: true, size: 9 };
    subHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    // Gray out Set 4/Set 5 in subheader to indicate rarely used
    subHeader.getCell(10).font = { italic: true, size: 9, color: { argb: 'FF999999' } };
    subHeader.getCell(11).font = { italic: true, size: 9, color: { argb: 'FF999999' } };
    // Gray out helper columns
    for (let col = 16; col <= 20; col++) {
        subHeader.getCell(col).font = { italic: true, size: 9, color: { argb: 'FF999999' } };
    }

    // Add 200 workout rows with formulas (data starts at row 10 after 7 legend rows + header + subheader)
    for (let i = 10; i <= 209; i++) {
        const row = workoutSheet.addRow(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);

        // Define main lifts for formula checks (includes Light Squat for tracking)
        const isMainLift = `OR(C${i}="Squat",C${i}="Bench Press",C${i}="Deadlift",C${i}="Overhead Press",C${i}="Power Clean",C${i}="Light Squat")`;

        // Scheme Formula - shows sets x reps for the exercise (3x10 for assistance)
        const schemeFormula = `IF(C${i}="","",IF(OR(C${i}="Squat",C${i}="Bench Press",C${i}="Overhead Press",C${i}="Light Squat"),"3x5",IF(C${i}="Deadlift","1x5",IF(C${i}="Power Clean","5x3","3x10"))))`;
        row.getCell(4).value = { formula: schemeFormula };

        // Target Weight Formula - calculates for main lifts including Light Squat (80% of Squat)
        // Light Squat: 80% of current Squat working weight OR last successful Light Squat + increment
        // Row 10 special case: No previous rows exist, so always return starting weight (avoids invalid C$10:C9 range)
        let targetFormula;
        if (i === 10) {
            // First data row - no previous data to look up, just return starting weight
            targetFormula = `IF(C10="","",IF(NOT(${isMainLift}),"—",IF(C10="Light Squat",0,IF(C10="Squat",Settings!$B$5,IF(C10="Bench Press",Settings!$B$6,IF(C10="Deadlift",Settings!$B$7,IF(C10="Overhead Press",Settings!$B$8,IF(C10="Power Clean",Settings!$B$9,0))))))))`;
        } else {
            // Target weight formula using SUMPRODUCT for Excel 2013+ compatibility (MAXIFS requires Excel 2016+)
            // SUMPRODUCT(MAX(...)) pattern works across all Excel versions
            const startingWeightLookup = `IF(C${i}="Squat",Settings!$B$5,IF(C${i}="Bench Press",Settings!$B$6,IF(C${i}="Deadlift",Settings!$B$7,IF(C${i}="Overhead Press",Settings!$B$8,IF(C${i}="Power Clean",Settings!$B$9,0)))))`;
            const incrementLookup = `IF(C${i}="Squat",Settings!$B$12,IF(C${i}="Bench Press",Settings!$B$13,IF(C${i}="Deadlift",Settings!$B$14,IF(C${i}="Overhead Press",Settings!$B$15,IF(C${i}="Power Clean",Settings!$B$16,0)))))`;

            // Count previous OK entries for this exercise using SUMPRODUCT (Excel 2013 compatible)
            const countOK = `SUMPRODUCT((C$10:C${i-1}=C${i})*(N$10:N${i-1}="OK")*1)`;

            // Get max weight from previous OK entries using SUMPRODUCT(MAX(...))
            const maxOKWeight = `SUMPRODUCT(MAX((C$10:C${i-1}=C${i})*(N$10:N${i-1}="OK")*(F$10:F${i-1})))`;

            // For main lifts: if no previous "OK" entries exist, use starting weight; otherwise use max OK weight + increment
            const mainLiftFormula = `IF(${countOK}=0,${startingWeightLookup},${maxOKWeight}+${incrementLookup})`;

            // Light Squat calculations
            const countLightSquat = `COUNTIF(C$10:C${i-1},"Light Squat")`;
            const countLightSquatOK = `SUMPRODUCT((C$10:C${i-1}="Light Squat")*(N$10:N${i-1}="OK")*1)`;
            const maxSquatOK = `SUMPRODUCT(MAX((C$10:C${i-1}="Squat")*(N$10:N${i-1}="OK")*(F$10:F${i-1})))`;
            const maxLightSquatOK = `SUMPRODUCT(MAX((C$10:C${i-1}="Light Squat")*(N$10:N${i-1}="OK")*(F$10:F${i-1})))`;
            const lightSquatFromSquat = `ROUND(${maxSquatOK}*Settings!$B$37/100/5,0)*5`;

            // Light Squat: 80% of best Squat, or previous Light Squat + increment
            const lightSquatFormula = `IF(${countLightSquat}=0,${lightSquatFromSquat},IF(${countLightSquatOK}=0,${lightSquatFromSquat},${maxLightSquatOK}+Settings!$B$38))`;

            targetFormula = `IF(C${i}="","",IF(NOT(${isMainLift}),"—",IF(C${i}="Light Squat",${lightSquatFormula},${mainLiftFormula})))`;
        }
        row.getCell(5).value = { formula: targetFormula };

        // Total Reps Formula
        row.getCell(12).value = { formula: `IF(C${i}="","",SUM(G${i}:K${i}))` };

        // Target Reps Formula - 30 for assistance (3x10), normal for main lifts (Light Squat = 15)
        const targetRepsFormula = `IF(C${i}="","",IF(NOT(${isMainLift}),30,IF(OR(C${i}="Squat",C${i}="Bench Press",C${i}="Overhead Press",C${i}="Light Squat"),15,IF(C${i}="Deadlift",5,IF(C${i}="Power Clean",15,0)))))`;
        row.getCell(13).value = { formula: targetRepsFormula };

        // Status Formula - OK/STALL for main lifts (including Light Squat), "—" for assistance
        row.getCell(14).value = { formula: `IF(C${i}="","",IF(NOT(${isMainLift}),"—",IF(L${i}>=M${i},"OK","STALL")))` };

        // Helper columns for Progress Summary - extract weight by exercise (P-T = columns 16-20)
        // These return the Actual Weight if the exercise matches, otherwise 0
        row.getCell(16).value = { formula: `IF($C${i}="Squat",$F${i},0)` };  // SquatWt
        row.getCell(17).value = { formula: `IF($C${i}="Bench Press",$F${i},0)` };  // BenchWt
        row.getCell(18).value = { formula: `IF($C${i}="Deadlift",$F${i},0)` };  // DeadliftWt
        row.getCell(19).value = { formula: `IF($C${i}="Overhead Press",$F${i},0)` };  // OHPWt
        row.getCell(20).value = { formula: `IF($C${i}="Power Clean",$F${i},0)` };  // CleanWt

        // Add borders to all cells
        for (let col = 1; col <= 20; col++) {
            row.getCell(col).border = thinBorder;
        }

        // Input cells highlighting (only the cells user needs to fill)
        [1, 2, 3, 6, 7, 8, 9].forEach(col => row.getCell(col).fill = inputFill);
        // Grayed out Set 4/Set 5 (rarely used)
        row.getCell(10).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
        row.getCell(11).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
    }

    // Freeze header rows (9 rows: 7 legend + header + subheader)
    workoutSheet.views = [{ state: 'frozen', ySplit: 9 }];

    // Hide helper columns (P-T = columns 16-20)
    for (let col = 16; col <= 20; col++) {
        workoutSheet.getColumn(col).hidden = true;
    }

    // Data validations
    workoutSheet.dataValidations.add('B10:B209', {
        type: 'list',
        allowBlank: true,
        formulae: ['"A,B"']
    });

    workoutSheet.dataValidations.add('C10:C209', {
        type: 'list',
        allowBlank: true,
        formulae: ['\'Assistance Exercises\'!$E$2:$E$18']  // References exercise list in Assistance Exercises sheet
    });

    // Conditional formatting for STALL
    workoutSheet.addConditionalFormatting({
        ref: 'N10:N209',
        rules: [
            {
                type: 'containsText',
                operator: 'containsText',
                text: 'STALL',
                style: {
                    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF6B6B' } },
                    font: { bold: true, color: { argb: 'FF8B0000' } }
                }
            },
            {
                type: 'containsText',
                operator: 'containsText',
                text: 'OK',
                style: {
                    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF92D050' } },
                    font: { color: { argb: 'FF006400' } }
                }
            }
        ]
    });

    // ==================== WARM-UP CALCULATOR SHEET ====================
    const warmupSheet = workbook.addWorksheet('Warm-Up Calculator', {
        properties: { tabColor: { argb: 'FF7030A0' } }
    });

    warmupSheet.columns = [
        { key: 'label', width: 25 },
        { key: 'value', width: 15 },
        { key: 'spacer', width: 5 },
        { key: 'setNum', width: 12 },
        { key: 'weight', width: 15 },
        { key: 'reps', width: 10 },
        { key: 'rest', width: 15 }
    ];

    const warmupTitle = warmupSheet.addRow(['WARM-UP CALCULATOR']);
    warmupTitle.font = { bold: true, size: 16, color: { argb: colors.headerBg } };
    warmupSheet.mergeCells('A1:G1');

    warmupSheet.addRow(['']);

    const inputHeader = warmupSheet.addRow(['Enter Your Working Weight:', '', '', '', '', '', '']);
    inputHeader.font = { bold: true };

    warmupSheet.addRow(['Working Weight (lbs):', 225, '', '', '', '', '']);
    warmupSheet.getCell('B4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.inputCell } };
    warmupSheet.getCell('B4').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
    };

    warmupSheet.addRow(['Bar Weight (lbs):', { formula: 'Settings!B2' }, '', '', '', '', '']);
    warmupSheet.addRow(['']);

    const protocolHeader = warmupSheet.addRow(['', '', '', 'Set', 'Weight (lbs)', 'Reps', 'Rest']);
    protocolHeader.font = { bold: true, color: { argb: colors.headerText } };
    protocolHeader.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    protocolHeader.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    protocolHeader.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    protocolHeader.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };

    const warmupSets = [
        ['', '', '', '1', { formula: 'Settings!B2' }, 5, 'None'],
        ['', '', '', '2', { formula: 'Settings!B2' }, 5, 'None'],
        ['', '', '', '3', { formula: 'ROUND(B4*0.4/5,0)*5' }, 5, '1 min'],
        ['', '', '', '4', { formula: 'ROUND(B4*0.6/5,0)*5' }, 3, '1 min'],
        ['', '', '', '5', { formula: 'ROUND(B4*0.8/5,0)*5' }, 2, '2 min'],
        ['', '', '', 'WORK', { formula: 'B4' }, '3x5 or 1x5', '3-5 min']
    ];

    warmupSets.forEach((row, index) => {
        const newRow = warmupSheet.addRow(row);
        if (index === warmupSets.length - 1) {
            newRow.font = { bold: true };
            newRow.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.success } };
            newRow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.success } };
            newRow.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.success } };
            newRow.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.success } };
        } else {
            newRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.lightBg } };
        }
    });

    warmupSheet.addRow(['']);
    warmupSheet.addRow(['']);

    const warmupInstructions = [
        ['WARM-UP GUIDELINES:'],
        ['- Empty bar sets: Focus on form and bar path'],
        ['- Increase speed slightly with each warm-up set'],
        ['- Do NOT fatigue yourself - warm-ups prepare, not tire'],
        ['- Rest just long enough to load the bar between warm-up sets'],
        ['- Take full rest (3-5 min) before work sets'],
        [''],
        ['NOTE: Weights are automatically rounded to nearest 5 lbs']
    ];

    warmupInstructions.forEach((row, index) => {
        const newRow = warmupSheet.addRow(row);
        if (index === 0) {
            newRow.font = { bold: true };
        }
    });

    // ==================== BODY WEIGHT LOG SHEET ====================
    const bodyWeightSheet = workbook.addWorksheet('Body Weight Log', {
        properties: { tabColor: { argb: 'FF00B0F0' } }
    });

    bodyWeightSheet.columns = [
        { key: 'date', width: 15 },
        { key: 'weight', width: 15 },
        { key: 'notes', width: 40 },
        { key: 'spacer', width: 5 },
        { key: 'stat', width: 20 },
        { key: 'value', width: 15 }
    ];

    const bwHeader = bodyWeightSheet.addRow(['Date', 'Weight (lbs)', 'Notes', '', 'Statistics', '']);
    bwHeader.font = { bold: true, color: { argb: colors.headerText } };
    bwHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    bwHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    bwHeader.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    bwHeader.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    bwHeader.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };

    const bwData = [
        ['', '', '', '', 'Starting Weight:', { formula: 'IF(COUNTA(B2:B200)>0,INDEX(B:B,MATCH(TRUE,B2:B200<>"",0)+1),"N/A")' }],
        ['', '', '', '', 'Current Weight:', { formula: 'IF(COUNTA(B2:B200)>0,LOOKUP(2,1/(B2:B200<>""),B2:B200),"N/A")' }],
        ['', '', '', '', 'Total Change:', { formula: 'IF(AND(F2<>"N/A",F3<>"N/A"),F3-F2,"N/A")' }],
        ['', '', '', '', 'Average Weight:', { formula: 'IF(COUNTA(B2:B200)>0,ROUND(AVERAGE(B2:B200),1),"N/A")' }],
        ['', '', '', '', 'Weigh-ins Count:', { formula: 'COUNTA(B2:B200)' }]
    ];

    bwData.forEach(row => {
        const newRow = bodyWeightSheet.addRow(row);
        newRow.getCell(5).font = { bold: true };
    });

    for (let i = 0; i < 100; i++) {
        const newRow = bodyWeightSheet.addRow(['', '', '', '', '', '']);
        for (let col = 1; col <= 3; col++) {
            newRow.getCell(col).fill = inputFill;
            newRow.getCell(col).border = thinBorder;
        }
    }

    bodyWeightSheet.views = [{ state: 'frozen', ySplit: 1 }];

    // ==================== PROGRESS SUMMARY SHEET ====================
    const summarySheet = workbook.addWorksheet('Progress Summary', {
        properties: { tabColor: { argb: 'FFED7D31' } }
    });

    summarySheet.columns = [
        { key: 'week', width: 12 },
        { key: 'date', width: 12 },
        { key: 'squat', width: 12 },
        { key: 'bench', width: 12 },
        { key: 'deadlift', width: 12 },
        { key: 'ohp', width: 12 },
        { key: 'clean', width: 12 },
        { key: 'bodyweight', width: 12 }
    ];

    const summaryTitle = summarySheet.addRow(['PROGRESS SUMMARY - Auto-Updated from Workout Log']);
    summaryTitle.font = { bold: true, size: 14, color: { argb: colors.headerBg } };
    summarySheet.mergeCells('A1:H1');

    summarySheet.addRow(['']);

    const summaryHeader = summarySheet.addRow(['Week', 'End Date', 'Squat', 'Bench', 'Deadlift', 'OHP', 'P.Clean', 'Body Wt']);
    summaryHeader.font = { bold: true, color: { argb: colors.headerText } };
    summaryHeader.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    });

    // Add DMAX criteria area (hidden column I)
    // Column I contains "Date" header + criteria values for each week
    summarySheet.getColumn(9).width = 12;
    summarySheet.getColumn(9).hidden = true;
    summarySheet.getCell('I3').value = 'Date';  // Header must match Workout Log column header exactly

    // Add 52 weeks of formulas (1 year of tracking)
    // Using DMAX with helper columns - reliable across all Excel versions
    // DMAX finds the maximum value in a database column that matches criteria
    for (let week = 1; week <= 52; week++) {
        const rowNum = week + 3; // Row 4 = Week 1, Row 5 = Week 2, etc.

        // Set cells directly to avoid row ordering issues
        summarySheet.getCell(`A${rowNum}`).value = week;
        summarySheet.getCell(`B${rowNum}`).value = ''; // End Date - user fills this in
        summarySheet.getCell(`C${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Workout Log'!$A$8:$T$209,"SquatWt",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`D${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Workout Log'!$A$8:$T$209,"BenchWt",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`E${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Workout Log'!$A$8:$T$209,"DeadliftWt",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`F${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Workout Log'!$A$8:$T$209,"OHPWt",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`G${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Workout Log'!$A$8:$T$209,"CleanWt",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`H${rowNum}`).value = { formula: `IF(B${rowNum}="","",IFERROR(DMAX('Body Weight Log'!$A$1:$B$200,"Weight (lbs)",$I$3:$I${rowNum}),0))` };
        summarySheet.getCell(`I${rowNum}`).value = { formula: `IF(B${rowNum}="","","<="&B${rowNum})` };

        // Style the row (columns A-H only, skip hidden column I)
        for (let col = 1; col <= 8; col++) {
            const cell = summarySheet.getCell(rowNum, col);
            cell.border = thinBorder;
            if (col === 2) cell.fill = inputFill;
        }
    }

    summarySheet.addRow(['']);
    const summaryNote = summarySheet.addRow(['NOTE: Values show your BEST successful lift for each exercise. Updates automatically as you log workouts.']);
    summaryNote.font = { italic: true, color: { argb: 'FF666666' } };

    // ==================== PROGRESS CHART SHEET ====================
    const chartSheet = workbook.addWorksheet('Progress Chart', {
        properties: { tabColor: { argb: 'FF00B050' } }
    });

    // Add chart title
    const chartTitle = chartSheet.addRow(['STRENGTH PROGRESSION CHART']);
    chartTitle.font = { bold: true, size: 16, color: { argb: colors.headerBg } };
    chartSheet.mergeCells('A1:J1');

    chartSheet.addRow(['']);
    chartSheet.addRow(['This chart updates automatically when you add workout data.']);
    chartSheet.addRow(['To view the chart: Select data in Progress Summary (A3:H55) → Insert → Chart → Line Chart']);
    chartSheet.addRow(['']);

    // Create a simple data reference for manual chart creation
    const chartInstructions = [
        ['MANUAL CHART CREATION STEPS:'],
        ['1. Go to Progress Summary sheet'],
        ['2. Select cells A3 through H55 (headers + data)'],
        ['3. Click Insert → Chart (or Insert → Recommended Charts)'],
        ['4. Choose "Line with Markers"'],
        ['5. The chart will auto-update as you add workout data'],
        [''],
        ['CURRENT PR (Personal Records):'],
    ];

    chartInstructions.forEach((row, index) => {
        const newRow = chartSheet.addRow(row);
        if (index === 0 || index === 7) {
            newRow.font = { bold: true };
        }
    });

    // Add PR summary - Using simple MAX on helper columns (P-T in Workout Log)
    // Helper columns contain the weight for each exercise (or 0), so MAX finds the highest
    const prData = [
        ['Squat PR:', { formula: 'IF(MAX(\'Workout Log\'!$P$10:$P$209)=0,"No data",MAX(\'Workout Log\'!$P$10:$P$209))' }, 'lbs'],
        ['Bench PR:', { formula: 'IF(MAX(\'Workout Log\'!$Q$10:$Q$209)=0,"No data",MAX(\'Workout Log\'!$Q$10:$Q$209))' }, 'lbs'],
        ['Deadlift PR:', { formula: 'IF(MAX(\'Workout Log\'!$R$10:$R$209)=0,"No data",MAX(\'Workout Log\'!$R$10:$R$209))' }, 'lbs'],
        ['OHP PR:', { formula: 'IF(MAX(\'Workout Log\'!$S$10:$S$209)=0,"No data",MAX(\'Workout Log\'!$S$10:$S$209))' }, 'lbs'],
        ['Power Clean PR:', { formula: 'IF(MAX(\'Workout Log\'!$T$10:$T$209)=0,"No data",MAX(\'Workout Log\'!$T$10:$T$209))' }, 'lbs'],
    ];

    prData.forEach(row => {
        const newRow = chartSheet.addRow(row);
        newRow.getCell(1).font = { bold: true };
        newRow.getCell(2).font = { bold: true, size: 14, color: { argb: colors.headerBg } };
    });

    // ==================== PROGRAM PHASE SHEET ====================
    const phaseSheet = workbook.addWorksheet('Program Phase', {
        properties: { tabColor: { argb: 'FFC00000' } }
    });

    phaseSheet.columns = [
        { key: 'label', width: 25 },
        { key: 'value', width: 18 },
        { key: 'spacer', width: 3 },
        { key: 'info', width: 55 }
    ];

    const phaseTitle = phaseSheet.addRow(['PROGRAM PHASE TRACKER - Auto-Detection']);
    phaseTitle.font = { bold: true, size: 16, color: { argb: colors.headerBg } };
    phaseSheet.mergeCells('A1:D1');

    phaseSheet.addRow(['']);

    // Current phase with formula
    const phaseRow = phaseSheet.addRow(['Current Phase:', { formula: 'IF(MAX(B9:B13)>=Settings!$B$21,"INTERMEDIATE","NOVICE")' }, '',
        { formula: 'IF(B3="NOVICE","Linear progression - adding weight every session","Weekly progression - recovery takes longer")' }]);
    phaseRow.getCell(1).font = { bold: true, size: 14 };
    phaseRow.getCell(2).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };

    phaseSheet.addRow(['']);

    // Overall status
    const statusRow = phaseSheet.addRow(['Overall Status:',
        { formula: 'IF(MAX(B9:B13)>=Settings!$B$21,"⚠️ TRANSITION NEEDED",IF(MAX(B9:B13)>=2,"⚡ WARNING","✓ ON TRACK"))' }, '', '']);
    statusRow.getCell(1).font = { bold: true };
    statusRow.getCell(2).font = { bold: true, size: 12 };

    phaseSheet.addRow(['']);

    // Stall tracking header
    const stallHeader = phaseSheet.addRow(['STALL TRACKING (Auto-Calculated)', '', '', '']);
    stallHeader.font = { bold: true, size: 12 };
    stallHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };

    const stallSubHeader = phaseSheet.addRow(['Exercise', 'Consecutive Stalls', '', 'Status']);
    stallSubHeader.font = { bold: true };

    // Stall count formulas for each exercise
    const exercises = ['Squat', 'Bench Press', 'Deadlift', 'Overhead Press', 'Power Clean'];

    exercises.forEach((exerciseName, index) => {
        const rowNum = 9 + index;
        const stallFormula = `SUMPRODUCT(('Workout Log'!C$10:C$209="${exerciseName}")*('Workout Log'!N$10:N$209="STALL")*1)`;
        const statusFormula = `IF(B${rowNum}>=Settings!$B$21,"🔴 TRANSITION",IF(B${rowNum}>=2,"🟡 WARNING",IF(B${rowNum}>=1,"🟠 MONITOR","🟢 ON TRACK")))`;

        const row = phaseSheet.addRow([
            exerciseName,
            { formula: stallFormula },
            '',
            { formula: statusFormula }
        ]);

        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.lightBg } };
        row.getCell(2).border = thinBorder;
        row.getCell(2).alignment = { horizontal: 'center' };
    });

    phaseSheet.addRow(['']);
    phaseSheet.addRow(['']);

    // Phase definitions
    const phaseDefHeader = phaseSheet.addRow(['PHASE DEFINITIONS', '', '', '']);
    phaseDefHeader.font = { bold: true, size: 12 };
    phaseDefHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };

    const phaseDefinitions = [
        [''],
        ['NOVICE', 'Linear Progression', '', 'Add weight every session. Recover between sessions (48-72 hrs).'],
        ['', '', '', 'Duration: 3-9 months depending on age, diet, sleep.'],
        [''],
        ['INTERMEDIATE', 'Weekly Progression', '', 'Add weight every week. Recovery takes longer than one session.'],
        ['', '', '', 'Programs: Texas Method, HLM, or similar periodization.'],
        [''],
        ['ADVANCED', 'Monthly Progression', '', 'Complex periodization. Most people never truly reach this.']
    ];

    phaseDefinitions.forEach(row => {
        const newRow = phaseSheet.addRow(row);
        if (row[0] === 'NOVICE' || row[0] === 'INTERMEDIATE' || row[0] === 'ADVANCED') {
            newRow.font = { bold: true };
        }
    });

    phaseSheet.addRow(['']);

    // Action guide
    const actionHeader = phaseSheet.addRow(['WHAT TO DO WHEN YOU STALL', '', '', '']);
    actionHeader.font = { bold: true, size: 12 };
    actionHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };

    const actionSteps = [
        [''],
        ['1 Stall:', '', '', 'Try again next session. Ensure sleep, food, and recovery are good.'],
        ['2 Stalls:', '', '', 'Deload 10% and work back up. Check form.'],
        ['3+ Stalls:', '', '', 'Time to consider Intermediate programming for this lift.'],
        [''],
        ['IMPORTANT:', '', '', 'Exhaust ALL novice gains first. Most people move too early.']
    ];

    actionSteps.forEach(row => {
        const newRow = phaseSheet.addRow(row);
        if (row[0] && row[0].includes(':')) {
            newRow.font = { bold: true };
        }
        if (row[0] === 'IMPORTANT:') {
            newRow.getCell(1).font = { bold: true, color: { argb: 'FFFF0000' } };
        }
    });

    // Add conditional formatting to phase cell
    phaseSheet.addConditionalFormatting({
        ref: 'B3',
        rules: [
            {
                type: 'containsText',
                operator: 'containsText',
                text: 'NOVICE',
                style: {
                    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF92D050' } }
                }
            },
            {
                type: 'containsText',
                operator: 'containsText',
                text: 'INTERMEDIATE',
                style: {
                    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC000' } }
                }
            }
        ]
    });

    // Save the workbook
    await workbook.xlsx.writeFile('Barbell_Strength_Tracker.xlsx');
    console.log('Barbell Strength Tracker v2.5 created successfully!');
    console.log('File: Barbell_Strength_Tracker.xlsx');
    console.log('');
    console.log('New in v2.5:');
    console.log('- Smart Exercise Introduction: Light Squat and Chin Ups appear based on training progress');
    console.log('- Light Squat: Appears when any lift reaches stall threshold (intermediate transition)');
    console.log('- Chin Ups: Appears after configurable number of weeks (default: 2 weeks)');
    console.log('- Session/Week tracking auto-calculated from Workout Log');
    console.log('- Configurable introduction timing in Settings');
    console.log('');
    console.log('Core Features:');
    console.log('- Auto-calculated target weights based on last successful lift');
    console.log('- Auto-detection of stalls (STALL/OK status)');
    console.log('- Program Phase auto-detects when to transition');
    console.log('- Conditional formatting for visual feedback');
}

createStartingStrengthTracker().catch(console.error);
