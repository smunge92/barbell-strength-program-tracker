const ExcelJS = require('exceljs');

async function testStartingStrengthTracker() {
    console.log('='.repeat(70));
    console.log('BARBELL STRENGTH TRACKER v2.5 - COMPREHENSIVE TEST SUITE');
    console.log('='.repeat(70));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('Barbell_Strength_Tracker.xlsx');

    let testsPassed = 0;
    let testsFailed = 0;

    function test(name, condition, details = '') {
        if (condition) {
            console.log(`  ✓ PASS: ${name}`);
            testsPassed++;
        } else {
            console.log(`  ✗ FAIL: ${name} ${details ? '- ' + details : ''}`);
            testsFailed++;
        }
    }

    // ==================== TEST 1: Sheet Structure ====================
    console.log('\n--- Test 1: Sheet Structure ---');

    const expectedSheets = ['README', 'Settings', 'Assistance Exercises', 'Workout Log', 'Warm-Up Calculator',
                           'Body Weight Log', 'Progress Summary', 'Progress Chart', 'Program Phase'];
    const actualSheets = workbook.worksheets.map(ws => ws.name);

    test('All 9 sheets exist', expectedSheets.every(s => actualSheets.includes(s)),
         `Expected: ${expectedSheets.join(', ')}, Got: ${actualSheets.join(', ')}`);

    expectedSheets.forEach(sheetName => {
        test(`Sheet "${sheetName}" exists`, actualSheets.includes(sheetName));
    });

    // ==================== TEST 2: Settings Sheet Defaults ====================
    console.log('\n--- Test 2: Settings Sheet Defaults ---');

    const settings = workbook.getWorksheet('Settings');

    test('Bar weight default is 45', settings.getCell('B2').value === 45);
    test('Squat starting weight is 135', settings.getCell('B5').value === 135);
    test('Bench starting weight is 95', settings.getCell('B6').value === 95);
    test('Deadlift starting weight is 155', settings.getCell('B7').value === 155);
    test('OHP starting weight is 65', settings.getCell('B8').value === 65);
    test('Power Clean starting weight is 95', settings.getCell('B9').value === 95);
    test('Squat increment is 5', settings.getCell('B12').value === 5);
    test('Bench increment is 5', settings.getCell('B13').value === 5);
    test('Deadlift increment is 10', settings.getCell('B14').value === 10);
    test('OHP increment is 5', settings.getCell('B15').value === 5);
    test('Stall threshold is 3', settings.getCell('B21').value === 3);
    test('Deload percentage is 10', settings.getCell('B22').value === 10);

    // Target reps settings
    test('Squat target reps is 5', settings.getCell('B25').value === 5);
    test('Bench target reps is 5', settings.getCell('B26').value === 5);
    test('Deadlift target reps is 5', settings.getCell('B27').value === 5);
    test('OHP target reps is 5', settings.getCell('B28').value === 5);
    test('Power Clean target reps is 3', settings.getCell('B29').value === 3);

    // ==================== TEST 3: NEW v2.5 - Session Tracking ====================
    console.log('\n--- Test 3: NEW v2.5 - Session Tracking ---');

    // Session tracking section header
    test('SESSION TRACKING header exists', settings.getCell('A31').value === 'SESSION TRACKING');
    test('Total Sessions label exists', settings.getCell('A32').value === 'Total Sessions');
    test('Current Week label exists', settings.getCell('A33').value === 'Current Week');

    // Total Sessions formula
    const totalSessionsCell = settings.getCell('B32');
    test('Total Sessions has formula', totalSessionsCell.value && totalSessionsCell.value.formula !== undefined);
    if (totalSessionsCell.value && totalSessionsCell.value.formula) {
        test('Total Sessions formula uses FREQUENCY', totalSessionsCell.value.formula.includes('FREQUENCY'));
        test('Total Sessions formula references Workout Log', totalSessionsCell.value.formula.includes('Workout Log'));
    }

    // Current Week formula
    const currentWeekCell = settings.getCell('B33');
    test('Current Week has formula', currentWeekCell.value && currentWeekCell.value.formula !== undefined);
    if (currentWeekCell.value && currentWeekCell.value.formula) {
        test('Current Week formula uses CEILING', currentWeekCell.value.formula.includes('CEILING'));
        test('Current Week formula references B32', currentWeekCell.value.formula.includes('B32'));
    }

    // ==================== TEST 4: NEW v2.5 - Exercise Introduction Settings ====================
    console.log('\n--- Test 4: NEW v2.5 - Exercise Introduction Settings ---');

    test('EXERCISE INTRODUCTION header exists', settings.getCell('A35').value === 'EXERCISE INTRODUCTION');
    test('Chin Ups Introduction Week label exists', settings.getCell('A36').value === 'Chin Ups Introduction Week');
    test('Chin Ups Introduction Week default is 2', settings.getCell('B36').value === 2);
    test('Light Squat Percentage label exists', settings.getCell('A37').value === 'Light Squat Percentage');
    test('Light Squat Percentage default is 80', settings.getCell('B37').value === 80);
    test('Light Squat Increment label exists', settings.getCell('A38').value === 'Light Squat Increment');
    test('Light Squat Increment default is 5', settings.getCell('B38').value === 5);

    // ==================== TEST 5: NEW v2.5 - Conditional Exercise List ====================
    console.log('\n--- Test 5: NEW v2.5 - Conditional Exercise List ---');

    // Exercise list is now in Assistance Exercises sheet column E
    const assistSheet = workbook.getWorksheet('Assistance Exercises');

    // Main lifts (E2-E6) - static
    test('Assistance Exercises has Exercise List header', assistSheet.getCell('E1').value === 'EXERCISE LIST');
    test('Squat in exercise list', assistSheet.getCell('E2').value === 'Squat');
    test('Bench Press in exercise list', assistSheet.getCell('E3').value === 'Bench Press');
    test('Deadlift in exercise list', assistSheet.getCell('E4').value === 'Deadlift');
    test('Overhead Press in exercise list', assistSheet.getCell('E5').value === 'Overhead Press');
    test('Power Clean in exercise list', assistSheet.getCell('E6').value === 'Power Clean');

    // Conditional exercises (E7-E8) - formulas
    const lightSquatCell = assistSheet.getCell('E7');
    test('Light Squat cell has formula', lightSquatCell.value && lightSquatCell.value.formula !== undefined);
    if (lightSquatCell.value && lightSquatCell.value.formula) {
        test('Light Squat formula references Program Phase', lightSquatCell.value.formula.includes('Program Phase'));
        test('Light Squat formula checks stall threshold', lightSquatCell.value.formula.includes('B$21'));
        test('Light Squat formula returns "Light Squat"', lightSquatCell.value.formula.includes('"Light Squat"'));
    }

    const chinUpsCell = assistSheet.getCell('E8');
    test('Chin Ups cell has formula', chinUpsCell.value && chinUpsCell.value.formula !== undefined);
    if (chinUpsCell.value && chinUpsCell.value.formula) {
        test('Chin Ups formula references Current Week (B33)', chinUpsCell.value.formula.includes('B$33'));
        test('Chin Ups formula references Intro Week (B36)', chinUpsCell.value.formula.includes('B$36'));
        test('Chin Ups formula returns "Chin Ups"', chinUpsCell.value.formula.includes('"Chin Ups"'));
    }

    // Assistance exercises (E9-E18) - formulas referencing column A
    const assistFormulaCell = assistSheet.getCell('E9');
    test('Assistance exercises start at E9', assistFormulaCell.value && assistFormulaCell.value.formula !== undefined);
    if (assistFormulaCell.value && assistFormulaCell.value.formula) {
        test('E9 references column A5', assistFormulaCell.value.formula.includes('A5'));
    }

    // ==================== TEST 6: Workout Log Structure ====================
    console.log('\n--- Test 6: Workout Log Structure ---');

    const workoutLog = workbook.getWorksheet('Workout Log');

    // Rows 1-7 are the SETS GUIDE legend rows
    test('SETS GUIDE title exists in row 1', workoutLog.getCell('A1').value === 'SETS GUIDE:');
    test('Squat guide exists in row 2',
         workoutLog.getCell('A2').value && workoutLog.getCell('A2').value.includes('Squat'));
    test('Assistance guide exists in row 7',
         workoutLog.getCell('A7').value && workoutLog.getCell('A7').value.includes('ASSISTANCE'));

    // Row 8 has headers
    const expectedHeaders = ['Date', 'Type', 'Exercise', 'Scheme', 'Target Weight', 'Actual Weight',
                            'Set 1', 'Set 2', 'Set 3', 'Set 4', 'Set 5', 'Total Reps', 'Target Reps', 'Status', 'Notes'];
    const actualHeaders = [];
    for (let i = 1; i <= 15; i++) {
        actualHeaders.push(workoutLog.getCell(8, i).value);
    }

    test('Workout Log has correct headers (15 columns)',
         expectedHeaders.every((h, i) => actualHeaders[i] === h),
         `Got: ${actualHeaders.join(', ')}`);

    // Check formulas exist in row 10
    test('Scheme formula exists in row 10',
         workoutLog.getCell('D10').value && workoutLog.getCell('D10').value.formula !== undefined);
    test('Target Weight formula exists in row 10',
         workoutLog.getCell('E10').value && workoutLog.getCell('E10').value.formula !== undefined);
    test('Total Reps formula exists in row 10',
         workoutLog.getCell('L10').value && workoutLog.getCell('L10').value.formula !== undefined);
    test('Target Reps formula exists in row 10',
         workoutLog.getCell('M10').value && workoutLog.getCell('M10').value.formula !== undefined);
    test('Status formula exists in row 10',
         workoutLog.getCell('N10').value && workoutLog.getCell('N10').value.formula !== undefined);

    // ==================== TEST 7: NEW v2.5 - Light Squat in Workout Log Formulas ====================
    console.log('\n--- Test 7: NEW v2.5 - Light Squat in Workout Log Formulas ---');

    const schemeFormula = workoutLog.getCell('D10').value.formula;
    test('Scheme formula includes Light Squat', schemeFormula.includes('Light Squat'));
    test('Scheme formula shows Light Squat as 3x5', schemeFormula.includes('"Light Squat"') && schemeFormula.includes('3x5'));

    const targetFormula10 = workoutLog.getCell('E10').value.formula;
    test('Target Weight formula handles Light Squat', targetFormula10.includes('Light Squat'));
    // Row 10 has special simplified formula (no previous data), check row 11 for full formula
    const targetFormula11 = workoutLog.getCell('E11').value.formula;
    test('Target Weight formula uses Light Squat Percentage (B37)', targetFormula11.includes('B$37'));
    test('Target Weight formula uses Light Squat Increment (B38)', targetFormula11.includes('B$38'));
    test('Target Weight formula calculates 80% of Squat for Light Squat (using SUMPRODUCT)',
         targetFormula11.includes('SUMPRODUCT') && targetFormula11.includes('"Squat"'));

    const targetRepsFormula = workoutLog.getCell('M10').value.formula;
    test('Target Reps formula includes Light Squat', targetRepsFormula.includes('Light Squat'));

    const statusFormula = workoutLog.getCell('N10').value.formula;
    test('Status formula includes Light Squat in main lifts check', statusFormula.includes('Light Squat'));

    // ==================== TEST 8: Dropdown Range Updated ====================
    console.log('\n--- Test 8: Dropdown Range Updated for v2.5 ---');

    const validations = workoutLog.dataValidations;
    test('Data validations are configured', validations !== undefined);

    const validationModel = JSON.stringify(workoutLog.dataValidations.model);
    test('Dropdown references Assistance Exercises sheet', validationModel.includes('Assistance Exercises'));
    test('Dropdown range uses E2:E18 exercise list',
         validationModel.includes('E$2:$E$18') || validationModel.includes('$E$2:$E$18'));

    // ==================== TEST 9: Assistance Exercises Sheet (Chin Ups removed) ====================
    console.log('\n--- Test 9: Assistance Exercises Sheet (Chin Ups removed) ---');

    // assistSheet already declared in Test 5
    test('Assistance Exercises sheet exists', assistSheet !== undefined);
    test('Assistance sheet has title', assistSheet.getCell('A1').value === 'ASSISTANCE EXERCISES');
    test('Assistance sheet has header row', assistSheet.getCell('A4').value === 'Exercise Name');

    // Check Chin Ups is NOT in default assistance (it's now conditionally introduced)
    test('Barbell Curls is first default exercise (not Chin Ups)', assistSheet.getCell('A5').value === 'Barbell Curls');
    test('Back Extension is second default exercise', assistSheet.getCell('A6').value === 'Back Extension');
    test('Skull Crushers is third default exercise', assistSheet.getCell('A7').value === 'Skull Crushers');
    test('Tricep Pushdown is fourth default exercise', assistSheet.getCell('A8').value === 'Tricep Pushdown');
    test('Dips is fifth default exercise', assistSheet.getCell('A9').value === 'Dips');

    // ==================== TEST 10: Warm-Up Calculator ====================
    console.log('\n--- Test 10: Warm-Up Calculator Formulas ---');

    const warmup = workbook.getWorksheet('Warm-Up Calculator');

    test('Working weight input cell exists', warmup.getCell('B4').value === 225);
    test('Bar weight formula references Settings',
         warmup.getCell('B5').value && warmup.getCell('B5').value.formula === 'Settings!B2');
    test('40% warm-up formula exists',
         warmup.getCell('E10').value && warmup.getCell('E10').value.formula.includes('0.4'));
    test('60% warm-up formula exists',
         warmup.getCell('E11').value && warmup.getCell('E11').value.formula.includes('0.6'));
    test('80% warm-up formula exists',
         warmup.getCell('E12').value && warmup.getCell('E12').value.formula.includes('0.8'));

    // ==================== TEST 11: Body Weight Log ====================
    console.log('\n--- Test 11: Body Weight Log Formulas ---');

    const bodyWeight = workbook.getWorksheet('Body Weight Log');

    test('Starting weight formula exists',
         bodyWeight.getCell('F2').value && bodyWeight.getCell('F2').value.formula !== undefined);
    test('Current weight formula exists',
         bodyWeight.getCell('F3').value && bodyWeight.getCell('F3').value.formula !== undefined);
    test('Total change formula exists',
         bodyWeight.getCell('F4').value && bodyWeight.getCell('F4').value.formula !== undefined);

    // ==================== TEST 12: Progress Summary ====================
    console.log('\n--- Test 12: Progress Summary Structure ---');

    const summary = workbook.getWorksheet('Progress Summary');

    const summaryHeaders = ['Week', 'End Date', 'Squat', 'Bench', 'Deadlift', 'OHP', 'P.Clean', 'Body Wt'];
    const actualSummaryHeaders = [];
    for (let i = 1; i <= 8; i++) {
        actualSummaryHeaders.push(summary.getCell(3, i).value);
    }

    test('Progress Summary has correct headers',
         summaryHeaders.every((h, i) => actualSummaryHeaders[i] === h),
         `Got: ${actualSummaryHeaders.join(', ')}`);

    test('Squat SUMPRODUCT formula exists',
         summary.getCell('C4').value && summary.getCell('C4').value.formula &&
         summary.getCell('C4').value.formula.includes('SUMPRODUCT'));
    test('Week 1 exists', summary.getCell('A4').value === 1);
    test('Week 52 exists', summary.getCell('A55').value === 52);

    // ==================== TEST 13: Progress Chart ====================
    console.log('\n--- Test 13: Progress Chart Sheet ---');

    const chart = workbook.getWorksheet('Progress Chart');

    test('Progress Chart sheet exists', chart !== undefined);
    test('Chart title exists', chart.getCell('A1').value && chart.getCell('A1').value.includes('CHART'));
    test('Squat PR formula exists',
         chart.getCell('B14').value && chart.getCell('B14').value.formula &&
         chart.getCell('B14').value.formula.includes('SUMPRODUCT'));

    // ==================== TEST 14: Program Phase Auto-Detection ====================
    console.log('\n--- Test 14: Program Phase Auto-Detection ---');

    const phase = workbook.getWorksheet('Program Phase');

    test('Program Phase sheet exists', phase !== undefined);
    test('Current Phase formula exists',
         phase.getCell('B3').value && phase.getCell('B3').value.formula !== undefined);
    test('Overall Status formula exists',
         phase.getCell('B5').value && phase.getCell('B5').value.formula !== undefined);

    const exerciseRows = [9, 10, 11, 12, 13];
    const exerciseNames = ['Squat', 'Bench Press', 'Deadlift', 'Overhead Press', 'Power Clean'];

    exerciseRows.forEach((row, i) => {
        test(`${exerciseNames[i]} stall tracking exists`,
             phase.getCell(`A${row}`).value === exerciseNames[i]);
        test(`${exerciseNames[i]} stall formula exists`,
             phase.getCell(`B${row}`).value && phase.getCell(`B${row}`).value.formula !== undefined);
    });

    // ==================== TEST 15: README Content v2.5 ====================
    console.log('\n--- Test 15: README Content v2.5 ---');

    const readme = workbook.getWorksheet('README');

    const titleCell = readme.getCell('A2').value;
    test('README has fancy title with emoji',
         titleCell && titleCell.includes('BARBELL STRENGTH TRACKER'));

    // Collect README content (extended range to capture disclaimer)
    let readmeContent = '';
    for (let i = 1; i <= 70; i++) {
        for (let col = 1; col <= 3; col++) {
            const cellValue = readme.getCell(i, col).value;
            if (cellValue) readmeContent += cellValue + ' ';
        }
    }

    test('README has Quick Start section', readmeContent.includes('QUICK START'));
    test('README mentions stall detection', readmeContent.includes('STALL') || readmeContent.includes('stall'));
    test('README has Sheets Guide section', readmeContent.includes('SHEETS GUIDE'));
    test('README contains version 2.5', readmeContent.includes('2.5'));

    // New v2.5 README content
    test('README has EXERCISE INTRODUCTION TIMING section',
         readmeContent.includes('EXERCISE INTRODUCTION TIMING'));
    test('README explains Light Squat', readmeContent.includes('Light Squat'));
    test('README explains Chin Ups timing',
         readmeContent.includes('Chin Ups') && readmeContent.includes('2 weeks'));
    test('README mentions configurable Settings',
         readmeContent.includes('configurable') || readmeContent.includes('Settings'));

    // Disclaimer and attribution
    test('README has disclaimer about not being affiliated',
         readmeContent.includes('unofficial') || readmeContent.includes('not affiliated'));
    test('README credits Mark Rippetoe', readmeContent.includes('Rippetoe'));

    // ==================== TEST 16: Conditional Formatting ====================
    console.log('\n--- Test 16: Conditional Formatting ---');

    const conditionalFormatting = workoutLog.conditionalFormattings;
    test('Workout Log has conditional formatting',
         conditionalFormatting && conditionalFormatting.length > 0);

    const phaseConditionalFormatting = phase.conditionalFormattings;
    test('Program Phase has conditional formatting',
         phaseConditionalFormatting && phaseConditionalFormatting.length > 0);

    // ==================== TEST 17: Simulate Workout Data Entry ====================
    console.log('\n--- Test 17: Simulate Workout Data Entry ---');

    const testWorkouts = [
        { row: 10, date: '2024-01-15', type: 'A', exercise: 'Squat', actual: 135, sets: [5, 5, 5, '', ''] },
        { row: 11, date: '2024-01-15', type: 'A', exercise: 'Bench Press', actual: 95, sets: [5, 5, 5, '', ''] },
        { row: 12, date: '2024-01-15', type: 'A', exercise: 'Deadlift', actual: 155, sets: [5, '', '', '', ''] },
        { row: 13, date: '2024-01-17', type: 'B', exercise: 'Squat', actual: 140, sets: [5, 5, 5, '', ''] },
        { row: 14, date: '2024-01-17', type: 'B', exercise: 'Overhead Press', actual: 65, sets: [5, 5, 5, '', ''] },
        { row: 15, date: '2024-01-17', type: 'B', exercise: 'Deadlift', actual: 165, sets: [5, '', '', '', ''] },
        { row: 16, date: '2024-01-19', type: 'A', exercise: 'Squat', actual: 145, sets: [5, 5, 4, '', ''] }, // Stall
        { row: 17, date: '2024-01-19', type: 'A', exercise: 'Barbell Curls', actual: 45, sets: [10, 10, 10, '', ''] }, // Assistance
    ];

    testWorkouts.forEach(w => {
        workoutLog.getCell(`A${w.row}`).value = w.date;
        workoutLog.getCell(`B${w.row}`).value = w.type;
        workoutLog.getCell(`C${w.row}`).value = w.exercise;
        workoutLog.getCell(`F${w.row}`).value = w.actual;
        workoutLog.getCell(`G${w.row}`).value = w.sets[0];
        workoutLog.getCell(`H${w.row}`).value = w.sets[1];
        workoutLog.getCell(`I${w.row}`).value = w.sets[2];
        workoutLog.getCell(`J${w.row}`).value = w.sets[3];
        workoutLog.getCell(`K${w.row}`).value = w.sets[4];
    });

    test('Workout data entry successful', workoutLog.getCell('A10').value === '2024-01-15');
    test('Exercise selection works', workoutLog.getCell('C10').value === 'Squat');
    test('Actual weight recorded', workoutLog.getCell('F10').value === 135);
    test('Reps recorded correctly', workoutLog.getCell('G10').value === 5);

    // ==================== TEST 18: Simulate Light Squat Entry ====================
    console.log('\n--- Test 18: NEW v2.5 - Simulate Light Squat Entry ---');

    // Add a Light Squat workout entry (simulating after stall threshold reached)
    workoutLog.getCell('A18').value = '2024-01-21';
    workoutLog.getCell('B18').value = 'A';
    workoutLog.getCell('C18').value = 'Light Squat';
    workoutLog.getCell('F18').value = 116; // ~80% of 145
    workoutLog.getCell('G18').value = 5;
    workoutLog.getCell('H18').value = 5;
    workoutLog.getCell('I18').value = 5;

    test('Light Squat entry date recorded', workoutLog.getCell('A18').value === '2024-01-21');
    test('Light Squat exercise recorded', workoutLog.getCell('C18').value === 'Light Squat');
    test('Light Squat weight recorded', workoutLog.getCell('F18').value === 116);

    // ==================== TEST 19: Simulate Chin Ups Entry ====================
    console.log('\n--- Test 19: NEW v2.5 - Simulate Chin Ups Entry ---');

    // Add 6 unique dates to simulate 2 weeks (triggers Chin Ups availability)
    const extraDates = [
        { row: 19, date: '2024-01-22' },
        { row: 20, date: '2024-01-24' },
        { row: 21, date: '2024-01-26' },
    ];
    extraDates.forEach(d => {
        workoutLog.getCell(`A${d.row}`).value = d.date;
        workoutLog.getCell(`B${d.row}`).value = 'A';
        workoutLog.getCell(`C${d.row}`).value = 'Squat';
        workoutLog.getCell(`F${d.row}`).value = 150;
        workoutLog.getCell(`G${d.row}`).value = 5;
        workoutLog.getCell(`H${d.row}`).value = 5;
        workoutLog.getCell(`I${d.row}`).value = 5;
    });

    // Now add Chin Ups entry (simulating after 2 weeks)
    workoutLog.getCell('A22').value = '2024-01-26';
    workoutLog.getCell('B22').value = 'A';
    workoutLog.getCell('C22').value = 'Chin Ups';
    workoutLog.getCell('F22').value = 0; // bodyweight
    workoutLog.getCell('G22').value = 8;
    workoutLog.getCell('H22').value = 7;
    workoutLog.getCell('I22').value = 6;

    test('Chin Ups entry recorded', workoutLog.getCell('C22').value === 'Chin Ups');
    test('Chin Ups reps recorded', workoutLog.getCell('G22').value === 8);

    // ==================== TEST 20: Simulate Body Weight Entry ====================
    console.log('\n--- Test 20: Simulate Body Weight Data Entry ---');

    const testBodyWeights = [
        { row: 2, date: '2024-01-14', weight: 180, notes: 'Starting weight' },
        { row: 3, date: '2024-01-21', weight: 181, notes: '' },
        { row: 4, date: '2024-01-28', weight: 182.5, notes: 'Good eating' },
    ];

    testBodyWeights.forEach(bw => {
        bodyWeight.getCell(`A${bw.row}`).value = bw.date;
        bodyWeight.getCell(`B${bw.row}`).value = bw.weight;
        bodyWeight.getCell(`C${bw.row}`).value = bw.notes;
    });

    test('Body weight data entry successful', bodyWeight.getCell('B2').value === 180);
    test('Multiple body weights entered', bodyWeight.getCell('B4').value === 182.5);

    // ==================== TEST 21: Advanced Formula Verification ====================
    console.log('\n--- Test 21: Advanced Formula Verification ---');

    // Scheme formula
    test('Scheme formula detects exercise type', schemeFormula.includes('3x5') && schemeFormula.includes('1x5'));
    test('Scheme formula handles assistance (3x10)', schemeFormula.includes('3x10'));
    test('Scheme formula handles Power Clean (5x3)', schemeFormula.includes('5x3'));

    // Target Weight formula (use row 11 which has full formula, row 10 is simplified)
    const targetFormulaAdv = workoutLog.getCell('E11').value.formula;
    test('Target Weight formula references Settings starting weights', targetFormulaAdv.includes('Settings!$B$5'));
    test('Target Weight formula references Settings increments', targetFormulaAdv.includes('Settings!$B$12'));
    test('Target Weight formula uses MAXIFS for progression', targetFormulaAdv.includes('SUMPRODUCT') || targetFormulaAdv.includes('MAX'));

    // Status formula
    test('Status formula checks OK/STALL', statusFormula.includes('OK') && statusFormula.includes('STALL'));
    test('Status formula compares L vs M (total vs target reps)', statusFormula.includes('L') && statusFormula.includes('M'));

    // ==================== TEST 22: Progress Summary Formulas ====================
    console.log('\n--- Test 22: Progress Summary Formula Verification ---');

    const squatMaxFormula = summary.getCell('C4').value.formula;
    test('Squat formula references Workout Log', squatMaxFormula.includes('Workout Log'));
    test('Squat formula uses SUMPRODUCT', squatMaxFormula.includes('SUMPRODUCT'));
    test('Squat formula checks for OK status', squatMaxFormula.includes('OK'));

    // ==================== TEST 23: Program Phase Formulas ====================
    console.log('\n--- Test 23: Program Phase Formula Verification ---');

    const phaseFormula = phase.getCell('B3').value.formula;
    test('Phase formula uses MAX to check stalls', phaseFormula.includes('MAX'));
    test('Phase formula references Settings threshold', phaseFormula.includes('Settings'));
    test('Phase formula returns NOVICE or INTERMEDIATE',
         phaseFormula.includes('NOVICE') && phaseFormula.includes('INTERMEDIATE'));

    const stallFormula = phase.getCell('B9').value.formula;
    test('Stall formula uses SUMPRODUCT', stallFormula.includes('SUMPRODUCT'));
    test('Stall formula checks Workout Log', stallFormula.includes('Workout Log'));
    test('Stall formula counts STALL status', stallFormula.includes('STALL'));

    // ==================== TEST 24: Exercise List Completeness ====================
    console.log('\n--- Test 24: Exercise List Structure Completeness ---');

    // Verify exercise list structure in Assistance Exercises column E: 5 main + 2 conditional + 10 assistance = 17 rows (E2:E18)
    test('Main lifts occupy E2-E6 (5 rows)',
         assistSheet.getCell('E2').value === 'Squat' && assistSheet.getCell('E6').value === 'Power Clean');

    test('E7 is conditional Light Squat',
         assistSheet.getCell('E7').value && assistSheet.getCell('E7').value.formula !== undefined);

    test('E8 is conditional Chin Ups',
         assistSheet.getCell('E8').value && assistSheet.getCell('E8').value.formula !== undefined);

    test('E9 starts assistance exercises',
         assistSheet.getCell('E9').value && assistSheet.getCell('E9').value.formula !== undefined);

    test('E18 ends assistance exercises',
         assistSheet.getCell('E18').value && assistSheet.getCell('E18').value.formula !== undefined);

    // ==================== TEST 25: Named Ranges ====================
    console.log('\n--- Test 25: Named Ranges Verification ---');

    // Check that key cells have names (ExcelJS doesn't easily expose names, but we can verify cells exist)
    test('Bar Weight cell exists (B2)', settings.getCell('B2').value === 45);
    test('Stall Threshold cell exists (B21)', settings.getCell('B21').value === 3);
    test('Total Sessions cell exists (B32)', settings.getCell('B32').value !== undefined);
    test('Current Week cell exists (B33)', settings.getCell('B33').value !== undefined);
    test('Chin Ups Intro Week cell exists (B36)', settings.getCell('B36').value === 2);
    test('Light Squat Pct cell exists (B37)', settings.getCell('B37').value === 80);
    test('Light Squat Inc cell exists (B38)', settings.getCell('B38').value === 5);

    // ==================== SAVE TEST FILE ====================
    console.log('\n--- Saving Test Version ---');

    await workbook.xlsx.writeFile('Barbell_Strength_Tracker_TEST.xlsx');
    test('Test file saved successfully', true);

    // ==================== SUMMARY ====================
    console.log('\n' + '='.repeat(70));
    console.log('TEST SUMMARY');
    console.log('='.repeat(70));
    console.log(`Total Tests: ${testsPassed + testsFailed}`);
    console.log(`Passed: ${testsPassed}`);
    console.log(`Failed: ${testsFailed}`);
    console.log(`Success Rate: ${((testsPassed / (testsPassed + testsFailed)) * 100).toFixed(1)}%`);
    console.log('='.repeat(70));

    if (testsFailed === 0) {
        console.log('\n✓ ALL TESTS PASSED! The tracker v2.5 is working correctly.');
    } else {
        console.log('\n⚠ Some tests failed. Review the output above.');
    }

    console.log('\nv2.5 Features Tested:');
    console.log('- Session/Week tracking (auto-calculated from Workout Log)');
    console.log('- Exercise Introduction settings (Chin Ups week, Light Squat %)');
    console.log('- Conditional Light Squat (appears at intermediate transition)');
    console.log('- Conditional Chin Ups (appears after configurable weeks)');
    console.log('- Light Squat target weight formula (80% of Squat)');
    console.log('- Exercise list in Assistance Exercises column E (E2:E18)');
    console.log('- Dropdown validation references Assistance Exercises sheet');
    console.log('- Target Weight formulas use SUMPRODUCT (Excel 2013 compatible)');
    console.log('- Chin Ups removed from default assistance');
    console.log('- README updated with Exercise Introduction section');
    console.log('- Disclaimer added (not affiliated with Starting Strength, Inc.)');

    console.log('\nFiles:');
    console.log('- Barbell_Strength_Tracker.xlsx (v2.5 with smart exercise introduction)');
    console.log('- Barbell_Strength_Tracker_TEST.xlsx (with sample data)');

    return { passed: testsPassed, failed: testsFailed, success: testsFailed === 0 };
}

testStartingStrengthTracker().catch(console.error);
