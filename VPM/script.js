function calculateVendorScores() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Vendor Data");
    const weightSheet = ss.getSheetByName("KPI Weights");

    if (!sheet || !weightSheet) {
        Logger.log("❌ Error: Required sheets not found.");
        return;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(header => header.toString().trim());
    const headers1 = data[1].map(header => header.toString().trim());

    const kpiColumns = [
        "AVERAGE PURCHASE PRICE",
        "AVG PRICE ADV FOR LARGER QTY",
        "DHU",
        "FIRST TIME PASS",
        "ON TIME DELIVERY %",
        "OTIF",
        "ORDER FULFILlMENT",
        "COMPLIANCE"
    ];

    const kpiIndices = kpiColumns.map(kpi => {
        const indexInHeaders = headers.findIndex(header => header.includes(kpi));
        const indexInHeaders1 = headers1.findIndex(header => header.includes(kpi));
        return indexInHeaders !== -1 ? indexInHeaders : indexInHeaders1;
    });

    if (kpiIndices.some(index => index === -1)) {
        Logger.log("❌ Error: One or more KPI columns not found. Check your headers.");
        Logger.log(`KPI Indices: ${kpiIndices}`);
        return;
    }

    const rawScoreCol = headers.indexOf("RAW SCORE") + 1;
    const maxRawScoreCol = headers.indexOf("MAX RAW SCORE") + 1;
    const finalScoreCol = headers.indexOf("FINAL SCORE") + 1;
    const finalRankCol = headers.indexOf("FINAL RANK") + 1;

    const weights = weightSheet.getRange("B2:B9").getValues().flat().map(w => parseFloat(w)); // Use weights as-is without dividing by 100
    const lastRow = sheet.getLastRow();
    const rawScores = [];

    for (let i = 2; i < lastRow; i++) { // Start from row 2 (data rows)
        let rawScore = 0;
        Logger.log(`Row ${i + 1}:`);

        kpiIndices.forEach((colIndex, weightIndex) => {
            let kpiValue = parseFloat(data[i][colIndex]);
            const weight = weights[weightIndex];

            if (!isNaN(kpiValue)) {
                // Convert percentage values to decimals only for specific KPIs
                if (["COMPLIANCE"].includes(kpiColumns[weightIndex])) {
                    kpiValue = kpiValue / 100; // Divide by 100 only for percentage-based KPIs
                }
                // Do not divide AVERAGE PURCHASE PRICE and DHU by 100
                rawScore += kpiValue * weight;
                Logger.log(`  KPI: ${kpiColumns[weightIndex]}, Value: ${kpiValue}, Weight: ${weight}, Contribution: ${kpiValue * weight}`);
            } else {
                Logger.log(`  KPI: ${kpiColumns[weightIndex]}, Value: N/A, Weight: ${weight}`);
            }
        });

        rawScores.push(rawScore);
        sheet.getRange(i + 1, rawScoreCol).setValue(rawScore.toFixed(2));
        Logger.log(`  Raw Score: ${rawScore.toFixed(2)}`);
    }

    if (rawScores.length === 0 || rawScores.some(score => isNaN(score))) {
        Logger.log("❌ Error: Invalid KPI values in raw scores.");
        return;
    }

    const maxRawScore = Math.max(...rawScores);
    sheet.getRange(2, maxRawScoreCol, lastRow - 1, 1).setValue(maxRawScore.toFixed(2));

    const minScore = Math.min(...rawScores);
    let maxScore = Math.max(...rawScores);
    if (minScore === maxScore) maxScore += 1;

    let index = 0;
    for (let i = 2; i < lastRow; i++) {
        if (rawScores[index] !== undefined) {
            const scaledScore = 50 + ((rawScores[index] - minScore) / (maxScore - minScore)) * 50;
            sheet.getRange(i + 1, finalScoreCol).setValue(scaledScore.toFixed(2));
            index++;
        }
    }

    rankVendors(); // Automatically calculate ranks after scores are calculated
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Vendor Scores & Rankings Updated", "Success", 3);
}

function rankVendors() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendor Data");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const finalScoreCol = headers.indexOf("FINAL SCORE");
    const finalRankCol = headers.indexOf("FINAL RANK");

    if (finalScoreCol === -1 || finalRankCol === -1) {
        Logger.log("❌ Error: 'FINAL SCORE' or 'FINAL RANK' column not found.");
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const scores = sheet.getRange(2, finalScoreCol + 1, lastRow - 1, 1).getValues().flat();

    const scoreList = scores.map((val, i) => ({
        row: i + 2,
        score: parseFloat(val)
    })).filter(item => !isNaN(item.score));

    scoreList.sort((a, b) => b.score - a.score);

    let currentRank = 1;
    let prevScore = null;
    let offset = 0;

    scoreList.forEach((entry, index) => {
        if (entry.score === prevScore) {
            offset++;
        } else {
            currentRank = index + 1;
            offset = 0;
        }
        sheet.getRange(entry.row, finalRankCol + 1).setValue(currentRank);
        prevScore = entry.score;
    });
}

function printColumnNames() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Vendor Data");

    if (!sheet) {
        Logger.log("❌ Error: 'Vendor Data' sheet not found.");
        return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Column Names:");
    headers.forEach((header, index) => {
        Logger.log(`${index + 1}: ${header}`);
    });
}











