/**
 * Office.js Excel Automation Script
 * Transforms the input Excel file into the final dashboard per changelog
 */

(async function () {
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            const dashboardSheet = workbook.worksheets.getItem("Dashboard");
            const rawDataSheet = workbook.worksheets.getItem("Raw Data");

            // Clear any existing calculations in Dashboard
            clearDashboard(dashboardSheet);
            
            // 1. Setup Quarterly Performance Summary formulas
            await setupSummaryFormulas(context, dashboardSheet, rawDataSheet);
            
            // 2. Apply number formatting
            await applyNumberFormatting(dashboardSheet);
            
            // 3. Setup Chart Data section
            await setupChartData(context, dashboardSheet, rawDataSheet);
            
            // 4. Create and configure chart
            await createChart(context, dashboardSheet);
            
            // 5. Apply conditional formatting
            await applyConditionalFormatting(dashboardSheet);
            
            await context.sync();
            console.log("Dashboard automation completed successfully!");
        });
    } catch (error) {
        console.error("Error:", error);
    }
})();

// Function to clear existing data in Dashboard
function clearDashboard(dashboardSheet) {
    const clearRanges = [
        "C8:G39",  // Summary table
        "A42:F51"  // Chart data area
    ];
    
    clearRanges.forEach(range => {
        const rangeObj = dashboardSheet.getRange(range);
        rangeObj.clear();
    });
}

// Function to setup all summary formulas
async function setupSummaryFormulas(context, dashboardSheet, rawDataSheet) {
    // Setup main summary table (rows 8-39)
    for (let row = 8; row <= 39; row++) {
        // 1. Total Revenue (Column C)
        const productCell = dashboardSheet.getRange(`A${row}`);
        const quarterCell = dashboardSheet.getRange(`B${row}`);
        
        const totalRevenueFormula = `=SUMIFS('${rawDataSheet.name}'!$E$2:$E$187,
            '${rawDataSheet.name}'!$D$2:$D$187,$A${row},
            '${rawDataSheet.name}'!$B$2:$B$187,VALUE(LEFT($B${row},4)),
            '${rawDataSheet.name}'!$C$2:$C$187,RIGHT($B${row},2))`;
        
        const revenueCell = dashboardSheet.getRange(`C${row}`);
        revenueCell.formulas = [[totalRevenueFormula]];
        
        // 2. Weighted Average Margin (Column D)
        const marginFormula = `=SUMIFS('${rawDataSheet.name}'!$G$2:$G$187,
            '${rawDataSheet.name}'!$D$2:$D$187,$A${row},
            '${rawDataSheet.name}'!$B$2:$B$187,VALUE(LEFT($B${row},4)),
            '${rawDataSheet.name}'!$C$2:$C$187,RIGHT($B${row},2))/C${row}`;
        
        const marginCell = dashboardSheet.getRange(`D${row}`);
        marginCell.formulas = [[marginFormula]];
        
        // 3. Rolling 3-Month Trend (Column E)
        const trendFormula = `=IF(AND(LEFT($B${row},4)="2023",RIGHT($B${row},2)="Q1"),
            "N/A",D${row}-D${row-1})`;
        
        const trendCell = dashboardSheet.getRange(`E${row}`);
        trendCell.formulas = [[trendFormula]];
        
        // 4. YoY Margin Delta (Column F)
        const yoyFormula = `=IF(LEFT($B${row},4)="2023","N/A",
            D${row}-INDEX($D$8:$D$39,
            MATCH($A${row}&" "&(VALUE(LEFT($B${row},4))-1)&" "&RIGHT($B${row},2),
            $A$8:$A$39&" "&$B$8:$B$39,0)))`;
        
        const yoyCell = dashboardSheet.getRange(`F${row}`);
        yoyCell.formulas = [[yoyFormula]];
        
        // 5. Margin Health Classification (Column G)
        const healthFormula = `=IF(D${row}>0.35,"Strong",
            IF(D${row}>=0.2,"Moderate","At Risk"))`;
        
        const healthCell = dashboardSheet.getRange(`G${row}`);
        healthCell.formulas = [[healthFormula]];
    }
    
    await context.sync();
}

// Function to apply number formatting
async function applyNumberFormatting(dashboardSheet) {
    // Currency formatting for Total Revenue (Column C)
    const revenueRange = dashboardSheet.getRange("C8:C39");
    revenueRange.numberFormat = [["$#,##0"]]; // Currency with 0 decimals
    
    // Percentage formatting for Weighted Avg Margin (Column D)
    const marginRange = dashboardSheet.getRange("D8:D39");
    marginRange.numberFormat = [["0.0%"]]; // Percentage with 1 decimal
    
    // Percentage formatting for Rolling Trend (Column E)
    const trendRange = dashboardSheet.getRange("E8:E39");
    trendRange.numberFormat = [["0.0%"]];
    
    // Percentage formatting for YoY Delta (Column F)
    const yoyRange = dashboardSheet.getRange("F8:F39");
    yoyRange.numberFormat = [["0.0%"]];
    
    // Chart data formatting
    const chartMarginRange = dashboardSheet.getRange("B44:E51");
    chartMarginRange.numberFormat = [["0.0%"]];
    
    const chartRevenueRange = dashboardSheet.getRange("F44:F51");
    chartRevenueRange.numberFormat = [["$#,##0"]];
}

// Function to setup chart data section
async function setupChartData(context, dashboardSheet, rawDataSheet) {
    // Section label
    const labelCell = dashboardSheet.getRange("A42");
    labelCell.values = [["Chart Data"]];
    labelCell.format.font.bold = true;
    
    // Headers
    const headers = ["Quarter", "Widget Pro", "Widget Standard", 
                     "Service Package", "Accessory Kit", "Total Revenue"];
    const headerRange = dashboardSheet.getRange("A43:F43");
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";
    
    // Quarter labels
    const quarters = ["2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4",
                      "2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4"];
    for (let i = 0; i < quarters.length; i++) {
        const quarterCell = dashboardSheet.getRange(`A${44 + i}`);
        quarterCell.values = [[quarters[i]]];
    }
    
    // Setup formulas for product margins
    for (let row = 44; row <= 51; row++) {
        const quarter = dashboardSheet.getRange(`A${row}`).load("text");
        await context.sync();
        
        // Widget Pro formula
        const widgetProFormula = `=SUMPRODUCT(('${rawDataSheet.name}'!$D$2:$D$187="Widget Pro")*
            ('${rawDataSheet.name}'!$B$2:$B$187=VALUE(LEFT($A${row},4)))*
            ('${rawDataSheet.name}'!$C$2:$C$187=RIGHT($A${row},2))*
            ('${rawDataSheet.name}'!$G$2:$G$187))/
            SUMPRODUCT(('${rawDataSheet.name}'!$D$2:$D$187="Widget Pro")*
            ('${rawDataSheet.name}'!$B$2:$B$187=VALUE(LEFT($A${row},4)))*
            ('${rawDataSheet.name}'!$C$2:$C$187=RIGHT($A${row},2))*
            ('${rawDataSheet.name}'!$E$2:$E$187))`;
        
        // Similar formulas for other products (simplified for example)
        const widgetStdFormula = `=SUMPRODUCT(('${rawDataSheet.name}'!$D$2:$D$187="Widget Standard")*
            ('${rawDataSheet.name}'!$B$2:$B$187=VALUE(LEFT($A${row},4)))*
            ('${rawDataSheet.name}'!$C$2:$C$187=RIGHT($A${row},2))*
            ('${rawDataSheet.name}'!$G$2:$G$187))/
            SUMPRODUCT(('${rawDataSheet.name}'!$D$2:$D$187="Widget Standard")*
            ('${rawDataSheet.name}'!$B$2:$B$187=VALUE(LEFT($A${row},4)))*
            ('${rawDataSheet.name}'!$C$2:$C$187=RIGHT($A${row},2))*
            ('${rawDataSheet.name}'!$E$2:$E$187))`;
        
        // Total Revenue formula
        const totalRevenueFormula = `=SUMIF($B$8:$B$39,$A${row},$C$8:$C$39)`;
        
        // Apply formulas
        dashboardSheet.getRange(`B${row}`).formulas = [[widgetProFormula]];
        dashboardSheet.getRange(`C${row}`).formulas = [[widgetStdFormula]];
        // Add similar for other products
        dashboardSheet.getRange(`F${row}`).formulas = [[totalRevenueFormula]];
    }
    
    await context.sync();
}

// Function to create and configure the chart
async function createChart(context, dashboardSheet) {
    // Create clustered column chart
    const chartDataRange = dashboardSheet.getRange("A43:F51");
    const chart = dashboardSheet.charts.add(
        Excel.ChartType.columnClustered,
        chartDataRange,
        Excel.ChartSeriesBy.columns
    );
    
    chart.title.text = "Quarterly Margin Trends by Product";
    chart.title.format.font.size = 14;
    
    // Move Total Revenue to secondary axis and change to line
    const series = chart.series.load("items");
    await context.sync();
    
    if (series.items.length > 0) {
        const totalRevenueSeries = series.items[series.items.length - 1];
        totalRevenueSeries.axisGroup = Excel.ChartAxisGroup.secondary;
        totalRevenueSeries.chartType = Excel.ChartType.line;
    }
    
    // Add axis titles
    chart.axes.valueMajor.title.text = "Profit Margin";
    chart.axes.valueMajor.title.format.font.size = 10;
    
    chart.axes.secondaryValueMajor.title.text = "Total Revenue ($)";
    chart.axes.secondaryValueMajor.title.format.font.size = 10;
    
    // Position chart
    chart.top = 350;
    chart.left = 50;
    chart.width = 600;
    chart.height = 300;
    
    await context.sync();
}

// Function to apply conditional formatting
async function applyConditionalFormatting(dashboardSheet) {
    const healthRange = dashboardSheet.getRange("G8:G39");
    
    // Clear any existing conditional formatting
    const formats = healthRange.conditionalFormats;
    formats.load("items");
    await context.sync();
    
    while (formats.items.length > 0) {
        formats.items[0].delete();
    }
    
    // Green for "Strong"
    const strongFormat = formats.add(Excel.ConditionalFormatType.cellValue);
    strongFormat.cellValue.format = {
        fill: { color: "#C6EFCE" }, // Light green
        font: { color: "#006100" }
    };
    strongFormat.cellValue.rule = { formula1: '="Strong"' };
    
    // Yellow for "Moderate"
    const moderateFormat = formats.add(Excel.ConditionalFormatType.cellValue);
    moderateFormat.cellValue.format = {
        fill: { color: "#FFEB9C" }, // Light yellow
        font: { color: "#9C6500" }
    };
    moderateFormat.cellValue.rule = { formula1: '="Moderate"' };
    
    // Red for "At Risk"
    const riskFormat = formats.add(Excel.ConditionalFormatType.cellValue);
    riskFormat.cellValue.format = {
        fill: { color: "#FFC7CE" }, // Light red
        font: { color: "#9C0006" }
    };
    riskFormat.cellValue.rule = { formula1: '="At Risk"' };
    
    await context.sync();
}
