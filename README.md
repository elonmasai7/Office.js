# Excel Dashboard Automation with Office.js

## Overview

A professional Office.js JavaScript solution that automatically transforms raw transactional data into an executive sales performance dashboard. This project demonstrates practical proficiency in Excel automation, dynamic calculations, and data visualization using the Office JavaScript API.

## Features

### Dynamic Calculations
- **Total Revenue** by product, year, and quarter using SUMIFS
- **Weighted Average Margin** calculations with SUMPRODUCT
- **Rolling 3-Month Trends** with conditional N/A handling
- **Year-over-Year Delta** comparisons using INDEX/MATCH
- **Margin Health Classification** with nested IF statements

### Professional Formatting
- Currency formatting with zero decimals
- Percentage formatting with one decimal place
- Conditional formatting (green/yellow/red) based on performance
- Bold headers with professional color schemes

### Interactive Visualizations
- Clustered column charts showing quarterly margins
- Dual-axis configuration (columns + line chart)
- Automated chart positioning and sizing
- Custom axis titles and legends

### Performance Optimized
- Batch operations minimize context.sync() calls
- Efficient range manipulation
- Modular function design for maintainability
- Comprehensive error handling

## Technical Stack

| Technology | Purpose | Version |
|------------|---------|---------|
| **Office.js** | Excel automation API | 1.1+ |
| **JavaScript** | Core programming language | ES6+ |
| **Excel** | Spreadsheet platform | 2016+ / Microsoft 365 |
| **VS Code** | Recommended IDE | Latest |

## Project Structure

```
excel-dashboard-automation/
├── README.md                # This documentation
├── dashboard-automation.js  # Main Office.js script
├── Excel_SFT_input_sample.xlsx      # Input file template
├── Excel_SFT_final_sample.xlsx      # Expected output
└── Excel_SFT_changelog_sample.pdf   # Implementation specifications
```

## Quick Start

### Prerequisites
1. **Microsoft Excel** (2016 or later, or Microsoft 365)
2. **Script Lab** add-in for Excel (or Office.js development environment)
3. Basic understanding of Excel formulas and JavaScript

### Installation & Setup

#### Method 1: Using Script Lab (Recommended)
1. Open Excel and install "Script Lab" from the Insert > Get Add-ins store
2. Create a new snippet in Script Lab
3. Paste the entire dashboard-automation.js code
4. Load your input Excel file (Excel_SFT_input_sample.xlsx)
5. Run the script

#### Method 2: Office Add-in Development
```bash
# Clone and set up a development environment
git clone https://github.com/elonmasai7/Office.js.git
cd Office.js

# Install dependencies (if using Yeoman generator)
npm install -g yo generator-office
yo office

# Follow prompts to create add-in project
```

### Running the Automation
1. Open the input Excel file in Excel
2. Load the script via Script Lab or as an add-in
3. Click "Run" to execute the automation
4. Watch as the dashboard builds automatically

## Core Functions

| Function | Purpose | Key Features |
|----------|---------|--------------|
| `clearDashboard()` | Prepares workspace | Clears previous calculations, resets formatting |
| `setupSummaryFormulas()` | Builds main table | Implements all required Excel formulas dynamically |
| `applyNumberFormatting()` | Professional formatting | Currency, percentage, decimal place control |
| `setupChartData()` | Chart data preparation | SUMPRODUCT formulas for margin aggregation |
| `createChart()` | Visualization creation | Dual-axis chart with custom styling |
| `applyConditionalFormatting()` | Visual indicators | Green/yellow/red based on margin thresholds |

## Change Log Implementation

The script implements all 34 steps from the changelog:

### Phase 1: Data Calculations (Steps 1-12)
- SUMIFS for total revenue (Step 1-3)
- SUMPRODUCT for weighted averages (Step 4-6)
- IF formulas for trend analysis (Step 7-9)
- INDEX/MATCH for YoY comparisons (Step 10-12)

### Phase 2: Classification & Formatting (Steps 13-15)
- Nested IF for health classification (Step 13-14)
- Conditional formatting with color scales (Step 15)

### Phase 3: Chart Data Preparation (Steps 16-26)
- Chart section setup (Step 16-18)
- SUMPRODUCT formulas for product margins (Step 19-24)
- Professional number formatting (Step 25-26)

### Phase 4: Visualization (Steps 27-34)
- Clustered column chart creation (Step 27-28)
- Dual-axis configuration (Step 29-30)
- Custom titles and positioning (Step 31-34)

## Advanced Usage

### Customizing Thresholds
```javascript
// Modify margin health thresholds
const HEALTH_THRESHOLDS = {
  STRONG: 0.35,    // Above 35%
  MODERATE: 0.20,  // 20% - 35%
  AT_RISK: 0.20    // Below 20%
};
```

### Extending Product Lines
```javascript
// Add new products to the analysis
const PRODUCTS = [
  "Widget Pro",
  "Widget Standard", 
  "Service Package",
  "Accessory Kit",
  "New Product"  // Add your product here
];
```

### Custom Date Ranges
```javascript
// Adjust quarter analysis
const QUARTERS = [
  "2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4",
  "2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4",
  "2025 Q1"  // Add future quarters
];
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **Script Lab not loading** | Enable add-ins in Excel Options > Trust Center |
| **Formulas not calculating** | Ensure workbook is in Automatic Calculation mode |
| **Chart not appearing** | Check if data range contains valid numbers |
| **Permission errors** | Grant necessary permissions when prompted |
| **Slow performance** | Reduce context.sync() calls, batch operations |

## Performance Metrics

| Operation | Time (approx) | Optimization |
|-----------|---------------|--------------|
| Formula setup | 2-3 seconds | Batch formula assignment |
| Formatting | 1-2 seconds | Range-based formatting |
| Chart creation | 1 second | Template-based charting |
| **Total runtime** | **4-6 seconds** | Efficient context management |

## Code Quality

### Best Practices Implemented
- **Modular Design**: Separated concerns with dedicated functions
- **Error Handling**: Comprehensive try-catch blocks
- **Performance**: Minimal context.sync() calls
- **Readability**: Clear variable names and comments
- **Maintainability**: Configurable constants and functions

### Testing
```javascript
// Example test function
async function testDashboard() {
  const testCases = [
    { product: "Widget Pro", quarter: "2024 Q1", expectedMargin: "35%" },
    { product: "Widget Standard", quarter: "2023 Q4", expectedRevenue: "$150,000" }
  ];
  
  // Run validations
  // Add your testing logic here
}
```

## Learning Resources

### Office.js Documentation
- [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/overview/excel)
- [Script Lab Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/explore-with-script-lab)
- [Excel JavaScript Tutorials](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial)

### Related Projects
- [Excel Formula Builder](https://github.com/OfficeDev/Excel-Custom-Functions)
- [Dashboard Templates](https://github.com/OfficeDev/Office-Add-in-samples)
- [Data Visualization Examples](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)

## Contributing

We welcome contributions! Here's how you can help:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **Push** to the branch (`git push origin feature/AmazingFeature`)
5. **Open** a Pull Request

### Areas for Contribution
- Additional chart types
- Export functionality (PDF/Excel)
- Real-time data connections
- Mobile-responsive design
- Unit tests and validation

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Microsoft Office.js team for the comprehensive API
- Excel community for formula patterns and best practices
- Contributors and testers who helped refine the solution

## Support

Having issues or questions?

1. **Check the Troubleshooting** section above
2. **Review the Office.js documentation**
3. **Open an Issue** on GitHub with:
   - Excel version
   - Error messages
   - Steps to reproduce

---

Built for Excel power users and Office.js developers

Transform your raw data into actionable insights with just one click
