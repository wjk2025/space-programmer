# Architecture Space Programming Tool

A comprehensive space programming calculator for architecture firms to determine space requirements based on staff counts, departmental allocations, and industry-standard factors.

## Quick Start

### Option 1: Web Interface (Recommended)
```bash
# Install dependencies
pip install flask openpyxl reportlab

# Run the web app
python app.py

# Open browser to http://localhost:5000
```

### Option 2: Command Line
```bash
# Install dependencies
pip install openpyxl reportlab

# Run interactive CLI
python space_programmer.py
```

### Option 3: Generate Demo Output
```bash
python generate_demo.py
```

## Features

### Core Functionality
- **Staff-Based Calculations**: Input total and departmental staff counts with various workspace types
- **Standard Space Sizes**: Pre-configured industry-standard sizes for workstations, offices, and support spaces
- **Usable Square Feet (USF)**: Calculates net assignable area plus circulation
- **Rentable Square Feet (RSF)**: Converts USF to RSF using configurable loss factor
- **Company Registration**: Track company name, location, project name, and preparer

### Remote Work Analysis
- Analyze impact of 6 different remote work policies on space requirements
- Compare scenarios from full on-site to fully remote
- Generate recommendations for hybrid workspace strategies
- Calculate potential square footage savings

### Export Formats
- **PDF Report**: Professional formatted report with executive summary
- **Excel Workbook (.xlsx)**: Multi-sheet workbook with:
  - Summary sheet
  - Department detail with quantities
  - Support spaces breakdown
  - Remote work scenario analysis
  - Space standards reference

### Data Persistence
- Save program data as JSON for future adjustments
- Load and modify existing programs
- Track version history

## Space Standards (Default)

### Workstations
| Type | Size (SF) | Description |
|------|-----------|-------------|
| Open Workstation | 48 | Open plan workstation |
| Standard Workstation | 64 | Standard cubicle |
| Large Workstation | 80 | Senior/large workstation |

### Private Offices
| Type | Size (SF) | Description |
|------|-----------|-------------|
| Small Office | 100 | Small enclosed office |
| Standard Office | 150 | Standard private office |
| Large Office | 200 | Executive/large office |
| Executive Suite | 300 | C-suite office |

### Support Spaces
| Type | Size (SF) | Description |
|------|-----------|-------------|
| Small Conference | 150 | 4-6 person meeting room |
| Medium Conference | 300 | 8-12 person conference |
| Large Conference | 500 | 16-20 person boardroom |
| Huddle Room | 80 | 2-4 person huddle |
| Phone Booth | 35 | Single person focus booth |
| Break Room | 200 | Kitchen/break area |
| Reception | 250 | Reception/waiting area |
| Copy/Print | 100 | Copy/print/mail area |
| Storage | 150 | General storage |
| Server Room | 120 | Server/IT closet |
| Wellness Room | 80 | Wellness/lactation room |
| Training Room | 600 | Large training room |
| Collaboration Area | 200 | Open collaboration zone |

## Industry Standard Factors

### Circulation Factor
- **Default**: 35% (0.35)
- **Typical Range**: 30-40%
- Added to net assignable SF to account for corridors, aisles, and circulation paths

### Loss Factor (USF to RSF Conversion)
- **Default**: 15% (0.15)
- **Typical Range**: 10-20%
- Accounts for building core, mechanical, structural columns, etc.

## Remote Work Policies

| Policy | Reduction | Description |
|--------|-----------|-------------|
| Full On-Site | 0% | 100% on-site, no reduction |
| Hybrid Light | 15% | 1 day remote/week |
| Hybrid Moderate | 30% | 2 days remote/week |
| Hybrid Heavy | 45% | 3 days remote/week |
| Remote First | 60% | 4 days remote/week |
| Fully Remote | 85% | Fully remote, hoteling only |

## Usage

### Interactive Mode
```bash
python space_programmer.py
```

Follow the menu prompts to:
1. Enter company information
2. Add departments with staff counts
3. Configure support spaces
4. Set circulation and loss factors
5. Select remote work policy
6. Calculate and export results

### Programmatic Usage
```python
from space_programmer import (
    SpaceProgram, Department, SupportSpaces,
    SpaceCalculator, RemoteWorkAnalyzer, DataManager,
    export_to_pdf, export_to_excel
)

# Create program
program = SpaceProgram(
    company_name="Your Company",
    location="City, State",
    circulation_factor=0.35,
    loss_factor=0.15,
    remote_work_policy="hybrid_moderate"
)

# Add departments
program.departments.append(Department(
    name="Engineering",
    standard_workstations=50,
    standard_offices=10,
    large_offices=2
))

# Configure support spaces
program.support_spaces = SupportSpaces(
    small_conference=4,
    medium_conference=2,
    huddle_rooms=6,
    break_rooms=2
)

# Calculate
calculator = SpaceCalculator(program)
results = calculator.calculate_totals()

# Analyze remote work impact
analyzer = RemoteWorkAnalyzer(program)
remote_analysis = analyzer.analyze_scenarios()

# Export
export_to_pdf(results, remote_analysis, program, "output.pdf")
export_to_excel(results, remote_analysis, program, "output.xlsx")

# Save for future modifications
data_manager = DataManager("./data")
data_manager.save_program(program)
```

## Output Files

### PDF Report Contents
- Executive summary with company info
- Space calculations breakdown
- Key metrics (SF per person)
- Department breakdown
- Remote work impact analysis
- Recommendations

### Excel Workbook Sheets
1. **Summary**: Overview with all calculations
2. **Department Detail**: Staff counts by workspace type per department
3. **Support Spaces**: Quantities and totals for support areas
4. **Remote Work Analysis**: All scenario comparisons
5. **Space Standards**: Reference sheet with all standard sizes

## Customization

### Custom Space Standards
Override default sizes in your program:
```python
program.custom_standards = {
    "workstation_open": {"name": "Open WS", "sf": 42},
    "office_standard": {"name": "Private Office", "sf": 180}
}
```

### Adjusting Factors
```python
program.circulation_factor = 0.40  # 40% circulation
program.loss_factor = 0.12  # 12% loss factor
```

## Requirements

- Python 3.8+
- openpyxl (Excel export)
- reportlab (PDF export)

Install dependencies:
```bash
pip install openpyxl reportlab
```

## File Structure

```
space_programmer/
├── app.py                 # Flask web application
├── space_programmer.py    # Core calculation engine
├── generate_demo.py       # Demo generation script
├── requirements.txt       # Python dependencies
├── Dockerfile            # Docker containerization
├── Procfile              # Heroku deployment
├── render.yaml           # Render.com deployment
├── railway.json          # Railway deployment
├── .gitignore            # Git ignore rules
├── data/                 # Saved program data (JSON)
└── outputs/              # Generated PDF/Excel files
```

## Deployment

### Render.com (Recommended - Free Tier)
1. Push to GitHub
2. Connect repo at render.com
3. It auto-detects `render.yaml`
4. Deploy!

### Railway
1. Push to GitHub
2. Connect repo at railway.app
3. It auto-detects `railway.json`
4. Deploy!

### Heroku
```bash
heroku create space-programmer
git push heroku main
```

### Docker
```bash
docker build -t space-programmer .
docker run -p 5000:5000 space-programmer
```

## Support

For questions or feature requests, contact your architecture firm's IT department or the tool maintainer.
