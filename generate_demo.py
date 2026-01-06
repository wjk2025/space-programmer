#!/usr/bin/env python3
"""
Demo script to generate sample space programming outputs
"""

import sys
sys.path.insert(0, '/home/claude/space_programmer')

from space_programmer import (
    SpaceProgram, Department, SupportSpaces, 
    SpaceCalculator, RemoteWorkAnalyzer, DataManager,
    export_to_pdf, export_to_excel
)

def create_demo_program():
    """Create a realistic demo program for a mid-size company"""
    
    program = SpaceProgram(
        company_name="Acme Technologies Inc.",
        location="Denver, Colorado",
        project_name="New HQ Space Planning",
        prepared_by="ABC Architecture Partners",
        circulation_factor=0.35,  # 35% circulation
        loss_factor=0.15,  # 15% loss factor
        remote_work_policy="hybrid_moderate",  # 2 days remote
        notes="Initial programming for new headquarters relocation. "
              "Client anticipates 10% growth over next 3 years. "
              "Hybrid work policy allows 2 days remote per week."
    )
    
    # Add departments
    program.departments = [
        Department(
            name="Executive Leadership",
            executive_offices=5,
            large_offices=3,
            standard_offices=2,
        ),
        Department(
            name="Finance & Accounting",
            standard_offices=4,
            small_offices=2,
            standard_workstations=12,
        ),
        Department(
            name="Human Resources",
            standard_offices=2,
            small_offices=3,
            standard_workstations=8,
        ),
        Department(
            name="Marketing & Sales",
            large_offices=2,
            standard_offices=6,
            standard_workstations=20,
            open_workstations=15,
        ),
        Department(
            name="Engineering",
            large_offices=4,
            standard_offices=8,
            large_workstations=25,
            standard_workstations=35,
        ),
        Department(
            name="Product Management",
            standard_offices=6,
            standard_workstations=12,
        ),
        Department(
            name="Customer Success",
            small_offices=2,
            standard_workstations=18,
            open_workstations=10,
        ),
        Department(
            name="IT & Operations",
            standard_offices=3,
            standard_workstations=8,
        ),
    ]
    
    # Add support spaces
    program.support_spaces = SupportSpaces(
        small_conference=6,      # 4-6 person rooms
        medium_conference=4,     # 8-12 person rooms
        large_conference=2,      # Board rooms
        huddle_rooms=8,          # 2-4 person
        phone_booths=12,         # Focus/call booths
        break_rooms=3,           # Kitchen/break areas
        reception_areas=1,       # Main reception
        copy_print_centers=2,    # Print/mail rooms
        storage_rooms=3,         # General storage
        server_rooms=1,          # IT closet
        wellness_rooms=2,        # Mother's/wellness
        training_rooms=1,        # Large training
        collaboration_areas=4,   # Open collab zones
    )
    
    return program


def main():
    print("="*60)
    print("SPACE PROGRAMMING TOOL - DEMO OUTPUT GENERATION")
    print("="*60)
    
    # Create demo program
    program = create_demo_program()
    
    # Initialize data manager
    data_manager = DataManager('/home/claude/space_programmer/data')
    
    # Calculate results
    calculator = SpaceCalculator(program)
    results = calculator.calculate_totals()
    
    # Analyze remote work scenarios
    analyzer = RemoteWorkAnalyzer(program)
    remote_analysis = analyzer.analyze_scenarios()
    
    # Display summary
    print(f"\nCompany: {program.company_name}")
    print(f"Location: {program.location}")
    print("-"*60)
    
    print(f"\nDEPARTMENT SUMMARY:")
    for dept in results["departments"]:
        print(f"  {dept['name']}: {dept['staff']} staff, {dept['breakdown']['total']:,.0f} SF")
    
    print(f"\n{'='*60}")
    print("SPACE CALCULATIONS")
    print("="*60)
    print(f"Total Headcount:       {results['totals']['total_staff']:>12,} persons")
    print(f"Department Space:      {results['totals']['department_sf']:>12,.0f} SF")
    print(f"Support Space:         {results['totals']['support_sf']:>12,.0f} SF")
    print(f"Net Assignable SF:     {results['totals']['net_assignable_sf']:>12,.0f} SF")
    print(f"Circulation (35%):     {results['totals']['circulation_sf']:>12,.0f} SF")
    print(f"Usable SF:             {results['totals']['usable_sf']:>12,.0f} SF")
    print(f"\nRemote Policy: {results['totals']['remote_work_description']}")
    print(f"Adjusted Usable SF:    {results['totals']['adjusted_usable_sf']:>12,.0f} SF")
    print(f"Loss Factor (15%):     {results['totals']['loss_sf']:>12,.0f} SF")
    print(f"-"*60)
    print(f"RENTABLE SF:           {results['totals']['rentable_sf']:>12,.0f} SF")
    print(f"="*60)
    
    print(f"\nKEY METRICS:")
    print(f"  SF per Person (Net):      {results['metrics']['sf_per_person_net']:>8.1f} SF")
    print(f"  SF per Person (Usable):   {results['metrics']['sf_per_person_usable']:>8.1f} SF")
    print(f"  SF per Person (Adjusted): {results['metrics']['sf_per_person_adjusted']:>8.1f} SF")
    print(f"  SF per Person (Rentable): {results['metrics']['sf_per_person_rentable']:>8.1f} SF")
    
    # Save program data
    json_path = data_manager.save_program(program, "Acme_Technologies_demo.json")
    print(f"\n✓ Program data saved: {json_path}")
    
    # Export to PDF
    pdf_path = "/mnt/user-data/outputs/Acme_Technologies_Space_Program.pdf"
    export_to_pdf(results, remote_analysis, program, pdf_path)
    print(f"✓ PDF report exported: {pdf_path}")
    
    # Export to Excel
    xlsx_path = "/mnt/user-data/outputs/Acme_Technologies_Space_Program.xlsx"
    export_to_excel(results, remote_analysis, program, xlsx_path)
    print(f"✓ Excel workbook exported: {xlsx_path}")
    
    print("\n" + "="*60)
    print("REMOTE WORK ANALYSIS - SCENARIO COMPARISON")
    print("="*60)
    print(f"Base Rentable SF (Full On-Site): {remote_analysis['base_rentable_sf']:,.0f} SF")
    print(f"\n{'Policy':<20} {'Reduction':>10} {'Rentable SF':>14} {'SF Saved':>12}")
    print("-"*60)
    for s in remote_analysis['scenarios']:
        print(f"{s['policy'].replace('_',' ').title():<20} "
              f"{s['percent_reduction']:>9.0f}% "
              f"{s['adjusted_rentable_sf']:>13,.0f} "
              f"{s['rentable_sf_saved']:>11,.0f}")
    
    print("\n" + "="*60)
    print("DEMO COMPLETE - Files ready for download")
    print("="*60)
    
    return pdf_path, xlsx_path


if __name__ == "__main__":
    main()
