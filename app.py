#!/usr/bin/env python3
"""
Flask Web Application for Space Programming Tool
Run with: python app.py
Access at: http://localhost:5000
"""

from flask import Flask, render_template, request, jsonify, send_file, session
import os
import json
from datetime import datetime
from pathlib import Path

from space_programmer import (
    SpaceProgram, Department, SupportSpaces,
    SpaceCalculator, RemoteWorkAnalyzer, DataManager,
    export_to_pdf, export_to_excel, DEFAULT_SPACE_STANDARDS, REMOTE_WORK_FACTORS
)

app = Flask(__name__)
app.secret_key = 'space-programmer-secret-key-change-in-production'

# Data directory
DATA_DIR = Path('./data')
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR = Path('./outputs')
OUTPUT_DIR.mkdir(exist_ok=True)

def get_program_from_session():
    """Reconstruct SpaceProgram from session data"""
    data = session.get('program', {})
    if not data:
        return SpaceProgram()
    
    departments = [Department(**d) for d in data.get('departments', [])]
    support_data = data.get('support_spaces', {})
    support_spaces = SupportSpaces(**support_data) if support_data else SupportSpaces()
    
    return SpaceProgram(
        company_name=data.get('company_name', ''),
        location=data.get('location', ''),
        project_name=data.get('project_name', ''),
        prepared_by=data.get('prepared_by', ''),
        date_created=data.get('date_created', datetime.now().strftime("%Y-%m-%d")),
        departments=departments,
        support_spaces=support_spaces,
        circulation_factor=data.get('circulation_factor', 0.35),
        loss_factor=data.get('loss_factor', 0.15),
        remote_work_policy=data.get('remote_work_policy', 'full_onsite'),
        notes=data.get('notes', ''),
    )

def save_program_to_session(program):
    """Save SpaceProgram to session"""
    from dataclasses import asdict
    session['program'] = {
        'company_name': program.company_name,
        'location': program.location,
        'project_name': program.project_name,
        'prepared_by': program.prepared_by,
        'date_created': program.date_created,
        'departments': [asdict(d) for d in program.departments],
        'support_spaces': asdict(program.support_spaces),
        'circulation_factor': program.circulation_factor,
        'loss_factor': program.loss_factor,
        'remote_work_policy': program.remote_work_policy,
        'notes': program.notes,
    }
    session.modified = True


# HTML Template embedded for simplicity
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Space Programming Tool</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        header { background: linear-gradient(135deg, #2E5090, #4472C4); color: white; padding: 30px; margin-bottom: 30px; border-radius: 10px; }
        header h1 { font-size: 2em; margin-bottom: 5px; }
        header p { opacity: 0.9; }
        .card { background: white; border-radius: 10px; padding: 25px; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .card h2 { color: #2E5090; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #e0e6ed; }
        .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; font-weight: 600; margin-bottom: 5px; color: #555; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; transition: border-color 0.2s; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: #4472C4; }
        .form-group input[type="number"] { text-align: right; }
        .btn { padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.2s; }
        .btn-primary { background: #2E5090; color: white; }
        .btn-primary:hover { background: #243d6e; }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #218838; }
        .btn-danger { background: #dc3545; color: white; }
        .btn-danger:hover { background: #c82333; }
        .btn-outline { background: white; border: 2px solid #2E5090; color: #2E5090; }
        .btn-outline:hover { background: #2E5090; color: white; }
        .btn-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #e0e6ed; }
        th { background: #f8f9fa; font-weight: 600; color: #555; }
        tr:hover { background: #f8f9fa; }
        .number { text-align: right; font-family: 'SF Mono', Monaco, monospace; }
        .total-row { background: #FFC000 !important; font-weight: 700; }
        .total-row td { border-top: 2px solid #333; }
        .results-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 20px; }
        .metric-card { background: linear-gradient(135deg, #f8f9fa, #e9ecef); padding: 20px; border-radius: 8px; text-align: center; }
        .metric-card .value { font-size: 2em; font-weight: 700; color: #2E5090; }
        .metric-card .label { color: #666; font-size: 0.9em; margin-top: 5px; }
        .metric-card.highlight { background: linear-gradient(135deg, #FFC000, #ffb300); }
        .metric-card.highlight .value { color: #333; }
        .tabs { display: flex; border-bottom: 2px solid #e0e6ed; margin-bottom: 20px; }
        .tab { padding: 12px 24px; cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; font-weight: 600; color: #666; }
        .tab:hover { color: #2E5090; }
        .tab.active { color: #2E5090; border-bottom-color: #2E5090; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .department-item { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; }
        .department-item .info { flex: 1; }
        .department-item .name { font-weight: 600; color: #333; }
        .department-item .details { color: #666; font-size: 0.9em; }
        .alert { padding: 15px; border-radius: 6px; margin-bottom: 20px; }
        .alert-info { background: #e7f3ff; border: 1px solid #b6d4fe; color: #084298; }
        .alert-success { background: #d1e7dd; border: 1px solid #badbcc; color: #0f5132; }
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; }
        .modal.active { display: flex; align-items: center; justify-content: center; }
        .modal-content { background: white; padding: 30px; border-radius: 10px; max-width: 600px; width: 90%; max-height: 90vh; overflow-y: auto; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .modal-header h3 { color: #2E5090; }
        .close-btn { background: none; border: none; font-size: 24px; cursor: pointer; color: #666; }
        .scenario-table tr.current { background: #e7f3ff; }
        .help-text { font-size: 0.85em; color: #666; margin-top: 3px; }
        @media (max-width: 768px) {
            .form-grid { grid-template-columns: 1fr; }
            .results-grid { grid-template-columns: 1fr 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>🏢 Space Programming Tool</h1>
            <p>Calculate space requirements based on staff counts and industry standards</p>
        </header>

        <div class="tabs">
            <div class="tab active" data-tab="company">Company Info</div>
            <div class="tab" data-tab="departments">Departments</div>
            <div class="tab" data-tab="support">Support Spaces</div>
            <div class="tab" data-tab="factors">Factors</div>
            <div class="tab" data-tab="results">Results</div>
        </div>

        <!-- Company Info Tab -->
        <div id="company" class="tab-content active">
            <div class="card">
                <h2>Company Information</h2>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Company Name</label>
                        <input type="text" id="company_name" placeholder="Enter company name">
                    </div>
                    <div class="form-group">
                        <label>Location</label>
                        <input type="text" id="location" placeholder="City, State">
                    </div>
                    <div class="form-group">
                        <label>Project Name</label>
                        <input type="text" id="project_name" placeholder="Optional project name">
                    </div>
                    <div class="form-group">
                        <label>Prepared By</label>
                        <input type="text" id="prepared_by" placeholder="Your name/firm">
                    </div>
                </div>
                <div class="form-group">
                    <label>Notes</label>
                    <textarea id="notes" rows="3" placeholder="Additional notes or assumptions"></textarea>
                </div>
                <div class="btn-group">
                    <button class="btn btn-primary" onclick="saveCompanyInfo()">Save & Continue</button>
                </div>
            </div>
        </div>

        <!-- Departments Tab -->
        <div id="departments" class="tab-content">
            <div class="card">
                <h2>Departments</h2>
                <div id="department-list"></div>
                <div class="btn-group">
                    <button class="btn btn-primary" onclick="openDeptModal()">+ Add Department</button>
                </div>
            </div>
        </div>

        <!-- Support Spaces Tab -->
        <div id="support" class="tab-content">
            <div class="card">
                <h2>Support & Amenity Spaces</h2>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Small Conference Rooms (4-6 person)</label>
                        <input type="number" id="small_conference" min="0" value="0">
                        <div class="help-text">150 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Medium Conference Rooms (8-12 person)</label>
                        <input type="number" id="medium_conference" min="0" value="0">
                        <div class="help-text">300 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Large Conference/Board Rooms (16-20)</label>
                        <input type="number" id="large_conference" min="0" value="0">
                        <div class="help-text">500 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Huddle Rooms (2-4 person)</label>
                        <input type="number" id="huddle_rooms" min="0" value="0">
                        <div class="help-text">80 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Phone/Focus Booths</label>
                        <input type="number" id="phone_booths" min="0" value="0">
                        <div class="help-text">35 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Break Rooms/Kitchens</label>
                        <input type="number" id="break_rooms" min="0" value="0">
                        <div class="help-text">200 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Reception Areas</label>
                        <input type="number" id="reception_areas" min="0" value="0">
                        <div class="help-text">250 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Copy/Print Centers</label>
                        <input type="number" id="copy_print_centers" min="0" value="0">
                        <div class="help-text">100 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Storage Rooms</label>
                        <input type="number" id="storage_rooms" min="0" value="0">
                        <div class="help-text">150 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Server/IT Rooms</label>
                        <input type="number" id="server_rooms" min="0" value="0">
                        <div class="help-text">120 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Wellness/Mother's Rooms</label>
                        <input type="number" id="wellness_rooms" min="0" value="0">
                        <div class="help-text">80 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Training Rooms</label>
                        <input type="number" id="training_rooms" min="0" value="0">
                        <div class="help-text">600 SF each</div>
                    </div>
                    <div class="form-group">
                        <label>Collaboration Areas</label>
                        <input type="number" id="collaboration_areas" min="0" value="0">
                        <div class="help-text">200 SF each</div>
                    </div>
                </div>
                <div class="btn-group">
                    <button class="btn btn-primary" onclick="saveSupportSpaces()">Save Support Spaces</button>
                </div>
            </div>
        </div>

        <!-- Factors Tab -->
        <div id="factors" class="tab-content">
            <div class="card">
                <h2>Calculation Factors</h2>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Circulation Factor (%)</label>
                        <input type="number" id="circulation_factor" min="0" max="100" value="35">
                        <div class="help-text">Industry standard: 30-40%. Added to net SF for corridors/aisles.</div>
                    </div>
                    <div class="form-group">
                        <label>Loss Factor (%) - USF to RSF</label>
                        <input type="number" id="loss_factor" min="0" max="50" value="15">
                        <div class="help-text">Typical: 10-20%. Accounts for building core, mechanical, etc.</div>
                    </div>
                    <div class="form-group">
                        <label>Remote Work Policy</label>
                        <select id="remote_work_policy">
                            <option value="full_onsite">Full On-Site (0% reduction)</option>
                            <option value="hybrid_light">Hybrid Light - 1 day remote (15% reduction)</option>
                            <option value="hybrid_moderate">Hybrid Moderate - 2 days remote (30% reduction)</option>
                            <option value="hybrid_heavy">Hybrid Heavy - 3 days remote (45% reduction)</option>
                            <option value="remote_first">Remote First - 4 days remote (60% reduction)</option>
                            <option value="fully_remote">Fully Remote - Hoteling only (85% reduction)</option>
                        </select>
                    </div>
                </div>
                <div class="btn-group">
                    <button class="btn btn-primary" onclick="saveFactors()">Save Factors</button>
                </div>
            </div>
        </div>

        <!-- Results Tab -->
        <div id="results" class="tab-content">
            <div class="card">
                <h2>Space Calculation Results</h2>
                <div id="results-content">
                    <div class="alert alert-info">
                        Click "Calculate Results" to see space requirements.
                    </div>
                </div>
                <div class="btn-group">
                    <button class="btn btn-primary" onclick="calculateResults()">Calculate Results</button>
                    <button class="btn btn-success" onclick="exportPDF()">Export PDF</button>
                    <button class="btn btn-success" onclick="exportExcel()">Export Excel</button>
                    <button class="btn btn-outline" onclick="viewRemoteAnalysis()">Remote Work Analysis</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Add Department Modal -->
    <div id="dept-modal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>Add Department</h3>
                <button class="close-btn" onclick="closeDeptModal()">&times;</button>
            </div>
            <div class="form-group">
                <label>Department Name</label>
                <input type="text" id="dept_name" placeholder="e.g., Engineering, Sales, HR">
            </div>
            <h4 style="margin: 20px 0 15px; color: #555;">Workspace Counts</h4>
            <div class="form-grid">
                <div class="form-group">
                    <label>Open Workstations (48 SF)</label>
                    <input type="number" id="dept_open_ws" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Standard Workstations (64 SF)</label>
                    <input type="number" id="dept_std_ws" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Large Workstations (80 SF)</label>
                    <input type="number" id="dept_lg_ws" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Small Offices (100 SF)</label>
                    <input type="number" id="dept_sm_office" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Standard Offices (150 SF)</label>
                    <input type="number" id="dept_std_office" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Large Offices (200 SF)</label>
                    <input type="number" id="dept_lg_office" min="0" value="0">
                </div>
                <div class="form-group">
                    <label>Executive Offices (300 SF)</label>
                    <input type="number" id="dept_exec_office" min="0" value="0">
                </div>
            </div>
            <div class="btn-group">
                <button class="btn btn-primary" onclick="saveDepartment()">Add Department</button>
                <button class="btn btn-outline" onclick="closeDeptModal()">Cancel</button>
            </div>
        </div>
    </div>

    <!-- Remote Analysis Modal -->
    <div id="remote-modal" class="modal">
        <div class="modal-content" style="max-width: 800px;">
            <div class="modal-header">
                <h3>Remote Work Impact Analysis</h3>
                <button class="close-btn" onclick="closeRemoteModal()">&times;</button>
            </div>
            <div id="remote-content"></div>
        </div>
    </div>

    <script>
        let departments = [];
        let currentProgram = {};

        // Tab switching
        document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                tab.classList.add('active');
                document.getElementById(tab.dataset.tab).classList.add('active');
            });
        });

        // Load saved data on page load
        window.onload = function() {
            fetch('/api/load')
                .then(r => r.json())
                .then(data => {
                    if (data.program) {
                        currentProgram = data.program;
                        loadFormData(data.program);
                    }
                });
        };

        function loadFormData(p) {
            document.getElementById('company_name').value = p.company_name || '';
            document.getElementById('location').value = p.location || '';
            document.getElementById('project_name').value = p.project_name || '';
            document.getElementById('prepared_by').value = p.prepared_by || '';
            document.getElementById('notes').value = p.notes || '';
            document.getElementById('circulation_factor').value = (p.circulation_factor || 0.35) * 100;
            document.getElementById('loss_factor').value = (p.loss_factor || 0.15) * 100;
            document.getElementById('remote_work_policy').value = p.remote_work_policy || 'full_onsite';
            
            departments = p.departments || [];
            renderDepartments();
            
            if (p.support_spaces) {
                Object.keys(p.support_spaces).forEach(key => {
                    const el = document.getElementById(key);
                    if (el) el.value = p.support_spaces[key];
                });
            }
        }

        function saveCompanyInfo() {
            const data = {
                company_name: document.getElementById('company_name').value,
                location: document.getElementById('location').value,
                project_name: document.getElementById('project_name').value,
                prepared_by: document.getElementById('prepared_by').value,
                notes: document.getElementById('notes').value
            };
            fetch('/api/save-company', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(data)
            }).then(() => {
                document.querySelector('[data-tab="departments"]').click();
            });
        }

        function openDeptModal() {
            document.getElementById('dept-modal').classList.add('active');
            document.getElementById('dept_name').value = '';
            ['dept_open_ws','dept_std_ws','dept_lg_ws','dept_sm_office','dept_std_office','dept_lg_office','dept_exec_office']
                .forEach(id => document.getElementById(id).value = 0);
        }

        function closeDeptModal() {
            document.getElementById('dept-modal').classList.remove('active');
        }

        function saveDepartment() {
            const dept = {
                name: document.getElementById('dept_name').value,
                open_workstations: parseInt(document.getElementById('dept_open_ws').value) || 0,
                standard_workstations: parseInt(document.getElementById('dept_std_ws').value) || 0,
                large_workstations: parseInt(document.getElementById('dept_lg_ws').value) || 0,
                small_offices: parseInt(document.getElementById('dept_sm_office').value) || 0,
                standard_offices: parseInt(document.getElementById('dept_std_office').value) || 0,
                large_offices: parseInt(document.getElementById('dept_lg_office').value) || 0,
                executive_offices: parseInt(document.getElementById('dept_exec_office').value) || 0
            };
            
            if (!dept.name) {
                alert('Please enter a department name');
                return;
            }
            
            departments.push(dept);
            renderDepartments();
            closeDeptModal();
            
            fetch('/api/save-departments', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({departments: departments})
            });
        }

        function deleteDepartment(index) {
            departments.splice(index, 1);
            renderDepartments();
            fetch('/api/save-departments', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({departments: departments})
            });
        }

        function renderDepartments() {
            const list = document.getElementById('department-list');
            if (departments.length === 0) {
                list.innerHTML = '<div class="alert alert-info">No departments added yet. Click "Add Department" to begin.</div>';
                return;
            }
            
            list.innerHTML = departments.map((d, i) => {
                const total = d.open_workstations + d.standard_workstations + d.large_workstations + 
                             d.small_offices + d.standard_offices + d.large_offices + d.executive_offices;
                return `
                    <div class="department-item">
                        <div class="info">
                            <div class="name">${d.name}</div>
                            <div class="details">${total} staff</div>
                        </div>
                        <button class="btn btn-danger" onclick="deleteDepartment(${i})">Remove</button>
                    </div>
                `;
            }).join('');
        }

        function saveSupportSpaces() {
            const support = {
                small_conference: parseInt(document.getElementById('small_conference').value) || 0,
                medium_conference: parseInt(document.getElementById('medium_conference').value) || 0,
                large_conference: parseInt(document.getElementById('large_conference').value) || 0,
                huddle_rooms: parseInt(document.getElementById('huddle_rooms').value) || 0,
                phone_booths: parseInt(document.getElementById('phone_booths').value) || 0,
                break_rooms: parseInt(document.getElementById('break_rooms').value) || 0,
                reception_areas: parseInt(document.getElementById('reception_areas').value) || 0,
                copy_print_centers: parseInt(document.getElementById('copy_print_centers').value) || 0,
                storage_rooms: parseInt(document.getElementById('storage_rooms').value) || 0,
                server_rooms: parseInt(document.getElementById('server_rooms').value) || 0,
                wellness_rooms: parseInt(document.getElementById('wellness_rooms').value) || 0,
                training_rooms: parseInt(document.getElementById('training_rooms').value) || 0,
                collaboration_areas: parseInt(document.getElementById('collaboration_areas').value) || 0
            };
            
            fetch('/api/save-support', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(support)
            }).then(() => {
                document.querySelector('[data-tab="factors"]').click();
            });
        }

        function saveFactors() {
            const data = {
                circulation_factor: parseFloat(document.getElementById('circulation_factor').value) / 100,
                loss_factor: parseFloat(document.getElementById('loss_factor').value) / 100,
                remote_work_policy: document.getElementById('remote_work_policy').value
            };
            
            fetch('/api/save-factors', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(data)
            }).then(() => {
                document.querySelector('[data-tab="results"]').click();
                calculateResults();
            });
        }

        function calculateResults() {
            fetch('/api/calculate')
                .then(r => r.json())
                .then(data => {
                    if (data.error) {
                        document.getElementById('results-content').innerHTML = 
                            `<div class="alert alert-info">${data.error}</div>`;
                        return;
                    }
                    renderResults(data);
                });
        }

        function renderResults(data) {
            const t = data.totals;
            const m = data.metrics;
            
            const html = `
                <div class="results-grid">
                    <div class="metric-card">
                        <div class="value">${t.total_staff.toLocaleString()}</div>
                        <div class="label">Total Headcount</div>
                    </div>
                    <div class="metric-card">
                        <div class="value">${Math.round(t.net_assignable_sf).toLocaleString()}</div>
                        <div class="label">Net Assignable SF</div>
                    </div>
                    <div class="metric-card">
                        <div class="value">${Math.round(t.usable_sf).toLocaleString()}</div>
                        <div class="label">Usable SF</div>
                    </div>
                    <div class="metric-card highlight">
                        <div class="value">${Math.round(t.rentable_sf).toLocaleString()}</div>
                        <div class="label">Rentable SF</div>
                    </div>
                </div>
                
                <table>
                    <tr><th>Component</th><th class="number">Square Feet</th></tr>
                    <tr><td>Department Space</td><td class="number">${Math.round(t.department_sf).toLocaleString()}</td></tr>
                    <tr><td>Support Space</td><td class="number">${Math.round(t.support_sf).toLocaleString()}</td></tr>
                    <tr><td>Net Assignable SF</td><td class="number">${Math.round(t.net_assignable_sf).toLocaleString()}</td></tr>
                    <tr><td>Circulation (${(t.circulation_factor*100).toFixed(0)}%)</td><td class="number">${Math.round(t.circulation_sf).toLocaleString()}</td></tr>
                    <tr><td>Usable SF</td><td class="number">${Math.round(t.usable_sf).toLocaleString()}</td></tr>
                    <tr><td>Remote Work Adjustment (${t.remote_work_description})</td><td class="number">-${Math.round(t.usable_sf - t.adjusted_usable_sf).toLocaleString()}</td></tr>
                    <tr><td>Adjusted Usable SF</td><td class="number">${Math.round(t.adjusted_usable_sf).toLocaleString()}</td></tr>
                    <tr><td>Loss Factor (${(t.loss_factor*100).toFixed(0)}%)</td><td class="number">${Math.round(t.loss_sf).toLocaleString()}</td></tr>
                    <tr class="total-row"><td>RENTABLE SQUARE FEET</td><td class="number">${Math.round(t.rentable_sf).toLocaleString()}</td></tr>
                </table>
                
                <h4 style="margin: 25px 0 15px;">Key Metrics</h4>
                <table>
                    <tr><th>Metric</th><th class="number">SF/Person</th></tr>
                    <tr><td>Net Assignable</td><td class="number">${m.sf_per_person_net.toFixed(1)}</td></tr>
                    <tr><td>Usable</td><td class="number">${m.sf_per_person_usable.toFixed(1)}</td></tr>
                    <tr><td>Adjusted (with Remote)</td><td class="number">${m.sf_per_person_adjusted.toFixed(1)}</td></tr>
                    <tr><td>Rentable</td><td class="number">${m.sf_per_person_rentable.toFixed(1)}</td></tr>
                </table>
            `;
            
            document.getElementById('results-content').innerHTML = html;
        }

        function viewRemoteAnalysis() {
            fetch('/api/remote-analysis')
                .then(r => r.json())
                .then(data => {
                    if (data.error) {
                        alert(data.error);
                        return;
                    }
                    
                    const currentPolicy = document.getElementById('remote_work_policy').value;
                    
                    let html = `
                        <p><strong>Base Rentable SF (Full On-Site):</strong> ${Math.round(data.base_rentable_sf).toLocaleString()} SF</p>
                        <table class="scenario-table" style="margin-top: 15px;">
                            <tr>
                                <th>Policy</th>
                                <th>Description</th>
                                <th class="number">Reduction</th>
                                <th class="number">Rentable SF</th>
                                <th class="number">SF Saved</th>
                            </tr>
                    `;
                    
                    data.scenarios.forEach(s => {
                        const isCurrent = s.policy === currentPolicy;
                        html += `
                            <tr class="${isCurrent ? 'current' : ''}">
                                <td>${s.policy.replace(/_/g, ' ').replace(/\\b\\w/g, l => l.toUpperCase())}</td>
                                <td>${s.description}</td>
                                <td class="number">${s.percent_reduction.toFixed(0)}%</td>
                                <td class="number">${Math.round(s.adjusted_rentable_sf).toLocaleString()}</td>
                                <td class="number">${Math.round(s.rentable_sf_saved).toLocaleString()}</td>
                            </tr>
                        `;
                    });
                    
                    html += '</table><h4 style="margin: 25px 0 15px;">Recommendations</h4>';
                    data.recommendations.forEach(r => {
                        html += `<p><strong>${r.category}:</strong> ${r.text}</p>`;
                    });
                    
                    document.getElementById('remote-content').innerHTML = html;
                    document.getElementById('remote-modal').classList.add('active');
                });
        }

        function closeRemoteModal() {
            document.getElementById('remote-modal').classList.remove('active');
        }

        function exportPDF() {
            window.location.href = '/api/export/pdf';
        }

        function exportExcel() {
            window.location.href = '/api/export/excel';
        }
    </script>
</body>
</html>
'''


@app.route('/')
def index():
    return HTML_TEMPLATE


@app.route('/api/load')
def api_load():
    program = get_program_from_session()
    from dataclasses import asdict
    return jsonify({
        'program': {
            'company_name': program.company_name,
            'location': program.location,
            'project_name': program.project_name,
            'prepared_by': program.prepared_by,
            'notes': program.notes,
            'departments': [asdict(d) for d in program.departments],
            'support_spaces': asdict(program.support_spaces),
            'circulation_factor': program.circulation_factor,
            'loss_factor': program.loss_factor,
            'remote_work_policy': program.remote_work_policy,
        }
    })


@app.route('/api/save-company', methods=['POST'])
def api_save_company():
    data = request.json
    program = get_program_from_session()
    program.company_name = data.get('company_name', '')
    program.location = data.get('location', '')
    program.project_name = data.get('project_name', '')
    program.prepared_by = data.get('prepared_by', '')
    program.notes = data.get('notes', '')
    save_program_to_session(program)
    return jsonify({'status': 'ok'})


@app.route('/api/save-departments', methods=['POST'])
def api_save_departments():
    data = request.json
    program = get_program_from_session()
    program.departments = [Department(**d) for d in data.get('departments', [])]
    save_program_to_session(program)
    return jsonify({'status': 'ok'})


@app.route('/api/save-support', methods=['POST'])
def api_save_support():
    data = request.json
    program = get_program_from_session()
    program.support_spaces = SupportSpaces(**data)
    save_program_to_session(program)
    return jsonify({'status': 'ok'})


@app.route('/api/save-factors', methods=['POST'])
def api_save_factors():
    data = request.json
    program = get_program_from_session()
    program.circulation_factor = data.get('circulation_factor', 0.35)
    program.loss_factor = data.get('loss_factor', 0.15)
    program.remote_work_policy = data.get('remote_work_policy', 'full_onsite')
    save_program_to_session(program)
    return jsonify({'status': 'ok'})


@app.route('/api/calculate')
def api_calculate():
    program = get_program_from_session()
    if not program.departments:
        return jsonify({'error': 'Please add at least one department first.'})
    
    calculator = SpaceCalculator(program)
    results = calculator.calculate_totals()
    return jsonify(results)


@app.route('/api/remote-analysis')
def api_remote_analysis():
    program = get_program_from_session()
    if not program.departments:
        return jsonify({'error': 'Please add at least one department first.'})
    
    analyzer = RemoteWorkAnalyzer(program)
    analysis = analyzer.analyze_scenarios()
    return jsonify(analysis)


@app.route('/api/export/pdf')
def api_export_pdf():
    program = get_program_from_session()
    if not program.departments:
        return "No data to export", 400
    
    calculator = SpaceCalculator(program)
    results = calculator.calculate_totals()
    analyzer = RemoteWorkAnalyzer(program)
    analysis = analyzer.analyze_scenarios()
    
    safe_name = "".join(c if c.isalnum() else "_" for c in (program.company_name or "Space_Program"))
    filename = f"{safe_name}_Space_Program.pdf"
    filepath = OUTPUT_DIR / filename
    
    export_to_pdf(results, analysis, program, str(filepath))
    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route('/api/export/excel')
def api_export_excel():
    program = get_program_from_session()
    if not program.departments:
        return "No data to export", 400
    
    calculator = SpaceCalculator(program)
    results = calculator.calculate_totals()
    analyzer = RemoteWorkAnalyzer(program)
    analysis = analyzer.analyze_scenarios()
    
    safe_name = "".join(c if c.isalnum() else "_" for c in (program.company_name or "Space_Program"))
    filename = f"{safe_name}_Space_Program.xlsx"
    filepath = OUTPUT_DIR / filename
    
    export_to_excel(results, analysis, program, str(filepath))
    return send_file(filepath, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    print("\n" + "="*50)
    print("SPACE PROGRAMMING WEB APPLICATION")
    print("="*50)
    print("\nStarting server...")
    print("Open your browser to: http://localhost:5000")
    print("\nPress Ctrl+C to stop\n")
    app.run(debug=True, port=5000)
