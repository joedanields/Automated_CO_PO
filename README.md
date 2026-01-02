# CO-PO Attainment Sheet Generator

## 📋 Project Overview

An automated system to generate **Course Outcome (CO) and Program Outcome (PO) Attainment Sheets** according to **Anna University standards** for multiple regulations (R17, R21, R24). The system processes multiple evaluation sheets and generates consolidated attainment reports using pre-defined Excel templates with built-in formulas.

---

## 🎯 Problem Statement

Academic institutions need to:
- Generate CO-PO attainment reports for accreditation (NAAC, NBA)
- Process multiple evaluation sheets (IA1, IA2, Model, Lab, Project)
- Follow different calculation rules for different regulations (R17, R21, R24)
- Maintain consistency across multiple course types (Theory, Analytical, Lab, Project)
- Handle separate requirements for Department courses vs Science & Humanities courses

**Manual Process Issues:**
- Time-consuming data entry
- Error-prone formula calculations
- Inconsistent formatting across departments
- Difficult to maintain different regulation standards

---

## 💡 Solution Approach

### Core Concept
**Use Excel Templates as "Formula Containers"** - Similar to how we reference image files in a project, we maintain Excel templates in a separate directory and use them as blueprints for attainment generation.

**Key Innovation:**
- Templates contain **all CO-PO calculation formulas** pre-configured
- Code only **extracts marks** from evaluation sheets and **fills them into templates**
- Formulas **auto-calculate** attainment percentages
- No need to replicate complex Anna University formulas in code

---

## 📁 Project Structure

```
CO PO Att Proj/
│
├── Attainment_Template/          # Template files with formulas
│   ├── Reg_17/
│   │   ├── Dept THEORY template_ R17 V3 AtSheet.xlsx
│   │   ├── Dept THEORY Analytical template_R17 V3 AtSheet.xlsx
│   │   ├── S&H THEORY template _R17 V3 AtSheet.xlsx
│   │   ├── S&H THEORY template Analytical_R17 V3 AtSheet.xlsx
│   │   ├── LAB template_R17 V3 AtSheet.xlsx
│   │   └── Project template_R17 V3 AtSheet.xlsx
│   │
│   ├── Reg_21/
│   │   └── [Similar templates for R21]
│   │
│   └── Reg_24/
│       └── [Similar templates for R24]
│
├── sample/                        # Sample input/output for testing
│   ├── input_R17/
│   │   ├── theory_eval/
│   │   │   ├── Dept_theory/      # IA1, IA2, Model sheets
│   │   │   └── S&H_theory/       # IA1, IA2, Model sheets
│   │   ├── analytical_eval/
│   │   │   └── S&H_analytical/   # IA1, IA2, Model sheets
│   │   ├── lab_eval/             # Single lab eval sheet
│   │   └── proj_eval/            # Review1, Review2, Review3 sheets
│   │
│   └── output_R17/               # Generated attainment sheets
│
├── uploads/                       # Temporary storage for user uploads
├── outputs/                       # Generated attainment files
├── utils/                         # Core logic modules
│   ├── excel_handler.py          # Excel read/write with formula preservation
│   ├── data_parser.py            # Extract marks from eval sheets
│   ├── validator.py              # Validate eval sheet consistency
│   └── template_mapper.py        # Map regulation+category to templates
│
├── templates/                     # Flask HTML templates
│   ├── index.html                # Upload interface
│   └── result.html               # Download page
│
├── app.py                         # Flask application
├── analyze_files.py              # Analysis utility script
├── requirements.txt              # Python dependencies
└── README.md                      # This file
```

---

## 🔄 Complete Workflow

### User Journey

```
┌─────────────────────────────────────────────────────────────┐
│  STEP 1: Select Regulation                                  │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐                     │
│  │  R17    │  │  R21    │  │  R24    │                     │
│  └─────────┘  └─────────┘  └─────────┘                     │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 2: Select Course Category                             │
│  ┌─────────┐  ┌────────────┐  ┌──────┐  ┌─────────┐       │
│  │ Theory  │  │ Analytical │  │ Lab  │  │ Project │       │
│  └─────────┘  └────────────┘  └──────┘  └─────────┘       │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 3: Select Department Type (Theory/Analytical only)    │
│  ┌───────────────────┐  ┌──────────────────────────┐       │
│  │  Department (Dept) │  │  Science & Humanities (S&H) │   │
│  └───────────────────┘  └──────────────────────────┘       │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 4: Upload Evaluation Sheets                           │
│                                                              │
│  For Theory/Analytical:                                      │
│    - IA1 Eval Sheet (.xlsx/.csv)                            │
│    - IA2 Eval Sheet (.xlsx/.csv)                            │
│    - Model Exam Eval Sheet (.xlsx/.csv)                     │
│                                                              │
│  For Lab:                                                    │
│    - Lab Eval Sheet (.xlsx/.csv)                            │
│                                                              │
│  For Project:                                                │
│    - Review 1 Eval Sheet (.xlsx/.csv)                       │
│    - Review 2 Eval Sheet (.xlsx/.csv)                       │
│    - Review 3 Eval Sheet (.xlsx/.csv)                       │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 5: System Validation                                  │
│  ✓ Course Code matches across all sheets                    │
│  ✓ Course Name matches across all sheets                    │
│  ✓ Faculty Name matches across all sheets                   │
│  ✓ Academic Year matches across all sheets                  │
│  ✓ Regulation matches across all sheets                     │
│  ✓ Department matches across all sheets                     │
│                                                              │
│  ❌ If validation fails → Show error and reject             │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 6: Template Selection                                 │
│  System loads: Attainment_Template/{regulation}/{type}.xlsx │
│                                                              │
│  Example: Reg_17/Dept THEORY template_ R17 V3 AtSheet.xlsx │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 7: Data Extraction & Mapping                          │
│  ┌──────────────┐                                           │
│  │  IA1 Sheet   │ → Extract: Student Reg No, Name, CO Marks │
│  │  IA2 Sheet   │ → Extract: Student Reg No, Name, CO Marks │
│  │  Model Sheet │ → Extract: Student Reg No, Name, CO Marks │
│  └──────────────┘                                           │
│                                                              │
│  Match students across sheets using: Reg No + Name          │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 8: Template Filling                                   │
│  1. Copy template to new file                               │
│  2. Fill student data (Reg No, Name)                        │
│  3. Fill CO marks from all eval sheets                      │
│  4. Formulas auto-calculate CO/PO attainment                │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  STEP 9: Generate Output                                    │
│  📥 Download: {CourseCode}_{CourseName}_Attainment.xlsx     │
│                                                              │
│  Format: Excel (.xlsx) with all formulas preserved          │
└─────────────────────────────────────────────────────────────┘
```

---

## 📊 Evaluation Sheet Structure

### Theory/Analytical Eval Sheet Format

**Example: IA1 Evaluation Sheet**

```
Row 1:  | SHEET INFO : CO EVALUATION SHEET                               |
Row 2:  | Course Code : C211                                              |
Row 3:  | Course Name : COMPUTER ARCHITECTURE                             |
Row 4:  | Faculty Name : ANANTHI M                                        |
Row 5:  | Academic Year : 2020-2021 (EVEN)                                |
Row 6:  | Class : B.TECH.IT (2ND YEAR)                                    |
Row 7:  | Regulation : R2017 - AUC                                        |
Row 8:  | Total No of Students : 62                                       |
Row 9:  | ASSESSMENT NAME : INTERNAL ASSESSMENT-1                         |
Row 10: | [Empty]                                                         |
Row 11: | [Empty] | [Empty] | QUESTION/ASSESSMENT NO | 1|2|3|4|5|6|7|8 | CO | CO | TOTAL |
Row 12: | [Empty] | [Empty] | COURSE OUTCOME NO      | 1|1|1|2|2|1|2|1 | 1  | 2  | TM    |
Row 13: | S.NO | REG. NO | NAME | MARKS | 2|2|2|2|2|16|16|8 | 30 | 20 | 50    |
Row 14: | 1 | 711719205002 | ADITHYA R | 2|2|1|1|1|16|10|5 | 26 | 12 | 38    |
Row 15: | 2 | 711719205003 | AGALYA R  | 2|2|2|2|2|12|16|7 | 25 | 20 | 45    |
...
```

**Key Data Points:**
- **Validation Fields**: Rows 2-7 (Course Code, Name, Faculty, Year, Regulation)
- **CO Mapping**: Row 12 (which questions map to which CO)
- **Student Data**: Row 14 onwards
- **Pre-calculated CO Totals**: Present in columns (no need to calculate from questions)

**CO Coverage by Assessment:**
- **IA1**: Covers CO1, CO2
- **IA2**: Covers CO3, CO4
- **Model**: Covers CO5 (and/or all COs depending on template)

---

## 🧮 CO-PO Calculation Logic

### How Templates Work

Templates contain **pre-defined formulas** that:
1. Calculate **CO Attainment %** based on student marks
2. Map **CO to PO** using Anna University correlation matrix
3. Calculate **final PO Attainment %** for accreditation

**Example Formula Flow:**
```
Student Marks (from eval) 
    ↓
CO1 Attainment = (Average of CO1 marks / Max CO1 marks) × 100
    ↓
PO1 Attainment = Weighted average of (CO1 × correlation factor)
    ↓
Final PO Attainment % (shown in template)
```

**Our Code's Responsibility:**
- ✅ Extract marks from eval sheets
- ✅ Fill marks into template cells
- ❌ **NOT** calculate formulas (templates do this automatically)

---

## 🗂️ Template Categories

### Regulation 17 Templates

| Category | Department Type | Template File | Input Required |
|----------|----------------|---------------|----------------|
| Theory | Department | `Dept THEORY template_ R17 V3 AtSheet.xlsx` | IA1, IA2, Model |
| Theory | S&H | `S&H THEORY template _R17 V3 AtSheet.xlsx` | IA1, IA2, Model |
| Analytical | Department | `Dept THEORY Analytical template_R17 V3 AtSheet.xlsx` | IA1, IA2, Model |
| Analytical | S&H | `S&H THEORY template Analytical_R17 V3 AtSheet.xlsx` | IA1, IA2, Model |
| Lab | N/A | `LAB template_R17 V3 AtSheet.xlsx` | Lab Eval |
| Project | N/A | `Project template_R17 V3 AtSheet.xlsx` | Review1, Review2, Review3 |

### Regulation 21 & 24 Templates

**Difference from R17:**
- **IA1, IA2, Integrated** (instead of IA1, IA2, Model)
- Different formula calculations
- Different CO-PO mapping matrices
- Separate lab evaluation structure

---

## 🔍 Validation Rules

All uploaded evaluation sheets **must match** on these fields:

| Field | Location in Eval Sheet | Validation Rule |
|-------|------------------------|-----------------|
| Course Code | Row 2, Column C | Must be identical across all sheets |
| Course Name | Row 3, Column C | Must be identical across all sheets |
| Faculty Name | Row 4, Column C | Must be identical across all sheets |
| Academic Year | Row 5, Column C | Must be identical across all sheets |
| Regulation | Row 7, Column C | Must match selected regulation |
| Department | Inferred from Row 6 | Must be consistent across sheets |

**If validation fails:**
- Show error message with mismatched fields
- Highlight which sheets have discrepancies
- Reject processing until fixed

---

## 🛠️ Technical Architecture

### Tech Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| Backend | **Python 3.9+** | Core logic, data processing |
| Web Framework | **Flask** | File upload interface, routing |
| Excel Handling | **openpyxl** | Read/write Excel with formula preservation |
| Data Processing | **pandas** | Parse CSV, data manipulation |
| File Storage | **Local Filesystem** | Store templates, uploads, outputs |
| Frontend | **HTML/CSS/JavaScript** | User interface |

### Core Modules

#### 1. **excel_handler.py**
```python
- load_template(regulation, category, dept_type)
  → Loads correct template from Attainment_Template/
  
- copy_template(template_path, output_path)
  → Creates copy while preserving formulas
  
- fill_student_data(workbook, student_data)
  → Fills student reg no, names, CO marks into template
  
- save_with_formulas(workbook, output_path)
  → Saves file with formulas intact (not values)
```

#### 2. **data_parser.py**
```python
- extract_validation_fields(eval_sheet)
  → Gets course code, name, faculty, year, regulation
  
- extract_student_data(eval_sheet)
  → Returns {reg_no: {name, co1, co2, co3, co4, co5}}
  
- merge_eval_data(ia1_data, ia2_data, model_data)
  → Combines marks from multiple evaluations per student
```

#### 3. **validator.py**
```python
- validate_consistency(eval_sheets_list)
  → Checks all validation fields match
  
- validate_student_match(eval_sheets_list)
  → Ensures same students across all sheets
  
- validate_marks_range(eval_sheet)
  → Checks marks are within valid limits
```

#### 4. **template_mapper.py**
```python
- get_template_path(regulation, category, dept_type)
  → Returns path to correct template file
  
- get_required_inputs(regulation, category)
  → Returns list of required eval sheets (IA1, IA2, Model, etc.)
```

---

## 🚀 Implementation Flow (Code Level)

### Main Processing Pipeline

```python
def generate_attainment(regulation, category, dept_type, uploaded_files):
    """
    Main function to generate attainment sheet
    """
    # Step 1: Validate uploaded files
    validation_result = validator.validate_consistency(uploaded_files)
    if not validation_result.is_valid:
        return {"error": validation_result.error_message}
    
    # Step 2: Get correct template
    template_path = template_mapper.get_template_path(
        regulation, category, dept_type
    )
    template = excel_handler.load_template(template_path)
    
    # Step 3: Parse all eval sheets
    student_data = {}
    for eval_file in uploaded_files:
        parsed_data = data_parser.extract_student_data(eval_file)
        student_data = data_parser.merge_data(student_data, parsed_data)
    
    # Step 4: Fill template with data
    output_workbook = excel_handler.copy_template(template)
    excel_handler.fill_student_data(output_workbook, student_data)
    
    # Step 5: Save output
    output_filename = f"{course_code}_{course_name}_Attainment.xlsx"
    output_path = f"outputs/{output_filename}"
    excel_handler.save_with_formulas(output_workbook, output_path)
    
    return {"success": True, "file": output_path}
```

---

## 📝 Key Design Decisions

### 1. **Why Separate Templates for S&H and Dept?**
- Science & Humanities courses have **different CO-PO mapping rules**
- Department courses have **different attainment thresholds**
- Anna University mandates different calculation methods

### 2. **Why Extract CO Totals Instead of Question-Wise Marks?**
- CO totals are **already calculated in eval sheets**
- Avoids reimplementing CO mapping logic in code
- Reduces errors from mismatched question-CO mappings
- Simpler and more maintainable

### 3. **Why Use openpyxl with data_only=False?**
- Preserves **formulas** (not just calculated values)
- When template is filled and opened in Excel, formulas auto-calculate
- No need to replicate complex Anna University formulas in Python
- Templates can be updated without code changes

### 4. **Why Local Filesystem Instead of Database?**
- Templates and outputs are **Excel files** (binary)
- No need for complex querying
- Simpler backup and version control
- Easy for non-technical users to update templates

---

## 🎯 Advantages of This Approach

✅ **Template-Driven Design**
- All calculation logic stays in Excel templates
- Code only handles data extraction and mapping
- Easy for faculty to update formulas without touching code

✅ **Regulation Flexibility**
- Adding new regulation = adding new template folder
- No code changes needed for formula updates
- Each regulation's rules isolated in its templates

✅ **Validation First**
- Ensures data consistency before processing
- Clear error messages for mismatched data
- Prevents garbage output

✅ **Scalability**
- Can handle multiple courses simultaneously
- Parallel processing possible (independent files)
- Local filesystem = no cloud costs

✅ **Maintainability**
- Clear separation of concerns
- Each module has single responsibility
- Easy to debug and test

---

## 🔮 Future Enhancements

### Phase 2 Features
- [ ] Batch processing (multiple courses at once)
- [ ] Email notification when attainment ready
- [ ] History/logs of generated attainments
- [ ] Preview before final generation
- [ ] Support for custom regulations

### Phase 3 Features
- [ ] PDF export of attainment sheets
- [ ] Dashboard showing CO-PO trends
- [ ] Comparison across semesters
- [ ] Cloud storage integration (Google Drive, OneDrive)
- [ ] Role-based access (Faculty, HOD, Principal)

---

## 📌 Important Notes

### For Users
1. **Eval sheets must follow exact format** (Row 2 = Course Code, etc.)
2. **All eval sheets must have matching metadata**
3. **Student Reg Numbers must be consistent** across all sheets
4. **Do not modify template files** in Attainment_Template/ directory

### For Developers
1. **Always use openpyxl with data_only=False** to preserve formulas
2. **Never hardcode cell positions** - make them configurable
3. **Validate before processing** - fail fast with clear errors
4. **Copy templates before filling** - never modify originals
5. **Log all operations** for debugging and audit trail

---

## 📖 Glossary

| Term | Full Form | Description |
|------|-----------|-------------|
| CO | Course Outcome | What students should learn from a course |
| PO | Program Outcome | Overall program goals/objectives |
| IA | Internal Assessment | Mid-semester exams (IA1, IA2) |
| Model | Model Exam | Pre-final exam before semester end |
| R17/R21/R24 | Regulation 2017/2021/2024 | Anna University curriculum versions |
| S&H | Science & Humanities | Non-core subjects (Math, Physics, English) |
| Dept | Department | Core technical subjects |
| Analytical | Analytical Course | Math-heavy courses requiring formula sheets |

---

## 🤝 Contributing

### Development Setup
```bash
# Clone repository
git clone <repo-url>
cd CO_PO_Att_Proj

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run Flask app
python app.py
```

### Testing
- Place sample eval sheets in `sample/input_R17/`
- Run generation process
- Compare output with `sample/output_R17/` expected results

---

## 📄 License

[Add your license here]

---

## 👥 Authors

[Add your name/team here]

---

**Last Updated:** January 2, 2026
