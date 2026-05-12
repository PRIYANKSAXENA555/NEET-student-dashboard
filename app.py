import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="NEET Student Performance Dashboard",
    page_icon="🩺",
    layout="wide"
)

st.title("🩺 NEET Student Performance Dashboard")
st.markdown("---")

# ============================================================
# FILE PATH - Put your Excel file in the same directory
# ============================================================
EXCEL_FILE_PATH = "NEET 2027 ALL RESULT.xlsx"

import os
if not os.path.exists(EXCEL_FILE_PATH):
    st.error(f"❌ Excel file not found at: {EXCEL_FILE_PATH}")
    st.info("Please upload the NEET 2027 ALL RESULT.xlsx file to the same directory as this app.")
    st.stop()

# ============================================================
# NEET CONFIGURATION
# ============================================================
# NEET 2027: Physics=180, Chemistry=180, Biology=360, Total=720
NEET_CONFIG = {
    "max_phy": 180,
    "max_chem": 180,
    "max_bio": 360,
    "max_total": 720,
    "subject_cols": {
        "phy": "PHY",
        "chem": "CHEM", 
        "bio": "BIO TOTAL"
    }
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def detect_test_type(sheet_name):
    """Detect test type from sheet name"""
    sheet_upper = sheet_name.upper()
    if "BRTEST" in sheet_upper:
        return "BRTEST"
    elif "GRAND" in sheet_upper:
        return "GRAND TEST"
    return "BTEST"

def normalize_name(name):
    """Normalize student names for consistent matching"""
    if pd.isna(name):
        return None
    name = str(name).upper().strip()
    # Remove extra spaces
    name = re.sub(r'\s+', ' ', name)
    # Fix common name variations
    replacements = {
        'ANTRA SHEGDAR': 'ANTRA SHEGDAR',
        'ANTARA SHEGDAR': 'ANTRA SHEGDAR',
        'ABHINAY HANNURE': 'ABHINAV HANNURE',
        'ABHINAV HANNURE': 'ABHINAV HANNURE',
        'SRUSHTI HARISHCHANDRA TIPE': 'SHRUSHTI HARISHCHANDRA TIPE',
        'SHRUSHTI TIPE': 'SHRUSHTI HARISHCHANDRA TIPE',
        'SHREEYA VIRENDRA BOLABATTIN': 'SHREEYA VIRENDRA BOLABATTIN',
        'SHREEYA BOLABATTIN': 'SHREEYA VIRENDRA BOLABATTIN',
        'AKSHATA HIPPARGE': 'AKSHATA BHARATRAJ HIPPARGE',
        'AKSHATA BHARATRAJ HIPPARGE': 'AKSHATA BHARATRAJ HIPPARGE',
        'VIRAJ DEVARNADGI': 'VIRAJ LAXMIKANT DEVARANADAGI',
        'VIRAJ LAXMIKANT DEVARANADAGI': 'VIRAJ LAXMIKANT DEVARANADAGI',
        'PRATHMESH SHINDE': 'PRATHMESH AMOL SHINDE',
        'PRATHMESH AMOL SHINDE': 'PRATHMESH AMOL SHINDE',
        'SRUSHTI HARISHCHANDRA TIPE': 'SHRUSHTI HARISHCHANDRA TIPE',
        'ANTRA SHEGDAR': 'ANTRA SHEGDAR',
        'SHRIVARDHAN BURA': 'SHRIVARDHAN BURA',
        'ABHIJIT A JADHAV': 'ABHIJIT ANNASAHEB JADHAV',
        'ABHIJIT ANNASAHEB JADHAV': 'ABHIJIT ANNASAHEB JADHAV',
        'SRUSHTI TIPE': 'SHRUSHTI HARISHCHANDRA TIPE',
        'SHWETA HASARMANI': 'SHWETA GURUBASSAPA HASARAMANI',
        'SHWETA GURUBASSAPA HASARAMANI': 'SHWETA GURUBASSAPA HASARAMANI',
        'OYES NADAF': 'OWAIS SIKANDAR NADAF',
        'OWAIS SIKANDAR NADAF': 'OWAIS SIKANDAR NADAF',
        'PALLAVI BISWAS': 'PALLABI ASHIS BISWAS',
        'SAKSHI JIDGE': 'SAKSHI SACHIN JIDGE',
        'AMULYA BOMDYAL': 'AMULYA SATISH BOMDYAL',
        'AMULYA SATISH BOMDYAL': 'AMULYA SATISH BOMDYAL',
        'SHRADDHA DUDHANIKAR': 'SHRADDHA SOMSHEKHAR DUDHANIKAR',
        'KANISHKA DHULAM': 'KANISHKA SUNIL DHULAM',
        'DEVAYANI DAREKAR': 'DEVYANI DAREKAR',
        'ANWESHA MANJUNATH CHINMALLI': 'ANWESHA MANJUNATH CHINMALLI',
        'NAGACHAITANYA SHRIRAM': 'NAGACHAITANYA SHRIRAM',
        'RITHIKA DYAWARKONDA': 'RITHIKA RENURAJA DYAWARKONDA',
    }
    name = replacements.get(name, name)
    if len(name) < 3:
        return None
    return name

def get_weakest_subject_neet(phy_rank, chem_rank, bio_rank):
    """Determine weakest subject based on rank (lower rank = better performance)"""
    if pd.isna(phy_rank) or pd.isna(chem_rank) or pd.isna(bio_rank):
        return "Absent"
    
    p, c, b = float(phy_rank), float(chem_rank), float(bio_rank)
    
    # If ranks are very close, consider balanced
    if max(p, c, b) - min(p, c, b) <= 30:
        return "Balanced"
    
    ranks = {'Physics': p, 'Chemistry': c, 'Biology': b}
    weakest = max(ranks, key=ranks.get)  # Higher rank number = worse performance
    other_avg = (sum(ranks.values()) - ranks[weakest]) / 2
    
    # Significant weakness if rank is more than 1.5x average of other subjects
    if ranks[weakest] > 1.5 * other_avg:
        return weakest
    return "Balanced"

@st.cache_data
def load_excel_data():
    """Load all NEET test data from Excel file"""
    all_student_data = {}
    test_metadata = {}
    
    try:
        xl = pd.ExcelFile(EXCEL_FILE_PATH)
        
        # Get all test sheets
        test_sheets = [s for s in xl.sheet_names if not s.upper().startswith("MASTER")]
        
        for sheet in test_sheets:
            test_type = detect_test_type(sheet)
            
            # Store metadata
            test_metadata[sheet] = {
                "type": test_type,
                "max_phy": NEET_CONFIG["max_phy"],
                "max_chem": NEET_CONFIG["max_chem"],
                "max_bio": NEET_CONFIG["max_bio"],
                "max_total": NEET_CONFIG["max_total"]
            }
            
            try:
                # Read the sheet
                df = pd.read_excel(xl, sheet_name=sheet)
                
                # Find the data start row (look for TOTAL RANK or STUDENT NAME)
                data_start_row = 0
                for idx, row in df.iterrows():
                    row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
                    if 'TOTAL RANK' in row_str.upper() or ('STUDENT NAME' in row_str.upper() and 'PHY' in row_str.upper()):
                        data_start_row = idx
                        break
                
                # Reload with correct start row
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=data_start_row)
                
                # Determine column indices based on headers
                headers = df.iloc[0].astype(str).str.upper().tolist() if len(df) > 0 else []
                
                # Find column indices
                col_map = {
                    'overall_rank': None,
                    'name': None,
                    'phy': None,
                    'phy_rank': None,
                    'chem': None,
                    'chem_rank': None,
                    'bio': None,
                    'bio_rank': None,
                    'total': None,
                    'branch': None
                }
                
                for i, h in enumerate(headers):
                    h_clean = str(h).upper().strip()
                    if 'TOTAL RANK' in h_clean and col_map['overall_rank'] is None:
                        col_map['overall_rank'] = i
                    elif 'STUDENT NAME' in h_clean and col_map['name'] is None:
                        col_map['name'] = i
                    elif h_clean == 'PHY' and col_map['phy'] is None:
                        col_map['phy'] = i
                    elif 'PHY RANK' in h_clean and col_map['phy_rank'] is None:
                        col_map['phy_rank'] = i
                    elif h_clean == 'CHEM' and col_map['chem'] is None:
                        col_map['chem'] = i
                    elif 'CHEM RANK' in h_clean and col_map['chem_rank'] is None:
                        col_map['chem_rank'] = i
                    elif 'BIO TOTAL' in h_clean and col_map['bio'] is None:
                        col_map['bio'] = i
                    elif 'BIO RANK' in h_clean and col_map['bio_rank'] is None:
                        col_map['bio_rank'] = i
                    elif h_clean == 'TOTAL' and col_map['total'] is None:
                        col_map['total'] = i
                    elif h_clean == 'BRANCH' and col_map['branch'] is None:
                        col_map['branch'] = i
                
                # Process rows
                for idx in range(1, len(df)):
                    row = df.iloc[idx]
                    
                    # Check if row has student data
                    name_val = row.iloc[col_map['name']] if col_map['name'] is not None else None
                    if pd.isna(name_val):
                        continue
                    
                    # Skip summary rows
                    name_str = str(name_val).upper()
                    if any(keyword in name_str for keyword in ['MEDIAN', 'MAX', 'TOP', 'STUDENT WHOSE', 'ABSENT']):
                        continue
                    
                    student_name = normalize_name(name_val)
                    if not student_name:
                        continue
                    
                    # Get marks
                    total = pd.to_numeric(row.iloc[col_map['total']] if col_map['total'] is not None else None, errors='coerce')
                    overall_rank = pd.to_numeric(row.iloc[col_map['overall_rank']] if col_map['overall_rank'] is not None else None, errors='coerce')
                    
                    if pd.isna(total) or total < 0:
                        continue
                    
                    # Initialize student if not exists
                    if student_name not in all_student_data:
                        all_student_data[student_name] = {
                            "name": student_name,
                            "branch": None,
                            "tests": {}
                        }
                    
                    # Get branch
                    if col_map['branch'] is not None and pd.notna(row.iloc[col_map['branch']]):
                        branch = str(row.iloc[col_map['branch']]).strip()
                        if branch and branch != 'nan':
                            all_student_data[student_name]["branch"] = branch
                    
                    # Get subject marks
                    phy = pd.to_numeric(row.iloc[col_map['phy']] if col_map['phy'] is not None else None, errors='coerce')
                    chem = pd.to_numeric(row.iloc[col_map['chem']] if col_map['chem'] is not None else None, errors='coerce')
                    bio = pd.to_numeric(row.iloc[col_map['bio']] if col_map['bio'] is not None else None, errors='coerce')
                    phy_rank = pd.to_numeric(row.iloc[col_map['phy_rank']] if col_map['phy_rank'] is not None else None, errors='coerce')
                    chem_rank = pd.to_numeric(row.iloc[col_map['chem_rank']] if col_map['chem_rank'] is not None else None, errors='coerce')
                    bio_rank = pd.to_numeric(row.iloc[col_map['bio_rank']] if col_map['bio_rank'] is not None else None, errors='coerce')
                    
                    # Store test data
                    all_student_data[student_name]["tests"][sheet] = {
                        "phy": phy if pd.notna(phy) else 0,
                        "chem": chem if pd.notna(chem) else 0,
                        "bio": bio if pd.notna(bio) else 0,
                        "phy_rank": phy_rank if pd.notna(phy_rank) else None,
                        "chem_rank": chem_rank if pd.notna(chem_rank) else None,
                        "bio_rank": bio_rank if pd.notna(bio_rank) else None,
                        "total": total,
                        "overall_rank": overall_rank if pd.notna(overall_rank) else None,
                        "type": test_type
                    }
                    
            except Exception as e:
                st.warning(f"Error reading sheet {sheet}: {str(e)}")
                continue
        
        return all_student_data, test_metadata
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None, None

# ============================================================
# LOAD DATA
# ============================================================
with st.spinner("🔄 Loading NEET student data..."):
    all_student_data, test_metadata = load_excel_data()

if all_student_data and len(all_student_data) > 0:
    st.success(f"✅ Loaded {len(all_student_data)} students and {len(test_metadata)} tests!")
    
    # Student selection
    student_options = sorted(all_student_data.keys())
    selected_student = st.selectbox("🎓 Select Student", student_options)
    
    if selected_student:
        student = all_student_data[selected_student]
        
        btest_results = []
        grand_test_results = []
        brtest_results = []
        
        # Track best/worst ranks
        all_overall_ranks = []
        all_phy_ranks = []
        all_chem_ranks = []
        all_bio_ranks = []
        
        for sheet, meta in test_metadata.items():
            if sheet not in student["tests"]:
                continue
            
            marks = student["tests"][sheet]
            pct = round((marks["total"] / meta["max_total"]) * 100, 1)
            
            overall_rank = marks.get("overall_rank")
            
            # Calculate rank if not available
            if overall_rank is None or pd.isna(overall_rank):
                all_scores = []
                for s_data in all_student_data.values():
                    if sheet in s_data["tests"]:
                        all_scores.append(s_data["tests"][sheet]["total"])
                overall_rank = sum(score > marks["total"] for score in all_scores) + 1
            else:
                overall_rank = int(overall_rank)
            
            all_overall_ranks.append(overall_rank)
            
            if marks.get("phy_rank") is not None and not pd.isna(marks.get("phy_rank")):
                all_phy_ranks.append(int(marks.get("phy_rank")))
            if marks.get("chem_rank") is not None and not pd.isna(marks.get("chem_rank")):
                all_chem_ranks.append(int(marks.get("chem_rank")))
            if marks.get("bio_rank") is not None and not pd.isna(marks.get("bio_rank")):
                all_bio_ranks.append(int(marks.get("bio_rank")))
            
            weakest = get_weakest_subject_neet(
                marks.get("phy_rank"), 
                marks.get("chem_rank"), 
                marks.get("bio_rank")
            )
            
            phy_pct = round((marks["phy"] / meta["max_phy"]) * 100, 1) if marks["phy"] > 0 else 0
            chem_pct = round((marks["chem"] / meta["max_chem"]) * 100, 1) if marks["chem"] > 0 else 0
            bio_pct = round((marks["bio"] / meta["max_bio"]) * 100, 1) if marks["bio"] > 0 else 0
            
            # Extract test number for sorting
            test_num = 0
            match = re.search(r'(\d+)', sheet)
            if match:
                test_num = int(match.group(1))
            
            result = {
                "S.No.": test_num,
                "Test Name": sheet,
                "Type": meta["type"],
                "Physics": marks["phy"],
                "Physics %": phy_pct,
                "Phy Rank": marks.get("phy_rank") if marks.get("phy_rank") is not None else '-',
                "Chemistry": marks["chem"],
                "Chemistry %": chem_pct,
                "Chem Rank": marks.get("chem_rank") if marks.get("chem_rank") is not None else '-',
                "Biology": marks["bio"],
                "Biology %": bio_pct,
                "Bio Rank": marks.get("bio_rank") if marks.get("bio_rank") is not None else '-',
                "Total": f"{marks['total']:.0f}/{meta['max_total']}",
                "%": f"{pct}%",
                "Overall Rank": overall_rank,
                "Weakest Subject": weakest
            }
            
            if meta["type"] == "BTEST":
                btest_results.append(result)
            elif meta["type"] == "GRAND TEST":
                grand_test_results.append(result)
            else:
                brtest_results.append(result)
        
        # Sort by test number
        btest_results.sort(key=lambda x: x["S.No."])
        grand_test_results.sort(key=lambda x: x["S.No."])
        brtest_results.sort(key=lambda x: x["S.No."])
        
        if not btest_results and not grand_test_results and not brtest_results:
            st.warning("No tests found for this student")
        else:
            all_tests = btest_results + grand_test_results + brtest_results
            avg_pct = round(np.mean([float(t["%"].replace("%", "")) for t in all_tests]), 1)
            
            best_overall_rank = min(all_overall_ranks) if all_overall_ranks else 'N/A'
            worst_overall_rank = max(all_overall_ranks) if all_overall_ranks else 'N/A'
            best_phy_rank = min(all_phy_ranks) if all_phy_ranks else 'N/A'
            worst_phy_rank = max(all_phy_ranks) if all_phy_ranks else 'N/A'
            best_chem_rank = min(all_chem_ranks) if all_chem_ranks else 'N/A'
            worst_chem_rank = max(all_chem_ranks) if all_chem_ranks else 'N/A'
            best_bio_rank = min(all_bio_ranks) if all_bio_ranks else 'N/A'
            worst_bio_rank = max(all_bio_ranks) if all_bio_ranks else 'N/A'
            
            weak_subjects = [t["Weakest Subject"] for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
            weak_count = Counter(weak_subjects)
            
            # ====================================================
            # STUDENT INFO HEADER
            # ====================================================
            branch_info = f" | Branch: {student['branch']}" if student.get('branch') else ""
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#0b3b3b,#1a5a5a,#2d7a7a);
                        padding:20px;
                        border-radius:15px;
                        margin-bottom:20px;">
                <h2 style="color:white; margin:0;">🩺 {selected_student}{branch_info}</h2>
                <p style="color:white; margin:5px 0 0 0;">NEET 2027 Aspirant | Physics:180 | Chemistry:180 | Biology:360 | Total:720</p>
            </div>
            """, unsafe_allow_html=True)
            
            # ====================================================
            # TEST SUMMARY METRICS
            # ====================================================
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📚 Tests Attempted", f"{len(all_tests)}/{len(test_metadata)}")
            with col2:
                st.metric("📊 Average Score", f"{avg_pct}%")
            with col3:
                st.metric("🏆 Best Overall Rank", best_overall_rank)
            with col4:
                st.metric("⚠️ Worst Overall Rank", worst_overall_rank)
            
            st.markdown("---")
            
            # ====================================================
            # SUBJECT RANK SUMMARY
            # ====================================================
            st.subheader("📊 Subject-wise Rank Summary (Lower is Better)")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.info("**🔬 Physics (180 marks)**")
                st.metric("Best Rank", best_phy_rank, delta="⭐")
                st.metric("Worst Rank", worst_phy_rank, delta="⚠️")
            
            with col2:
                st.info("**⚗️ Chemistry (180 marks)**")
                st.metric("Best Rank", best_chem_rank, delta="⭐")
                st.metric("Worst Rank", worst_chem_rank, delta="⚠️")
            
            with col3:
                st.info("**🧬 Biology (360 marks)**")
                st.metric("Best Rank", best_bio_rank, delta="⭐")
                st.metric("Worst Rank", worst_bio_rank, delta="⚠️")
            
            st.markdown("---")
            
            # ====================================================
            # WEAK SUBJECT SUMMARY
            # ====================================================
            if weak_count:
                st.subheader("⚠️ Weak Subject Analysis")
                weak_cols = st.columns(len(weak_count))
                for i, (subject, count) in enumerate(weak_count.items()):
                    with weak_cols[i]:
                        st.warning(f"⚠️ {subject}")
                        st.metric("Weak in", f"{count} test(s)")
                st.markdown("---")
            
            # ====================================================
            # BTEST/Grand TEST RESULTS TABLE
            # ====================================================
            regular_tests = btest_results + grand_test_results
            if regular_tests:
                with st.expander("📘 BTEST & GRAND TESTS (Full NEET Format - 720 marks)", expanded=True):
                    display_cols = ['Test Name', 'Type', 'Physics', 'Phy Rank', 'Chemistry', 'Chem Rank', 
                                   'Biology', 'Bio Rank', 'Total', '%', 'Overall Rank', 'Weakest Subject']
                    regular_df = pd.DataFrame(regular_tests)
                    st.dataframe(regular_df[display_cols], use_container_width=True)
            
            # ====================================================
            # BRTEST RESULTS TABLE
            # ====================================================
            if brtest_results:
                with st.expander("📘 BRTEST TESTS (Revision Tests)", expanded=True):
                    display_cols = ['Test Name', 'Physics', 'Phy Rank', 'Chemistry', 'Chem Rank', 
                                   'Biology', 'Bio Rank', 'Total', '%', 'Overall Rank', 'Weakest Subject']
                    brtest_df = pd.DataFrame(brtest_results)
                    st.dataframe(brtest_df[display_cols], use_container_width=True)
            
            st.markdown("---")
            
            # ====================================================
            # SUBJECT MARKS TRENDS - Regular Tests
            # ====================================================
            if regular_tests:
                st.subheader("📊 Subject Marks Trends - BTEST & GRAND TESTS")
                
                test_names = [t['Test Name'][:25] for t in regular_tests]
                max_marks = [NEET_CONFIG["max_phy"], NEET_CONFIG["max_chem"], NEET_CONFIG["max_bio"]]
                
                fig_marks = go.Figure()
                fig_marks.add_trace(go.Scatter(x=test_names, y=[t['Physics'] for t in regular_tests], 
                                                mode='lines+markers', name='Physics (max 180)', 
                                                line=dict(color='#3498DB', width=2), marker=dict(size=8)))
                fig_marks.add_trace(go.Scatter(x=test_names, y=[t['Chemistry'] for t in regular_tests], 
                                                mode='lines+markers', name='Chemistry (max 180)', 
                                                line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
                fig_marks.add_trace(go.Scatter(x=test_names, y=[t['Biology'] for t in regular_tests], 
                                                mode='lines+markers', name='Biology (max 360)', 
                                                line=dict(color='#2ECC71', width=2), marker=dict(size=8)))
                
                # Add target lines (70% marks)
                fig_marks.add_hline(y=126, line_dash="dash", line_color="#3498DB", opacity=0.5, annotation_text="Physics 70% (126)")
                fig_marks.add_hline(y=126, line_dash="dash", line_color="#9B59B6", opacity=0.5, annotation_text="Chemistry 70% (126)")
                fig_marks.add_hline(y=252, line_dash="dash", line_color="#2ECC71", opacity=0.5, annotation_text="Biology 70% (252)")
                
                fig_marks.update_layout(title="Subject Marks Across Tests", height=450, hovermode='x unified')
                st.plotly_chart(fig_marks, use_container_width=True)
            
            # ====================================================
            # SUBJECT RANK TRENDS - Regular Tests
            # ====================================================
            if regular_tests:
                st.subheader("🏆 Subject Rank Trends - Regular Tests (Lower is Better)")
                
                fig_ranks = go.Figure()
                test_names_rank = [t['Test Name'][:25] for t in regular_tests]
                phy_ranks = [t['Phy Rank'] if t['Phy Rank'] != '-' else None for t in regular_tests]
                chem_ranks = [t['Chem Rank'] if t['Chem Rank'] != '-' else None for t in regular_tests]
                bio_ranks = [t['Bio Rank'] if t['Bio Rank'] != '-' else None for t in regular_tests]
                
                fig_ranks.add_trace(go.Scatter(x=test_names_rank, y=phy_ranks, 
                                                mode='lines+markers', name='Physics Rank',
                                                line=dict(color='#3498DB', width=2), marker=dict(size=8)))
                fig_ranks.add_trace(go.Scatter(x=test_names_rank, y=chem_ranks, 
                                                mode='lines+markers', name='Chemistry Rank',
                                                line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
                fig_ranks.add_trace(go.Scatter(x=test_names_rank, y=bio_ranks, 
                                                mode='lines+markers', name='Biology Rank',
                                                line=dict(color='#2ECC71', width=2), marker=dict(size=8)))
                
                fig_ranks.update_layout(title="Subject Rank Comparison (Lower is Better)", 
                                        yaxis=dict(autorange="reversed"), height=400)
                st.plotly_chart(fig_ranks, use_container_width=True)
            
            # ====================================================
            # OVERALL RANK TREND - Regular Tests
            # ====================================================
            if regular_tests:
                st.subheader("🏆 Overall Rank Trend - Regular Tests (Lower is Better)")
                
                fig_rank = go.Figure()
                overall_ranks = [t['Overall Rank'] for t in regular_tests]
                
                fig_rank.add_trace(go.Scatter(x=test_names_rank, y=overall_ranks,
                                               mode='lines+markers', name='Overall Rank',
                                               line=dict(color='#E74C3C', width=3), marker=dict(size=10)))
                best_rank = min(overall_ranks) if overall_ranks else None
                if best_rank:
                    fig_rank.add_hline(y=best_rank, line_dash="dash", line_color="green", 
                                      annotation_text=f"Best Rank: {best_rank}")
                
                fig_rank.update_layout(title="Overall Rank Performance (Lower is Better)", 
                                       yaxis=dict(autorange="reversed"), height=400)
                st.plotly_chart(fig_rank, use_container_width=True)
            
            # ====================================================
            # BRTEST SUBJECT MARKS TRENDS
            # ====================================================
            if brtest_results:
                st.subheader("📊 Subject Marks Trends - BRTEST TESTS (Revision Tests)")
                
                brtest_names = [t['Test Name'][:25] for t in brtest_results]
                
                fig_marks_br = go.Figure()
                fig_marks_br.add_trace(go.Scatter(x=brtest_names, y=[t['Physics'] for t in brtest_results], 
                                                   mode='lines+markers', name='Physics (max 180)', 
                                                   line=dict(color='#3498DB', width=2), marker=dict(size=8)))
                fig_marks_br.add_trace(go.Scatter(x=brtest_names, y=[t['Chemistry'] for t in brtest_results], 
                                                   mode='lines+markers', name='Chemistry (max 180)', 
                                                   line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
                fig_marks_br.add_trace(go.Scatter(x=brtest_names, y=[t['Biology'] for t in brtest_results], 
                                                   mode='lines+markers', name='Biology (max 360)', 
                                                   line=dict(color='#2ECC71', width=2), marker=dict(size=8)))
                
                fig_marks_br.add_hline(y=126, line_dash="dash", line_color="#3498DB", opacity=0.5)
                fig_marks_br.add_hline(y=126, line_dash="dash", line_color="#9B59B6", opacity=0.5)
                fig_marks_br.add_hline(y=252, line_dash="dash", line_color="#2ECC71", opacity=0.5)
                
                fig_marks_br.update_layout(title="Subject Marks in Revision Tests (BRTEST)", height=400)
                st.plotly_chart(fig_marks_br, use_container_width=True)
            
            # ====================================================
            # OVERALL PERCENTAGE TREND (All Tests)
            # ====================================================
            st.subheader("📈 Overall Percentage Trend")
            
            all_test_names = [t['Test Name'][:25] for t in all_tests]
            all_pcts = [float(t["%"].replace("%", "")) for t in all_tests]
            
            # Color code by test type
            colors = []
            for t in all_tests:
                if t['Type'] == 'BTEST':
                    colors.append('#3498DB')
                elif t['Type'] == 'GRAND TEST':
                    colors.append('#9B59B6')
                else:
                    colors.append('#E67E22')
            
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(
                x=all_test_names,
                y=all_pcts,
                mode='lines+markers',
                name='Percentage Score',
                line=dict(color='#1ABC9C', width=3),
                marker=dict(size=10, color=colors, symbol='circle')
            ))
            
            # Add target lines
            fig_trend.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="Target (75%)")
            fig_trend.add_hline(y=85, line_dash="dot", line_color="orange", annotation_text="Excellence (85%)")
            
            fig_trend.update_layout(
                title="Percentage Score Across All NEET Tests",
                height=450,
                xaxis=dict(tickangle=45),
                hovermode='x unified'
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            
            # ====================================================
            # DETAILED WEAKNESS INSIGHTS
            # ====================================================
            st.subheader("📊 Detailed Weakness Analysis for NEET")
            
            subject_weakness = {'Physics': 0, 'Chemistry': 0, 'Biology': 0}
            total_weak = len(weak_subjects)
            
            for ws in weak_subjects:
                if ws in subject_weakness:
                    subject_weakness[ws] += 1
            
            if total_weak > 0:
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Subject-wise weakness breakdown:**")
                    for subject, count in sorted(subject_weakness.items(), key=lambda x: x[1], reverse=True):
                        pct = (count / total_weak) * 100
                        st.write(f"• {subject}: {count} times ({pct:.0f}%)")
                        st.progress(int(pct))
                
                with col2:
                    st.write("**🔍 NEET Strategy Recommendations:**")
                    weak_tests_list = [t for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
                    if weak_tests_list:
                        st.write("**Focus Areas:**")
                        for subject, count in sorted(subject_weakness.items(), key=lambda x: x[1], reverse=True):
                            if count > 0:
                                if subject == 'Physics':
                                    st.info(f"• {subject}: Focus on Numericals & Problem Solving")
                                elif subject == 'Chemistry':
                                    st.info(f"• {subject}: Focus on Organic/Inorganic Reactions & NCERT")
                                else:
                                    st.info(f"• {subject}: Focus on NCERT Line-by-Line & Diagrams")
            
            # Balanced performance
            balanced_tests_list = [t for t in all_tests if t["Weakest Subject"] == "Balanced"]
            if balanced_tests_list:
                st.write(f"**✅ Balanced performance in {len(balanced_tests_list)} tests:**")
                for t in balanced_tests_list[:5]:
                    st.write(f"• {t['Test Name'][:50]}")
                if len(balanced_tests_list) > 5:
                    st.write(f"... and {len(balanced_tests_list) - 5} more")
            
            # Missing tests
            attempted = set([t['Test Name'] for t in all_tests])
            missing = [t for t in test_metadata.keys() if t not in attempted]
            if missing:
                st.write(f"**⚠️ ABSENT/NO DATA for {len(missing)} tests:**")
                for m in missing[:10]:
                    st.write(f"• {m} ({test_metadata[m]['type']})")
            
            # ====================================================
            # NEET PERFORMANCE METRICS
            # ====================================================
            st.markdown("---")
            st.subheader("🎯 NEET Performance Metrics")
            
            # Calculate moving average and improvement
            if len(all_pcts) >= 3:
                col1, col2 = st.columns(2)
                
                with col1:
                    # Recent trend
                    recent_avg = np.mean(all_pcts[-3:]) if len(all_pcts) >= 3 else avg_pct
                    st.metric("Last 3 Tests Average", f"{recent_avg:.1f}%")
                    
                    # Improvement from first to last
                    if len(all_pcts) >= 2 and all_pcts[0] > 0:
                        improvement = all_pcts[-1] - all_pcts[0]
                        delta_color = "normal" if improvement >= 0 else "inverse"
                        st.metric("Overall Improvement", f"{improvement:+.1f}%", delta=f"{improvement:+.1f}%")
                
                with col2:
                    # Consistency (standard deviation)
                    std_dev = np.std(all_pcts)
                    consistency = "High" if std_dev < 10 else "Medium" if std_dev < 15 else "Low"
                    st.metric("Consistency Score", consistency, delta=f"σ = {std_dev:.1f}")
                    
                    # Best performance
                    best_pct = max(all_pcts)
                    st.metric("Best Performance", f"{best_pct:.1f}%")
            
            st.markdown("---")
            st.caption("✅ Dashboard Complete | Data Source: NEET 2027 Result Sheets | Rank Analysis: Lower number = Better performance")
            
else:
    st.error("❌ No student data found in the file.")