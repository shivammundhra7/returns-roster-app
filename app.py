import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import math
import io

# ==========================================
# 1. APP UI DESIGN
# ==========================================
st.set_page_config(page_title="Putaway Roster Generator", page_icon="📅", layout="centered")

st.title("📅 Automated Putaway Roster")
st.markdown("""
Welcome! Upload your populated **Putaway_Roster.xlsx** file below. 
The system will automatically calculate optimal WOs, balance Day/Night shifts, and enforce all operational constraints.
""")

st.divider()

# File Uploader
uploaded_file = st.file_uploader("Upload Input Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Generate Roster", use_container_width=True):
        with st.spinner("Crunching the numbers and balancing shifts... Please wait."):
            try:
                # ==========================================
                # 2. LOAD DATA FROM UPLOAD
                # ==========================================
                df_may = pd.read_excel(uploaded_file, sheet_name='May_Attendance')
                df_targets = pd.read_excel(uploaded_file, sheet_name='June_Daily_Targets')
                df_leaves = pd.read_excel(uploaded_file, sheet_name='June_Planned_Leaves')
                df_emp = pd.read_excel(uploaded_file, sheet_name='June_Employee_Master')

                try:
                    df_prefs = pd.read_excel(uploaded_file, sheet_name='June_Night_Preferences')
                    df_prefs['Emp_ID'] = df_prefs['Emp_ID'].astype(str).str.strip()
                    df_prefs = df_prefs.drop_duplicates(subset=['Emp_ID'], keep='last')
                    pref_dict = df_prefs.set_index('Emp_ID')['Night_Shift_Pref'].to_dict()
                except Exception:
                    pref_dict = {}

                df_emp['Role'] = df_emp['Role'].astype(str).str.strip()
                df_targets['Role'] = df_targets['Role'].astype(str).str.strip()
                df_emp['Emp_ID'] = df_emp['Emp_ID'].astype(str).str.strip()
                df_may['Emp_ID'] = df_may['Emp_ID'].astype(str).str.strip()

                df_targets['Date'] = pd.to_datetime(df_targets['Date'])
                df_leaves['Date'] = pd.to_datetime(df_leaves['Date'])
                days_in_june = pd.date_range(start='2026-06-01', end='2026-06-30')

                # ==========================================
                # 3. INITIALIZE EMPLOYEES
                # ==========================================
                emp_state = {}
                may_date_cols = [f"{d}-May" for d in range(1, 32)]

                for _, row in df_emp.iterrows():
                    emp_id = row['Emp_ID']
                    may_data = df_may[df_may['Emp_ID'] == emp_id]
                    
                    target_wos = 4
                    if pd.notna(row['Date_of_Joining']):
                        try:
                            doj = pd.to_datetime(row['Date_of_Joining'])
                            if doj >= pd.to_datetime('2026-06-01'):
                                active_days = max(0, (pd.to_datetime('2026-06-30') - doj).days + 1)
                                target_wos = active_days // 6
                        except:
                            pass 
                            
                    last_shift_state = 'FREE' 
                    streak = 0
                    
                    if not may_data.empty:
                        may_row = may_data.iloc[0]
                        for d in reversed(may_date_cols):
                            if d not in may_row: continue
                            val = str(may_row[d]).strip().upper()
                            
                            if val in ['P-M', 'P-E', 'P-D', 'E', 'M', 'D']:
                                if streak == 0: last_shift_state = 'D'
                                streak += 1
                            elif val in ['P-N', 'N']:
                                if streak == 0: last_shift_state = 'N'
                                streak += 1
                            elif val in ['WOD', 'L-D', 'L-N', 'A-D', 'A-N']:
                                if streak == 0: last_shift_state = 'FREE'
                                break 
                                
                    min_n = 0
                    max_n = 20
                    if row['Gender'] == 'Male':
                        pref_str = str(pref_dict.get(emp_id, '')).strip()
                        if '-' in pref_str:
                            try:
                                parts = pref_str.split('-')
                                min_n = int(parts[0])
                                max_n = int(parts[1])
                            except:
                                pass
                    else:
                        max_n = 0 

                    emp_state[emp_id] = {
                        'Role': row['Role'],
                        'Gender': row['Gender'],
                        'WOs_Remaining': target_wos,
                        'Starting_WOs': target_wos,
                        'Current_Streak': streak,
                        'Night_Count': 0, 
                        'Min_Nights': min_n,
                        'Max_Nights': max_n,
                        'Lock_State': last_shift_state,
                        'Schedule': {}
                    }

                # ==========================================
                # 4. ROSTER GENERATION LOGIC (V9 Engine)
                # ==========================================
                roles = df_emp['Role'].unique()

                for day in days_in_june:
                    days_left = (pd.to_datetime('2026-06-30') - day).days + 1
                    
                    for role in roles:
                        role_emps = [e for e, data in emp_state.items() if data['Role'] == role]
                        if not role_emps: continue
                        
                        target_row = df_targets[(df_targets['Date'] == day) & (df_targets['Role'] == role)]
                        is_zero_wo = target_row.iloc[0]['Zero_WO_Day'] if not target_row.empty else False
                        
                        active_emps = []
                        for emp in role_emps:
                            leave_df = df_leaves[(df_leaves['Emp_ID'] == emp) & (df_leaves['Date'] == day)]
                            if not leave_df.empty:
                                emp_state[emp]['Schedule'][day] = 'L'
                                emp_state[emp]['Current_Streak'] = 0
                                emp_state[emp]['Lock_State'] = 'FREE'
                            else:
                                active_emps.append(emp)
                                
                        total_active = len(active_emps)
                        must_wo = []
                        total_wos_remaining_for_role = sum([emp_state[e]['WOs_Remaining'] for e in active_emps])
                        
                        if days_left > 0:
                            target_run_rate = math.ceil(total_wos_remaining_for_role / days_left)
                        else:
                            target_run_rate = total_wos_remaining_for_role

                        min_wo = int(total_active * 0.08)
                        max_wo = max(1, math.ceil(total_active * 0.17)) 
                        
                        if is_zero_wo:
                            dynamic_wo_target = 0
                        else:
                            dynamic_wo_target = max(min_wo, min(target_run_rate, max_wo))

                        for emp in active_emps:
                            state = emp_state[emp]
                            if state['WOs_Remaining'] > 0:
                                if state['WOs_Remaining'] >= days_left: 
                                    must_wo.append(emp)
                                elif state['Current_Streak'] >= 9: 
                                    must_wo.append(emp)
                                elif state['Lock_State'] == 'N' and state['Night_Count'] >= state['Max_Nights']: 
                                    must_wo.append(emp)
                                
                        assigned_wos = must_wo.copy()
                        
                        if not is_zero_wo and len(assigned_wos) < dynamic_wo_target:
                            planned_mp_est = total_active - dynamic_wo_target
                            target_n_est = round(planned_mp_est * 0.48)
                            target_d_est = planned_mp_est - target_n_est
                            
                            rem_active = [e for e in active_emps if e not in assigned_wos]
                            curr_locked_n = len([e for e in rem_active if emp_state[e]['Lock_State'] == 'N'])
                            curr_locked_d = len([e for e in rem_active if emp_state[e]['Lock_State'] == 'D'])
                            
                            surplus_n = curr_locked_n - target_n_est
                            surplus_d = curr_locked_d - target_d_est

                            candidates = []
                            for e in rem_active:
                                if emp_state[e]['WOs_Remaining'] > 0:
                                    if emp_state[e]['Lock_State'] == 'N' and emp_state[e]['WOs_Remaining'] == 1:
                                        continue 
                                    candidates.append(e)
                                    
                            def wo_sort_key(x):
                                state = emp_state[x]
                                balance_boost = 0
                                if surplus_n > 0 and state['Lock_State'] == 'N':
                                    balance_boost = 10000
                                elif surplus_d > 0 and state['Lock_State'] == 'D':
                                    balance_boost = 10000
                                
                                cooldown_penalty = 0
                                if state['Current_Streak'] < 5:
                                    cooldown_penalty = -5000
                                    
                                return (balance_boost + cooldown_penalty + (state['Current_Streak'] * 10) + state['WOs_Remaining'])
                                
                            candidates.sort(key=wo_sort_key, reverse=True)
                            shortfall = dynamic_wo_target - len(assigned_wos)
                            assigned_wos.extend(candidates[:shortfall])

                        working_emps = []
                        for emp in active_emps:
                            if emp in assigned_wos:
                                emp_state[emp]['Schedule'][day] = 'WOD'
                                emp_state[emp]['WOs_Remaining'] -= 1
                                emp_state[emp]['Current_Streak'] = 0
                                emp_state[emp]['Lock_State'] = 'FREE' 
                            else:
                                working_emps.append(emp)
                                
                        planned_mp = len(working_emps)
                        target_n = round(planned_mp * 0.48) 
                        
                        locked_n = [e for e in working_emps if emp_state[e]['Lock_State'] == 'N']
                        locked_d = [e for e in working_emps if emp_state[e]['Lock_State'] == 'D']
                        free_pool = [e for e in working_emps if emp_state[e]['Lock_State'] == 'FREE']
                        
                        for emp in locked_n:
                            emp_state[emp]['Schedule'][day] = 'N'
                            emp_state[emp]['Night_Count'] += 1
                            emp_state[emp]['Current_Streak'] += 1
                            
                        for emp in locked_d:
                            emp_state[emp]['Schedule'][day] = 'D'
                            emp_state[emp]['Current_Streak'] += 1
                            
                        shortfall_n = max(0, target_n - len(locked_n))
                        
                        eligible_free_males = []
                        for e in free_pool:
                            if emp_state[e]['Gender'] == 'Male' and emp_state[e]['Night_Count'] < emp_state[e]['Max_Nights']:
                                if emp_state[e]['WOs_Remaining'] == 0:
                                    if emp_state[e]['Night_Count'] + days_left <= emp_state[e]['Max_Nights']:
                                        eligible_free_males.append(e)
                                else:
                                    eligible_free_males.append(e)
                                    
                        eligible_free_males.sort(key=lambda x: (
                            -(max(0, emp_state[x]['Min_Nights'] - emp_state[x]['Night_Count'])), 
                            emp_state[x]['Night_Count']
                        ))
                        
                        assigned_free_n = 0
                        for emp in free_pool:
                            if assigned_free_n < shortfall_n and emp in eligible_free_males:
                                emp_state[emp]['Schedule'][day] = 'N'
                                emp_state[emp]['Night_Count'] += 1
                                emp_state[emp]['Current_Streak'] += 1
                                emp_state[emp]['Lock_State'] = 'N' 
                                assigned_free_n += 1
                            else:
                                emp_state[emp]['Schedule'][day] = 'D'
                                emp_state[emp]['Current_Streak'] += 1
                                emp_state[emp]['Lock_State'] = 'D' 

                # ==========================================
                # 5. PREPARE APP DOWNLOAD
                # ==========================================
                output_data = []
                for emp_id, data in emp_state.items():
                    row_dict = {'Emp_ID': emp_id, 'Role': data['Role']}
                    for day in days_in_june:
                        row_dict[day.strftime('%d-%b')] = data['Schedule'].get(day, '')
                    
                    row_dict['Total_June_WOs'] = data['Starting_WOs'] - data['WOs_Remaining']
                    row_dict['Total_Night_Shifts'] = data['Night_Count']
                    row_dict['Target_Pref_Range'] = f"{data['Min_Nights']}-{data['Max_Nights']}" if data['Gender'] == 'Male' else "N/A"
                    
                    output_data.append(row_dict)

                final_df = pd.DataFrame(output_data).sort_values('Role')
                
                # Write to Memory (Hidden from User)
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Generated_Roster')
                
                st.success("✅ Roster successfully generated!")
                
                # Download Button
                st.download_button(
                    label="⬇️ Download Final Roster (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name="June_Final_Roster_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"⚠️ An error occurred. Please check your Excel formatting: {e}")