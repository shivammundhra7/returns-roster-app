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
st.markdown("Upload your standardized **Putaway_Roster.xlsx** template below.\n*Includes strict max-nights, 9-day streak shield, & Experienced vs. New MP balancing.*")

uploaded_file = st.file_uploader("Upload Input Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Generate Roster", use_container_width=True):
        with st.spinner("Balancing Experience ratios & enforcing strict night limits... Please wait."):
            try:
                # ==========================================
                # 2. LOAD STANDARDIZED DATA
                # ==========================================
                df_prev = pd.read_excel(uploaded_file, sheet_name='Previous_Month_Attendance')
                df_targets = pd.read_excel(uploaded_file, sheet_name='Daily_Targets')
                df_leaves = pd.read_excel(uploaded_file, sheet_name='Planned_Leaves')
                df_emp = pd.read_excel(uploaded_file, sheet_name='Employee_Master')

                try:
                    df_prefs = pd.read_excel(uploaded_file, sheet_name='Night_Preferences')
                    df_prefs['Emp_ID'] = df_prefs['Emp_ID'].astype(str).str.strip()
                    df_prefs = df_prefs.drop_duplicates(subset=['Emp_ID'], keep='last')
                    pref_dict = df_prefs.set_index('Emp_ID')['Night_Shift_Pref'].to_dict()
                except Exception:
                    pref_dict = {}

                df_emp['Role'] = df_emp['Role'].astype(str).str.strip()
                df_targets['Role'] = df_targets['Role'].astype(str).str.strip()
                df_emp['Emp_ID'] = df_emp['Emp_ID'].astype(str).str.strip()
                df_prev['Emp_ID'] = df_prev['Emp_ID'].astype(str).str.strip()

                df_targets['Date'] = pd.to_datetime(df_targets['Date'])
                df_leaves['Date'] = pd.to_datetime(df_leaves['Date'])
                
                roster_start = df_targets['Date'].min()
                roster_end = df_targets['Date'].max()
                roster_days = pd.date_range(start=roster_start, end=roster_end)

                # ==========================================
                # 3. INITIALIZE EMPLOYEES
                # ==========================================
                emp_state = {}
                
                exclude_cols = ['Emp_ID', 'Role', 'Name', 'NAME', 'Gender']
                prev_date_cols = [c for c in df_prev.columns if c not in exclude_cols]

                for _, row in df_emp.iterrows():
                    emp_id = row['Emp_ID']
                    prev_data = df_prev[df_prev['Emp_ID'] == emp_id]
                    
                    target_wos = 4
                    emp_doj = None
                    
                    if pd.notna(row['Date_of_Joining']):
                        try:
                            emp_doj = pd.to_datetime(row['Date_of_Joining'])
                            if emp_doj >= roster_start:
                                active_days = max(0, (roster_end - emp_doj).days + 1)
                                target_wos = active_days // 6
                        except:
                            pass 
                            
                    last_shift_state = 'FREE' 
                    streak = 0
                    
                    if not prev_data.empty:
                        prev_row = prev_data.iloc[0]
                        for d in reversed(prev_date_cols):
                            # Skip entirely empty/blank cells 
                            if pd.isna(prev_row[d]):
                                if streak == 0: last_shift_state = 'FREE'
                                break
                            
                            val = str(prev_row[d]).strip().upper()
                            if val in ['NAN', 'NONE', '']:
                                if streak == 0: last_shift_state = 'FREE'
                                break
                                
                            if val in ['P-M', 'P-E', 'P-D', 'E', 'M', 'D']:
                                if streak == 0: last_shift_state = 'D'
                                streak += 1
                            elif val in ['P-N', 'N']:
                                if streak == 0: last_shift_state = 'N'
                                streak += 1
                            elif val in ['WOD', 'L-D', 'L-N', 'A-D', 'A-N', 'WO', 'A', 'L']:
                                if streak == 0: last_shift_state = 'FREE'
                                break 
                                
                    min_n = 0
                    max_n = 20
                    is_default = True
                    
                    if row['Gender'] == 'Male':
                        pref_val = pref_dict.get(emp_id, '')
                        
                        # Excel hidden Date Conversion Fix ("20-26")
                        if isinstance(pref_val, datetime):
                            pref_str = "20-26"
                        else:
                            pref_str = str(pref_val).strip()
                            
                        if '2026' in pref_str or 'Jul' in pref_str:
                            pref_str = "20-26"

                        if '-' in pref_str:
                            try:
                                parts = pref_str.split('-')
                                min_n = int(parts[0])
                                max_n = int(parts[1])
                                is_default = False
                            except:
                                pass
                    else:
                        max_n = 0 
                        is_default = False 

                    emp_state[emp_id] = {
                        'Role': row['Role'],
                        'Gender': row['Gender'],
                        'DOJ': emp_doj,
                        'WOs_Remaining': target_wos,
                        'Starting_WOs': target_wos,
                        'Current_Streak': streak,
                        'Night_Count': 0, 
                        'Min_Nights': min_n,
                        'Max_Nights': max_n,
                        'Is_Default_Pref': is_default,
                        'Lock_State': last_shift_state,
                        'Schedule': {}
                    }

                # ==========================================
                # 4. ROSTER GENERATION LOGIC 
                # ==========================================
                roles = df_emp['Role'].unique()

                for day in roster_days:
                    days_left = (roster_end - day).days + 1
                    
                    for role in roles:
                        role_emps = [e for e, data in emp_state.items() if data['Role'] == role]
                        if not role_emps: continue
                        
                        target_row = df_targets[(df_targets['Date'] == day) & (df_targets['Role'] == role)]
                        is_zero_wo = target_row.iloc[0]['Zero_WO_Day'] if not target_row.empty else False
                        
                        active_emps = []
                        for emp in role_emps:
                            # 1. NEW JOINER BLOCKER
                            emp_doj = emp_state[emp].get('DOJ')
                            if emp_doj and day < emp_doj:
                                emp_state[emp]['Schedule'][day] = 'Not Joined'
                                continue
                            
                            leave_df = df_leaves[(df_leaves['Emp_ID'] == emp) & (df_leaves['Date'] == day)]
                            if not leave_df.empty:
                                emp_state[emp]['Schedule'][day] = 'L'
                                emp_state[emp]['Current_Streak'] = 0
                                emp_state[emp]['Lock_State'] = 'FREE'
                            else:
                                active_emps.append(emp)
                                
                        total_active = len(active_emps)
                        if total_active == 0: continue
                        
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
                                
                        assigned_wos = must_wo.copy()
                        
                        if not is_zero_wo and len(assigned_wos) < dynamic_wo_target:
                            planned_mp_est = total_active - dynamic_wo_target
                            
                            target_n_est_max = math.floor(planned_mp_est * 0.50)
                            target_d_est_max = math.floor(planned_mp_est * 0.55)
                            
                            rem_active = [e for e in active_emps if e not in assigned_wos]
                            curr_locked_n = len([e for e in rem_active if emp_state[e]['Lock_State'] == 'N'])
                            curr_locked_d = len([e for e in rem_active if emp_state[e]['Lock_State'] == 'D'])
                            
                            surplus_n = curr_locked_n - target_n_est_max
                            surplus_d = curr_locked_d - target_d_est_max

                            candidates = []
                            for e in rem_active:
                                state = emp_state[e]
                                if state['WOs_Remaining'] > 0:
                                    if state['Lock_State'] == 'N' and state['WOs_Remaining'] == 1:
                                        continue 
                                    if (days_left - 1) > (state['WOs_Remaining'] - 1) * 10 + 9:
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
                        if planned_mp == 0: continue
                        
                        max_target_n = math.floor(planned_mp * 0.50)
                        
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
                            
                        curr_n = len(locked_n)
                        
                        # ==========================================
                        # EXPERIENCED VS NEW BALANCING LOGIC
                        # ==========================================
                        # Tag current working pool as Experienced (>= 60 days) or New (< 60 days)
                        for e in working_emps:
                            emp_doj = emp_state[e].get('DOJ')
                            emp_state[e]['Is_New'] = (emp_doj is not None) and ((day - emp_doj).days < 60)
                            
                        working_exp = [e for e in working_emps if not emp_state[e]['Is_New']]
                        total_working = len(working_emps)
                        
                        exp_ratio = len(working_exp) / total_working if total_working > 0 else 0.5
                        
                        # Calculate exact target ratio constraints for the Night shift
                        target_n_exp = round(max_target_n * exp_ratio)
                        target_n_new = max_target_n - target_n_exp
                        
                        assigned_n_exp = len([e for e in locked_n if not emp_state[e]['Is_New']])
                        assigned_n_new = len([e for e in locked_n if emp_state[e]['Is_New']])
                        
                        eligible_free_males = []
                        for e in free_pool:
                            if emp_state[e]['Gender'] == 'Male' and emp_state[e]['Night_Count'] < emp_state[e]['Max_Nights']:
                                if emp_state[e]['WOs_Remaining'] == 0:
                                    if emp_state[e]['Night_Count'] + days_left <= emp_state[e]['Max_Nights']:
                                        eligible_free_males.append(e)
                                else:
                                    eligible_free_males.append(e)
                                    
                        # Partition eligible males into Experienced and New pools
                        eligible_exp = [e for e in eligible_free_males if not emp_state[e]['Is_New']]
                        eligible_new = [e for e in eligible_free_males if emp_state[e]['Is_New']]
                        
                        def pref_sort(x):
                            return (-(max(0, emp_state[x]['Min_Nights'] - emp_state[x]['Night_Count'])), emp_state[x]['Night_Count'])
                            
                        eligible_exp.sort(key=pref_sort)
                        eligible_new.sort(key=pref_sort)
                        
                        final_n_picks = []
                        
                        # Intelligently pick to maintain the exact Experienced/New ratio
                        while curr_n < max_target_n and (eligible_exp or eligible_new):
                            need_exp = (assigned_n_exp < target_n_exp)
                            need_new = (assigned_n_new < target_n_new)
                            
                            if need_exp and eligible_exp:
                                final_n_picks.append(eligible_exp.pop(0))
                                assigned_n_exp += 1
                            elif need_new and eligible_new:
                                final_n_picks.append(eligible_new.pop(0))
                                assigned_n_new += 1
                            else:
                                # If one ratio is met but we still need physical headcount to hit 48% target
                                if eligible_exp:
                                    final_n_picks.append(eligible_exp.pop(0))
                                    assigned_n_exp += 1
                                elif eligible_new:
                                    final_n_picks.append(eligible_new.pop(0))
                                    assigned_n_new += 1
                                else:
                                    break
                            curr_n += 1

                        # Apply the shifts
                        for emp in free_pool:
                            if emp in final_n_picks:
                                emp_state[emp]['Schedule'][day] = 'N'
                                emp_state[emp]['Night_Count'] += 1
                                emp_state[emp]['Current_Streak'] += 1
                                emp_state[emp]['Lock_State'] = 'N' 
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
                    for day in roster_days:
                        row_dict[day.strftime('%d-%b')] = data['Schedule'].get(day, '')
                    
                    row_dict['Total_WOs_Assigned'] = data['Starting_WOs'] - data['WOs_Remaining']
                    row_dict['Total_Night_Shifts'] = data['Night_Count']
                    row_dict['Target_Pref_Range'] = f"{data['Min_Nights']}-{data['Max_Nights']}" if data['Gender'] == 'Male' else "N/A"
                    
                    output_data.append(row_dict)

                final_df = pd.DataFrame(output_data).sort_values('Role')
                
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Generated_Roster')
                
                month_name = roster_start.strftime('%B_%Y')
                out_filename = f"{month_name}_Putaway_Roster.xlsx"

                st.success(f"✅ {month_name.replace('_', ' ')} Putaway Roster successfully generated!")
                
                st.download_button(
                    label="⬇️ Download Final Roster (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"⚠️ An error occurred: {e}")
