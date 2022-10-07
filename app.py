from cmath import nan
from requests import Session, session
import streamlit as st
import numpy as np
import pandas as pd
import altair as alt
import pulp as pu
import xlsxwriter

st.set_page_config(
    page_title="จัดตารางเรียนตารางสอน ICDI CMU",
    page_icon="📚",
)

st.header('ยินดีต้อนรับเข้าสู่โปรแกรมช่วยจัดตารางสอน ICDI CMU')
st.write("โปรแกรมนี้จัดทำโดย อ.ดร. ศรณย์เศรษฐ์ โสกันธิกา โดยมีความคาดหวังว่าจะช่วยเพิ่มประสิทธิภาพในการทำงาน และแบ่งเบาภาระของผู้ที่ต้องจัดตารางสอนให้กับวิทยาลัยฯ")
st.write("ถ้าพบปัญหาหรือต้องการเพิ่ม options สามารถติดต่อมาได้ทุกเมื่อ 😉")
st.markdown("# อัพโหลดข้อมูล")
# st.sidebar.header("อัพโหลดข้อมูล")

st.write("หลังจากที่คณาจารย์ได้ประชุมเพื่อตกลงภาระงานสอนแล้ว โปรดคลิ๊กปุ่มด้านล่างเพื่อดาวน์โหลดแม่แบบสำหรับกรอกข้อมูล")

with open('Undergraduate_1_2022.xlsx', 'rb') as my_file:
    st.download_button(
        label = '📥 ดาวน์โหลดแม่แบบ 📥', 
        data = my_file, 
        file_name = 'Undergraduate_1_2022.xlsx', 
        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')      

st.write("หลังจากที่ได้ดาวน์โหลดแม่แบบแล้ว ถ้ายังไม่เคยใช้โปรแกรมมาก่อน โปรดลองนำไฟล์แม่แบบที่ดาวน์โหลดมาอัพโหลดด้านล่างเพื่อศึกษาวิธีการใช้งานโดยภาพรวม")
st.write("เมื่อพร้อมที่จะจัดตารางสอนของโรงเรียนแล้ว โปรดอ่านคำแนะนำดังนี้")
st.write("  1. โปรดเปลี่ยนชื่อไฟล์เป็น Undergraduate_เทอม_ปีการศึกษา เช่น Undergraduate_1_2022")
st.write("  2. เลขที่ปรากฎในตารางในชีท teachers หมายถึงจำนวนสัดส่วนการสอนในวิชานั้น ๆ โดยเลขจะต้องอยู่ระหว่าง 0 ถึง 1")
st.write("  3. ในชีท students นักศึกษาของวิทยาลัยนานาชาติฯ ต้องเริ่มชื่อด้วย DIN เท่านั้น")
st.write("  4. ถ้ามีวิชาที่ต้องการระบุวันและเวลาที่สอนด้วยตนเอง โปรดกำหนดในชีท manual")
st.write("หมายเหตุ: โปรแกรมนี้ใช้สำหรับจัดตารางสอนตั้งแต่วันจันทร์ถึงศุกร์ โดยแต่ละวันกำหนดไว้ว่ามี 5 คาบ ดังต่อไปนี้ 8.00-9.30, 9.30-11.00, 11.00-12.30, 13.00-14.30, 14.30-16.00")

# st.write("  3. เลขที่ปรากฎในตารางหมายถึงจำนานคาบที่สอนในวิชาและชั้นเรียนนั้น ๆ ในหนึ่งสัปดาห์ เช่น ในชีทแรก ครูสมหมาย สอนวิชาภาษาไทย ป.4 จำนวน 4 คาบ ใน 1 สัปดาห์")


# check and stop program with error_counter
error_counter = False

uploaded_file = st.file_uploader("อัพโหลดแม่แบบ")
# upload_status = False

if uploaded_file is not None:
    # upload_status = True
    df = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
    try:
        semester = uploaded_file.name.split('_')[1]
        year = uploaded_file.name.split('_')[2]
    except:
        st.error("โปรดเปลี่ยนชื่อไฟล์เป็น **ชื่อโรงเรียน_เทอม_ปีการศึกษา** เช่น โรงเรียนบ้านห่างไกล_2_2565")
        st.stop()

    # Create Data structure
    Days = ['Mon','Tue','Wed','Thu','Fri']
    # OnDays = ['Mon','Tue','Wed','Thu','Fri']
    # OffDays = ['Sat','Sun']
    # Days = OnDays+OffDays


    
    sessions = ['8.00-9.30', '9.30-11.00', '11.00-12.30', '13.00-14.30', '14.30-16.00']
    
    num_sessions_per_day = len(sessions)
    # num_sessions_per_day = st.number_input('จำนวนคาบต่อวัน',min_value = 4, max_value = 6, value = 5, key='sessionsaday')
    
    Timeslot = [(i,j) for i in Days for j in sessions]
    # st.write(Timeslot)

    Teachers = [i for i in df['teachers'].iloc[2:,0]]

    Students = [i for i in df['students'].iloc[2:,0]]
    Dinstudents = [ i for i in Students if i[:3].upper() == "DIN"]

    MDateAndTime = [i for i in df['manual'].iloc[2:,0]]

    SDummy = df['teachers'].iloc[:2,1:]
    TSubjects = ['{} sec {}'.format(str(int(i)).zfill(6),str(int(j)).zfill(3)) if type(j) != str else '{} sec {}'.format(str(int(i)),j) for i,j in zip(SDummy.iloc[0],SDummy.iloc[1])]
    
    SDummy = df['students'].iloc[:2,1:]
    SSubjects = ['{} sec {}'.format(str(int(i)).zfill(6),str(int(j)).zfill(3)) if type(j) != str else '{} sec {}'.format(str(int(i)),j) for i,j in zip(SDummy.iloc[0],SDummy.iloc[1])]
    OutSubjects = [s for s in SSubjects if int(s[:6]) < 888000 or  int(s[:6]) > 889000]

    MDummy = df['manual'].iloc[:2,1:]
    MSubjects = ['{} sec {}'.format(str(int(i)).zfill(6),str(int(j)).zfill(3)) if type(j) != str else '{} sec {}'.format(str(int(i)),j) for i,j in zip(MDummy.iloc[0],MDummy.iloc[1])]
    
    Subjects = SSubjects
    for s in TSubjects:
        if s not in Subjects:
            Subjects.append(s)
    

    # functions used for retrive information in dataframe
    def Teacherindex(t):
        return Teachers.index(t)+2
    def Tsubjectindex(s):
        return TSubjects.index(s)+1
    def Studentindex(s):
        return Students.index(s)+2
    def Ssubjectindex(s):
        return SSubjects.index(s)+1
    def Manualindex(s):
        return MDateAndTime.index(s)+2
    def Msubjectindex(s):
        return MSubjects.index(s)+1

    #Adjusting Conditions
    st.write('**ตัวเลือกเพิ่มเติมในการจัดตารางสอน**')
    # Morning and afternoon preference
    with st.expander('กำหนดอาจารย์ที่อยากสอนช่วงเช้าหรือบ่าย'):
        morning_class = st.multiselect('อาจารย์ที่อยากสอนตอนเช้า (เรียงตามความสำคัญ)', pd.Series(Teachers),[], key='morningteacher')
        afternoon_class = st.multiselect('อาจารย์ที่อยากสอนตอนบ่าย (เรียงตามความสำคัญ)', pd.Series(Teachers),[], key='afternoonteacher')
    # Days preference
    with st.expander('กำหนดอาจารย์ที่มีวันไม่อยากสอน (เช่น อาจสอนวันเสาร์อาทิตย์แล้วไม่อยากสอนในวันจันทร์)'):
        teacher_avoid_day = st.multiselect('เลือกอาจารย์', pd.Series(Teachers),[], key='teacheravoidday')
        teacher_avoid_vars = {t:[] for t in teacher_avoid_day}
        for t in teacher_avoid_day:
            teacher_avoid_vars[t] = st.multiselect('วันที่อาจารย์{}ไม่อยากสอน'.format(t), pd.Series(Days), [], key = '{}avoidday'.format(t))

    # Subject which not teach in MTh and TuF format
    with st.expander('กำหนดวิชาที่ไม่จำเป็นต้องสอนตามรูปแบบ MTh/TuF '):
        # st.write('วิชาที่ไม่จำเป็นต้องสอนตามรูปแบบ MTh/TuF โดยไม่กำหนดเวลาด้วยตนเอง')
        # st.write('(หากต้องการกำหนดวันและเวลาเอง โปรดใส่ข้อมูลในชีท manual))')
        nonformat = st.multiselect('เลือกวิชา', pd.Series(Subjects),[], key='nonformatsubject')
        nonformat_vars = {s:[] for s in nonformat}
        for s in nonformat:
            nonformat_vars[s] = st.selectbox('จำนวนคาบของวิชา {} ใน 1 สัปดาห์'.format(s), range(1,4), index = 1, key = '{}sessionsnumbers'.format(s))

    # Actual Classes

    Subjects_sessions_num = {s:[2] for s in Subjects}

    # manual subject
    manual_plan = {s:[] for s in MSubjects}
    for s in MSubjects:
        if df['manual'].notna().iloc[Manualindex('Date1'),Msubjectindex(s)] and df['manual'].notna().iloc[Manualindex('Time1'),Msubjectindex(s)]:
            manual_plan[s].append((df['manual'].iloc[Manualindex('Date1'),Msubjectindex(s)], df['manual'].iloc[Manualindex('Time1'),Msubjectindex(s)]))
        if df['manual'].notna().iloc[Manualindex('Date2'),Msubjectindex(s)] and df['manual'].notna().iloc[Manualindex('Time2'),Msubjectindex(s)]:
            manual_plan[s].append((df['manual'].iloc[Manualindex('Date2'),Msubjectindex(s)], df['manual'].iloc[Manualindex('Time2'),Msubjectindex(s)]))
        # st.write(manual_plan[s])
        # st.write([len(manual_plan[s])])
        Subjects_sessions_num[s][0] = len(manual_plan[s])
    
    # for s in Subjects:
    #         st.write(Subjects_sessions_num[s])

    # Change session number of non format subjects
    for s in st.session_state.nonformatsubject:
        Subjects_sessions_num[s] = nonformat_vars[s]

    DinClasses = []
    for t in Teachers:
        for s in TSubjects:
            if s in SSubjects:            
                if df['teachers'].iloc[Teacherindex(t),Tsubjectindex(s)] > 0 and df['teachers'].iloc[Teacherindex(t),Tsubjectindex(s)] <= 1:
                    for i in Students:                        
                        if df['students'].iloc[Studentindex(i),Ssubjectindex(s)] > 0 and df['students'].iloc[Studentindex(i),Ssubjectindex(s)] <= 1:
                            if s not in st.session_state:
                                for k in range(Subjects_sessions_num[s][0]):
                                    DinClasses.append((t,i,s,k))
    # Check manual sheet
    Dummy = []
    for s in MSubjects:
        for c in DinClasses:
            if c[2] == s:
                for slot in manual_plan[s]:
                    if (c[0],c[3],slot) not in Dummy:
                        Dummy.append((c[0],c[3],slot))
                    else:
                        st.error("วิชา{} ของอาจารย์{} สอนตรงกับวิชาอื่น โปรดเช็คข้อมูลใน sheet manual อีกครั้ง".format(s,c[0]))
                        error_counter = True

    OutClasses = []
    for i in Students: 
        for s in OutSubjects:                               
            if df['students'].iloc[Studentindex(i),Ssubjectindex(s)] > 0 and df['students'].iloc[Studentindex(i),Ssubjectindex(s)] <= 1:
                for k in range(Subjects_sessions_num[s][0]):
                    OutClasses.append(('outside',i,s,k))

    Classes = DinClasses
    for c in OutClasses:
        if c not in Classes:
            Classes.append(c)

    # teaching plan 
    teacher_plan = {t:[] for t in Teachers}
    for t in Teachers:
        for c in Classes:
            if c[0] == t:
                teacher_plan[t].append(c)
                

    # student plan
    student_plan = {s:[] for s in Students}
    for s in Students:
        for c in Classes:
            if c[1] == s:
                student_plan[s].append(c)
    
    # subject plan
    subject_plan = {s:[] for s in SSubjects}
    for s in SSubjects:
        for c in Classes:
            if c[2] == s:
                subject_plan[s].append(c)

    if error_counter == False:
        
        #Scheduling
        st.markdown("# เริ่มการจัดตารางสอน")

  
# ''' Model formulation'''
    
        p = pu.LpProblem('ICDI Timetabling', pu.LpMinimize)

        # Add variables
        var = pu.LpVariable.dicts("ClassAtTime", (Classes,Timeslot), lowBound = 0, upBound = None, cat='Binary')

        # '''Soft Constraints'''  
        Penelty_weight = 100
        Reward = -10

        # Morning Class Preference
        Constraint_var = []
        for t in st.session_state.morningteacher:
            Dummy = []
            for j in sessions:
                Dummy.append(pu.lpSum(var[c][(i,j)] for i in Days for c in teacher_plan[t]))
            Constraint_var.append(Dummy)

        Penelty_distribution = []
        for t in st.session_state.morningteacher:
            Dummy = []
            for i in sessions:
                if sessions.index(i) <= 2:
                    Dummy.append(1/np.power(2,sessions.index(i)+st.session_state.morningteacher.index(t))*Reward)
                else:
                    Dummy.append(1/np.power(2,st.session_state.morningteacher.index(t))*np.power(Penelty_weight,sessions.index(i)))
            Penelty_distribution.append(Dummy)

        Morning_Teacher_Penelty = pu.lpSum(np.dot(Penelty_distribution[i],Constraint_var[i]) for i in range(len(st.session_state.morningteacher)))
        
        # Afternoon Class Preference
        Constraint_var = []
        for t in st.session_state.afternoonteacher:
            Dummy = []
            for j in sessions:
                Dummy.append(pu.lpSum(var[c][(i,j)] for i in Days for c in teacher_plan[t]))
            Constraint_var.append(Dummy)

        Penelty_distribution = []
        for t in st.session_state.afternoonteacher:
            Dummy = []
            for i in sessions:
                if sessions.index(i) <= 1:
                    Dummy.insert(0,1/np.power(2,sessions.index(i)+st.session_state.afternoonteacher.index(t))*Reward)
                else:
                    Dummy.insert(0,1/np.power(2,st.session_state.afternoonteacher.index(t))*np.power(Penelty_weight,sessions.index(i)))
            Penelty_distribution.append(Dummy)

        Afternoon_Teacher_Penelty = pu.lpSum(np.dot(Penelty_distribution[i],Constraint_var[i]) for i in range(len(st.session_state.afternoonteacher)))
        
        
        # Teacher avoid someday
        Dummy = []
        for t in st.session_state.teacheravoidday:
            for d in teacher_avoid_vars[t]:
                for c in teacher_plan[t]:
                    for i in sessions:
                        Dummy.append(np.power(Penelty_weight,5)*var[c][(d,i)])

        Teacher_Avoid_Day_Penelty =pu.lpSum(i for i in Dummy)

        # Student not learn in Campus or in ICDI alternatively




        # Add Objective function
        p += (Morning_Teacher_Penelty+Afternoon_Teacher_Penelty+Teacher_Avoid_Day_Penelty,"Sum_of_Total_Penalty",)

        # Hard Constraint

        # Teaching According to the Curriculum
        for c in Classes:
            p += (pu.lpSum(var[c][s] for s in Timeslot) == 1)

        # Teachers teach one class at a time        
        for t in Teachers:  
            for s in Timeslot:
                p += (pu.lpSum(var[c][s] for c in teacher_plan[t]) <= 1)

        # Each Students attend one class at a time
        for i in Students:
            for s in Timeslot:
                p += (pu.lpSum(var[c][s] for c in student_plan[i]) <= 1)
        
        # Subjects in standard format MTh TuF
        for s in Subjects:
            if s not in MSubjects and s not in nonformat:
                for c in subject_plan[s]:
                    if c[3] == 0:
                        for j in sessions: 
                            p+= (var[(c[0],c[1],c[2],0)][(Days[0],j)] == var[(c[0],c[1],c[2],1)][(Days[3],j)])
                            p+= (var[(c[0],c[1],c[2],1)][(Days[0],j)] == var[(c[0],c[1],c[2],0)][(Days[3],j)])
                            p+= (var[(c[0],c[1],c[2],0)][(Days[1],j)] == var[(c[0],c[1],c[2],1)][(Days[4],j)])
                            p+= (var[(c[0],c[1],c[2],1)][(Days[1],j)] == var[(c[0],c[1],c[2],0)][(Days[4],j)])
                            p+= (var[(c[0],c[1],c[2],0)][(Days[2],j)] == 0)
                            p+= (var[(c[0],c[1],c[2],1)][(Days[2],j)] == 0)
        
        # self-manage-timeslot subjects
        for s in MSubjects:
            for slot in manual_plan[s]:
                for c in Classes:
                    if c[2] == s and c[3] == manual_plan[s].index(slot): 
                        # st.write(c,slot)
                        p += (var[c][slot] == 1)
        
        # Lectuer teach least than 3 Classes a day
        for t in Teachers:
            for d in Days:
                p += (pu.lpSum(var[c][(d,p)] for c in teacher_plan[t] for p in sessions) <= 2)

        # Teacher avoid have classes before and after lunch in the same day
        for t in Teachers:
            for d in Days:
                for c1 in teacher_plan[t]:
                    for c2 in teacher_plan[t]:                
                        p+= (var[c1][(d,sessions[2])]+var[c2][(d,sessions[3])] <= 1)

        # '''Solve'''
        st.write('เมื่อเตรียมข้อมูลพร้อมแล้ว สามารถกดปุ่มด้านล่างเพื่อเริ่มสร้างตารางสอนได้เลย :sunglasses:')

        solve_button = st.button('เริ่มการสร้างตารางสอน',key='solve')

        if solve_button:
            with st.spinner('โปรแกรมกำลังประมวลผลโดยจะใช้เวลาไม่เกิน 1 นาที...โปรดรอ...'):
                p.solve(pu.PULP_CBC_CMD(maxSeconds=60, msg=1, fracGap=0))

                if pu.LpStatus[p.status] == 'Infeasible':
                    st.error("โปรแกรมไม่สามารถหาคำตอบที่สอดคล้องกับทุกเงื่อนไขได้ โปรดตรวจสอบข้อมูลที่ให้ หรืออาจปรับเงื่อนไขลดลง")
                else:
                    colums_name = ['{}'.format(p) for p in sessions]
                    
                    # Teacher Schedule Dataframe
                    df_teacher = { t : pd.DataFrame(index=Days, columns= colums_name) for t in Teachers}   
                    
                    for t in Teachers:
                        for i in Days:
                            for j in sessions:
                                Dummy = [round(var[c][(i,j)].varValue) for c in teacher_plan[t]] #somehow, some solution is not exactly one.
                                if sum(Dummy) == 0:
                                    df_teacher[t].at[i,j] = ''
                                else:
                                    for c in teacher_plan[t]:
                                        if round(var[c][(i,j)].varValue) == 1:
                                            df_teacher[t].at[i,j] = c[2]
                        # Insert Lunch Time
                        df_teacher[t].insert(3,'Lunch Time',['','','','','']) 
                    
                    # Din Student Schedule Dataframe
                    df_student = { g : pd.DataFrame(index=Days, columns= colums_name) for g in Dinstudents}   

                    for g in Dinstudents:
                        for i in Days:
                            for j in sessions:
                                Dummy = [round(var[c][(i,j)].varValue) for c in student_plan[g]] #somehow, some solution is not exactly one.
                                if sum(Dummy) == 0:
                                    df_student[g].at[i,j] = ''
                                else:
                                    for c in student_plan[g]:
                                        if round(var[c][(i,j)].varValue) == 1:
                                            df_student[g].at[i,j] = c[2]
                    # Insert Lunch Time
                        df_student[g].insert(3,'Lunch Break',['','','','','']) 
                    
                    # Show result
                    with st.expander("ตรวจดูตารางเรียนและตารางสอน"):
                            for g in Dinstudents:
                                st.write('นักศึกษา {}'.format(g))
                                st.write(df_student[g])
                            for t in Teachers:
                                st.write('อาจารย์ {}'.format(t))
                                st.write(df_teacher[t])

                     # '''Create a Pandas Excel writer using XlsxWriter as the engine.'''
                        
                    writer = pd.ExcelWriter('ICDI timetabing semester {} Academic year {}.xlsx'.format(semester,year), engine='xlsxwriter')

                    # Write each dataframe to a different worksheet.
                    for g in Dinstudents:
                            df_student[g].to_excel(writer, sheet_name='{}'.format(g))
                    
                    for t in Teachers:
                        df_teacher[t].to_excel(writer, sheet_name='{}'.format(t))
                    

                    # Close the Pandas Excel writer and output the Excel file.
                    writer.save()

                    with open('ICDI timetabing semester {} Academic year {}.xlsx'.format(semester,year), 'rb') as my_file:
                        st.download_button(
                            label = '📥 ดาวน์โหลดตารางสอน 📥', 
                            data = my_file, 
                            file_name = 'ICDI timetabing semester {} Academic year {}.xlsx'.format(semester,year), 
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            key='download')
