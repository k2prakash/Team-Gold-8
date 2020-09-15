# This is a Python script written to fulfill the team assignment to aid with our research on  
# "What Makes an Employee go from "Good to Great"
# The script takes an excel sheet as an input which contains the responses of all the employees in two columns
# The scrip performs data processing on the responses and created dictionaries of related concepts
# As part of post processing the dictionaries are further processes to extract data and build relatiionships between 
# them and the three motiational buckets, namely "Personal needs, Organizational Culture and Job Satisfaction"
# The results are then plotted for further analysis using OB behavior principles
# Author: Kaushik Prakash
# Last Modified: 09/14/2020

import pandas as pd
import re
import matplotlib.pyplot as plt
import numpy as np
import math

file = r'C:\Users\kp409917\Desktop\Team project\Jay\Jay-Team project - Interview question sheet.xlsx'

interviewee_dict = {}
common_benefits = {"Retirement Plan", "Medical Insurance", "ESPP", "Vision", "Dental", "Vacation/Sick time",
                   "Life Insurance", "Disability Insurance", "Severance", "Assistance Plans", "Parental Leave",
                   "Employee Wellness", "Work Flexibility", "Pet Insurance", "Company Car", "Air Travel",
                   "Workplace Amenities", "Accommodation Allowance", "Work Compensation", "Education Allowance"}

common_motivators = {'Salary': {}, 'Religion': {}, 'Work': {}, 'Growth': {}, 'People': {}, 'Influence': {},
                     'Leadership': {}, 'Integrity': {}, 'Appreciation': {}, 'Family': {}, 'Learning': {},
                     'Reward': {}, 'Mentorship': {}, 'Competition': {}, 'Success': {}, 'Fulfillment': {}, 'Role': {},
                     'Helping': {}, 'Goals': {}, 'Interaction': {}, 'Action': {}, 'Self': {}, 'Miscellaneous': {},
                     'Competence': {}, 'Autonomy': {}}

combined_factors = {}

satisfaction_levels = {'Less Satisfied': {}, 'Moderately Satisfied': {}, 'Highly Satisfied': {}, 'Satisfied': {}}

education_levels = {'Doctoral Degree': {}, 'Master\'s Degree': {}, 'Bachelor\'s Degree': {}, 'Associate\'s Degree': {},
                    'Post-Secondary, non-Degree award': {}, 'GED': {}, 'High school': {},
                    'Diploma': {}, 'No formal Education': {}}

employee_valuations = {'Good': {}, 'Great': {}, 'Unknown': {}, 'Neither': {}}

common_buckets = {'Organizational Factors': {}, 'Personal needs': {}, 'Overcoming Inertia': {}}

pay_satisfaction_levels = {'Less Satisfied': {}, 'Moderately Satisfied': {}, 'Highly Satisfied': {}, 'Satisfied': {},
                           'NA': {}}

leadership_satisfaction_levels = {'Less Satisfied': {}, 'Moderately Satisfied': {}, 'Highly Satisfied': {},
                                  'Satisfied': {}, 'NA': {}}


def get_seniority_level(title):
    if re.search(r'.*senior', title.lower()):
        return "Senior level"
    elif re.search(r'.*mid', title.lower()):
        return "Mid level"
    elif re.search(r'.*executive', title.lower()):
        return "Executive level"
    elif re.search(r'.*manager', title.lower()):
        return "Manager level"
    elif re.search(r'.*entree', title.lower()):
        return "Entry level"


def process_personal_satisfaction(text, name):
    if re.search('^not.*|^no.*|.*negat.*', text.lower(), re.IGNORECASE):
        satisfaction_levels['Less Satisfied'][name] = []
        satisfaction_levels['Less Satisfied'][name].append(text)
    elif re.search('.*medi.*|.*moderat.*', text.lower()):
        satisfaction_levels['Moderately Satisfied'][name] = []
        satisfaction_levels['Moderately Satisfied'][name].append(text)
    elif re.search('.*yes.*|.*sure.*|^satisfied.*', text.lower()):
        satisfaction_levels['Satisfied'][name] = []
        satisfaction_levels['Satisfied'][name].append(text)
    elif re.search('.*very.*|.*high|.*extrem.*|.*increas', text.lower()):
        satisfaction_levels['Highly Satisfied'][name] = []
        satisfaction_levels['Highly Satisfied'][name].append(text)


def process_personal_motivators(text, name):
    motivators = text.split(';')
    for m in motivators:
        if re.search('.*compete.*', m.lower()) or re.search('.*capab.*', m.lower()) or \
                re.search('.*profici.*', m.lower()) or re.search('.*expert.*', m.lower()) or \
                re.search('.*skill.*', m.lower()) or re.search('.*effect.*', m.lower()) or \
                re.search('.*producti.*', m.lower()) or re.search('.*effect.*', m.lower()) or \
                re.search('.*confiden.*', m.lower()) or re.search('.*determination.*', m.lower()) or \
                re.search('.*ability.*', m.lower()):
            if name not in common_motivators['Competence'].keys():
                common_motivators['Competence'][name] = []
                common_motivators['Competence'][name].append(m.strip())
            else:
                common_motivators['Competence'][name].append(m.strip())
        elif re.search('.*live my life.*', m.lower()) or re.search('.*live the life i want.*', m.lower()) \
                or re.search('.*Myself*', m.lower()) or re.search('.*\bme\b.*', m.lower()) \
                or re.search('.*time off.*', m.lower()) or re.search('.*downtime.*', m.lower()) \
                or re.search('.*do what i love.*', m.lower()):
            if name not in common_motivators['Self'].keys():
                common_motivators['Self'][name] = []
                common_motivators['Self'][name].append(m.strip())
            else:
                common_motivators['Self'][name].append(m.strip())
        elif re.search('.*autonom*', m.lower()) or re.search('.*indepen.*', m.lower()) \
                or re.search('.*freedom*', m.lower()) or re.search('.freewill*', m.lower()):
            if name not in common_motivators['Autonomy'].keys():
                common_motivators['Autonomy'][name] = []
                common_motivators['Autonomy'][name].append(m.strip())
            else:
                common_motivators['Autonomy'][name].append(m.strip())
        elif re.search('.*salary.*', m.lower()) or re.search('.*pay.*', m.lower()) or \
                re.search('.*money.*', m.lower()) or re.search('.*hike.*', m.lower()) or \
                re.search('.*incentive.*', m.lower()) or re.search('.*paid.*', m.lower()) or \
                re.search('.*compensation.*', m.lower()):
            if name not in common_motivators['Salary'].keys():
                common_motivators['Salary'][name] = []
                common_motivators['Salary'][name].append(m.strip())
            else:
                common_motivators['Salary'][name].append(m.strip())
        elif re.search('.*role.*', m.lower()) or re.search('.*occupation.*', m.lower()) \
                or re.search('.*profess.*', m.lower()) or re.search('.*post.*', m.lower()) \
                or re.search('.*position.*', m.lower()) or re.search('.*assignment.*', m.lower()) \
                or re.search('.*activity.*', m.lower()):
            if name not in common_motivators['Role'].keys():
                common_motivators['Role'][name] = []
                common_motivators['Role'][name].append(m.strip())
            else:
                common_motivators['Role'][name].append(m.strip())
        elif re.search('.*fulfil.*', m.lower()) or re.search('.*accomplish.*', m.lower()) \
                or re.search('.*satisf.*', m.lower()):
            if name not in common_motivators['Fulfillment'].keys():
                common_motivators['Fulfillment'][name] = []
                common_motivators['Fulfillment'][name].append(m.strip())
            else:
                common_motivators['Fulfillment'][name].append(m.strip())
        elif re.search('.*religion.*', m.lower()):
            if name not in common_motivators['Religion'].keys():
                common_motivators['Religion'][name] = []
                common_motivators['Religion'][name].append(m.strip())
            else:
                common_motivators['Religion'][name].append(m.strip())
        elif re.search('.*value.*', m.lower()) or re.search('.*appreicat.*', m.lower()) or \
                re.search('.*receive credit', m.lower()) or re.search('.*receive praise', m.lower()):
            if name not in common_motivators['Appreciation'].keys():
                common_motivators['Appreciation'][name] = []
                common_motivators['Appreciation'][name].append(m.strip())
            else:
                common_motivators['Appreciation'][name].append(m.strip())
        elif re.search('.*help.*', m.lower()) or re.search('.*assist.*', m.lower()) or \
                re.search('.*aid.*', m.lower()) or re.search('.*service.*', m.lower()) or \
                re.search('.*support.*', m.lower()) or re.search('.*contribut.*', m.lower()) \
                or re.search('.*compassion.*', m.lower()):
            if name not in common_motivators['Helping'].keys():
                common_motivators['Helping'][name] = []
                common_motivators['Helping'][name].append(m.strip())
            else:
                common_motivators['Helping'][name].append(m.strip())
        elif re.search('.*growth.*', m.lower()) or re.search('.*promot.*', m.lower()) or \
                re.search('.*progress.*', m.lower()):
            if name not in common_motivators['Growth'].keys():
                common_motivators['Growth'][name] = []
                common_motivators['Growth'][name].append(m.strip())
            else:
                common_motivators['Growth'][name].append(m.strip())
        elif re.search('.*team.*', m.lower()) or re.search('.*colleag', m.lower()) or \
                re.search('.*partner', m.lower()) or re.search('.*member', m.lower()) or \
                re.search('.*co.*work', m.lower()) or re.search('.*patient', m.lower()):
            if name not in common_motivators['People'].keys():
                common_motivators['People'][name] = []
                common_motivators['People'][name].append(m.strip())
            else:
                common_motivators['People'][name].append(m.strip())
        elif re.search('.*work.*', m.lower()) or re.search('.*tasks.*', m.lower()) \
                or re.search('.*industry.*', m.lower()) or re.search('.*produc.*', m.lower()) \
                or re.search('.*job.*', m.lower()) or re.search('.*process.*', m.lower()):
            if name not in common_motivators['Work'].keys():
                common_motivators['Work'][name] = []
                common_motivators['Work'][name].append(m.strip())
            else:
                common_motivators['Work'][name].append(m.strip())
        elif re.search('.*influenc.*', m.lower()) or re.search('.*difference.*', m.lower()) \
                or re.search('.*betterment.*', m.lower()) or re.search('.*bettering.*', m.lower()):
            if name not in common_motivators['Influence'].keys():
                common_motivators['Influence'][name] = []
                common_motivators['Influence'][name].append(m.strip())
            else:
                common_motivators['Influence'][name].append(m.strip())
        elif re.search('.*mentor.*', m.lower()) or re.search('.*guide.*', m.lower()) \
                or re.search('.*advice.*', m.lower()) or re.search('.*counsel.*', m.lower()) or \
                re.search('.*teach.*', m.lower()):
            if name not in common_motivators['Mentorship'].keys():
                common_motivators['Mentorship'][name] = []
                common_motivators['Mentorship'][name].append(m.strip())
            else:
                common_motivators['Mentorship'][name].append(m.strip())
        elif re.search('.*leader.*', m.lower()) or re.search('.*mentor.*', m.lower()) \
                or re.search('.*supervisor.*', m.lower()) or re.search('.*manag.*', m.lower()):
            if name not in common_motivators['Leadership'].keys():
                common_motivators['Leadership'][name] = []
                common_motivators['Leadership'][name].append(m.strip())
            else:
                common_motivators['Leadership'][name].append(m.strip())
        elif re.search('.*integrity.*', m.lower()) or re.search('.*honest.*', m.lower()) or \
                re.search('.*transparen.*', m.lower()) or re.search('.*trust.*', m.lower()):
            if name not in common_motivators['Integrity'].keys():
                common_motivators['Integrity'][name] = []
                common_motivators['Integrity'][name].append(m.strip())
            else:
                common_motivators['Integrity'][name].append(m.strip())
        elif re.search('.*family.*', m.lower()) or re.search('.*wife.*', m.lower()) \
                or re.search('.*child.*', m.lower()):
            if name not in common_motivators['Integrity'].keys():
                common_motivators['Family'][name] = []
                common_motivators['Family'][name].append(m.strip())
            else:
                common_motivators['Family'][name].append(m.strip())
        elif re.search('.*learn.*', m.lower()) or re.search('.*educat.*', m.lower()):
            if name not in common_motivators['Learning'].keys():
                common_motivators['Learning'][name] = []
                common_motivators['Learning'][name].append(m.strip())
            else:
                common_motivators['Learning'][name].append(m.strip())
        elif re.search('.*rewar.*', m.lower()) or re.search('.*award.*', m.lower()):
            if name not in common_motivators['Rewards'].keys():
                common_motivators['Rewards'][name] = []
                common_motivators['Rewards'][name].append(m.strip())
            else:
                common_motivators['Rewards'][name].append(m.strip())
        elif re.search('.*competition.*', m.lower()) or re.search('.*rivalry.*', m.lower()):
            if name not in common_motivators['Competition'].keys():
                common_motivators['Competition'][name] = []
                common_motivators['Competition'][name].append(m.strip())
            else:
                common_motivators['Competition'][name].append(m.strip())
        elif re.search('.*success.*', m.lower()) or re.search('.*accomplish.*', m.lower()) \
                or re.search('.*achieve.*', m.lower()) or re.search('.*victor.*', m.lower()) \
                or re.search('.*win.*', m.lower()) or re.search('.*progress.*', m.lower()):
            if name not in common_motivators['Success'].keys():
                common_motivators['Success'][name] = []
                common_motivators['Success'][name].append(m.strip())
            else:
                common_motivators['Success'][name].append(m.strip())
        elif re.search('.*goal.*', m.lower()) or re.search('.*intention.*', m.lower()) \
                or re.search('.*sale.*', m.lower()) or re.search('.*target.*', m.lower()) \
                or re.search('.*objective.*', m.lower()) or re.search('.*desire.*', m.lower()):
            if name not in common_motivators['Goals'].keys():
                common_motivators['Goals'][name] = []
                common_motivators['Goals'][name].append(m.strip())
            else:
                common_motivators['Goals'][name].append(m.strip())
        elif re.search('.*interact.*', m.lower()) or re.search('.*communicat.*', m.lower()) \
                or re.search('.*talk.*', m.lower()) or re.search('.*exchange.*', m.lower()) \
                or re.search('.*reciproc.*', m.lower()) or re.search('.*relationship.*', m.lower()):
            if name not in common_motivators['Interaction'].keys():
                common_motivators['Interaction'][name] = []
                common_motivators['Interaction'][name].append(m.strip())
            else:
                common_motivators['Interaction'][name].append(m.strip())
        elif re.search('.*implement.*', m.lower()) or re.search('.*perform.*', m.lower()) \
                or re.search('.*carry.*out.*', m.lower()) or re.search('.*deed.*', m.lower()) \
                or re.search('.*execut.*', m.lower()):
            if name not in common_motivators['Action'].keys():
                common_motivators['Action'][name] = []
                common_motivators['Action'][name].append(m.strip())
            else:
                common_motivators['Action'][name].append(m.strip())
        elif re.search('.*relax*', m.lower()) or re.search('.*music*', m.lower()):
            if name not in common_motivators['Miscellaneous'].keys():
                common_motivators['Miscellaneous'][name] = []
                common_motivators['Miscellaneous'][name].append(m.strip())
            else:
                common_motivators['Miscellaneous'][name].append(m.strip())

    return


def process_pay_satisfaction_levels(text, name):
    if re.search('^not.*|^no.*|.*negat.*|.*lack.*|.*unfair.*|.*below.*|.*underpaid.*|.*small.*|.*less.*', text.lower(), re.IGNORECASE):
        pay_satisfaction_levels['Less Satisfied'][name] = []
        pay_satisfaction_levels['Less Satisfied'][name].append(text)
    elif re.search('.*medi.*|.*moderat.*|.*paid more.*|.*reasonable.*|.*ok.*|.*okay|.*could be better.*', text.lower()):
        pay_satisfaction_levels['Moderately Satisfied'][name] = []
        pay_satisfaction_levels['Moderately Satisfied'][name].append(text)
    elif re.search('^yes.*|.*sure.*|^satisfied.*|.*fair.*|.*general.*', text.lower()):
        pay_satisfaction_levels['Satisfied'][name] = []
        pay_satisfaction_levels['Satisfied'][name].append(text)
    elif re.search('.*very.*|.*high|.*extrem.*|.*increas|.*great|.*competetive', text.lower()):
        pay_satisfaction_levels['Highly Satisfied'][name] = []
        pay_satisfaction_levels['Highly Satisfied'][name].append(text)
    elif re.search('.*n\/a.*|.*\bna\b.*', text.lower()):
        pay_satisfaction_levels['NA'][name] = []
        pay_satisfaction_levels['NA'][name].append(text)


def process_leadership_satisfaction_levels(text, name):
    if re.search('^not.*|^no.*|.*negat.*|.*lack.*|.*unfair.*|.*below.*|.*underpaid.*|.*small.*|.*less.*|.*minimal.*',
                 text.lower(), re.IGNORECASE):
        leadership_satisfaction_levels['Less Satisfied'][name] = []
        leadership_satisfaction_levels['Less Satisfied'][name].append(text)
    elif re.search('.*medi.*|.*moderat.*|.*average.|.*somewhat.*|.*situational.*|.*ok.*|.*okay', text.lower()):
        leadership_satisfaction_levels['Moderately Satisfied'][name] = []
        leadership_satisfaction_levels['Moderately Satisfied'][name].append(text)
    elif re.search('.*very.*|.*high|.*everytime.*|.*great|.*100%.*', text.lower()):
        leadership_satisfaction_levels['Highly Satisfied'][name] = []
        leadership_satisfaction_levels['Highly Satisfied'][name].append(text)
    elif re.search('.*yes.*|.*sure.*|^satisfied.*|.*fair.*|.*general.*', text.lower()):
        leadership_satisfaction_levels['Satisfied'][name] = []
        leadership_satisfaction_levels['Satisfied'][name].append(text)
    elif re.search('.*n\/a.*|.*\bna\b.*', text.lower()):
        leadership_satisfaction_levels['NA'][name] = []
        leadership_satisfaction_levels['NA'][name].append(text)
    else:
        leadership_satisfaction_levels['NA'][name] = []
        leadership_satisfaction_levels['NA'][name].append(text)


def process_employee_valuation(text, question, name):
    if question == 'good':
        if len(text.strip()) > 0:
            if re.search('.*good.*', text.lower()) or re.search('.*better.*', text.lower()) \
                    or re.search('.*satisfactory.*', text.lower()) or re.search('.*acceptable.*', text.lower()) \
                    or re.search('.*yes.*', text.lower()):
                if name not in employee_valuations['Good'].keys():
                    employee_valuations['Good'][name] = []
                    employee_valuations['Good'][name].append(text.strip())
            elif re.search('.*no.*', text.lower()) or re.search('.*negat.*', text.lower()):
                if name not in employee_valuations['Neither'].keys():
                    employee_valuations['Neither'][name] = []
                    employee_valuations['Neither'][name].append(text.strip())
            elif re.search('nan', text.lower()):
                if name not in employee_valuations['Unknown'].keys():
                    employee_valuations['Unknown'][name] = []
                    employee_valuations['Unknown'][name].append(text.strip())
    elif question == 'great':
        if len(text.strip()) > 0:
            if re.search('.*great.*', text.lower()) or re.search('.*excellent.*', text.lower()) \
                    or re.search('.*exceptional.*', text.lower()) or re.search('.*extraordina.*', text.lower()) \
                    or re.search('.*amazing.*', text.lower()) or re.search('.*yes.*', text.lower()):
                if name not in employee_valuations['Great'].keys():
                    employee_valuations['Great'][name] = []
                    employee_valuations['Great'][name].append(text.strip())
                    if name in employee_valuations['Good'].keys():
                        employee_valuations['Good'].pop(name)
        else:
            return


def process_education_level(text, name):
    if re.search('no.*educ', text.lower()):
        education_levels['No formal Education'][name] = []
        education_levels['No formal Education'][name].append(text)
    elif re.search('diploma.*', text.lower()):
        education_levels['Diploma'][name] = []
        education_levels['Diploma'][name].append(text)
    elif re.search('high sch.*', text.lower()) or re.search('highsch.*', text.lower()):
        education_levels['High school'][name] = []
        education_levels['High school'][name].append(text)
    elif re.search('college.*no-deg', text.lower()) or re.search('.*ged', text.lower()):
        education_levels['GED'][name] = []
        education_levels['GED'][name].append(text)
    elif re.search('post-second.*non-deg', text.lower()):
        education_levels['Post-Secondary, non-Degree award'][name] = []
        education_levels['Post-Secondary, non-Degree award'][name].append(text)
    elif re.search('.*associa', text.lower()):
        education_levels['Associate\'s Degree'][name] = []
        education_levels['Associate\'s Degree'][name].append(text)
    elif re.search('.*bachelor.*|.*b\.s|.*bs|.*btech|.*b\.tech|.*bfa|.*b\.f\.a', text.lower()):
        education_levels['Bachelor\'s Degree'][name] = []
        education_levels['Bachelor\'s Degree'][name].append(text)
    elif re.search('.*master.*|.*mba.*|.*m\.b\.a|.*j\.d|.*jd|.*juris|.*m\.s|.*ms|.*m\.a|.*ma', text.lower()):
        education_levels['Master\'s Degree'][name] = []
        education_levels['Master\'s Degree'][name].append(text)
    elif re.search('.*doctor.*|phd|.*post', text.lower()):
        education_levels['Doctoral Degree'][name] = []
        education_levels['Doctoral Degree'][name].append(text)


def process_excel():
    excel_file = pd.ExcelFile(file)
    # print(excel_file.sheet_names)
    for name in excel_file.sheet_names:
        df = pd.read_excel(file, name)
        for index in df.index:
            m1 = re.search(".*nan.*", str(df['Responses'][index]))
            m2 = re.search(r'what is your name.*', str(df['Questions'][index]).lower())
            m3 = re.search(r'what is your job title.*', str(df['Questions'][index]).lower())
            m4 = re.search(r'at what level in your company.*', str(df['Questions'][index]).lower())
            m5 = re.search(r'age range.*', str(df['Questions'][index]).lower())
            m6 = re.search(r'race.*', str(df['Questions'][index]).lower())
            m7 = re.search(r'.*married.*', str(df['Questions'][index]).lower())
            m8 = re.search(r'gender.*', str(df['Questions'][index]).lower())
            m10 = re.search(r'do you believe you\'re a good employee.*', str(df['Questions'][index]).lower())
            m12 = re.search(r'what motivates you.*', str(df['Questions'][index]).lower())
            m13 = re.search(r'are you satisfied in.*', str(df['Questions'][index]).lower())
            m14 = re.search(r'what is your highest level of educ', str(df['Questions'][index]).lower())
            m16 = re.search(r'do you believe you\'re a great employee.*', str(df['Questions'][index]).lower())
            m17 = re.search(r'how do you feel about your pay.*', str(df['Questions'][index]).lower())
            m18 = re.search(r'do you receive support from you direct leader.*', str(df['Questions'][index]).lower())
            if m2:
                interviewee_dict[name] = {}
            elif m3:
                interviewee_dict[name]['Title'] = str(df['Responses'][index])
            elif m4:
                interviewee_dict[name]['Level'] = get_seniority_level(str(df['Responses'][index]))
            elif m5:
                interviewee_dict[name]['Age-Group'] = str(df['Responses'][index])
            elif m6:
                interviewee_dict[name]['Race'] = str(df['Responses'][index])
            elif m7:
                interviewee_dict[name]['Marital Status'] = str(df['Responses'][index])
            elif m8:
                interviewee_dict[name]['Gender'] = str(df['Responses'][index])
            elif m10:
                process_employee_valuation(str(df['Responses'][index]), 'good', name)
            elif m12:
                process_personal_motivators(str(df['Responses'][index]), name)
            elif m13:
                process_personal_satisfaction(str(df['Responses'][index]), name)
            elif m14:
                interviewee_dict[name]['Education Level'] = str(df['Responses'][index])
                process_education_level(str(df['Responses'][index]), name)
            elif m16:
                process_employee_valuation(str(df['Responses'][index]), 'great', name)
            elif m17:
                process_pay_satisfaction_levels(str(df['Responses'][index]), name)
            elif m18:
                process_leadership_satisfaction_levels(str(df['Responses'][index]), name)
            elif m1:
                continue


def get_count_of_attribute(attribute_dict):
    count = 0
    for name in attribute_dict:
        l = len(attribute_dict[name])
        count = count + l
    return count


def post_process():
    
    print("These are the details of the people who were interviewed.")
    for person in interviewee_dict.keys():
        print("Name: {0}".format(person))
        print("Gender: {0}".format(interviewee_dict[person]['Gender']))
        print("Title: {0}".format(interviewee_dict[person]['Title']))
        print("Organizational position: {0}".format(interviewee_dict[person]['Level']))
        print("Race: {0}".format(interviewee_dict[person]['Race']))
        print("Age group: {0}".format(interviewee_dict[person]['Age-Group']))
        print("Marital Status: {0}".format(interviewee_dict[person]['Marital Status']))
        print("Education Level: {0}".format(interviewee_dict[person]['Education Level']))
        print("\n")
    print("Interviewee responses for \"Are you satisfied in your role?\"")
    print("-------------------------------------------------------------------------------------------")
    satisfact_levels_xaxis = []
    satisfact_count_yaxis = []
    for key in satisfaction_levels.keys():
        print("\"{0}\" was mentioned {1} time(s).".format(key, get_count_of_attribute(satisfaction_levels[key])))
        satisfact_levels_xaxis.append(key)
        satisfact_count_yaxis.append(get_count_of_attribute(satisfaction_levels[key]))
        for name in satisfaction_levels[key]:
            print("From {0}'s response: ".format(name))
            for level in satisfaction_levels[key][name]:
                print("\t\"{0}\"".format(level))
        print("\n")
    cmaps_edu = ['red', 'green', 'orange', 'cyan', 'brown', 'grey', 'blue', 'indigo']
    x_pos = np.arange(0, len(satisfact_levels_xaxis))
    y_pos = np.arange(0, math.ceil(max(satisfact_count_yaxis)) + 5, step=5)
    plt.figure(figsize=(13, 13))
    plt.bar(x_pos, satisfact_count_yaxis, align='center', alpha=0.5, color=cmaps_edu)
    plt.xticks(x_pos, satisfact_levels_xaxis, rotation=25, fontsize=14)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    for index, data in enumerate(satisfact_count_yaxis):
        plt.text(x=index, y=data+1, s=f"{data}", fontsize='20')
    plt.title('Analysis of Job Satisfaction Levels', fontsize=20)
    plt.savefig(r'C:\Users\kp409917\Desktop\satisfaction_levels.png')

    print("Education Levels Analysis")
    print("-------------------------------------------------------------------------------------------")
    edu_levels_xaxis = []
    edu_level_count_yaxis = []
    for key in education_levels.keys():
        # if get_count_of_attribute(education_levels[key]) != 0:
        if key == 'Post-Secondary, non-Degree award':
            edu_levels_xaxis.append('Post-Secondary N.D')
        else:
            edu_levels_xaxis.append(key)
        edu_level_count_yaxis.append(get_count_of_attribute(education_levels[key]))
        print("\"{0}\" was mentioned {1} time(s).".format(key, get_count_of_attribute(education_levels[key])))
        for name in education_levels[key]:
            print("{0}'s Education Level: ".format(name))
            for level in education_levels[key][name]:
                print("\t\"{0}\"".format(level))
        print("\n")
    cmaps_edu = ['red', 'green', 'orange', 'cyan', 'brown', 'grey', 'blue', 'indigo']
    x_pos = np.arange(0, len(edu_levels_xaxis))
    y_pos = np.arange(0, math.ceil(max(edu_level_count_yaxis)) + 10, step=5)
    plt.figure(figsize=(15, 15))
    plt.bar(x_pos, edu_level_count_yaxis, align='center', alpha=0.5, color=cmaps_edu)
    plt.xticks(x_pos, edu_levels_xaxis, rotation=40, fontsize=15)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    plt.xlabel('\n*N.D -> Non Degree')
    for index, data in enumerate(edu_level_count_yaxis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.title('Analysis of Education Levels', fontsize=20)
    plt.savefig(r'C:\Users\kp409917\Desktop\education_leves.png')

    print("Employee valuation analysis")
    print("-------------------------------------------------------------------------------------------")
    valuation_labels_xaxis = []
    valuation_count_yaxis = []
    for key in employee_valuations.keys():
        valuation_labels_xaxis.append(key)
        valuation_count_yaxis.append(get_count_of_attribute(employee_valuations[key]))
        print("\"{0}\" was mentioned {1} time(s).".format(key, get_count_of_attribute(employee_valuations[key])))
        for name in employee_valuations[key]:
            print("From {0}'s response: ".format(name))
            for valuation in employee_valuations[key][name]:
                print("\t\"{0}\"".format(valuation))
        print("\n")

    cmaps = ['red', 'green', 'orange', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'beige', 'yellow',
             'purple', 'pink', 'maroon']

    x_pos = np.arange(0, len(valuation_labels_xaxis))
    y_pos = np.arange(0, math.ceil(max(valuation_count_yaxis)) + 10, step=5)
    plt.figure(figsize=(10, 10))
    plt.bar(x_pos, valuation_count_yaxis, align='center', alpha=0.5, color=cmaps)
    plt.xticks(x_pos, valuation_labels_xaxis, rotation=45, fontsize=14)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    plt.title('Analysis of Employee \"Good vs Great\" Evaluation', fontsize=20)
    for index, data in enumerate(valuation_count_yaxis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.savefig(r'C:\Users\kp409917\Desktop\Employee_Valuation.png')

    print("Employee Pay evaluation analysis")
    print("-------------------------------------------------------------------------------------------")
    pay_valuation_labels_xaxis = []
    pay_valuation_count_yaxis = []
    for key in pay_satisfaction_levels.keys():
        pay_valuation_labels_xaxis.append(key)
        pay_valuation_count_yaxis.append(get_count_of_attribute(pay_satisfaction_levels[key]))
        print("\"{0}\" was mentioned {1} time(s).".format(key, get_count_of_attribute(pay_satisfaction_levels[key])))
        for name in pay_satisfaction_levels[key]:
            print("From {0}'s response: ".format(name))
            for valuation in pay_satisfaction_levels[key][name]:
                print("\t\"{0}\"".format(valuation))
        print("\n")

    cmaps = ['red', 'green', 'orange', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'beige', 'yellow',
             'purple', 'pink', 'maroon']

    x_pos = np.arange(0, len(pay_valuation_labels_xaxis))
    y_pos = np.arange(0, math.ceil(max(pay_valuation_count_yaxis)) + 10, step=5)
    plt.figure(figsize=(10, 10))
    plt.bar(x_pos, pay_valuation_count_yaxis, align='center', alpha=0.5, color=cmaps)
    plt.xticks(x_pos, pay_valuation_labels_xaxis, rotation=30, fontsize=12)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    plt.title('Analysis of Employee Pay Evaluation', fontsize=20)
    for index, data in enumerate(pay_valuation_count_yaxis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.savefig(r'C:\Users\kp409917\Desktop\Employee Pay Valuation.png')

    print("Employee Leadership evaluation analysis")
    print("-------------------------------------------------------------------------------------------")
    leadership_evaluation_labels_xaxis = []
    leadership_evaluation_count_yaxis = []
    for key in leadership_satisfaction_levels.keys():
        leadership_evaluation_labels_xaxis.append(key)
        leadership_evaluation_count_yaxis.append(get_count_of_attribute(leadership_satisfaction_levels[key]))
        print("\"{0}\" was mentioned {1} time(s).".format(key, get_count_of_attribute(leadership_satisfaction_levels[key])))
        for name in leadership_satisfaction_levels[key]:
            print("From {0}'s response: ".format(name))
            for valuation in leadership_satisfaction_levels[key][name]:
                print("\t\"{0}\"".format(valuation))
        print("\n")

    cmaps = ['red', 'green', 'orange', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'beige', 'yellow',
             'purple', 'pink', 'maroon']

    x_pos = np.arange(0, len(leadership_evaluation_labels_xaxis))
    y_pos = np.arange(0, math.ceil(max(leadership_evaluation_count_yaxis)) + 10, step=5)
    plt.figure(figsize=(10, 10))
    plt.bar(x_pos, leadership_evaluation_count_yaxis, align='center', alpha=0.5, color=cmaps)
    plt.xticks(x_pos, leadership_evaluation_labels_xaxis, rotation=30, fontsize=12)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    plt.title('Analysis of Employee Leadership Evaluation', fontsize=20)
    for index, data in enumerate(leadership_evaluation_count_yaxis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.savefig(r'C:\Users\kp409917\Desktop\Employee Leadership Evaluation.png')

    print("Interviewee responses for \"What motivates you?\"")
    print("-------------------------------------------------------------------------------------------")
    common_bucket = {'Personal Needs': {}, 'Organizational Culture': {}, 'Job Satisfaction': {}}
    for key in common_motivators.keys():
        if key == "Salary":
            common_bucket['Personal Needs']['Salary'] = common_motivators[key]
        elif key == "Religion":
            common_bucket['Personal Needs']['Religion'] = common_motivators[key]
        elif key == "Growth":
            common_bucket['Personal Needs']['Growth'] = common_motivators[key]
        elif key == "Integrity":
            common_bucket['Personal Needs']['Integrity'] = common_motivators[key]
        elif key == "Family":
            common_bucket['Personal Needs']['Family'] = common_motivators[key]
        elif key == "Self":
            common_bucket['Personal Needs']['Self'] = common_motivators[key]
        elif key == "Learning":
            common_bucket['Personal Needs']['Learning'] = common_motivators[key]
        elif key == "Work":
            common_bucket['Job Satisfaction']['Work'] = common_motivators[key]
        elif key == "People":
            common_bucket['Organizational Culture']['People'] = common_motivators[key]
        elif key == "Leadership":
            common_bucket['Organizational Culture']['Leadership'] = common_motivators[key]
        elif key == "Role":
            common_bucket['Organizational Culture']['Role'] = common_motivators[key]
        elif key == "Company Culture":
            common_bucket['Organizational Culture']['Company Culture'] = common_motivators[key]
        elif key == "Goals":
            common_bucket['Job Satisfaction']['Goals'] = common_motivators[key]
        elif key == "Autonomy":
            common_bucket['Job Satisfaction']['Autonomy'] = common_motivators[key]
        elif key == "Influence":
            common_bucket['Organizational Culture']['Influence'] = common_motivators[key]
        elif key == "Appreciation":
            common_bucket['Job Satisfaction']['Appreciation'] = common_motivators[key]
        elif key == "Mentorship":
            common_bucket['Organizational Culture']['Mentorship'] = common_motivators[key]
        elif key == "Competition":
            common_bucket['Job Satisfaction']['Competition'] = common_motivators[key]
        elif key == "Fulfillment":
            common_bucket['Job Satisfaction']['Fulfillment'] = common_motivators[key]
        elif key == "Success":
            common_bucket['Job Satisfaction']['Success'] = common_motivators[key]
        elif key == "Helping":
            common_bucket['Organizational Culture']['Helping'] = common_motivators[key]
        elif key == "Interaction":
            common_bucket['Organizational Culture']['Interaction'] = common_motivators[key]
        elif key == "Reward":
            common_bucket['Job Satisfaction']['Reward'] = common_motivators[key]
        elif key == "Competence":
            common_bucket['Job Satisfaction']['Competence'] = common_motivators[key]

    common_bucket_labels_x_axis = []
    common_bucket_counts_y_axis = []
    needs_count = 0
    org_culture_count = 0
    job_satisfaction_count = 0
    for key in common_bucket.keys():
        if key == 'Personal Needs':
            print("Analyzing Motivating factors under Personal Needs")
            common_bucket_labels_x_axis.append(key)
            personal_needs_labels_x_axis = []
            personal_needs_counts_y_axis = []
            for k in common_bucket[key]:
                needs_count += get_count_of_attribute(common_bucket[key][k])
                personal_needs_labels_x_axis.append(k)
                personal_needs_counts_y_axis.append(get_count_of_attribute(common_bucket[key][k]))
                if get_count_of_attribute(common_bucket[key][k]) != 0:
                    print(
                        "\t\"{0}\" was mentioned {1} time(s).".format(k, get_count_of_attribute(common_bucket[key][k])))
                    for name in common_bucket[key][k]:
                        print("\t\tFrom {0}'s response: ".format(name))
                        for motivator in common_bucket[key][k][name]:
                            print("\t\t\t\"{0}\"".format(motivator))
            print("---------------------------------------------------------------")
            print("Personal Needs reported {0} entries".format(needs_count))
            print("---------------------------------------------------------------")
            common_bucket_counts_y_axis.append(needs_count)

            cmaps = ['red', 'green', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'orange',
                     'purple', 'pink', 'maroon']
            x_pos = np.arange(0, len(personal_needs_labels_x_axis))
            y_pos = np.arange(0, math.ceil(max(personal_needs_counts_y_axis)) + 10, step=5)
            plt.figure(figsize=(10, 10))
            plt.bar(x_pos, personal_needs_counts_y_axis, align='center', alpha=0.5, color=cmaps)
            plt.xticks(x_pos, personal_needs_labels_x_axis, rotation=25, fontsize=12)
            plt.yticks(y_pos, fontsize=15)
            plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
            for index, data in enumerate(personal_needs_counts_y_axis):
                plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
            plt.title('Analysis of Personal Needs Motivators', fontsize='20')
            plt.savefig(r'C:\Users\kp409917\Desktop\Personal Needs motivators.png')
            print("\n")
        elif key == 'Organizational Culture':
            print("Analyzing Motivating factors under Organizational Culture")
            common_bucket_labels_x_axis.append(key)
            org_culture_labels_x_axis = []
            org_culture_counts_y_axis = []
            for k in common_bucket[key]:
                org_culture_count += get_count_of_attribute(common_bucket[key][k])
                org_culture_labels_x_axis.append(k)
                org_culture_counts_y_axis.append(get_count_of_attribute(common_bucket[key][k]))
                if get_count_of_attribute(common_bucket[key][k]) != 0:
                    print(
                        "\t\"{0}\" was mentioned {1} time(s).".format(k, get_count_of_attribute(common_bucket[key][k])))
                    for name in common_bucket[key][k]:
                        print("\t\tFrom {0}'s response: ".format(name))
                        for motivator in common_bucket[key][k][name]:
                            print("\t\t\t\"{0}\"".format(motivator))
            print("---------------------------------------------------------------")
            print("Organizational Culture was reported {0} entries".format(org_culture_count))
            print("---------------------------------------------------------------")
            common_bucket_counts_y_axis.append(org_culture_count)

            cmaps = ['red', 'green', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'orange',
                     'purple', 'pink', 'maroon']
            x_pos = np.arange(0, len(org_culture_labels_x_axis))
            y_pos = np.arange(0, math.ceil(max(org_culture_counts_y_axis)) + 10, step=5)
            plt.figure(figsize=(10, 10))
            plt.bar(x_pos, org_culture_counts_y_axis, align='center', alpha=0.5, color=cmaps)
            plt.xticks(x_pos, org_culture_labels_x_axis, rotation=25, fontsize=12)
            plt.yticks(y_pos, fontsize=15)
            plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
            for index, data in enumerate(org_culture_counts_y_axis):
                plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
            plt.title('Analysis of Organization Culture Motivators', fontsize='20')
            plt.savefig(r'C:\Users\kp409917\Desktop\Organizational Culture motivators.png')
            print("\n")
        elif key == 'Job Satisfaction':
            print("Analyzing Motivating factors under Overcoming Inertia")
            common_bucket_labels_x_axis.append(key)
            job_satisfaction_labels_x_axis = []
            job_satisfaction_counts_y_axis = []
            for k in common_bucket[key]:
                job_satisfaction_count += get_count_of_attribute(common_bucket[key][k])
                job_satisfaction_labels_x_axis.append(k)
                job_satisfaction_counts_y_axis.append(get_count_of_attribute(common_bucket[key][k]))
                if get_count_of_attribute(common_bucket[key][k]) != 0:
                    print(
                        "\t\"{0}\" was mentioned {1} time(s).".format(k, get_count_of_attribute(common_bucket[key][k])))
                    for name in common_bucket[key][k]:
                        print("\t\tFrom {0}'s response: ".format(name))
                        for motivator in common_bucket[key][k][name]:
                            print("\t\t\t\"{0}\"".format(motivator))
            print("---------------------------------------------------------------")
            print("Job Satisfaction was reported {0} entries".format(job_satisfaction_count))
            print("---------------------------------------------------------------")

            cmaps = ['red', 'green', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'orange',
                     'purple', 'pink', 'maroon']

            x_pos = np.arange(0, len(job_satisfaction_labels_x_axis))
            y_pos = np.arange(0, math.ceil(max(job_satisfaction_counts_y_axis)) + 10, step=5)
            plt.figure(figsize=(10, 10))
            plt.bar(x_pos, job_satisfaction_counts_y_axis, align='center', alpha=0.5, color=cmaps)
            plt.xticks(x_pos, job_satisfaction_labels_x_axis, rotation=25, fontsize=12)
            plt.yticks(y_pos, fontsize=15)
            plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
            for index, data in enumerate(job_satisfaction_counts_y_axis):
                plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
            plt.title('Analysis of Job Satisfaction Motivators', fontsize='20')
            plt.savefig(r'C:\Users\kp409917\Desktop\Job Satisfaction motivators.png')

            common_bucket_counts_y_axis.append(job_satisfaction_count)

    cmaps = ['red', 'green', 'cyan', 'brown', 'grey', 'blue', 'indigo', 'orange',
             'purple', 'pink', 'maroon']
   
    x_pos = np.arange(0, len(common_bucket_labels_x_axis))
    y_pos = np.arange(0, math.ceil(max(common_bucket_counts_y_axis))+10, step=5)
    plt.figure(figsize=(10, 10))
    plt.bar(x_pos, common_bucket_counts_y_axis, align='center', alpha=0.5, color=cmaps)
    plt.xticks(x_pos, common_bucket_labels_x_axis, rotation=25, fontsize=12)
    plt.yticks(y_pos, fontsize=15)
    plt.ylabel('Interviewee Count', fontsize=15, labelpad=10)
    for index, data in enumerate(common_bucket_counts_y_axis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.title('Analysis of Employee Motivators', fontsize='20')
    plt.savefig(r'C:\Users\kp409917\Desktop\motivators.png')


''' 
The Main script which will run the program 
'''
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    process_excel()
    post_process()


