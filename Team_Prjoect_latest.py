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


good_to_great_factors = {'Work': {}, 'Feedback': {}, 'Appreciation': {}, 'Opportunities': {}, 'Mentoring': {},
                         'Culture': {}, 'Recognition': {}, 'Rewards': {}, 'Versatility': {},
                         'Commitment': {}, 'Performance': {}, 'Competence': {}, 'Support': {}, 'Intelligence': {},
                         'Creativity': {}}
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

primary_motivators = {'Salary': {}, 'Religion': {}, 'Work': {}, 'Growth': {}, 'People': {}, 'Influence': {},
                      'Leadership': {}, 'Integrity': {}, 'Appreciation': {}, 'Family': {}, 'Learning': {},
                      'Reward': {}, 'Mentorship': {}, 'Competition': {}, 'Success': {}, 'Fulfillment': {}, 'Role': {},
                      'Company Culture': {}, 'Helping': {}, 'Goals': {}, 'Interaction': {}, 'Action': {}, 'Self': {},
                      'Miscellaneous': {}, 'Competence': {}, 'Autonomy': {}}

satisfaction_levels = {'Less Satisfied': {}, 'Moderately Satisfied': {}, 'Highly Satisfied': {}, 'Satisfied': {}}

education_levels = {'Doctoral Degree': {}, 'Master\'s Degree': {}, 'Bachelor\'s Degree': {}, 'Associate\'s Degree': {},
                    'Post-Secondary, non-Degree award': {}, 'GED': {}, 'High school': {},
                    'Diploma': {}, 'No formal Education': {}}

employee_valuations = {'Good': {}, 'Great': {}, 'Unknown': {}, 'Neither': {}}

common_buckets = {'Organizational Factors': {}, 'Personal needs': {}, 'Overcoming Inertia': {}}


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


def process_primary_motivators(text, name):
    motivators = text.split(';')
    for m in motivators:
        if re.search('.*compete.*', m.lower()) or re.search('.*capab.*', m.lower()) or \
                re.search('.*profici.*', m.lower()) or re.search('.*expert.*', m.lower()) or \
                re.search('.*skill.*', m.lower()) or re.search('.*effect.*', m.lower()) or \
                re.search('.*producti.*', m.lower()) or re.search('.*effect.*', m.lower()) or \
                re.search('.*confiden.*', m.lower()) or re.search('.*determination.*', m.lower()) or \
                re.search('.*ability.*', m.lower()):
            if name not in primary_motivators['Competence'].keys():
                primary_motivators['Competence'][name] = []
                primary_motivators['Competence'][name].append(m.strip())
            else:
                primary_motivators['Competence'][name].append(m.strip())
        elif re.search('.*live my life.*', m.lower()) or re.search('.*live the life i want.*', m.lower()) \
                or re.search('.*Myself*', m.lower()) or re.search('.*\bme\b.*', m.lower()) \
                or re.search('.*employee.*', m.lower()):
            if name not in primary_motivators['Self'].keys():
                primary_motivators['Self'][name] = []
                primary_motivators['Self'][name].append(m.strip())
            else:
                primary_motivators['Self'][name].append(m.strip())
        elif re.search('.*autonom*', m.lower()) or re.search('.*indepen.*', m.lower()) \
                or re.search('.*freedom*', m.lower()) or re.search('.freewill*', m.lower()):
            if name not in primary_motivators['Autonomy'].keys():
                primary_motivators['Autonomy'][name] = []
                primary_motivators['Autonomy'][name].append(m.strip())
            else:
                primary_motivators['Autonomy'][name].append(m.strip())
        elif re.search('.*culture.*', m.lower()) or re.search('.*principles.*', m.lower()) or \
                re.search('.*ethos.*', m.lower()) or re.search('.*philosoph.*', m.lower()) or \
                re.search('.*mission.*', m.lower()) or re.search('.*values.*', m.lower()):
            if name not in primary_motivators['Company Culture'].keys():
                primary_motivators['Company Culture'][name] = []
                primary_motivators['Company Culture'][name].append(m.strip())
            else:
                primary_motivators['Company Culture'][name].append(m.strip())
        elif re.search('.*salary.*', m.lower()) or re.search('.*pay.*', m.lower()) or \
                re.search('.*money.*', m.lower()) or re.search('.*hike.*', m.lower()) or \
                re.search('.*incentive.*', m.lower()) or re.search('.*paid.*', m.lower()):
            if name not in primary_motivators['Salary'].keys():
                primary_motivators['Salary'][name] = []
                primary_motivators['Salary'][name].append(m.strip())
            else:
                primary_motivators['Salary'][name].append(m.strip())
        elif re.search('.*role.*', m.lower()) or re.search('.*occupation.*', m.lower()) \
                or re.search('.*profess.*', m.lower()) or re.search('.*post.*', m.lower()) \
                or re.search('.*position.*', m.lower()) or re.search('.*assignment.*', m.lower()) \
                or re.search('.*activity.*', m.lower()):
            if name not in primary_motivators['Role'].keys():
                primary_motivators['Role'][name] = []
                primary_motivators['Role'][name].append(m.strip())
            else:
                primary_motivators['Role'][name].append(m.strip())
        elif re.search('.*fullfil.*', m.lower()) or re.search('.*accomplish.*', m.lower()) \
                or re.search('.*satisf.*', m.lower()):
            if name not in primary_motivators['Fulfillment'].keys():
                primary_motivators['Fulfillment'][name] = []
                primary_motivators['Fulfillment'][name].append(m.strip())
            else:
                primary_motivators['Fulfillment'][name].append(m.strip())
        elif re.search('.*religion.*', m.lower()):
            if name not in primary_motivators['Religion'].keys():
                primary_motivators['Religion'][name] = []
                primary_motivators['Religion'][name].append(m.strip())
            else:
                primary_motivators['Religion'][name].append(m.strip())
        elif re.search('.*influenc.*', m.lower()) or re.search('.*difference.*', m.lower()) \
                or re.search('.*betterment.*', m.lower()) or re.search('.*bettering.*', m.lower()):
            if name not in primary_motivators['Influence'].keys():
                primary_motivators['Influence'][name] = []
                primary_motivators['Influence'][name].append(m.strip())
            else:
                primary_motivators['Influence'][name].append(m.strip())
        elif re.search('.*growth.*', m.lower()) or re.search('.*promot.*', m.lower()) or \
                re.search('.*progress.*', m.lower()):
            if name not in primary_motivators['Growth'].keys():
                primary_motivators['Growth'][name] = []
                primary_motivators['Growth'][name].append(m.strip())
            else:
                primary_motivators['Growth'][name].append(m.strip())
        elif re.search('.*work.*', m.lower()) or re.search('.*tasks.*', m.lower()) \
                or re.search('.*industry.*', m.lower()) or re.search('.*produc.*', m.lower()) \
                or re.search('.*job.*', m.lower()) or re.search('.*process.*', m.lower()):
            if name not in primary_motivators['Work'].keys():
                primary_motivators['Work'][name] = []
                primary_motivators['Work'][name].append(m.strip())
            else:
                primary_motivators['Work'][name].append(m.strip())
        elif re.search('.*team.*', m.lower()) or re.search('.*colleag', m.lower()) or \
                re.search('.*partner', m.lower()) or re.search('.*member', m.lower()) or \
                re.search('.*co.*work', m.lower()) or re.search('.*patient', m.lower()):
            if name not in primary_motivators['People'].keys():
                primary_motivators['People'][name] = []
                primary_motivators['People'][name].append(m.strip())
            else:
                primary_motivators['People'][name].append(m.strip())

        elif re.search('.*mentor.*', m.lower()) or re.search('.*guide.*', m.lower()) \
                or re.search('.*advice.*', m.lower()) or re.search('.*counsel.*', m.lower()) or \
                re.search('.*teach.*', m.lower()):
            if name not in primary_motivators['Mentorship'].keys():
                primary_motivators['Mentorship'][name] = []
                primary_motivators['Mentorship'][name].append(m.strip())
            else:
                primary_motivators['Mentorship'][name].append(m.strip())
        elif re.search('.*leader.*', m.lower()) or re.search('.*mentor.*', m.lower()) \
                or re.search('.*supervisor.*', m.lower()) or re.search('.*manag.*', m.lower()):
            if name not in primary_motivators['Leadership'].keys():
                primary_motivators['Leadership'][name] = []
                primary_motivators['Leadership'][name].append(m.strip())
            else:
                primary_motivators['Leadership'][name].append(m.strip())
        elif re.search('.*value.*', m.lower()) or re.search('.*appreicat.*', m.lower()):
            if name not in primary_motivators['Appreciation'].keys():
                primary_motivators['Appreciation'][name] = []
                primary_motivators['Appreciation'][name].append(m.strip())
            else:
                primary_motivators['Appreciation'][name].append(m.strip())
        elif re.search('.*help.*', m.lower()) or re.search('.*assist.*', m.lower()) or \
                re.search('.*aid.*', m.lower()) or re.search('.*service.*', m.lower()) or \
                re.search('.*support.*', m.lower()) or re.search('.*contribut.*', m.lower()):
            if name not in primary_motivators['Helping'].keys():
                primary_motivators['Helping'][name] = []
                primary_motivators['Helping'][name].append(m.strip())
            else:
                primary_motivators['Helping'][name].append(m.strip())
        elif re.search('.*integrity.*', m.lower()) or re.search('.*honest.*', m.lower()):
            if name not in primary_motivators['Integrity'].keys():
                primary_motivators['Integrity'][name] = []
                primary_motivators['Integrity'][name].append(m.strip())
            else:
                primary_motivators['Integrity'][name].append(m.strip())
        elif re.search('.*family.*', m.lower()) or re.search('.*wife.*', m.lower()) \
                or re.search('.*child.*', m.lower()):
            if name not in primary_motivators['Integrity'].keys():
                primary_motivators['Family'][name] = []
                primary_motivators['Family'][name].append(m.strip())
            else:
                primary_motivators['Family'][name].append(m.strip())
        elif re.search('.*learn.*', m.lower()) or re.search('.*educat.*', m.lower()):
            if name not in primary_motivators['Learning'].keys():
                primary_motivators['Learning'][name] = []
                primary_motivators['Learning'][name].append(m.strip())
            else:
                primary_motivators['Learning'][name].append(m.strip())
        elif re.search('.*rewar.*', m.lower()) or re.search('.*award.*', m.lower()):
            if name not in primary_motivators['Rewards'].keys():
                primary_motivators['Rewards'][name] = []
                primary_motivators['Rewards'][name].append(m.strip())
            else:
                primary_motivators['Rewards'][name].append(m.strip())
        elif re.search('.*competition.*', m.lower()) or re.search('.*rivalry.*', m.lower()):
            if name not in primary_motivators['Competition'].keys():
                primary_motivators['Competition'][name] = []
                primary_motivators['Competition'][name].append(m.strip())
            else:
                primary_motivators['Competition'][name].append(m.strip())
        elif re.search('.*success.*', m.lower()) or re.search('.*accomplish.*', m.lower()) \
                or re.search('.*achieve.*', m.lower()) or re.search('.*victor.*', m.lower()) \
                or re.search('.*win.*', m.lower()) or re.search('.*progress.*', m.lower()):
            if name not in primary_motivators['Success'].keys():
                primary_motivators['Success'][name] = []
                primary_motivators['Success'][name].append(m.strip())
            else:
                primary_motivators['Success'][name].append(m.strip())
        elif re.search('.*goal.*', m.lower()) or re.search('.*intention.*', m.lower()) \
                or re.search('.*sale.*', m.lower()) or re.search('.*target.*', m.lower()) \
                or re.search('.*objective.*', m.lower()) or re.search('.*desire.*', m.lower()):
            if name not in primary_motivators['Goals'].keys():
                primary_motivators['Goals'][name] = []
                primary_motivators['Goals'][name].append(m.strip())
            else:
                primary_motivators['Goals'][name].append(m.strip())
        elif re.search('.*interact.*', m.lower()) or re.search('.*communicat.*', m.lower()) \
                or re.search('.*talk.*', m.lower()) or re.search('.*exchange.*', m.lower()) \
                or re.search('.*reciproc.*', m.lower()) or re.search('.*relationship.*', m.lower()):
            if name not in primary_motivators['Interaction'].keys():
                primary_motivators['Interaction'][name] = []
                primary_motivators['Interaction'][name].append(m.strip())
            else:
                primary_motivators['Interaction'][name].append(m.strip())
        elif re.search('.*implement.*', m.lower()) or re.search('.*perform.*', m.lower()) \
                or re.search('.*carry.*out.*', m.lower()) or re.search('.*deed.*', m.lower()) \
                or re.search('.*execut.*', m.lower()):
            if name not in primary_motivators['Action'].keys():
                primary_motivators['Action'][name] = []
                primary_motivators['Action'][name].append(m.strip())
            else:
                primary_motivators['Action'][name].append(m.strip())

    return


def process_recommendation(text, name):
    recommendations = text.split(';')
    for r in recommendations:
        if re.search('.*versatil.*', r.lower()) or re.search('.*flexib.*', r.lower()) or \
                re.search('.*round.*', r.lower()) or re.search('.*adapt.*', r.lower()):
            if name not in good_to_great_factors['Versatility'].keys():
                good_to_great_factors['Versatility'][name] = []
                good_to_great_factors['Versatility'][name].append(r.strip())
            else:
                good_to_great_factors['Versatility'][name].append(r.strip())
        elif re.search('.*rewar.*', r.lower()) or re.search('.*award.*', r.lower()):
            if name not in good_to_great_factors['Rewards'].keys():
                good_to_great_factors['Rewards'][name] = []
                good_to_great_factors['Rewards'][name].append(r.strip())
            else:
                good_to_great_factors['Rewards'][name].append(r.strip())
        elif re.search('.*recogni.*', r.lower()) or re.search('.*fame.*', r.lower()) \
                or re.search('.*greatness.*', r.lower()) or re.search('.*importance.*', r.lower()):
            if name not in good_to_great_factors['Recognition'].keys():
                good_to_great_factors['Recognition'][name] = []
                good_to_great_factors['Recognition'][name].append(r.strip())
            else:
                good_to_great_factors['Recognition'][name].append(r.strip())
        elif re.search('.*cult.*', r.lower()) or re.search('.*workplace.*', r.lower()):
            if name not in good_to_great_factors['Culture'].keys():
                good_to_great_factors['Culture'][name] = []
                good_to_great_factors['Culture'][name].append(r.strip())
            else:
                good_to_great_factors['Culture'][name].append(r.strip())
        elif re.search('.*feedback.*', r.lower()) or re.search('.*assess.*', r.lower()):
            if name not in good_to_great_factors['Feedback'].keys():
                good_to_great_factors['Feedback'][name] = []
                good_to_great_factors['Feedback'][name].append(r.strip())
            else:
                good_to_great_factors['Feedback'][name].append(r.strip())
        elif re.search('.*work.*', r.lower()) or re.search('.*challen.*', r.lower()) \
                or re.search('.*challen.*', r.lower()):
            if name not in good_to_great_factors['Work'].keys():
                good_to_great_factors['Work'][name] = []
                good_to_great_factors['Work'][name].append(r.strip())
            else:
                good_to_great_factors['Work'][name].append(r.strip())
        elif re.search('.*value.*', r.lower()) or re.search('.*appreciat.*', r.lower()):
            if name not in good_to_great_factors['Appreciation'].keys():
                good_to_great_factors['Appreciation'][name] = []
                good_to_great_factors['Appreciation'][name].append(r.strip())
            else:
                good_to_great_factors['work'][name].append(r.strip())
        elif re.search('.*opportun.*', r.lower()) or re.search('.*roles.*', r.lower()):
            if name not in good_to_great_factors['Opportunities'].keys():
                good_to_great_factors['Opportunities'][name] = []
                good_to_great_factors['Opportunities'][name].append(r.strip())
            else:
                good_to_great_factors['Opportunities'][name].append(r.strip())
        elif re.search('.*mentor.*', r.lower()) or re.search('.*encour.*', r.lower()) or \
                re.search('.*motiv.*', r.lower()):
            if name not in good_to_great_factors['Mentoring'].keys():
                good_to_great_factors['Mentoring'][name] = []
                good_to_great_factors['Mentoring'][name].append(r.strip())
            else:
                good_to_great_factors['Mentoring'][name].append(r.strip())
        elif re.search('.*commit.*', r.lower()) or re.search('.*trust.*', r.lower()) or \
                re.search('.*loyal.*', r.lower()) or re.search('.*dependab.*', r.lower()) or \
                re.search('.*dedicat.*', r.lower()) or re.search('.*account.*', r.lower()) or \
                re.search('.*follow through.*', r.lower()):
            if name not in good_to_great_factors['Commitment'].keys():
                good_to_great_factors['Commitment'][name] = []
                good_to_great_factors['Commitment'][name].append(r.strip())
            else:
                good_to_great_factors['Commitment'][name].append(r.strip())
        elif re.search('.*perform.*', r.lower()) or re.search('.*excel.*', r.lower()) or \
                re.search('.*exceed.*', r.lower()) or re.search('.*surpass.*', r.lower()) or \
                re.search('.*going above and beyond.*', r.lower()):
            if name not in good_to_great_factors['Performance'].keys():
                good_to_great_factors['Performance'][name] = []
                good_to_great_factors['Performance'][name].append(r.strip())
            else:
                good_to_great_factors['Performance'][name].append(r.strip())
        elif re.search('.*compete.*', r.lower()) or re.search('.*capab.*', r.lower()) or \
                re.search('.*profici.*', r.lower()) or re.search('.*expert.*', r.lower()) or \
                re.search('.*skill.*', r.lower()) or re.search('.*effect.*', r.lower()) or \
                re.search('.*producti.*', r.lower()) or re.search('.*effect.*', r.lower()) or \
                re.search('.*confiden.*', r.lower() or re.search('.*determination.*', r.lower())):
            if name not in good_to_great_factors['Competence'].keys():
                good_to_great_factors['Competence'][name] = []
                good_to_great_factors['Competence'][name].append(r.strip())
            else:
                good_to_great_factors['Competence'][name].append(r.strip())
        elif re.search('.*support.*', r.lower()) or re.search('.*assist.*', r.lower()) or \
                re.search('.*help.*', r.lower()):
            if name not in good_to_great_factors['Support'].keys():
                good_to_great_factors['Support'][name] = []
                good_to_great_factors['Support'][name].append(r.strip())
            else:
                good_to_great_factors['Support'][name].append(r.strip())
        elif re.search('.*intellig.*', r.lower()) or re.search('.*intellect.*', r.lower()) or \
                re.search('.*acumen.*', r.lower()) or re.search('.*aptitude.*', r.lower()):
            if name not in good_to_great_factors['Intelligence'].keys():
                good_to_great_factors['Intelligence'][name] = []
                good_to_great_factors['Intelligence'][name].append(r.strip())
            else:
                good_to_great_factors['Intelligence'][name].append(r.strip())
        elif re.search('.*creat.*', r.lower()) or re.search('.*outside the box.*', r.lower()):
            if name not in good_to_great_factors['Creativity'].keys():
                good_to_great_factors['Creativity'][name] = []
                good_to_great_factors['Creativity'][name].append(r.strip())
            else:
                good_to_great_factors['Creativity'][name].append(r.strip())
    # for key in good_to_great_factors:
    #    for attributes in good_to_great_factors[key]:
    #        print(attributes)


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


def get_employee_benefits(text):
    employee_benefits = {}
    items = text.split(',')
    retirement_pattern = re.compile('.*40.*|roth|ira|prov.*|.*retir|.*pension|.*gratuity')
    med_insurance_pattern = re.compile('.*medi.*|.*health.*')
    espp_pattern = re.compile('.*espp|.*share.*')
    life_insurance_pattern = re.compile('.*life.*')
    disability_insurance_pattern = re.compile('.*disability.*')
    vision_pattern = re.compile('.*vision.*')
    dental_pattern = re.compile('.*dental.*')
    vacation_pattern = re.compile('.*pto|.*sick|.*paid|.*volun.*')
    parental_pattern = re.compile('.*maternal|.*paternal|.*parental')
    emp_wellness_pattern = re.compile('.*wellness|.*yoga|.*exerci|.*hiit.*|.*tabata|.*fit')
    pet_insurance_pattern = re.compile('.*pet.*')
    company_car_pattern = re.compile('.*company car.*|.*company vehicle')
    air_travel_pattern = re.compile('.*air*|.*aero|.*plane')
    work_amenities_pattern = re.compile('.*cafeteria|.*gym')
    accom_allowance_pattern = re.compile('.*accom*')
    flex_work_pattern = re.compile('.*flex*')
    severance_pattern = re.compile('.*sever*')
    work_comp_pattern = re.compile('.*compens*')
    education_allowance_pattern = re.compile('.*educ*|.*tuition')
    entertainment_pattern = re.compile('.*entertainment*')
    for benefit in items:
        if retirement_pattern.match(benefit.lower()):
            if 'Retirement Plan' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Retirement Plan'] = 0
        elif med_insurance_pattern.match(benefit.lower()):
            if 'Medical Insurance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Medical Insurance'] = 0
        elif entertainment_pattern.match(benefit.lower()):
            if 'Entertainment' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Entertainment'] = 0
        elif espp_pattern.match(benefit.lower()):
            if 'ESPP' in employee_benefits.keys():
                continue
            else:
                employee_benefits['ESPP'] = 0
        elif life_insurance_pattern.match(benefit.lower()):
            if 'Life Insurance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Life Insurance'] = 0
        elif disability_insurance_pattern.match(benefit.lower()):
            if 'Disability Insurance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Disability Insurance'] = 0
        elif vision_pattern.match(benefit.lower()):
            if 'Vision' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Vision'] = 0
        elif dental_pattern.match(benefit.lower()):
            if 'Dental' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Dental'] = 0
        elif vacation_pattern.match(benefit.lower()):
            if 'Paid Vacation' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Paid Vacation'] = 0
        elif parental_pattern.match(benefit.lower()):
            if 'Parental Leave' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Parental Leave'] = 0
        elif emp_wellness_pattern.match(benefit.lower()):
            if 'Employee Wellness' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Employee Wellness'] = 0
        elif pet_insurance_pattern.match(benefit.lower()):
            if 'Pet Insurance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Pet Insurance'] = 0
        elif company_car_pattern.match(benefit.lower()):
            if 'Company Car' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Company Car'] = 0
        elif air_travel_pattern.match(benefit.lower()):
            if 'Air Travel' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Air Travel'] = 0
        elif work_amenities_pattern.match(benefit.lower()):
            if 'Workplace Amenities' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Workplace Amenities'] = 0
        elif accom_allowance_pattern.match(benefit.lower()):
            if 'Accommodation Allowance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Accommodation Allowance'] = 0
        elif flex_work_pattern.match(benefit.lower()):
            if 'Work Flexibility' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Work Flexibility'] = 0
        elif severance_pattern.match(benefit.lower()):
            if 'Severance' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Severance'] = 0
        elif work_comp_pattern.match(benefit.lower()):
            if 'Work Compensation' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Work Compensation'] = 0
        elif education_allowance_pattern.match(benefit.lower()):
            if 'Education' in employee_benefits.keys():
                continue
            else:
                employee_benefits['Education'] = 0
    return employee_benefits


def process_excel():
    excel_file = pd.ExcelFile(file)
    # print(excel_file.sheet_names)
    for name in excel_file.sheet_names:
        # if name == 'Marissa Prakash' or name == 'Cate Hite' or name == 'Gary Johnson' or name == 'Kimberly Johnson' or \
        #        name == 'Ravi Prakash':
        # if name == 'MM':

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
            m11 = re.search(r'what benefits are current.*', str(df['Questions'][index]).lower())
            m12 = re.search(r'what motivates you.*', str(df['Questions'][index]).lower())
            m13 = re.search(r'are you satisfied in.*', str(df['Questions'][index]).lower())
            m14 = re.search(r'what is your highest level of educ', str(df['Questions'][index]).lower())
            m69 = re.search(r'Any recommendations on what.*', str(df['Questions'][index]))
            m15 = re.search(r'what is your primary motivator.*', str(df['Questions'][index]).lower())
            m16 = re.search(r'do you believe you\'re a great employee.*', str(df['Questions'][index]).lower())

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
            elif m11:
                interviewee_dict[name]['Benefits'] = get_employee_benefits(str(df['Responses'][index]))
            elif m12:
                process_personal_motivators(str(df['Responses'][index]), name)
            elif m13:
                process_personal_satisfaction(str(df['Responses'][index]), name)
            elif m14:
                interviewee_dict[name]['Education Level'] = str(df['Responses'][index])
                process_education_level(str(df['Responses'][index]), name)
            elif m69:
                process_recommendation(str(df['Responses'][index]), name)
            elif m15:
                process_primary_motivators(str(df['Responses'][index]), name)
            elif m16:
                process_employee_valuation(str(df['Responses'][index]), 'great', name)
            elif m1:
                continue


def get_count_of_attribute(attribute_dict):
    count = 0
    for name in attribute_dict:
        l = len(attribute_dict[name])
        count = count + l
    return count


def post_process():
    benefits_count_dict = {}
    retirement_count = 0
    medical_insurance_count = 0
    espp_count = 0
    vision_count = 0
    dental_count = 0
    vacation_count = 0
    life_insurance_count = 0
    disability_count = 0
    severance_count = 0
    parental_count = 0
    emp_wellness_count = 0
    flexwork_count = 0
    pet_insurance_count = 0
    company_car_count = 0
    air_travel_count = 0
    work_amen_count = 0
    accom_allow_count = 0
    work_comp_count = 0
    educ_allow_count = 0
    entertainment_count = 0
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
        for benefit in interviewee_dict[person]['Benefits']:
            if benefit == 'Retirement Plan':
                retirement_count += 1
            elif benefit == 'Medical Insurance':
                medical_insurance_count += 1
            elif benefit == 'ESPP':
                espp_count += 1
            elif benefit == 'ESPP':
                espp_count += 1
            elif benefit == 'Vision':
                vision_count += 1
            elif benefit == 'Dental':
                dental_count += 1
            elif benefit == 'Paid Vacation':
                vacation_count += 1
            elif benefit == 'Life Insurance':
                life_insurance_count += 1
            elif benefit == 'Disability Insurance':
                disability_count += 1
            elif benefit == 'Severance':
                severance_count += 1
            elif benefit == 'Parental Leave':
                parental_count += 1
            elif benefit == 'Employee Wellness':
                emp_wellness_count += 1
            elif benefit == 'Work Flexibility':
                flexwork_count += 1
            elif benefit == 'Pet Insurance':
                pet_insurance_count += 1
            elif benefit == 'Company Car':
                company_car_count += 1
            elif benefit == 'Air Travel':
                air_travel_count += 1
            elif benefit == 'Workplace Amenities':
                work_amen_count += 1
            elif benefit == 'Accommodation Allowance':
                accom_allow_count += 1
            elif benefit == 'Work Compensation':
                work_comp_count += 1
            elif benefit == 'Education':
                educ_allow_count += 1
            elif benefit == 'Entertainment':
                entertainment_count += 1

    benefits_count_dict['Retirement Plan'] = retirement_count
    benefits_count_dict['Medical Insurance'] = medical_insurance_count
    benefits_count_dict['ESPP'] = espp_count
    benefits_count_dict['Vision'] = vision_count
    benefits_count_dict['Dental'] = dental_count
    benefits_count_dict['Paid Vacation'] = vacation_count
    benefits_count_dict['Life Insurance'] = life_insurance_count
    benefits_count_dict['Disability Insurance'] = disability_count
    benefits_count_dict['Severance'] = severance_count
    benefits_count_dict['Parental Leave'] = parental_count
    benefits_count_dict['Employee Wellness'] = emp_wellness_count
    benefits_count_dict['Work Flexibility'] = flexwork_count
    benefits_count_dict['Pet Insurance'] = pet_insurance_count
    benefits_count_dict['Company Car'] = company_car_count
    benefits_count_dict['Air Travel'] = air_travel_count
    benefits_count_dict['Workplace Amenities'] = work_amen_count
    benefits_count_dict['Accommodation Allowance'] = accom_allow_count
    benefits_count_dict['Work Compensation'] = work_comp_count
    benefits_count_dict['Education'] = educ_allow_count
    benefits_count_dict['Entertainment'] = entertainment_count

    print("Benefits Analysis")
    print("-------------------------------------------------------------------------------------------")

    for key in benefits_count_dict:
        if benefits_count_dict[key] > 0:
            print("{0} benefits were brought up {1} time(s)".format(key, benefits_count_dict[key]))
    print("\n")
    sorted_benefits_by_count = sorted(benefits_count_dict.items(), key=lambda x: x[1],
                                      reverse=True)
    print("Benefits listed in the order of highest to the least counts")
    for i in sorted_benefits_by_count:
        if i[1] > 0:
            print("\t{0} {1}".format(i[0], i[1]))
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
    plt.title('Analysis of Employee Valuation', fontsize=20)
    for index, data in enumerate(valuation_count_yaxis):
        plt.text(x=index, y=data + 1, s=f"{data}", fontsize='20')
    plt.savefig(r'C:\Users\kp409917\Desktop\Employee_Valuation.png')

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

 


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    process_excel()
    post_process()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
