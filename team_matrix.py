#!/usr/bin/env python3

'''This module will programmatically create the Team Matrix Excel file'''

import xlsxwriter
import strengths

def build_names_and_strengths(wb, ws, info):
    # write the header row
    ws.write_string(0, 0, 'Name')
    for theme in range(1, 35):
        ws.write(0, theme, 'Theme {}'.format(theme))
    for row, person in enumerate(info):
        for col in range(len(person)):
            ws.write_string(row+1, col, person[col])
    # leave a blank row and write the lookup row
    sname_row = len(info) + 2
    center_fmt = wb.add_format({'align': 'center'})
    [ws.write_string(sname_row, col+1, strength, center_fmt) for col, strength in enumerate(strengths.all34)]
    # build the top 10 counts
    row_start = len(info) + 3
    for row in range(row_start, row_start+len(info)):
        # the first column is the name reference
        source_row = row-row_start+2
        ws.write_formula(row, 0, '=A{0:d}'.format(source_row))
        # the next columns count if the Strength occurs in the first 10 columns (top 10)
        for col in range(1, 35):
            formula = '=COUNTIF(B{0:d}:K{0:d}, "{1:s}")'.format(source_row, strengths.all34[col-1])
            ws.write_formula(row, col, formula, center_fmt)
    # autofit the columns based on the longest word in each column
    for col in range(35):  # name + 34 strengths
        width = max([len(s[col]) for s in info])
        ws.set_column(col, col, width)


def generate(fname, info):
    '''generate the team matrix Excel document
    @arg fname: the filename to save the Excel doc todo
    @arg info: a sequence of Name, Strengths[1:35]
    '''
    wb = xlsxwriter.Workbook(fname)
    ws_matrix = wb.add_worksheet('Matrix')
    ws_ns = wb.add_worksheet('Names and Strengths')
    build_names_and_strengths(wb, ws_ns, info)
    wb.close()

if __name__ == '__main__':
    info = [['Adam Williams', 'Achiever', 'Relator', 'Activator', 'Arranger', 'Belief', 'Responsibility', 'Strategic', 'Learner', 'Analytical', 'Intellection', 'Developer', 'Maximizer', 'Positivity', 'Futuristic', 'Command', 'Connectedness', 'Focus', 'Competition', 'Individualization', 'Context', 'Ideation', 'Self-Assurance', 'Input', 'Deliberative', 'Significance', 'Restorative', 'Harmony', 'Discipline', 'Empathy', 'Communication', 'Consistency', 'Adaptability', 'Includer', 'Woo'], \
            ['Adrianna Bustamante', 'Activator', 'Maximizer', 'Communication', 'Woo', 'Relator', 'Futuristic', 'Positivity', 'Command', 'Self-Assurance', 'Ideation', 'Strategic', 'Focus', 'Achiever', 'Learner', 'Competition', 'Individualization', 'Responsibility', 'Arranger', 'Empathy', 'Includer', 'Belief', 'Adaptability', 'Developer', 'Significance', 'Connectedness', 'Input', 'Intellection', 'Discipline', 'Deliberative', 'Analytical', 'Context', 'Consistency', 'Harmony', 'Restorative'], \
            ['Antoinette Lindon', 'Learner', 'Ideation', 'Futuristic', 'Analytical', 'Achiever', 'Intellection', 'Relator', 'Command', 'Competition', 'Activator', 'Input', 'Strategic', 'Self-Assurance', 'Deliberative', 'Maximizer', 'Individualization', 'Responsibility', 'Connectedness', 'Focus', 'Belief', 'Arranger', 'Discipline', 'Adaptability', 'Significance', 'Developer', 'Restorative', 'Includer', 'Context', 'Empathy', 'Consistency', 'Communication', 'Positivity', 'Harmony', 'Woo'], \
            ['Elizabeth Marrone', 'Responsibility', 'Deliberative', 'Relator', 'Learner', 'Discipline', 'Analytical', 'Individualization', 'Intellection', 'Restorative', 'Achiever', 'Harmony', 'Input', 'Consistency', 'Arranger', 'Belief', 'Connectedness', 'Developer', 'Focus', 'Command', 'Empathy', 'Maximizer', 'Futuristic', 'Adaptability', 'Strategic', 'Positivity', 'Self-Assurance', 'Competition', 'Ideation', 'Activator', 'Communication', 'Significance', 'Context', 'Woo', 'Includer'], \
            ['Janice Krpan', 'Strategic', 'Relator', 'Activator', 'Belief', 'Significance', 'Futuristic', 'Connectedness', 'Empathy', 'Positivity', 'Achiever', 'Arranger', 'Developer', 'Discipline', 'Ideation', 'Maximizer', 'Responsibility', 'Deliberative', 'Command', 'Woo', 'Communication', 'Analytical', 'Focus', 'Self-Assurance', 'Consistency', 'Adaptability', 'Context', 'Competition', 'Restorative', 'Harmony', 'Includer', 'Input', 'Intellection', 'Individualization', 'Learner'], \
            ['Judy Vansell', 'Achiever', 'Maximizer', 'Arranger', 'Individualization', 'Input', 'Relator', 'Communication', 'Includer', 'Focus', 'Connectedness', 'Strategic', 'Learner', 'Responsibility', 'Activator', 'Significance', 'Woo', 'Discipline', 'Analytical', 'Futuristic', 'Intellection', 'Positivity', 'Competition', 'Empathy', 'Self-Assurance', 'Ideation', 'Harmony', 'Developer', 'Consistency', 'Belief', 'Deliberative', 'Adaptability', 'Command', 'Context', 'Restorative'], \
            ['Lisa Mclin', 'Restorative', 'Relator', 'Learner', 'Developer', 'Positivity', 'Achiever', 'Arranger', 'Empathy', 'Includer', 'Belief', 'Strategic', 'Focus', 'Responsibility', 'Connectedness', 'Discipline', 'Futuristic', 'Consistency', 'Self-Assurance', 'Context', 'Woo', 'Input', 'Harmony', 'Intellection', 'Communication', 'Competition', 'Analytical', 'Individualization', 'Significance', 'Deliberative', 'Maximizer', 'Adaptability', 'Activator', 'Command', 'Ideation'], \
            ['Matt Stoyka', 'Responsibility', 'Relator', 'Activator', 'Futuristic', 'Achiever', 'Positivity', 'Self-Assurance', 'Arranger', 'Command', 'Strategic', 'Communication', 'Includer', 'Woo', 'Significance', 'Learner', 'Competition', 'Focus', 'Analytical', 'Belief', 'Input', 'Ideation', 'Maximizer', 'Intellection', 'Connectedness', 'Individualization', 'Developer', 'Discipline', 'Consistency', 'Restorative', 'Adaptability', 'Deliberative', 'Context', 'Harmony', 'Empathy'], \
            ['Peter Fitzgibbon', 'Individualization', 'Arranger', 'Restorative', 'Achiever', 'Relator', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''], \
            ['Peter Pixton', 'Context', 'Learner', 'Relator', 'Achiever', 'Analytical', 'Input', 'Competition', 'Intellection', 'Activator', 'Belief', 'Individualization', 'Responsibility', 'Self-Assurance', 'Focus', 'Arranger', 'Deliberative', 'Significance', 'Restorative', 'Discipline', 'Harmony', 'Command', 'Futuristic', 'Strategic', 'Ideation', 'Maximizer', 'Developer', 'Consistency', 'Connectedness', 'Woo', 'Communication', 'Adaptability', 'Positivity', 'Includer', 'Empathy'], \
            ['Rachel Woodson', 'Learner', 'Restorative', 'Individualization', 'Command', 'Input', 'Relator', 'Analytical', 'Arranger', 'Achiever', 'Intellection', 'Self-Assurance', 'Adaptability', 'Deliberative', 'Futuristic', 'Ideation', 'Focus', 'Responsibility', 'Developer', 'Empathy', 'Connectedness', 'Harmony', 'Strategic', 'Woo', 'Activator', 'Positivity', 'Context', 'Belief', 'Significance', 'Maximizer', 'Competition', 'Discipline', 'Consistency', 'Includer', 'Communication']]
    generate('test.xlsx', info)
