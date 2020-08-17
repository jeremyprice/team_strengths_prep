#!/usr/bin/env python3

'''This module will programmatically create the Team Matrix Excel file'''

import xlsxwriter
import strengths
from xlsxwriter.utility import xl_rowcol_to_cell

def write_domain_header(wb, ws, row, start_col, domain):
    first_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                               'rotation': 60, 'font_color': 'white', 'bg_color': '#25457E', 'top': 2,
                               'left': 2})
    mid_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                             'rotation': 60, 'font_color': 'white', 'bg_color': '#25457E', 'top': 2})
    last_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                              'rotation': 60, 'font_color': 'white', 'bg_color': '#25457E', 'top': 2,
                              'right': 2})
    last = len(domain) - 1
    for idx, s in enumerate(domain):
        if idx == 0:
            fmt = first_fmt
        elif idx == last:
            fmt = last_fmt
        else:
            fmt = mid_fmt
        ws.write_string(row, start_col + idx, s, fmt)

def build_matrix(wb, ws, info):
    # put in the header info and setup the entire worksheet
    fmt = wb.add_format({'valign': 'vcenter', 'bold': True, 'font_size': 16, 'align': 'left'})
    ws.merge_range(0, 0, 1, 2, 'Team Strengths Matrix', fmt)
    fmt = wb.add_format({'valign': 'vcenter', 'bold': True, 'font_size': 11, 'align': 'center'})
    ws.merge_range(2, 0, 2, 2, 'Team Name', fmt)
    ws.hide_gridlines(2)
    # build the domain headers
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12,
                         'valign': 'bottom', 'top': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Executing'),
                         'font_color': strengths.domain_txt_color('Executing')})
    col_end = 2 + len(strengths.executing)
    ws.merge_range(1, 3, 1, col_end, 'Executing', fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12,
                         'valign': 'bottom', 'top': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Strategic Thinking'),
                         'font_color': strengths.domain_txt_color('Strategic Thinking')})
    col_start = col_end + 2
    col_end = col_start + len(strengths.strategic_thinking) - 1
    ws.merge_range(1, col_start, 1, col_end, 'Strategic Thinking', fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12,
                         'valign': 'bottom', 'top': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Influencing'),
                         'font_color': strengths.domain_txt_color('Influencing')})
    col_start = col_end + 2
    col_end = col_start + len(strengths.influencing) - 1
    ws.merge_range(1, col_start, 1, col_end, 'Influencing', fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12,
                         'valign': 'bottom', 'top': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Relationship Building'),
                         'font_color': strengths.domain_txt_color('Relationship Building')})
    col_start = col_end + 2
    col_end = col_start + len(strengths.relationship_building) - 1
    ws.merge_range(1, col_start, 1, col_end, 'Relationship Building', fmt)
    # build the header row of strengths with the black spacer columns
    fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                         'rotation': 60, 'font_color': 'white', 'bg_color': '#25457E',
                         'top': 2, 'left': 2, 'right': 2})
    ws.write_string(3, 0, 'Name', fmt)
    spacer_fmt = wb.add_format({'bg_color': 'black', 'top': 2, 'left': 2, 'right': 2})
    start_col = 1
    write_domain_header(wb, ws, 3, start_col, strengths.executing)
    start_col += len(strengths.executing)
    for row in range(3, len(info)+5):
        ws.write_string(row, start_col, '', spacer_fmt)
    start_col += 1
    write_domain_header(wb, ws, 3, start_col, strengths.strategic_thinking)
    start_col += len(strengths.strategic_thinking)
    for row in range(3, len(info)+5):
        ws.write_string(row, start_col, '', spacer_fmt)
    start_col += 1
    write_domain_header(wb, ws, 3, start_col, strengths.influencing)
    start_col += len(strengths.influencing)
    for row in range(3, len(info)+5):
        ws.write_string(row, start_col, '', spacer_fmt)
    start_col += 1
    write_domain_header(wb, ws, 3, start_col, strengths.relationship_building)
    start_col += len(strengths.relationship_building)
    # set the strengths column width - we add one extra column for the domain headers
    ws.set_column(1, start_col+1, 4)
    # build the name + strengths rows
    name_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                              'font_color': 'white', 'bg_color': '#5A88D6'})
    odd_fmt =  wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'center',
                              'font_color': 'black', 'bg_color': 'white', 'top': 1, 'bottom': 1})
    even_fmt =  wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'center',
                               'font_color': 'black', 'bg_color': '#D9D9D9', 'top': 1, 'bottom': 1})
    count_fmt = wb.add_format({'left': 2})
    for row in range(4, len(info)+4):
        if row % 2 == 0:
            fmt = even_fmt
        else:
            fmt = odd_fmt
        name = "=IF(ISBLANK('Names and Strengths'!A{0:d}), \"\", 'Names and Strengths'!A{0:d})".format(row-2)
        ws.write_formula(row, 0, name, name_fmt)
        bit_row = len(info) + row
        formula = "=IF('Names and Strengths'!{0:s}{1:d}=1,\"x\",\" \")"
        start_col = 1
        for idx in range(len(strengths.executing)):
            col = chr(ord('B') + idx)
            ws.write_formula(row, idx+start_col, formula.format(col, bit_row), fmt)
        start_col += len(strengths.executing) + 1  # space column
        for idx in range(len(strengths.strategic_thinking)):
            col = chr(ord('K') + idx)
            ws.write_formula(row, idx+start_col, formula.format(col, bit_row), fmt)
        start_col += len(strengths.strategic_thinking) + 1  # space column
        for idx in range(len(strengths.influencing)):
            col = chr(ord('S') + idx)
            ws.write_formula(row, idx+start_col, formula.format(col, bit_row), fmt)
        start_col += len(strengths.influencing) + 1  # space column
        for idx in range(len(strengths.relationship_building)):
            col = 'A' + chr(ord('A') + idx)
            ws.write_formula(row, idx+start_col, formula.format(col, bit_row), fmt)
        start_col += len(strengths.relationship_building)
        # the count-check column at the far right
        formula = '=COUNTIF(B{0:d}:AL{0:d},"x")'.format(row + 1)
        ws.write_formula(row, start_col, formula, count_fmt)
    row = len(info) + 4
    total_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'left',
                               'font_color': 'white', 'bg_color': '#5A88D6', 'bottom': 2})
    num_fmt = wb.add_format({'valign': 'bottom', 'bold': True, 'font_size': 10, 'align': 'center',
                             'font_color': 'white', 'bg_color': '#5A88D6', 'border': 2})
    start_col = 0
    ws.write_string(row, start_col, 'Totals', total_fmt)
    start_col = 1
    row_range = (4, len(info)+3)
    insert_totals(ws, row, start_col, start_col+len(strengths.executing), row_range, num_fmt)
    start_col += len(strengths.executing) + 1  # space column
    insert_totals(ws, row, start_col, start_col+len(strengths.strategic_thinking), row_range, num_fmt)
    start_col += len(strengths.strategic_thinking) + 1  # space column
    insert_totals(ws, row, start_col, start_col+len(strengths.influencing), row_range, num_fmt)
    start_col += len(strengths.influencing) + 1  # space column
    insert_totals(ws, row, start_col, start_col+len(strengths.relationship_building), row_range, num_fmt)
    # put the team and domain counts at the bottom left
    row = 33 + len(info)
    ws.write_string(row, 0, 'team n')
    c1 = xl_rowcol_to_cell(4, 1)
    c2 = xl_rowcol_to_cell(3 + len(info), 1)
    formula = '=SUMPRODUCT(--(LEN({0:s}:{1:s})>0))'.format(c1, c2)
    ws.write_formula(row, 1, formula)
    team_n = xl_rowcol_to_cell(row, 1)
    row += 1
    ws.write_string(row, 0, 'Executing n')
    ws.write_formula(row, 1, '=9*{}'.format(team_n))
    exec_n = xl_rowcol_to_cell(row, 1)
    row += 1
    ws.write_string(row, 0, 'Strat Thinking n')
    ws.write_formula(row, 1, '=8*{}'.format(team_n))
    st_think_n = xl_rowcol_to_cell(row, 1)
    row += 1
    ws.write_string(row, 0, 'Influencing n')
    ws.write_formula(row, 1, '=8*{}'.format(team_n))
    influencing_n = xl_rowcol_to_cell(row, 1)
    row += 1
    ws.write_string(row, 0, 'Rel Building n')
    ws.write_formula(row, 1, '=9*{}'.format(team_n))
    rel_build_n = xl_rowcol_to_cell(row, 1)
    # add the domain breakdown for the chart
    row += 1
    ws.write_string(row, 1, 'Domain Breakdown')
    db_cell = xl_rowcol_to_cell(row, 1)
    row += 1
    percent_fmt = wb.add_format({'num_format': '0%'})
    p_col = 3
    for offset, domain in enumerate(strengths.domains):
        ws.write_string(row + offset, 0, domain)
        p_addr = xl_rowcol_to_cell(2, p_col)
        ws.write_formula(row + offset, 1, '={}'.format(p_addr), percent_fmt)
        p_col += strengths.domain_counts[domain] + 1
    label_cells = (xl_rowcol_to_cell(row, 0), xl_rowcol_to_cell(row+3, 0))
    p_cells = (xl_rowcol_to_cell(row, 1), xl_rowcol_to_cell(row+3, 1))
    # add the chart
    chart = wb.add_chart({'type': 'doughnut'})
    chart.add_series({
        'name': '=Matrix!{}'.format(db_cell),
        'categories': "=Matrix!{}:{}".format(*label_cells),
        'values': "=Matrix!{}:{}".format(*p_cells),
        'points': [{'fill': {'color': strengths.domain_color(domain)}} for domain in strengths.domain_order]
    })
    chart.set_legend({'none': True})
    chart.set_style(10)
    chart_loc = xl_rowcol_to_cell(len(info)+10, 0)
    ws.insert_chart(chart_loc, chart)
    # put the domain header calculations up top - TODO: need to fix this
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12, 'num_format': '0%',
                         'valign': 'bottom', 'bottom': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Executing'),
                         'font_color': strengths.domain_txt_color('Executing')})
    col_end = 2 + len(strengths.executing)
    r = (xl_rowcol_to_cell(4, 1), xl_rowcol_to_cell(3+len(info), len(strengths.executing)))
    ws.merge_range(2, 3, 2, col_end, '=(COUNTIF({}:{},"x"))/{}'.format(r[0], r[1], exec_n), fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12, 'num_format': '0%',
                         'valign': 'bottom', 'bottom': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Strategic Thinking'),
                         'font_color': strengths.domain_txt_color('Strategic Thinking')})
    col_start = col_end + 2
    col_end += len(strengths.executing)
    box_col_start = len(strengths.executing) + 2
    box_col_end = box_col_start + len(strengths.strategic_thinking) - 1
    r = (xl_rowcol_to_cell(4, box_col_start), xl_rowcol_to_cell(3+len(info), box_col_end))
    ws.merge_range(2, col_start, 2, col_end, '=(COUNTIF({}:{},"x"))/{}'.format(r[0], r[1], st_think_n), fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12, 'num_format': '0%',
                         'valign': 'bottom', 'bottom': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Influencing'),
                         'font_color': strengths.domain_txt_color('Influencing')})
    col_start = col_end + 2
    col_end += len(strengths.strategic_thinking) + 1
    box_col_start += len(strengths.strategic_thinking) + 1
    box_col_end = box_col_start + len(strengths.influencing) - 1
    r = (xl_rowcol_to_cell(4, box_col_start), xl_rowcol_to_cell(3+len(info), box_col_end))
    ws.merge_range(2, col_start, 2, col_end, '=(COUNTIF({}:{},"x"))/{}'.format(r[0], r[1], influencing_n), fmt)
    fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 12, 'num_format': '0%',
                         'valign': 'bottom', 'bottom': 2, 'left': 2, 'right': 2,
                         'bg_color': strengths.domain_color('Relationship Building'),
                         'font_color': strengths.domain_txt_color('Relationship Building')})
    col_start = col_end + 2
    col_end += len(strengths.influencing) + 2
    box_col_start += len(strengths.influencing) + 1
    box_col_end = box_col_start + len(strengths.relationship_building) - 1
    r = (xl_rowcol_to_cell(4, box_col_start), xl_rowcol_to_cell(3+len(info), box_col_end))
    ws.merge_range(2, col_start, 2, col_end, '=(COUNTIF({}:{},"x"))/{}'.format(r[0], r[1], rel_build_n), fmt)
    # put the percentages at the bottom of each Strength column
    percent_fmt = wb.add_format({'bold': True, 'font_size': 10, 'num_format': '0%'})
    p_row = len(info) + 5
    start_col = 1
    for col in range(start_col, start_col + len(strengths.executing)):
        total = xl_rowcol_to_cell(p_row-1, col)
        ws.write_formula(p_row, col, '={}/{}'.format(total, team_n), percent_fmt)
    start_col += len(strengths.executing) + 1
    for col in range(start_col, start_col + len(strengths.strategic_thinking)):
        total = xl_rowcol_to_cell(p_row-1, col)
        ws.write_formula(p_row, col, '={}/{}'.format(total, team_n), percent_fmt)
    start_col += len(strengths.strategic_thinking) + 1
    for col in range(start_col, start_col + len(strengths.influencing)):
        total = xl_rowcol_to_cell(p_row-1, col)
        ws.write_formula(p_row, col, '={}/{}'.format(total, team_n), percent_fmt)
    start_col += len(strengths.influencing) + 1
    for col in range(start_col, start_col + len(strengths.relationship_building)):
        total = xl_rowcol_to_cell(p_row-1, col)
        ws.write_formula(p_row, col, '={}/{}'.format(total, team_n), percent_fmt)
    # build the description boxes
    d_row = len(info) + 7
    start_col = 1
    for domain in strengths.domain_order:
        end_col = start_col + strengths.domain_counts[domain] - 1
        build_description_box(wb, ws, domain, d_row, start_col, end_col)
        start_col = end_col + 2
    # autofit the name column based on the longest name
    width = max([len(s[0]) for s in info])
    ws.set_column(0, 0, width)

def build_description_box(wb, ws, domain, row, start_col, end_col):
    title_fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 11,
                               'valign': 'vcenter', 'border': 2,
                               'bg_color': strengths.domain_color(domain),
                               'font_color': strengths.domain_txt_color(domain)})
    short_fmt = wb.add_format({'align': 'center', 'bold': True, 'font_size': 10,
                               'valign': 'vcenter', 'border': 2, 'bg_color': '#5A88D6',
                               'font_color': 'white'})
    long_fmt = wb.add_format({'align': 'left', 'bold': False, 'font_size': 10,
                              'valign': 'top', 'border': 2, 'text_wrap': True,
                              'bg_color': 'white', 'font_color': 'black'})
    ws.merge_range(row, start_col, row, end_col, domain.upper(), title_fmt)
    row += 1
    ws.merge_range(row, start_col, row, end_col,
                   strengths.domain_short_description[domain], short_fmt)
    row += 1
    ws.merge_range(row, start_col, row, end_col,
                   strengths.domain_long_description[domain], long_fmt)
    ws.set_row(row, 15 * 4)  # 15 is the default row height

def insert_totals(ws, row, start_col, end_col, row_range, fmt):
    formula = '=COUNTIF({0:s}:{1:s},"x")'
    for col in range(start_col, end_col):
        c1 = xl_rowcol_to_cell(row_range[0], col)
        c2 = xl_rowcol_to_cell(row_range[1], col)
        ws.write_formula(row, col, formula.format(c1, c2), fmt)

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
    sordered = strengths.executing + strengths.strategic_thinking + strengths.influencing + strengths.relationship_building
    [ws.write_string(sname_row, col+1, strength, center_fmt) for col, strength in enumerate(sordered)]
    # build the top 10 counts
    row_start = len(info) + 3
    for row in range(row_start, row_start+len(info)):
        # the first column is the name reference
        source_row = row-row_start+2
        ws.write_formula(row, 0, '=A{0:d}'.format(source_row))
        # the next columns count if the Strength occurs in the first 10 columns (top 10)
        col = 1
        for s in sordered:
            formula = '=COUNTIF(B{0:d}:K{0:d}, "{1:s}")'.format(source_row, s)
            ws.write_formula(row, col, formula, center_fmt)
            col += 1
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
    build_matrix(wb, ws_matrix, info)
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
