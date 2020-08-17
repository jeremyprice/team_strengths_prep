#!/usr/bin/env python3

import sys
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
# from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib import pagesizes
from reportlab.lib.units import inch
import reportlab.platypus
from PIL import Image

PAGE_WIDTH, PAGE_HEIGHT = pagesizes.LETTER
text_start_y = PAGE_HEIGHT / 2.0 - (1*inch)
large_line_spacing = 45
small_line_spacing = 35
name_only_spacing = 15
RACKSPACE_RED_RGB = (0.88, 0.15, 0.18)
RACKSPACE_GREY_RGB = (0.2, 0.2, 0.2)
RACKSPACE_TEXT_RGB = (0.25, 0.25, 0.25)
LEFT_X = 0.75 * inch
LEFT_2ND_COL = 4.0 * inch
IMAGE_MAX_WIDTH = LEFT_2ND_COL - LEFT_X
IMAGE_MAX_HEIGHT = 2.5 * inch


def load_fonts():
    folder = 'fonts'
    fs_ttf = os.path.join(folder, 'FiraSans-Book.ttf')
    fsb_ttf = os.path.join(folder, 'FiraSans-Bold.ttf')
    pdfmetrics.registerFont(TTFont("FiraSans", fs_ttf))
    pdfmetrics.registerFont(TTFont("FiraSans Bold", fsb_ttf))


def set_pdf_file_defaults(canvas):
    canvas.setAuthor('Rackspace University')
    canvas.setCreator('Rackspace University')
    canvas.setSubject('Top CliftonStrengths Talents')
    canvas.setTitle('Top Talents Name Tent')
    canvas.setPageSize(pagesizes.LETTER)


def create_pdf_canvas(fname):
    canvas = Canvas(fname)
    set_pdf_file_defaults(canvas)
    return canvas


def print_name(name, canvas):
    canvas.setFont('FiraSans', 36)
    canvas.setStrokeColorRGB(*RACKSPACE_RED_RGB)
    canvas.setFillColorRGB(*RACKSPACE_RED_RGB)
    y = text_start_y - name_only_spacing
    canvas.drawString(LEFT_X, y, name)


def print_name_and_title(name, title, canvas):
    # print name
    canvas.setFont('FiraSans', 36)
    canvas.setStrokeColorRGB(*RACKSPACE_RED_RGB)
    canvas.setFillColorRGB(*RACKSPACE_RED_RGB)
    y = text_start_y
    canvas.drawString(LEFT_X, y, name)
    # print title
    canvas.setFont('FiraSans', 24)
    canvas.setStrokeColorRGB(*RACKSPACE_RED_RGB)
    canvas.setFillColorRGB(*RACKSPACE_RED_RGB)
    y = text_start_y - small_line_spacing
    canvas.drawString(LEFT_X, y, title)

def print_lines(canvas):
    # draw the fold line in the middle
    canvas.setStrokeColorRGB(*RACKSPACE_GREY_RGB)
    canvas.setFillColorRGB(*RACKSPACE_GREY_RGB)
    canvas.setLineWidth(1)
    y = PAGE_HEIGHT / 2
    canvas.line(0, y, PAGE_WIDTH, y)
    # draw the line between the name and the strengths
    canvas.setStrokeColorRGB(*RACKSPACE_RED_RGB)
    canvas.setFillColorRGB(*RACKSPACE_RED_RGB)
    canvas.setLineWidth(2)
    y = text_start_y - large_line_spacing
    canvas.line(LEFT_X, y, PAGE_WIDTH - LEFT_X, y)


def print_image(image, canvas, scaling=1.0, font_scaling=1.0):
    im = Image.open(image)
    im_w, im_h = im.size
    new_height = IMAGE_MAX_HEIGHT
    scale = new_height / im_h
    new_width = im_w * scale
    if new_width > IMAGE_MAX_WIDTH:
        # image is too wide
        new_width = IMAGE_MAX_WIDTH
        scale = new_width / im_w
        new_height = im_h * scale
    offset = IMAGE_MAX_HEIGHT/2 + new_height/2
    y = text_start_y - offset - (2 * small_line_spacing)
    x = LEFT_X
    # TODO: center picture and scale appropriately
    reportlab.platypus.Image(image, width=new_width, height=new_height).drawOn(canvas, x, y)


def print_talents(talents, canvas, right=False):
    canvas.setFont('FiraSans', 24)
    canvas.setStrokeColorRGB(*RACKSPACE_TEXT_RGB)
    canvas.setFillColorRGB(*RACKSPACE_TEXT_RGB)
    y = text_start_y - (2 * large_line_spacing)
    if right:
        x = LEFT_2ND_COL
    else:
        x = LEFT_X
    for idx, talent in enumerate(talents[:10]):
        canvas.drawString(x, y, talent)
        y -= small_line_spacing
        if idx == 4:
            x = LEFT_2ND_COL
            y = text_start_y - (2 * large_line_spacing)


def print_logo(logo, canvas):
    im = Image.open(logo)
    im_w, im_h = im.size
    new_height = 0.3 * inch
    scale = new_height / im_h
    new_width = im_w * scale
    y = 0.4 * inch
    x = 0.55 * inch
    # TODO: center picture and scale appropriately
    reportlab.platypus.Image(logo, width=new_width, height=new_height).drawOn(canvas, x, y)


class MultiPageNameTent(object):
    def __init__(self, fname):
        load_fonts()
        self.canvas = create_pdf_canvas(fname)

    def create_page(self, name, talents, logo=True):
        # print the right side up side
        print_talents(talents, self.canvas, right=False)
        print_name(name, self.canvas)
        print_lines(self.canvas)
        if logo:
            print_logo('images/rackspace-logo.png', self.canvas)
        # print the upside down side
        self.canvas.saveState()
        self.canvas.translate(PAGE_WIDTH, PAGE_HEIGHT)
        self.canvas.rotate(180)

        print_talents(talents, self.canvas, right=False)
        print_name(name, self.canvas)
        print_lines(self.canvas)
        if logo:
            print_logo('images/rackspace-logo.png', self.canvas)
        self.canvas.restoreState()
        self.canvas.showPage()

    def done(self):
        self.canvas.save()


def create_one_file_nametents(fname, info):
    mpnt = MultiPageNameTent(fname)
    for person in info:
        name = person[0]
        talents = person[1:]
        mpnt.create_page(name, talents)
    mpnt.done()

def main():
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
    create_one_file_nametents('test.pdf', info)


if __name__ == '__main__':
    main()
else:
    load_fonts()



if __name__ == '__main__':
    main()
