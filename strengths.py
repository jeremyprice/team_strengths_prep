#!/usr/bin/env python3

all34 = ('Achiever', 'Activator', 'Adaptability', 'Analytical', 'Arranger', 'Belief', 'Command',
         'Communication', 'Competition', 'Connectedness', 'Consistency', 'Context', 'Deliberative',
         'Developer', 'Discipline', 'Empathy', 'Focus', 'Futuristic', 'Harmony', 'Ideation', 'Includer',
         'Individualization', 'Input', 'Intellection', 'Learner', 'Maximizer', 'Positivity', 'Relator',
         'Responsibility', 'Restorative', 'Self-Assurance', 'Significance', 'Strategic', 'Woo')
strategic_thinking = ('Analytical', 'Context', 'Futuristic', 'Ideation', 'Input', 'Intellection',
                      'Learner', 'Strategic')
executing = ('Achiever', 'Arranger', 'Consistency', 'Deliberative', 'Belief', 'Discipline', 'Focus',
             'Responsibility', 'Restorative')
influencing = ('Activator', 'Command', 'Communication', 'Competition', 'Maximizer', 'Self-Assurance',
               'Significance', 'Woo')
relationship_building = ('Adaptability', 'Connectedness', 'Developer', 'Empathy', 'Harmony', 'Includer',
                         'Individualization', 'Positivity', 'Relator')

use_old_colors = True

old_domain_colors = {'Executing': '#48255A', 'Strategic Thinking': '#840003',
                     'Influencing':'#EEA823', 'Relationship Building': '#1D3661'}
new_domain_colors = {'Executing': '#7B2481', 'Strategic Thinking': '#00945D',
                     'Influencing':'#E97200', 'Relationship Building': '#0070CD'}

def domain_color(domain):
    if use_old_colors:
        return old_domain_colors[domain]
    else:
        return new_domain_colors[domain]
