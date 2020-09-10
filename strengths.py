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

use_old_colors = False

old_domain_colors = {'Executing': '#48255A', 'Strategic Thinking': '#840003',
                     'Influencing':'#EEA823', 'Relationship Building': '#1D3661'}
old_domain_txt_colors = {'Executing': 'white', 'Strategic Thinking': 'white',
                         'Influencing':'black', 'Relationship Building': 'white'}
new_domain_colors = {'Executing': '#7B2481', 'Strategic Thinking': '#00945D',
                     'Influencing':'#E97200', 'Relationship Building': '#0070CD'}
new_domain_txt_colors = {'Executing': 'white', 'Strategic Thinking': 'white',
                         'Influencing':'white', 'Relationship Building': 'white'}

def domain_color(domain):
    if use_old_colors:
        return old_domain_colors[domain]
    else:
        return new_domain_colors[domain]

def domain_txt_color(domain):
    if use_old_colors:
        return old_domain_txt_colors[domain]
    else:
        return new_domain_txt_colors[domain]

domains = {'Executing': executing,
           'Strategic Thinking': strategic_thinking,
           'Influencing': influencing,
           'Relationship Building': relationship_building}

domain_order = ('Executing', 'Strategic Thinking', 'Influencing', 'Relationship Building')

domain_counts = {d: len(domains[d]) for d in domain_order}

domain_short_description = {'Executing': 'Making things happen',
                            'Strategic Thinking': 'Focused on what could be',
                            'Influencing':'Reaching a broader audience',
                            'Relationship Building': 'The glue that binds'}

domain_long_description = {'Executing': 'When the team needs to implement a solution, they will work tirelessly to get it done. Teams strong in the Executing domain have the ability to "catch" an idea and make it a reality.',
                           'Strategic Thinking': 'Teams strong in Strategic Thinking are constantly absorbing and analyzing information and helping each other make better decisions, continually stretching their thinking for the future.',
                           'Influencing': "Always selling the team's ideas inside and outside the organization. Quick to take charge, speak up, and make sure the group is heard.",
                           'Relationship Building': 'Rather than a composite of individuals, teams strong in this domain have the unique ability to create groups and organizations that are much bigger than the sum of their parts.'}
