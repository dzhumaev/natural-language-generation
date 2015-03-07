# -*- coding: utf-8 -*-

from collections import defaultdict
from pprint import pprint

import openpyxl


def make_player(ws, row_index):
    i = str(row_index)
    last_name_ru, first_name_ru = ws['E' + i].value.split()
    last_name_en, first_name_en = ws['C' + i].value.split()
    return {
        'first_name_ru': first_name_ru,
        'last_name_ru': last_name_ru,
        'first_name_en': first_name_en,
        'last_name_en': last_name_en,
        'role_en': ws['D' + i].value,
    }


def load_logs(file_name):
    logs = []
    wb = openpyxl.load_workbook(file_name)
    for ws in wb:
        log = {}
        logs.append(log)
        log['date'] = ws['A2'].value
        log['arena'] = ws['B2'].value
        log['city'] = ws['C2'].value
        log['attendance'] = ws['D2'].value
        log['home-team'] = ws['E2'].value
        log['guest-team'] = ws['F2'].value
        log['players'] = defaultdict(dict)
        home_team_last_row = 4
        for i in range(4, 31):
            if ws['A'+str(i)].value == 'guest-team':
                home_team_last_row = i - 2
                break
        guest_team_last_row = home_team_last_row + 2
        for i in range(guest_team_last_row, 61):
            if ws['A'+str(i)].value == 'Play':
                guest_team_last_row = i - 2
                break
        for i in range(5, home_team_last_row+1):
            log['players'][log['home-team']][ws['B'+str(i)].value] = make_player(ws, i)
        for i in range(home_team_last_row+3, guest_team_last_row+1):
            log['players'][log['guest-team']][ws['B'+str(i)].value] = make_player(ws, i)

        row = ws.get_highest_row()
        log['score'] = [ws['B'+str(row-1)].value, ws['B'+str(row)].value]
        
        log['goals'] = []
        for i in range(guest_team_last_row+3, row-2):
            if ws['J'+str(i)].value in {'scored', 'powerplay', 'shorthanded'}:
                team = ws['D'+str(i)].value
                author = log['players'][team][ws['E'+str(i)].value]
                log['goals'].append({
                    'minute': ws['B'+str(i)].value,
                    'team': team,
                    'author': author,
                    'type': ws['J'+str(i)].value,
                })
        process_log(log)
    return logs


def process_log(log):
    if log['score'][0] > log['score'][1]:
        log['winner'], log['looser'] = log['home-team'], log['guest-team']
    else:
        log['looser'], log['winner'] = log['home-team'], log['guest-team']
    log['score'] = '-'.join(map(str, sorted(log['score'], reverse=True)))


class Event:
    def __init__(self, log):
        self.log = log

    def is_applicable(self):
        return True

    def wrap(self, message):
        return message

    def gen_wrapped_russian(self):
        return self.wrap(self.gen_russian())

    def gen_wrapped_english(self):
        return self.wrap(self.gen_english())

    def gen_russian(self):
        return ''

    def gen_english(self):
        return ''


class Pause(Event):
    def wrap(self, message):
        return '<break time="1000ms"/>'


class HighPitchEvent(Event):
    # def wrap(self, message):
    #     return '<prosody rate="slow" pitch="high">' + message + '</prosody>'
    pass


class WinnerEvent(HighPitchEvent):
    def gen_russian(self):
        return (self.log['winner'] + ' обыграл ' + self.log['looser'] + ' со счётом '
                + self.log['score'])

    def gen_english(self):
        return ('The ' + self.log['winner'] + ' picked up a ' + self.log['score']
                + ' win against the ' + self.log['looser'])


def join_with_and(words, and_word):
    return ', '.join(words[:-1]) + ' ' + and_word + ' ' + words[-1]


def join_with_and_ru(words):
    return join_with_and(words, 'и')


def join_with_and_en(words):
    return join_with_and(words, 'and')


class GoalsSummaryEvent(HighPitchEvent):
    def __init__(self, log):
        super(GoalsSummaryEvent, self).__init__(log)
        self.goals_by_team_ru = defaultdict(set)
        self.goals_by_team_en = defaultdict(set)
        for goal in log['goals']:
            self.goals_by_team_ru[goal['team']].add(goal['author']['last_name_ru'])
            self.goals_by_team_en[goal['team']].add(goal['author']['last_name_en'])
        for k, v in self.goals_by_team_ru.items():
            self.goals_by_team_ru[k] = sorted(v)
        for k, v in self.goals_by_team_en.items():
            self.goals_by_team_en[k] = sorted(v)

    def gen_russian(self):
        return ('В составе ' + self.log['winner'] + ' отличились '
                + join_with_and_ru(self.goals_by_team_ru[self.log['winner']]) + ', за ' + self.log['looser']
                + ' голы забили ' + join_with_and_ru(self.goals_by_team_ru[self.log['looser']]))

    def gen_english(self):
        return (join_with_and_en(self.goals_by_team_en[self.log['winner']])
                + ' were on target for the ' + self.log['winner']  + ', while '
                + join_with_and_en(self.goals_by_team_en[self.log['looser']])
                + ' replied for the ' + self.log['looser'])


EVENT_CLASSES = [WinnerEvent, GoalsSummaryEvent, Pause]


def form_report(log):
    russian_report = []
    english_report = []
    reports = [russian_report, english_report]
    for event_class in EVENT_CLASSES:
        event = event_class(log)
        if event.is_applicable():
            russian_report.append(event.gen_wrapped_russian() + '. ')
            english_report.append(event.gen_wrapped_english() + '. ')
    return [''.join(report) for report in reports]


def main():
    logs = load_logs('Hockey_Log.xlsx')
    pprint(logs)
    for log in logs:
        for report in form_report(log):
            print(report)


if __name__ == '__main__':
    main()
