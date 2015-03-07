# -*- coding: utf-8 -*-

from pprint import pprint

import openpyxl

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
        log['players'] = [{}, {}]
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
            log['players'][0][ws['B'+str(i)].value] = [ws['C'+str(i)].value, ws['D'+str(i)].value]
        for i in range(home_team_last_row+3, guest_team_last_row+1):
            log['players'][1][ws['B'+str(i)].value] = [ws['C'+str(i)].value, ws['D'+str(i)].value]

        row = ws.get_highest_row()
        log['score'] = [ws['B'+str(row-1)].value, ws['B'+str(row)].value]
        
        log['goals'] = []
        for i in range(guest_team_last_row+3, row-2):
            if ws['J'+str(i)].value == 'scored':
                log['goals'].append([ws['B'+str(i)].value, ws['C'+str(i)].value])
    return logs


class Event:
    def is_applicable(self):
        return True

    def wrap(self, message):
        return message

    def gen_russian(self):
        raise NotImplementedError()

    def gen_english(self):
        raise NotImplementedError()


class WinnerEvent(Event):
    def __init__(self, log):
        if log['score'][0] > log['score'][1]:
            self.winner, self.looser = log['home-team'], log['guest-team']
        else:
            self.looser, self.winner = log['home-team'], log['guest-team']
        self.score = '-'.join(map(str, sorted(log['score'], reverse=True)))

    def wrap(self, message):
        return '<prosody rate="slow" pitch="high">' + message + '</prosody>'

    def gen_russian(self):
        return self.wrap(self.winner + ' обыграл ' + self.looser + ' со счётом ' + self.score)

    def gen_english(self):
        return self.wrap('The ' + self.winner + ' picked up a ' + self.score + ' win against the '
                         + self.looser)


EVENT_CLASSES = [WinnerEvent]


def form_report(log):
    russian_report = []
    english_report = []
    reports = [russian_report, english_report]
    for event_class in EVENT_CLASSES:
        event = event_class(log)
        if event.is_applicable():
            russian_report.append(event.gen_russian() + '. ')
            english_report.append(event.gen_english() + '. ')
    return [''.join(report) for report in reports]


def main():
    logs = load_logs('Hockey_Log.xlsx')
    pprint(logs)
    for log in logs:
        for report in form_report(log):
            print(report)


if __name__ == '__main__':
    main()
