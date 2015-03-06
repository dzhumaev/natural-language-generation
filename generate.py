# -*- coding: utf-8 -*-

import openpyxl

def load_logs(file_name):
    logs = []
    wb = openpyxl.load_workbook(file_name)
    log_number = -1
    for ws in wb:
        log_number += 1
        logs.append({})
        logs[log_number]['date'] = ws['A2'].value
        logs[log_number]['arena'] = ws['B2'].value
        logs[log_number]['city'] = ws['C2'].value
        logs[log_number]['attendance'] = ws['D2'].value
        logs[log_number]['home-team'] = ws['E2'].value
        logs[log_number]['guest-team'] = ws['F2'].value
        logs[log_number]['players'] = [{}, {}]
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
            logs[log_number]['players'][0][ws['B'+str(i)].value] = [ws['C'+str(i)].value, ws['D'+str(i)].value]
        for i in range(home_team_last_row+3, guest_team_last_row+1):
            logs[log_number]['players'][1][ws['B'+str(i)].value] = [ws['C'+str(i)].value, ws['D'+str(i)].value]

        row = ws.get_highest_row()
        logs[log_number]['score'] = [ws['B'+str(row-1)].value, ws['B'+str(row)].value]
        
        logs[log_number]['goals'] = []
        for i in range(guest_team_last_row+3, row-2):
            if ws['J'+str(i)].value == 'scored':
                logs[log_number]['goals'].append([ws['B'+str(i)].value, ws['C'+str(i)].value])
    return logs

def form_report(logs):
    reports = []
    for log in logs:
        reports.append(log['home-team']+' '+log['guest-team'])
    return reports
    
def main():
    logs = load_logs('Hockey_Log.xlsx')
    reports = form_report(logs)

if __name__ == '__main__':
    main()
