# -*- coding: utf-8 -*-
from pnid import PnID
from pprint import PrettyPrinter
import time

def get_type_code(type_name):
    if type_name == 'GATE':
        return 'GT'
    if type_name == 'CHECK':
        return 'CH'
    if type_name == 'BALL':
        return 'BL'
    if type_name == 'GLOBE':
        return 'GB'
    if type_name == 'BUTTERFLY':
        return 'BU'
    if type_name == 'NEEDLE':
        return 'NV'
    return None

start_time = time.time()
pnid = PnID(r'D:\Work\Project\XY2019P02-KAYAN.25MMSCFD.LNG\PnID\IssueForHazop\KAYAN.GENERATORS.P&ID2020.0612H.dwg')
valves = dict()
counter = 0
for valve in pnid.valves:
    dwg_number = pnid.locate_dwg_no(valve.pos)
    if dwg_number:
        dwg_number = dwg_number[-4::]
    else:
        continue
    if dwg_number not in valves:
        valves[dwg_number] = dict()
    if valve.type_name not in valves[dwg_number]:
        valves[dwg_number][valve.type_name] = []
    valves[dwg_number][valve.type_name].append(valve)
    if type_code := get_type_code(valve.type_name):
        valve.v_type = type_code
        valve.number = f'{dwg_number}{len(valves[dwg_number][valve.type_name]):02d}'
        # tagging, exclude legends
        if int(dwg_number) > 99:
            tag_handle = pnid.doc.HandleToObject(valve.tag_handle)
            tag_handle.TextString = valve.tag
            tag_handle.Height = 2.5
            tag_handle.ScaleFactor = 0.6
            counter += 1
print(f'{counter} valves processed.')
end_time = time.time()
time_spent = end_time - start_time
print('Time spent: %.2fs' % time_spent)
# pp = PrettyPrinter()
# pp.pprint(valves)
