# This is a Python script to compare DC-PAMT and AMAG CAS data for access control in DC Blonie

# Copyright masti 2020

import os
import re
import sys
from datetime import datetime

# global translations
from typing import Dict, Any, Union

PAMT2CAS: Dict[Union[str, Any], Union[str, Any]] = {
    'C01 Main Entrance': 'C01+C02 Main Entrance',
    'C02 Lobby': 'C01+C02 Main Entrance',
    'C02 Main entry + lobby': 'C01+C02 Main Entrance',
    'C03 Conference room': 'C03 CONFERENCE ROOM',
    'C04 Customer office': 'C04,C05,C06,C13 CUSTOM.OF',
    'C05 Customer office eService': 'C04,C05,C06,C13 CUSTOM.OF',
    'C06 Customer office': 'C04,C05,C06,C13 CUSTOM.OF',
    'C13 Customer office PEK': 'C04,C05,C06,C13 CUSTOM.OF',
    'D01 Material lock': 'D01 MATERIAL LOCK',
    'D04 Delivery area': 'D04 DELIVERY AREA',
    'FM Chillers/Generators': 'FM CHILLERSGENERATORS',
    'FM Fire buffer zone': 'FM FIRE BUFFER ZONE',
    'FM Fuel tanks': 'FM FUEL TANKS',
    'FM01 Internal FM corridor': 'FM01 INTERNAL FM CORRIDOR',
    'FM02 BMS': 'FM02 BMS',
    'FM03 FM Office': 'FM03 FM OFFICE',
    'FM04 Storage FM': 'FM04 STORAGE FM',
    'FM06 PDUs&CRACs Area - A': 'FM06 PDUs + CRACs Area A',
    'FM07 PDUs&CRACs Area - B': 'FM07 PDUs + CRACs Area B',
    'FM08 Gas Bottles Room': 'FM08 GAS BOTTLES ROOM',
    'FM09 LV Switchboards/UPSs - A': 'FM09 FM15 LV SWITCHB.UPS',
    'FM10 Batteries Room - A': 'FM10 FM18 BATTERIES A, B',
    'FM13 Mechanical plant': 'FM13 MECHANICAL PLANT',
    'FM15 LV Switchboards/UPSs - B': 'FM09 FM15 LV SWITCHB.UPS',
    'FM18 Batteries Room - B': 'FM10 FM18 BATTERIES A, B',
    'FM21 Sprinkler Pump Room': 'FM21 SPRINKLER PUMP ROOM',
    'IT01 Internal IT corridor': 'IT01 INTERNAL IT CORRIDOR',
    'IT02 Telco room': 'IT02 TELCO ROOM',
    'IT03 Network room': 'IT03 LAN ROOM',
    'IT04 Office IT': 'IT04 OFFICE IT',
    'IT05 Pekao Storage': 'IT05 STORAGE',
    'IT06 Tape storage': 'IT06 TAPE STORAGE',
    'IT07 Network room': 'IT07 LAN ROOM',
    'IT08 Telco room': 'IT08 WAN ROOM - TELCO',
    'IT09 Main storage': 'IT09 STORAGE',
    'IT10 DB': 'IT10 DB',
    'IT10 eService': 'IT10 eService',
    'IT10 Main DC room': 'IT10 IT HALL-MAIN DC ROOM',
    'IT10 PEK': 'IT10 PEK',
    'S01 Security': 'S01 IBM SECURITY',
}

CAS2PAMT = {
    'C01+C02 Main Entrance': 'C02 Main entry + lobby',
    'C03 CONFERENCE ROOM': 'C03 Conference room',
    'D01 MATERIAL LOCK': 'D01 Material lock',
    'D04 DELIVERY AREA': 'D04 Delivery area',
    'FM CHILLERSGENERATORS': 'FM Chillers/Generators',
    'FM FIRE BUFFER ZONE': 'FM Fire buffer zone',
    'FM FUEL TANKS': 'FM Fuel tanks',
    'FM01 INTERNAL FM CORRIDOR': 'FM01 Internal FM corridor',
    'FM02 BMS': 'FM02 BMS',
    'FM03 FM OFFICE': 'FM03 FM Office',
    'FM04 STORAGE FM': 'FM04 Storage FM',
    'FM06 PDUs + CRACs Area A': 'FM06 PDUs&CRACs Area - A',
    'FM07 PDUs + CRACs Area B': 'FM07 PDUs&CRACs Area - B',
    'FM08 GAS BOTTLES ROOM': 'FM08 Gas Bottles Room',
    'FM13 MECHANICAL PLANT': 'FM13 Mechanical plant',
    'FM21 SPRINKLER PUMP ROOM': 'FM21 Sprinkler Pump Room',
    'IT01 INTERNAL IT CORRIDOR': 'IT01 Internal IT corridor',
    'IT02 TELCO ROOM': 'IT02 Telco room',
    'IT03 LAN ROOM': 'IT03 Network room',
    'IT04 OFFICE IT': 'IT04 Office IT',
    'IT05 STORAGE': 'IT05 Pekao Storage',
    'IT06 TAPE STORAGE': 'IT06 Tape storage',
    'IT07 LAN ROOM': 'IT07 Network room',
    'IT08 WAN ROOM - TELCO': 'IT08 Telco room',
    'IT09 STORAGE': 'IT09 Main storage',
    'IT10 DB': 'IT10 DB',
    'IT10 eService': 'IT10 eService',
    'IT10 IT HALL-MAIN DC ROOM': 'IT10 Main DC room',
    'IT10 PEK': 'IT10 PEK',
    'S01 IBM SECURITY': 'S01 Security',
}

PAMT2Room = {
    'C01 Main Entrance': ['C01'],
    'C02 Lobby': ['C02'],
    'C02 Main entry + lobby': ['C02'],
    'C03 Conference room': ['C03'],
    'C04 Customer office': ['C04'],
    'C05 Customer office eService': ['C05'],
    'C06 Customer office': ['C06'],
    'C13 Customer office PEK': ['C13'],
    'D01 Material lock': ['D01'],
    'D04 Delivery area': ['D04'],
    'FM Chillers/Generators': ['FMCHI'],
    'FM Fire buffer zone': ['FMBUF'],
    'FM Fuel tanks': ['FMFUEL'],
    'FM01 Internal FM corridor': ['FM01'],
    'FM02 BMS': ['FM02'],
    'FM03 FM Office': ['FM03'],
    'FM04 Storage FM': ['FM04'],
    'FM06 PDUs&CRACs Area - A': ['FM06'],
    'FM07 PDUs&CRACs Area - B': ['FM07'],
    'FM08 Gas Bottles Room': ['FM08'],
    'FM09 LV Switchboards/UPSs - A': ['FM09'],
    'FM10 Batteries Room - A': ['FM10'],
    'FM13 Mechanical plant': ['FM13'],
    'FM15 LV Switchboards/UPSs - B': ['FM09'],
    'FM18 Batteries Room - B': ['FM10'],
    'FM21 Sprinkler Pump Room': ['FM21'],
    'IT01 Internal IT corridor': ['IT01'],
    'IT02 Telco room': ['FM01'],
    'IT03 Network room': ['IT03'],
    'IT04 Office IT': ['IT04'],
    'IT05 Pekao Storage': ['IT05'],
    'IT06 Tape storage': ['IT06'],
    'IT07 Network room': ['IT07'],
    'IT08 Telco room': ['IT08'],
    'IT09 Main storage': ['IT09'],
    'IT10 DB': ['IT10DB'],
    'IT10 eService': ['IT10ES'],
    'IT10 Main DC room': ['IT10'],
    'IT10 PEK': ['IT10PEK'],
    'S01 Security': ['S01'],
}

CAS2Room = {
    'C01+C02 Main Entrance': ['C01', 'C02'],
    'C03 CONFERENCE ROOM': ['C03'],
    'C04,C05,C06,C13 CUSTOM.OF': ['C04', 'C05', 'C06', 'C13'],
    'D01 MATERIAL LOCK': ['D01'],
    'D04 DELIVERY AREA': ['D04'],
    'FM CHILLERSGENERATORS': ['FMCHI'],
    'FM FIRE BUFFER ZONE': ['FMBUF'],
    'FM FUEL TANKS': ['FMFUEL'],
    'FM01 INTERNAL FM CORRIDOR': ['FM01'],
    'FM02 BMS': ['FM02'],
    'FM03 FM OFFICE': ['FM03'],
    'FM04 STORAGE FM': ['FM04'],
    'FM06 PDUs + CRACs Area A': ['FM06'],
    'FM07 PDUs + CRACs Area B': ['FM07'],
    'FM08 GAS BOTTLES ROOM': ['FM08'],
    'FM09 FM15 LV SWITCHB.UPS': ['FM09'],
    'FM10 FM18 BATTERIES A, B': ['FM10'],
    'FM13 MECHANICAL PLANT': ['FM13'],
    'FM21 SPRINKLER PUMP ROOM': ['FM21'],
    'IT01 INTERNAL IT CORRIDOR': ['IT01'],
    'IT02 TELCO ROOM': ['FM01'],
    'IT03 LAN ROOM': ['IT03'],
    'IT04 OFFICE IT': ['IT04'],
    'IT05 STORAGE': ['IT05'],
    'IT06 TAPE STORAGE': ['IT06'],
    'IT07 LAN ROOM': ['IT07'],
    'IT08 WAN ROOM - TELCO': ['IT08'],
    'IT09 STORAGE': ['IT09'],
    'IT10 DB': ['IT10DB'],
    'IT10 eService': ['IT10ES'],
    'IT10 IT HALL-MAIN DC ROOM': ['IT10'],
    'IT10 PEK': ['IT10PEK'],
    'S01 IBM SECURITY': ['S01'],
}

Room2CAS = {
    'C01': 'C01+C02 Main Entrance',
    'C02': 'C01+C02 Main Entrance',
    'C03': 'C03 CONFERENCE ROOM',
    'C04': 'C04,C05,C06,C13 CUSTOM.OF',
    'C05': 'C04,C05,C06,C13 CUSTOM.OF',
    'C06': 'C04,C05,C06,C13 CUSTOM.OF',
    'C13': 'C04,C05,C06,C13 CUSTOM.OF',
    'D01': 'D01 MATERIAL LOCK',
    'D04': 'D04 DELIVERY AREA',
    'FMCHI': 'FM CHILLERSGENERATORS',
    'FMBUF': 'FM FIRE BUFFER ZONE',
    'FMFUEL': 'FM FUEL TANKS',
    'FM01': 'FM01 INTERNAL FM CORRIDOR',
    'FM02': 'FM02 BMS',
    'FM03': 'FM03 FM OFFICE',
    'FM04': 'FM04 STORAGE FM',
    'FM06': 'FM06 PDUs + CRACs Area A',
    'FM07': 'FM07 PDUs + CRACs Area B',
    'FM08': 'FM08 GAS BOTTLES ROOM',
    'FM09': 'FM09 FM15 LV SWITCHB.UPS',
    'FM10': 'FM10 FM18 BATTERIES A, B',
    'FM13': 'FM13 MECHANICAL PLANT',
    'FM15': 'FM09 FM15 LV SWITCHB.UPS',
    'FM18': 'FM10 FM18 BATTERIES A, B',
    'FM21': 'FM21 SPRINKLER PUMP ROOM',
    'IT01': 'IT01 INTERNAL IT CORRIDOR',
    'IT02': 'IT02 TELCO ROOM',
    'IT03': 'IT03 LAN ROOM',
    'IT04': 'IT04 OFFICE IT',
    'IT05': 'IT05 STORAGE',
    'IT06': 'IT06 TAPE STORAGE',
    'IT07': 'IT07 LAN ROOM',
    'IT08': 'IT08 WAN ROOM - TELCO',
    'IT09': 'IT09 STORAGE',
    'IT10DB': 'IT10 DB',
    'IT10ES': 'IT10 eService',
    'IT10': 'IT10 IT HALL-MAIN DC ROOM',
    'IT10PEK': 'IT10 PEK',
    'S01': 'S01 IBM SECURITY',
}

# regexes
pamtR = re.compile(r"<table.*?<table.*?</table>.*?</table>")
roomsR = re.compile(r"Rooms:</small></td><td style='width: 585px;'>(?P<rooms>.*?)</td></tr>")
roomR = re.compile(r"(?P<room>.*?)<br>")
dateR = re.compile(
    r"Date:</small></td><td style='width: 585px;'>(?P<start>[\d\.]*)&nbsp;-&nbsp;(?P<end>[\d\.]*)</td></tr><tr>")
visitorR = re.compile(
    r"Visitor:</small></td><td style='width: 585px;'>(?P<visitor>.*?)&nbsp;<small>(?P<comment>.*?)</small><br></td></tr>")
guestsR = re.compile(r"Guests:</small></td><td style='width: 585px;'>(?P<guests>.*?)</td></tr>")
guestR = re.compile(r"(?P<name>.*?)&nbsp;.*?</small><br>")

roomAccessCAS = {}
personsAccess = {}
roomAccessPAMT = {}
resultAccess = {
    'add': {},
    'remove': {}
}
cardList = {}


def names(name):
    # extract last name form name
    l = name.split(', ')[0]
    f = name.split(', ')[1]
    return (l, f)


def cleanupCasLine(last, first):
    # remove Contractor, keycard
    # reverse firstname & name
    if 'contractor' in last:
        last = last.replace('contractor ', '')
        f, l = tuple(last.split())
    else:
        f = first.split()[0]
        l = last.split()[0]
    # print(f'Last:{l} First:{f}')
    return ("{0}, {1}".format(l.capitalize(), f.capitalize()))


def addRoomAccessCAS(name, card, to, room):
    # add access by room
    l, f = names(name)
    if not name in roomAccessCAS[room].keys():
        roomAccessCAS[room][name] = {
            'card': card,
            'to': to,
            'Last': l,
            'First': f,
        }
    elif roomAccessCAS[room][name]['to'] < to:
        roomAccessCAS[room][name]['to'] = to
    return


def treatCasLine(line, rooms):
    # process one line from CAS file
    # return dict

    l, f, c, i, a, to = tuple(line.split('\t'))
    n = cleanupCasLine(l, f)
    try:
        to = datetime.strptime(to, "%d/%m/%Y")
    except ValueError:
        print('^^^ VALUE ERROR in room:{0} ^^^'.format(rooms))
        return
    for r in rooms:
        addRoomAccessCAS(n, c, to, r)


def processCAS(fname, rooms):
    # process CAS file fname
    print('Processing CAS file:{0}'.format(fname))
    lcount = 0
    with open(fname, 'r') as f:
        for l in f.readlines()[6:-5]:
            treatCasLine(l.strip('\n').lower(), rooms)
            lcount += 1
        print('Processed {0} lines'.format(lcount))
    return


def prepareCASRooms(room):
    for r in CAS2Room[room]:
        if r not in roomAccessCAS.keys():
            roomAccessCAS[r] = {}


def preparePAMTRooms():
    for r in PAMT2Room:
        if r not in roomAccessPAMT.keys():
            roomAccessPAMT[r] = {}


def getRoomsList(line):
    # return list of rooms
    roomList = []
    rooms = roomsR.search(line)
    if rooms:
        rcount = 1
        for r in roomR.finditer(rooms.group('rooms')):
            rcount += 1
            if r.group('room') in PAMT2Room:
                roomList.extend(PAMT2Room[r.group('room')])
    return (roomList)


def getRequestDates(line):
    # return start & end date of request as tuple
    start = None
    end = None
    date = dateR.search(line)
    if date:
        try:
            start = datetime.strptime(date.group('start'), "%d.%m.%Y")
            end = datetime.strptime(date.group('end'), "%d.%m.%Y")
        except ValueError:
            print('^^^ VALUE ERROR in dates:{0} ^^^'.format(date))
            return (None, None)
    return (start, end)


def cleanGuestName(name):
    # return guest name in a form Surname, Name
    l = re.split(", | ", name)
    return ("{0}, {1}".format(l[0].lower().capitalize(), l[1].lower().capitalize()))


def getVisitor(line):
    # return visitor name for the request
    # None if does not need access
    v = visitorR.search(line)
    if v:
        visitor = v.group('visitor')
        comment = v.group('comment')
        if 'no action' in comment.lower():
            return (None)
        else:
            return (cleanGuestName(visitor))
    return (None)


def getGuests(line):
    # return list of guest names for the request
    # None if does not need access
    guests = guestsR.search(line)
    guestsList = []
    if guests:
        gcount = 1
        for g in guestR.finditer(guests.group('guests')):
            gcount += 1
            guestsList.append(cleanGuestName(g.group('name')))
    return (guestsList)


def addRoomAccessPAMT(name, to, room):
    # add access by room from PAMT
    if room not in roomAccessPAMT.keys():
        roomAccessPAMT[room] = {}
    if name not in roomAccessPAMT[room].keys():
        l, f = names(name)
        roomAccessPAMT[room][name] = {
            'to': to,
            'Last': l,
            'First': f
        }
    elif roomAccessPAMT[room][name]['to'] < to:
        # change access end date to later
        roomAccessPAMT[room][name]['to'] = to
    return


def treatPAMTLine(line):
    # process one line from PAMT report = 1 record
    rooms = getRoomsList(line)
    startDate, endDate = getRequestDates(line)
    visitor = getVisitor(line)
    if visitor:
        visitorList = [visitor, ]
    else:
        visitorList = []
    visitorList.extend(getGuests(line))
    for r in rooms:
        for v in visitorList:
            addRoomAccessPAMT(v, endDate, r)


def processPAMT(fname):
    # process PAMT file fname
    print('Processing PAMT file:{0}'.format(fname))
    lcount = 0
    with open(fname, 'r') as f:
        file = f.read()

    for l in pamtR.findall(file):
        treatPAMTLine(l)
        lcount += 1
    print('Processed {0} lines'.format(lcount))


def PAMT2CAScheck():
    # compare PAMT 2 CAS entries
    for r in roomAccessPAMT.keys():
        for n in roomAccessPAMT[r].keys():
            try:
                casAccess = roomAccessCAS[r][n]
            except KeyError:
                if r not in resultAccess['add'].keys():
                    resultAccess['add'][r] = {}
                resultAccess['add'][r][n] = roomAccessPAMT[r][n]['to']


def CAS2PAMTcheck():
    # compare CAS 2 PAMT entries
    for r in roomAccessCAS.keys():
        for n in roomAccessCAS[r].keys():
            try:
                pamtAccess = roomAccessPAMT[r][n]
            except KeyError:
                if r not in resultAccess['remove'].keys():
                    resultAccess['remove'][r] = {}
                resultAccess['remove'][r][n] = roomAccessCAS[r][n]['card']


def writeResults(text, fname):
    with open(fname, 'w') as f:
        f.write(text)


def generateResultsFiles(reslist, fname):
    # list to add entries
    print("GENERATE RESULTS ADD")

    res = '*****************************************************\n'
    res += '*****************  Entries to ADD   *****************\n'
    res += '*****************************************************\n'
    res += 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for r in reslist['add'].keys():
        res += "ROOM:{0} (CAS:{1})\n".format(r, Room2CAS[r])
        res += '*****************************************************\n'
        res += 'ACTION\tROOM\tNAME\tEND DATE\tCARD NUMBER\n'
        res += '*****************************************************\n'
        for name in reslist['add'][r].keys():
            if name not in cardList.keys():
                res += "ADD:\t{0}\t{1}\t{2}\t*** MISSING CARD NUMBER ***\n".format(r, name,
                                                                                   reslist['add'][r][name].strftime(
                                                                                       "%Y-%m-%d"))
            else:
                res += "ADD:\t{0}\t{1}\t{2}\t{3}\n".format(r, name, reslist['add'][r][name].strftime("%Y-%m-%d"),
                                                           cardList[name])

        res += '*****************************************************\n\n'

    print("GENERATE RESULTS REMOVE")
    res += '\n\n'
    res += '*****************************************************\n'
    res += '***************** Entries to REMOVE *****************\n'
    res += '*****************************************************\n'
    res += 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    for r in reslist['remove'].keys():
        res += "\n\nRoom:{0} (CAS:{1})\n".format(r, Room2CAS[r])
        res += 'NAME\tCARD NUMBER\n'
        for name in reslist['remove'][r].keys():
            res += "REMOVE:\t{0}\t{1}\n".format(name, reslist['remove'][r][name])
        res += '*****************************************************\n'

    with open(fname, 'w') as f:
        f.write(res)


def generateFullAccessFiles(reslist, fname):
    # list all PAMT entries
    print("Generating FULL ACCESS LIST")
    # return

    res = '*****************************************************\n'
    res += '*****************  Entries LIST     *****************\n'
    res += '*****************************************************\n'
    res += 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for r in reslist.keys():
        res += "\n\nROOM:{0} (CAS:{1})\n".format(r, Room2CAS[r])
        res += 'ACTION\tNAME\tEND DATE\tCARD NUMBER\n'
        for name in reslist[r].keys():
            if name not in cardList.keys():
                res += "SET:\t{0}\t{1}\t*** MISSING CARD NUMBER ***\n".format(name, reslist[r][name]['to'].strftime(
                    "%Y-%m-%d"))
            else:
                res += "SET:\t{0}\t{1}\t{2}\n".format(name, reslist[r][name]['to'].strftime("%Y-%m-%d"), cardList[name])
        res += '*****************************************************\n'

    with open(fname, 'w') as f:
        f.write(res)


def generateCardNumberFiles(reslist, fname):
    # list all cards per CAS
    print("Generating CARDS LIST")

    res = '*****************************************************\n'
    res += '*****************    Cards LIST     *****************\n'
    res += '*****************************************************\n'
    res += 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for r in reslist.keys():
        for name in reslist[r].keys():
            cardList[name] = reslist[r][name]['card']

    res += 'NAME\tCARD#\n'
    for c in sorted(cardList.keys()):
        res += "{0}\t{1}\n".format(c, cardList[c])

    with open(fname, 'w') as f:
        f.write(res)


def generateMissingCardList(fname):
    # list persons without card number
    print("Generating MISSIG CARDS LIST")
    missing = []
    res = '*****************************************************\n'
    res += '*************    Missing Cards LIST     *************\n'
    res += '*****************************************************\n'
    res += 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for r in roomAccessPAMT.keys():
        for n in roomAccessPAMT[r].keys():
            if n not in cardList.keys() and n not in missing:
                missing.append(n)

    for n in sorted(missing):
        res += "{0}\n".format(n)

    with open(fname, 'w') as f:
        f.write(res)


def run():
    # start execution
    counter = 0
    with os.scandir('CAS-PAMT') as entries:
        for entry in entries:
            if not entry.is_file():
                continue
            counter += 1
            if entry.name.endswith('.txt'):
                room = re.sub("\.txt", "", entry.name)
                room = re.sub(r"PL 3216 \(EH7I\) ", "", room)
                print('#{0}:{1} -> ROOM:{2}'.format(counter, entry.name, room))
                prepareCASRooms(room)
                processCAS('CAS-PAMT/' + entry.name, CAS2Room[room])
            elif entry.name.endswith('.html'):
                processPAMT('CAS-PAMT/' + entry.name)

    PAMT2CAScheck()
    CAS2PAMTcheck()

    generateCardNumberFiles(roomAccessCAS, 'AMAG card list.log')
    generateResultsFiles(resultAccess, 'comparison results by room.log')
    generateFullAccessFiles(roomAccessPAMT, 'all rooms access.log')
    generateMissingCardList('missing cards.log')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('Number of arguments:', len(sys.argv), 'arguments.')

    run()
