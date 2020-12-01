# This is a Python script to compare DC-PAMT and AMAG CAS data for access control in DC Blonie

# Copyright masti 2020

import os
import re
import sys
from datetime import datetime
from fpdf import FPDF, HTMLMixin

# global translations
from typing import Dict, Any, Union

#TODO: make translation tables dynamic from a file
PAMT2CAS: Dict[Union[str, Any], Union[str, Any]] = {
    'C01 Main Entrance': 'C01+C02 Main Entrance',
    'C02 Lobby': 'C01+C02 Main Entrance',
    'C02 Main entry + lobby': 'C01+C02 Main Entrance',
    'C03 Conference room': 'C03 CONFERENCE ROOM',
    'C04 Customer office': 'C04',
    'C05 Customer office eService': 'C05',
    'C06 Customer office': 'C06',
    'C13 Customer office PEK': 'C13',
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

#TODO: make translation tables dynamic from a file
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

#TODO: make translation tables dynamic from a file
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

#TODO: make translation tables dynamic from a file
CAS2Room = {
    'C01+C02 Main Entrance': ['C01', 'C02'],
    'C03 CONFERENCE ROOM': ['C03'],
    'C04': ['C04'],
    'C05': ['C05'],
    'C06': ['C06'],
    'C13': ['C13'],
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

#TODO: make translation tables dynamic from a file
Room2CAS = {
    'C01': 'C01+C02 Main Entrance',
    'C02': 'C01+C02 Main Entrance',
    'C03': 'C03 CONFERENCE ROOM',
    'C04': 'C04',
    'C05': 'C05',
    'C06': 'C06',
    'C13': 'C13',
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

# regexes to extract info from DC-PAMT specific HTML export files
#TODO: prepare for other tools
pamtR = re.compile(r"(?s)<table.*?<table.*?</table>.*?</table>")
roomsR = re.compile(r"(?s)Rooms:</small></td><td style='width: 585px;'>(?P<rooms>.*?)</td></tr>")
roomR = re.compile(r"(?P<room>.*?)<br>")
dateR = re.compile(
    r"(?s)Date:</small></td><td style='width: 585px;'>(?P<start>[\d\.]*)&nbsp;-&nbsp;(?P<end>[\d\.]*)</td></tr><tr>")
visitorR = re.compile(
    r"(?s)Visitor:</small></td><td style='width: 585px;'>(?P<visitor>.*?)&nbsp;<small>(?P<comment>.*?)</small><br></td></tr>")
guestsR = re.compile(r"(?s)Guests:</small></td><td style='width: 585px;'>(?P<guests>.*?)</td></tr>")
guestR = re.compile(r"(?P<name>.*?)&nbsp;.*?</small><br>")

roomAccessCAS = {}
personsAccess = {}
roomAccessPAMT = {}
resultAccess = dict(add={}, remove={})
cardList = {}


class dataInput():


    def names(self, name):
        # extract last name form name
        l = name.split(', ')[0]
        f = name.split(', ')[1]
        return (l, f)


    def cleanupCasLine(self, last, first):
        # remove Contractor, keycard
        # reverse firstname & name
        #return normalized name LastName, FirstName

        print("LAST:{0}".format(last))
        if 'contractor' in last:
            last = last.replace('contractor ', '')
            try:
                f, l = tuple(last.split())
            except ValueError:
                f, l = ('contractor','keycard')
        else:
            f = first.split()[0]
            l = last.split()[0]
        # print(f'Last:{l} First:{f}')
        return ("{0}, {1}".format(l.capitalize(), f.capitalize()))


    def addRoomAccessCAS(self, name, card, to, room):
        # add access by room
        l, f = self.names(name)
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


    def treatCasLine(self, line, rooms):
        # process one line from CAS file
        # return dict

        print("LINE:{0}".format(line))
        l, f, c, i, a, to = tuple(line.split('\t'))
        n = self.cleanupCasLine(l, f)
        try:
            to = datetime.strptime(to, "%d/%m/%Y")
        except ValueError:
            print('^^^ VALUE ERROR in room:{0} ^^^'.format(rooms))
            return
        for r in rooms:
            self.addRoomAccessCAS(n, c, to, r)


    def processCAS(self, fname, rooms):
        # process CAS file fname
        print('Processing CAS file:{0}'.format(fname))
        lcount = 0
        with open(fname, 'r') as f:
            for l in f.readlines()[6:-5]:
                self.treatCasLine(l.strip('\n').lower(), rooms)
                lcount += 1
            print('Processed {0} lines'.format(lcount))
        return


    def prepareCASRooms(self, room):
        for r in CAS2Room[room]:
            if r not in roomAccessCAS.keys():
                roomAccessCAS[r] = {}


    def preparePAMTRooms(self, ):
        for r in PAMT2Room:
            if r not in roomAccessPAMT.keys():
                roomAccessPAMT[r] = {}


    def getRoomsList(self, line):
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


    def getRequestDates(self, line):
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


    def cleanGuestName(self, name):
        # return guest name in a form Surname, Name
        l = re.split(", | ", name)
        return ("{0}, {1}".format(l[0].lower().capitalize(), l[1].lower().capitalize()))


    def getVisitor(self, line):
        # return visitor name for the request
        # None if does not need access
        v = visitorR.search(line)
        if v:
            visitor = v.group('visitor')
            comment = v.group('comment')
            if 'no action' in comment.lower():
                return (None)
            else:
                return (self.cleanGuestName(visitor))
        return (None)


    def getGuests(self, line):
        # return list of guest names for the request
        # None if does not need access
        guests = guestsR.search(line)
        guestsList = []
        if guests:
            gcount = 1
            for g in guestR.finditer(guests.group('guests')):
                gcount += 1
                guestsList.append(self.cleanGuestName(g.group('name')))
        return (guestsList)


    def addRoomAccessPAMT(self, name, to, room):
        # add access by room from PAMT
        if room not in roomAccessPAMT.keys():
            roomAccessPAMT[room] = {}
        if name not in roomAccessPAMT[room].keys():
            l, f = self.names(name)
            roomAccessPAMT[room][name] = {
                'to': to,
                'Last': l,
                'First': f
            }
        elif roomAccessPAMT[room][name]['to'] < to:
            # change access end date to later
            roomAccessPAMT[room][name]['to'] = to
        return


    def treatPAMTLine(self, line):
        # process one line from PAMT report = 1 record
        rooms = self.getRoomsList(line)
        startDate, endDate = self.getRequestDates(line)
        visitor = self.getVisitor(line)
        if visitor:
            visitorList = [visitor, ]
        else:
            visitorList = []
        visitorList.extend(self.getGuests(line))
        for r in rooms:
            for v in visitorList:
                self.addRoomAccessPAMT(v, endDate, r)


    def processPAMT(self, fname):
        # process PAMT file fname
        print('Processing PAMT file:{0}'.format(fname))
        lcount = 0
        with open(fname, 'r') as f:
            file = f.read()

        for l in pamtR.findall(file):
            self.treatPAMTLine(l)
            lcount += 1
        print('Processed {0} lines'.format(lcount))


    def PAMT2CAScheck(self):
        # compare PAMT 2 CAS entries
        for r in roomAccessPAMT.keys():
            for n in roomAccessPAMT[r].keys():
                try:
                    casAccess = roomAccessCAS[r][n]
                except KeyError:
                    if r not in resultAccess['add'].keys():
                        resultAccess['add'][r] = {}
                    resultAccess['add'][r][n] = roomAccessPAMT[r][n]['to']


    def PAMTCASbyPerson(self):
        """

        :rtype: object
        """
        # use resultAccess['add'] and resultAccess['remove'] to create by person view
        print("PAMTCASbyPerson ADD")
        print(resultAccess['add'])
        for r in resultAccess['add'].keys(): # iterate by room
            for n in resultAccess['add'][r].keys(): #iterate by person
                if n not in personsAccess.keys():
                    personsAccess[n] = {'add':[], 'remove':[]}
                personsAccess[n]['add'].append({'room':r, 'to':resultAccess['add'][r][n]})
        print("PAMTCASbyPerson REMOVE")
        print(resultAccess['remove'])
        for r in resultAccess['remove'].keys(): # iterate by room
            for n in resultAccess['remove'][r].keys(): #iterate by person
                if n not in personsAccess.keys():
                    personsAccess[n] = {'add':[], 'remove':[]}
                personsAccess[n]['remove'].append({'room':r, 'card':resultAccess['remove'][r][n]})
        print("PAMTCASbyPerson RESULT")
        print(personsAccess)
        return(personsAccess)

    def CAS2PAMTcheck(self):
        # compare CAS 2 PAMT entries
        for r in roomAccessCAS.keys():
            for n in roomAccessCAS[r].keys():
                try:
                    pamtAccess = roomAccessPAMT[r][n]
                except KeyError:
                    if r not in resultAccess['remove'].keys():
                        resultAccess['remove'][r] = {}
                    resultAccess['remove'][r][n] = roomAccessCAS[r][n]['card']


class MyFPDF(FPDF, HTMLMixin):

    def __init__(self, title):
        self.title = title
        super().__init__()

    def header(self):
        # Arial bold 15
        self.set_font('Arial', 'B', 12)
        self.set_text_color(128)
        # Title
        self.cell(0, 9, self.title, 0, 0, 'C')
        # Line break
        self.ln(10)

    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'B', 8)
        # Text color in gray
        self.set_text_color(128)
        # Page number
        self.cell(30, 10, '' , 0, 0, 'C')
        self.cell(130, 10, 'Generated on: {0}\n\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")) , 0, 0, 'C')
        self.cell(30, 10, 'Page ' + str(self.page_no()) + ' of {nb}' , 0, 0, 'R')

    def chapter_title(self, num, label):
        # Arial 12
        self.set_font('Arial', '', 12)
        # Background color
        self.set_fill_color(200, 220, 255)
        # Title
        self.cell(0, 6, 'Chapter %d : %s' % (num, label), 0, 1, 'L', 1)
        # Line break
        self.ln(4)

    def chapter_body(self, name):
        # Read text file
        with open(name, 'rb') as fh:
            txt = fh.read().decode('latin-1')
        # Times 12
        self.set_font('Arial', '', 12)
        # Output justified text
        self.multi_cell(0, 5, txt)
        # Line break
        self.ln()
        # Mention in italics
        self.set_font('', 'I')
        self.cell(0, 5, '(end of excerpt)')

    def print_chapter(self, num, title, name):
        self.add_page()
        self.chapter_title(num, title)
        self.chapter_body(name)

    def table_header(self):
        self.set_font('Arial', 'B', 10)

    def table_row(self):
        self.set_font('Arial', '', 10)

    def section_header(self):
        self.set_font('Arial', 'B', 16)

    def section_subheader(self):
        self.set_font('Arial', 'B', 12)

class dataOutput():

    def result_files_by_room(self, reslist, fname):
        # list to add entries
        print("GENERATE RESULTS ADD")

        pdf = MyFPDF(fname)
        pdf.alias_nb_pages()
        pdf.add_page()

        pdf.set_creator('Marek Stelmasik (marek.stelmasik@pl.ibm.com)')
        pdf.set_subject('fname')
        pdf.set_title(fname)
        pdf.set_display_mode('fullpage', 'default')

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, fname, 'B', 1, 'C')
        pdf.ln()

        pdf.cell(190, 10, 'CARDS TO ADD', 'B', 1, 'C')
        pdf.ln()

        for r in reslist['add'].keys():
            pdf.section_header()
            pdf.cell(190, 10, 'ROOM: {0}'.format(r), 'B', 1, 'L')
            pdf.section_subheader()
            pdf.cell(190, 10, 'CAS: {0}'.format(Room2CAS[r]), 'B', 1, 'L')

            pdf.table_header()
            pdf.cell(20, 6, 'ACTION', 'B', 0, 'L')
            pdf.cell(20, 6, 'ROOM', 'B', 0, 'L')
            pdf.cell(50, 6, 'NAME', 'B', 0, 'L')
            pdf.cell(40, 6, 'DATE', 'B', 0, 'L')
            pdf.cell(25, 6, 'CARD', 'B', 0, 'L')
            pdf.ln()
            pdf.table_row()

            for name in reslist['add'][r].keys():
                if name not in cardList.keys():
                    pdf.cell(20, 6, 'ADD', 0, 0, 'L')
                    pdf.cell(20, 6, '[{0}]'.format(r), 0, 0, 'L')
                    pdf.cell(50, 6, name, 0, 0, 'L')
                    pdf.cell(40, 6, reslist['add'][r][name].strftime("%Y-%m-%d"), 0, 0, 'L')
                    pdf.cell(25, 6, '- - -', 0, 0, 'C')
                else:
                    pdf.cell(20, 6, 'ADD', 0, 0, 'L')
                    pdf.cell(20, 6, '[{0}]'.format(r), 0, 0, 'L')
                    pdf.cell(50, 6, name, 0, 0, 'L')
                    pdf.cell(40, 6, reslist['add'][r][name].strftime("%Y-%m-%d"), 0, 0, 'L')
                    pdf.cell(25, 6, cardList[name], 0, 0, 'L')
                pdf.ln()

            pdf.ln()

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, 'CARDS TO REMOVE', 'B', 1, 'C')
        pdf.ln()

        print("GENERATE RESULTS REMOVE")


        for r in reslist['remove'].keys():
            pdf.section_header()
            pdf.cell(190, 10, 'ROOM: {0}'.format(r), 'B', 1, 'L')
            pdf.section_subheader()
            pdf.cell(190, 10, 'CAS: {0}'.format(Room2CAS[r]), 'B', 1, 'L')

            pdf.table_header()
            pdf.cell(20, 6, 'ACTION', 'B', 0, 'L')
            pdf.cell(20, 6, 'ROOM', 'B', 0, 'L')
            pdf.cell(50, 6, 'NAME', 'B', 0, 'L')
            pdf.cell(25, 6, 'CARD', 'B', 0, 'L')
            pdf.ln()
            pdf.table_row()
            for name in reslist['remove'][r].keys():
                pdf.cell(20, 6, 'REMOVE', 0, 0, 'L')
                pdf.cell(20, 6, '[{0}]'.format(r), 0, 0, 'L')
                pdf.cell(50, 6, name, 0, 0, 'L')
                pdf.cell(25, 6, cardList[name], 0, 0, 'L')
                pdf.ln()
            pdf.ln()

        # pdf.write_html(html)
        pdf.output(fname + '.pdf', 'F')

    def result_files_by_person(self, reslist, fname):
        # list to add entries
        print("GENERATE RESULTS BY PERSON")

        pdf = MyFPDF(fname)
        pdf.alias_nb_pages()
        pdf.add_page()

        pdf.set_creator('Marek Stelmasik (marek.stelmasik@pl.ibm.com)')
        pdf.set_subject('fname')
        pdf.set_title(fname)
        pdf.set_display_mode('fullpage', 'default')

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, fname, 'B', 1, 'C')
        pdf.ln()

        for n in reslist.keys():
            pdf.section_header()
            pdf.cell(190, 10, 'NAME: {0}'.format(n), 'B', 1, 'L')
            pdf.section_subheader()
            if n not in cardList.keys():
                pdf.cell(190, 10, 'CARD: - - -', 'B', 1, 'L')
            else:
                pdf.cell(190, 10, 'CARD: {0}'.format(cardList[n]), 'B', 1, 'L')

            pdf.table_header()
            pdf.cell(20, 6, 'ACTION', 'B', 0, 'L')
            pdf.cell(80, 6, 'ROOM', 'B', 0, 'L')
            pdf.cell(40, 6, 'DATE', 'B', 0, 'L')
            pdf.ln()
            pdf.table_row()
            for a in reslist[n].keys():
                if a == 'add':
                    for r in reslist[n][a]:
                        pdf.cell(20, 5, 'ADD:', 0, 0, 'L')
                        pdf.cell(80, 5, 'ROOM: [{0}] ({1})'.format(r['room'], Room2CAS[r['room']]), 0, 0, 'L')
                        pdf.cell(40, 5, r['to'].strftime("%Y-%m-%d"), 0, 0, 'L')
                        pdf.ln()
                else:
                    for r in reslist[n][a]:
                        pdf.cell(20, 5, 'REMOVE:', 0, 0, 'L')
                        pdf.cell(80, 5, 'ROOM: [{0}] ({1})'.format(r['room'], Room2CAS[r['room']]), 0, 0, 'L')
                        pdf.ln()

            pdf.ln()

        # pdf.write_html(html)
        pdf.output(fname + '.pdf', 'F')


    def full_access_setup_file(self, reslist, fname):
        # list all PAMT entries
        print("Generating FULL ACCESS LIST")
        # return

        pdf = MyFPDF(fname)
        pdf.alias_nb_pages()
        pdf.add_page()


        pdf.set_creator('Marek Stelmasik (marek.stelmasik@pl.ibm.com)')
        pdf.set_subject('fname')
        pdf.set_title(fname)
        pdf.set_display_mode('fullpage', 'default')

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, fname, 'B', 1, 'C')
        pdf.ln()


        for r in reslist.keys():
            pdf.section_header()
            pdf.cell(190, 10, 'ROOM: {0}'.format(r), 'B', 1, 'L')

            pdf.section_subheader()
            pdf.cell(190, 10, 'CAS: {0}'.format(Room2CAS[r]), 'B', 1, 'L')

            pdf.table_header()
            pdf.cell(20, 6, 'ACTION', 'B', 0, 'L')
            pdf.cell(80, 6, 'NAME', 'B', 0, 'L')
            pdf.cell(40, 6, 'DATE', 'B', 0, 'L')
            pdf.cell(25, 6, 'CARD', 'B', 0, 'L')
            pdf.ln()
            pdf.table_row()
            for name in reslist[r].keys():
                if name not in cardList.keys():
                    pdf.cell(20, 5, 'SET:', 0, 0, 'L')
                    pdf.cell(80, 5, name, 0, 0, 'L')
                    pdf.cell(40, 5, reslist[r][name]['to'].strftime("%Y-%m-%d"), 0, 0, 'L')
                    pdf.cell(25, 5, '- - -', 0, 1, 'C')
                else:
                    pdf.cell(20, 5, 'SET:', 0, 0, 'L')
                    pdf.cell(80, 5, name, 0, 0, 'L')
                    pdf.cell(40, 5, reslist[r][name]['to'].strftime("%Y-%m-%d"), 0, 0, 'L')
                    pdf.cell(25, 5, cardList[name], 0, 1, 'L')

            pdf.add_page()

        # pdf.write_html(html)
        pdf.output(fname + '.pdf', 'F')


    def card_numbers_CAS(self, reslist, fname):
        # list all cards per CAS
        print("Generating CARDS LIST")

        pdf = MyFPDF(fname)
        pdf.alias_nb_pages()
        pdf.add_page()


        pdf.set_creator('Marek Stelmasik (marek.stelmasik@pl.ibm.com)')
        pdf.set_subject(fname)
        pdf.set_title(fname)
        pdf.set_display_mode('fullpage')

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, 'AMAG cards list', 'B', 1, 'C')
        pdf.ln()


        for r in reslist.keys():
            for name in reslist[r].keys():
                cardList[name] = reslist[r][name]['card']

        pdf.table_header()
        pdf.cell(80, 6, 'NAME', 'B', 0, 'L')
        pdf.cell(25, 6, 'CARD', 'B', 0, 'L')
        pdf.ln()
        pdf.table_row()
        for c in sorted(cardList.keys()):
            pdf.cell(80, 6, c, 0, 0, 'L')
            pdf.cell(25, 6, cardList[c], 0, 0, 'L')
            pdf.ln()

        # pdf.write_html(html)
        pdf.output(fname+'.pdf', 'F')



    def missing_card_list(self, fname):
        # list persons without card number
        print("Generating MISSIG CARDS LIST")
        missing = []
        pdf = MyFPDF(fname)
        pdf.alias_nb_pages()
        pdf.add_page()

        pdf.set_creator('Marek Stelmasik (marek.stelmasik@pl.ibm.com)')
        pdf.set_subject(fname)
        pdf.set_title(fname)
        pdf.set_display_mode('fullpage')

        pdf.set_font('Arial', 'B', 24)
        pdf.cell(190, 10, 'AMAG missing cards list', 'B', 1, 'C')
        pdf.ln()

        for r in roomAccessPAMT.keys():
            for n in roomAccessPAMT[r].keys():
                if n not in cardList.keys() and n not in missing:
                    missing.append(n)

        pdf.table_row()
        for n in sorted(missing):
            pdf.cell(80, 6, n, 0, 0, 'L')
            pdf.ln()

        # pdf.write_html(html)
        pdf.output(fname + '.pdf', 'F')


def run():
    # start execution
    input = dataInput()
    output = dataOutput()

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
                input.prepareCASRooms(room)
                input.processCAS('CAS-PAMT/' + entry.name, CAS2Room[room])
            elif entry.name.endswith('.html'):
                input.processPAMT('CAS-PAMT/' + entry.name)

    input.PAMT2CAScheck()
    input.CAS2PAMTcheck()
    input.PAMTCASbyPerson()

    output.card_numbers_CAS(roomAccessCAS, 'AMAG card list')
    output.result_files_by_room(resultAccess, 'comparison results by room')
    output.full_access_setup_file(roomAccessPAMT, 'all rooms access')
    output.result_files_by_person(personsAccess, 'comparison results by person')
    output.missing_card_list('AMAG missing cards')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('Number of arguments:', len(sys.argv), 'arguments.')

    run()
