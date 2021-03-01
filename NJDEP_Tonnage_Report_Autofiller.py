import re

import csv

from elevate import elevate


def intro_sequence():
    #Introduce user to the program in plain English, keeping in mind
    #they might not be used to working in the terminal. Ask for response
    #to confirm they're reading the on-screen dialogue.
    
    print('''

Hello.

I am a program designed to help you fill out the NJDEP's
Solid Waste Facility Monthly Disposal and Materials Recovery Report.

Instead of printing a report and typing it into the DEP's fillable
pdf form by hand, just run the report in PC Scale, save it as a CSV,
and give it to me.

I'll sort the data and use it to fill out the form automatically.

Sound good?
''')

    while True:
        print('''
        (Enter "Y" to continue)
        ''')
        response = input()
        if str(response) == 'Y' or 'y':
            break

    #Confirm the user is present, has their files ready, and is able
    #to give commands to the terminal
    while True:
        print('''
Great, let's start by creating the CSV files and getting them ready.

Go to PC Scale's "Site Specific" program and run reports "Form 1" and "Form 2a"
for the entire month you're reporting on. Once they're done processing, click the
"Export" button (envelope with a red arrow) and save the report as a CSV file 
with the default settings. You can name the report anything you want, but make
sure you don't save it to a secured folder I won't have access to. Do this for
both reports.
        
Once you're done, type "B" and press Enter to continue.
        ''')
        response = input()
        if str(response) == 'B' or 'b':
            break



def parse_files():

    #The following lists will help the program sort each row in the csv to the right sublist.
    #master_list is a "list of lists", each list containing the info for the corresponding
    #county of origin.

    out_of_state_munis = ['PENNSYLVANIA', 'DELAWARE', 'NEW YORK']

    burl_county_munis = ['BASS RIVER TWP', 'BEVERLY CITY', 'BORDENTOWN CITY', 'BORDENTOWN TWP',
                         'BURLINGTON CITY', 'BURLINGTON TWP', 'CHESTERFIELD TWP', 'CINNAMINSON TWP',
                         'DELANCO TWP', 'DELRAN TWP', 'EASTAMPTON TWP', 'EDGEWATER PARK TWP', 'EVESHAM TWP',
                         'FIELDSBORO BOROUGH', 'FLORENCE TWP', 'HAINESPORT TWP', 'LUMBERTON TWP', 'MANSFIELD TWP',
                         'MAPLE SHADE TWP', 'MEDFORD TWP', 'MEDFORD LAKES BOROUGH', 'MOORESTOWN TWP', 'MOUNT HOLLY TWP',
                         'MOUNT LAUREL TWP', 'NEW HANOVER TWP', 'NORTH HANOVER TWP', 'PALMYRA BOROUGH',
                         'PEMBERTON BOROUGH', 'PEMBERTON TWP', 'RIVERSIDE TWP', 'RIVERTON BOROUGH', 'SHAMONG TWP',
                         'SOUTHAMPTON TWP', 'SPRINGFIELD TWP', 'TABERNACLE TWP', 'WASHINGTON TWP', 'WESTAMPTON TWP',
                         'WILLINGBORO TWP', 'WOODLAND TWP', 'WRIGHTSTOWN BOROUGH']

    camd_county_munis = ['AUDUBON BOROUGH', 'AUDUBON PARK BOROUGH', 'BARRINGTON BOROUGH', 'BELLMAWR BOROUGH', 'BERLIN BOROUGH',
                         'BERLIN TWP', 'BROOKLAWN BOROUGH', 'CAMDEN CITY', 'CHERRY HILL TWP', 'CHESILHURST BOROUGH', 
                         'CLEMENTON BOROUGH', 'COLLINGSWOOD BOROUGH', 'GIBBSBORO BOROUGH', 'GLOUCESTER CITY',
                         'GLOUCESTER TWP', 'HADDON TWP', 'HADDONFIELD BOROUGH', 'HADDON HEIGHTS BOROUGH', 'HI-NELLA BOROUGH',
                         'LAUREL SPRINGS BOROUGH', 'LAWNSIDE BOROUGH', 'LINDENWOLD BOROUGH', 'MAGNOLIA BOROUGH', 
                         'MERCHANTVILLE BOROUGH', 'MOUNT EPHRAIM BOROUGH', 'OAKLYN BOROUGH', 'PENNSAUKEN TWP', 'PINE HILL BOROUGH',
                         'PINE VALLEY BOROUGH', 'RUNNEMEDE BOROUGH', 'SOMERDALE BOROUGH', 'STRATFORD BOROUGH', 'TAVISTOCK BOROUGH',
                         'VOORHEES TWP', 'WATERFORD TWP', 'WINSLOW TWP', 'WOODLYNNE BOROUGH']

    glou_county_munis = ['CLAYTON BOROUGH', 'DEPTFORD TWP', 'EAST GREENWICH TWP', 'ELK TWP', 'FRANKLIN TWP', 'GLASSBORO BOROUGH',
                         'GREENWICH TWP', 'HARRISON TWP', 'LOGAN TWP', 'MANTUA TWP', 'MONROE TWP', 'NATIONAL PARK BOROUGH',
                         'NEWFIELD BOROUGH', 'PAULSBORO BOROUGH', 'PITMAN BOROUGH', 'SOUTH HARRISON TWP', 'SWEDESBORO BOROUGH',
                         'WASHINGTON TWP', 'WENONAH BOROUGH', 'WEST DEPTFORD TWP', 'WESTVILLE BOROUGH', 'WOODBURY CITY',
                         'WOODBURY HEIGHTS BOROUGH', 'WOOLWICH TWP']


        #Ask for location of source data and explain how to specify

    print('''
        Where is the Form 1 CSV saved? Include the entire file path,
        including the filename and extension.
        ''')

    form_one_CSV_path = str(input())

    print('''
        Now do the same for Form 2a.
        ''')

    form_two_CSV_path = str(input())

    print('''
    What's the name of the template form? (Default: blank report)
    ''')
    template_name = input()
    if template_name == '':
        template_name = "blank report"

    print('''
    And where can I find the template form? (Hit enter for default)
    ''')
    template_location = input()
    if template_location == '':
        template_location = 'D:\\Files for Py\\Testing\\NJDEP SW REPORT\\'

    print('''
    What do you want me to name the completed form? (Default: newform)
    ''')

    completed_form_name = input()
    if completed_form_name == '':
        completed_form_name = "newform"


    print('''
    And where should I save the completed form? (Full path)
    ''')
    finished_product_location = input()

    print('''
    Okay, we're almost done. The only thing left is the info for
    the form headers. I'm going to list off the fields and you can
    just type in the answer and hit "enter". (Leave it blank to
    use the default answer).

    First... What is the facility name? (Default: Pennsauken
    Sanitary Landfill)
    ''')
    facility_name = input()
    if facility_name == '':
        facility_name = 'Pennsauken Sanitary Landfill'

    print('''
    Facility Location (Municipality/County/State): (Default:
    Pennsauken/Camden/NJ)
    ''')
    facility_location = input()
    if facility_location == '':
        facility_location = 'Pennsauken/Camden/NJ'

    print('''
    Submitted By: (Default: Bob Iuliucci)
    ''')
    submitted_by = input()
    if submitted_by == '':
        submitted_by = 'Bob Iuliucci'

    print('''
    Title of Submitter: (Default: Deputy Director)
    ''')
    submitter_title = input()
    if submitter_title == '':
        submitter_title = 'Deputy Director'

    print('''
    Phone Number: (Default: 856-665-8787)
    ''')
    phone_number = input()
    if phone_number == '':
        phone_number = '856-665-8787'

    print('''
    Program Interest Number: (Default: 132037)
    ''')
    interest_number = input()
    if interest_number == '':
        interest_number = '132037'

    print('''
    Reporting Period Month (Jan = 1, Nov = 11, etc.): (NO DEFAULT)
    ''')
    reporting_period_month = input()

    print('''
    Reporting Period Year (YYYY): (NO DEFAULT)
    ''')
    reporting_period_year = input()

    print('''
    Today's Date (MM/DD/YYYY): (NO DEFAULT)
    ''')
    today_date = input()

    print('''
    Alright - your work here is done. Give me a few seconds to fill out
    the form. I'll leave the finished copy where you told me.
    ''')

    header_dict = {}
    header_dict['Facility Name'] = facility_name
    header_dict['Facility Location'] = facility_location
    header_dict['Submitted By'] = submitted_by
    header_dict['Submitter Title'] = submitter_title
    header_dict['Phone Number'] = phone_number
    header_dict['Program Interest Number'] = interest_number
    header_dict['Reporting Period Month'] = reporting_period_month
    header_dict['Reporting Period Year'] = reporting_period_year
    header_dict["Today's Date"] = today_date
    header_dict['Finished Product Path'] = finished_product_location
    header_dict['Template Name'] = template_name
    header_dict['Finished Product Name'] = completed_form_name
    header_dict['Template Location'] = template_location


    elevate()


    #Using the built-in csv library to read the raw csv file for Form 1
    #and assign it to variable "raw_CSV_one".

    with open(form_one_CSV_path, 'r', encoding = 'utf-8') as file:
        raw_CSV_one = csv.reader(file)


    #Now we take the entire csv and whittle it down to just the information
    #we need. It's important to note the csv file comes to us poorly
    #formatted. The full table we want info from is expressed in the first
    #40 delimited items or so, but there are no new lines distinguishing rows.
    #The first new line in the csv comes after the end of the "table" After
    #the new line, the table is repeated. This is done hundreds of times,
    #causing the two lines above to return a list comprised of hundreds of
    #identical sublists, each containing the entire table.
    #
    #Also, the 13 numerical tonnage values we need are stored in the same cell,
    #so we'll need to fix that as well. 

        #First we separate the first item (a copy of the whole table) from the csv.
        #Then, in that table, we find the cell with the 13 numerical values we need.
        #For this report, they happen to be in the 25th cell (item[26]), which I
        #expect to always be true, but because I'm a bit paranoid about that,
        #I'll have the program first try to locate that cell by idenfifying the descriptor
        #cell which always immediately precedes it. If that fails, it'll just specify
        #item[26] as a plan B.

        CSV_one_data = ''
        n = 0

        for item in raw_CSV_one:
            if n < 1:
                p = 0
                while True:
                    item_content = item[p]        ################
                    if p + 1 == len(item):
                        CSV_one_data = item[26]
                        break
                    elif not item_content.startswith('10\t\t\t\tHousehold & Municipal\n'):
                        p += 1
                    else:
                        CSV_one_data = item[p+1]
                        break
                n += 1
            else:
                break
        CSV_one_data = CSV_one_data.replace(',','')

    #We are now done parsing the Form 1 CSV. The variable "form_one_tonnage_list"
    #contains a list of tonnages in this order: [type 10, 13, 13c, 23, 25, 27, 27A,
    #27I, Other, Total Inbound, Total Disposed, Total Recovered for Recycling,
    #Total Recovered for Beneficial Reuse]  

    form_one_tonnage_list = CSV_one_data.split()



    
    #Time to start parsing the Form 2 CSV, which will be more involved. The first
    #thing to note is that we're again given a "list of lists". This time, each
    #item is a list containing all the tonnage information for a given municipality
    #along with all the auxillary information contained in the spreadsheet. Just
    #like with Form 1, we need to make a new list organizing the tonnage in a more
    #practical way.

    with open(form_two_CSV_path, 'r', encoding = 'utf-8') as file:              #Use csv library to read csv file specified by user
        raw_CSV_two = csv.reader(file)

        sub_dict = {}                                                           #Create a dictionary with k = municipality name and v = list of tonnages
        for item in raw_CSV_two:
            t = item.index('MUNICIPALITY')
            municipality = item[t + 2]
            tonnages = item[(t + 3) : (t + 13)]
            s = "-"
            s = s.join(tonnages)
            s = s.replace(',','')
            s = s.split("-")
            tonnages = s
            sub_dict[municipality] = tonnages
            
        
        oos_dict = {}                                                           #Sort dict entries by key into different dicts by origin
        burl_dict = {}
        camd_dict = {}
        glou_dict = {}
        for item in sub_dict:
            if item in out_of_state_munis:
                oos_dict[item] = sub_dict[item]
            elif item in burl_county_munis:
                burl_dict[item] = sub_dict[item]
            elif item in camd_county_munis:
                camd_dict[item] = sub_dict[item]
            elif item in glou_county_munis:
                glou_dict[item] = sub_dict[item]


        oos_totals = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]              #Ugly and verbose code to add a "Total" row to each origin's dict
        for entry in oos_dict:
            n = 0
            for tonnage in oos_dict[entry]:
                oos_totals[n] += float(tonnage)
                n += 1
                if n > 8:
                    break
        oos_totals.append(sum(oos_totals[0:8]))
        counter = 0
        for x in oos_totals:
            y = '{:.2f}'.format(x)
            oos_totals[counter] = y
            counter += 1
        oos_dict['Totals'] = oos_totals

        burl_totals = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        for entry in burl_dict:
            n = 0
            for tonnage in burl_dict[entry]:
                burl_totals[n] += float(tonnage)
                n += 1
                if n > 8:
                    break
        burl_totals.append(sum(burl_totals[0:8]))
        counter = 0
        for x in burl_totals:
            y = '{:.2f}'.format(x)
            burl_totals[counter] = y
            counter += 1
        burl_dict['Totals'] = burl_totals

        camd_totals = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        for entry in camd_dict:
            n = 0
            for tonnage in camd_dict[entry]:
                camd_totals[n] += float(tonnage)
                n += 1
                if n > 8:
                    break
        camd_totals.append(sum(camd_totals[0:8]))
        counter = 0
        for x in camd_totals:
            y = '{:.2f}'.format(x)
            camd_totals[counter] = y
            counter += 1
        camd_dict['Totals'] = camd_totals
        for entry in camd_dict:
            for v in camd_dict[entry]:
                v = v.replace(',','')

        glou_totals = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        for entry in glou_dict:
            n = 0
            for tonnage in glou_dict[entry]:
                glou_totals[n] += float(tonnage)
                n += 1
                if n > 8:
                    break
        glou_totals.append(sum(glou_totals[0:8]))
        counter = 0
        for x in glou_totals:
            y = '{:.2f}'.format(x)
            glou_totals[counter] = y
            counter += 1
        glou_dict['Totals'] = glou_totals

        global finished_dict
        finished_dict = {}                                                      #Create the final dict out of info gathered
        finished_dict['Header Info'] = header_dict
        finished_dict['Form 1 Totals'] = form_one_tonnage_list
        finished_dict['Out of State'] = oos_dict
        finished_dict['Burlington County'] = burl_dict
        finished_dict['Camden County'] = camd_dict
        finished_dict['Gloucester County'] = glou_dict
        


        


def populate_form():

    form_one_input_list = [finished_dict['Header Info']['Facility Name'],               #Using info from finished_dict, put the desired form one input into a simple ordered list
                           finished_dict['Header Info']['Program Interest Number'],
                           finished_dict['Header Info']['Submitted By'],
                           finished_dict['Header Info']['Phone Number'],
                           finished_dict['Header Info']['Reporting Period Month'],
                           finished_dict['Header Info']['Reporting Period Year']]
    for i in finished_dict['Form 1 Totals'][0:8]:
        form_one_input_list.append(i)       ###################
    form_one_input_list.append('')
    for i in finished_dict['Form 1 Totals'][8:13]:
        form_one_input_list.append(i)       ##################
    form_one_input_list.append(finished_dict['Header Info']['Submitter Title'])
    form_one_input_list.append('')
    form_one_input_list.append(finished_dict['Header Info']["Today's Date"])


    form_two_oos_sublist = []
    if len(finished_dict['Out of State']) < 24:                                        #If there's 23 rows or less, we can fit all the info on one sheet
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_oos_sublist.append('Out of State')

        helper_list = []
        counter = 1
        for i in finished_dict['Out of State']:
            if counter < len(finished_dict['Out of State']):
                helper_list.append(i)
                for tonnage in finished_dict['Out of State'][i]:
                    helper_list.append(tonnage)
            counter += 1
        while len(helper_list) < 242:
            helper_list.append('')
        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        for i in finished_dict['Out of State']['Totals']:
            helper_list.append(i)

        for i in helper_list:
            form_two_oos_sublist.append(str(i))

    else:                                                                               #If there's more than 23 rows, we'll need to play with the list to fit the info on two pages
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_oos_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_oos_sublist.append('Out of State')

        helper_list = []
        counter = 1
        for i in finished_dict['Out of State']:
            if counter < 23:                    #fill up the first page
                helper_list.append(i)
                for tonnage in finished_dict['Out of State'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == 23:                 #bottom "total" row of sheet one is left blank to avoid confusion
                for i in range(10):
                    helper_list.append('')
                helper_list.append(str(finished_dict['Header Info']['Facility Name']))      #We also need to fill in another header for the second sheet
                helper_list.append(str(finished_dict['Header Info']['Facility Location']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Month']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Year']))
                helper_list.append('Out of State')
                counter += 1
            elif counter > 23 and counter < (len(finished_dict['Out of State']) - 1):    #Now we fill in the rows normally until...
                helper_list.append(i)
                for tonnage in finished_dict['Out of State'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == len(finished_dict['Out of State']) - 1:     #After the second to last row, we fill in the rest of sheet two with blank cells
                helper_list.append(i)
                for tonnage in finished_dict['Out of State'][i]:
                    helper_list.append(tonnage)
                while len(helper_list) < 499:
                    helper_list.append('')
                counter += 1
            elif counter == len(finished_dict['Out of State']):         #To finish, we fill in the very last row of sheet two with our totals (Leaving out the "Totals" key)
                for tonnage in finished_dict['Out of State'][i]:
                    helper_list.append(tonnage)

        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        helper_list.insert(258, helper_list.pop(279))   #Again for page two

        for i in helper_list:
            form_two_oos_sublist.append(str(i))


    form_two_burl_sublist = []
    if len(finished_dict['Burlington County']) < 24:                                        #If there's 23 rows or less, we can fit all the info on one sheet
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_burl_sublist.append('Burlington County')

        helper_list = []
        counter = 1
        for i in finished_dict['Burlington County']:
            if counter < len(finished_dict['Burlington County']):
                helper_list.append(i)
                for tonnage in finished_dict['Burlington County'][i]:
                    helper_list.append(tonnage)
            counter += 1
        while len(helper_list) < 242:
            helper_list.append('')
        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        for i in finished_dict['Burlington County']['Totals']:
            helper_list.append(i)

        for i in helper_list:
            form_two_burl_sublist.append(str(i))

    else:                                                                               #If there's more than 23 rows, we'll need to play with the list to fit the info on two pages
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_burl_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_burl_sublist.append('Burlington County')

        helper_list = []
        counter = 1
        for i in finished_dict['Burlington County']:
            if counter < 23:                    #fill up the first page
                helper_list.append(i)
                for tonnage in finished_dict['Burlington County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == 23:                 #bottom "total" row of sheet one is left blank to avoid confusion
                for i in range(10):
                    helper_list.append('')
                helper_list.append(str(finished_dict['Header Info']['Facility Name']))      #We also need to fill in another header for the second sheet
                helper_list.append(str(finished_dict['Header Info']['Facility Location']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Month']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Year']))
                helper_list.append('Burlington County')
                counter += 1
            elif counter > 23 and counter < (len(finished_dict['Burlington County']) - 1):    #Now we fill in the rows normally until...
                helper_list.append(i)
                for tonnage in finished_dict['Burlington County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == len(finished_dict['Burlington County']) - 1:     #After the second to last row, we fill in the rest of sheet two with blank cells
                helper_list.append(i)
                for tonnage in finished_dict['Burlington County'][i]:
                    helper_list.append(tonnage)
                while len(helper_list) < 499:
                    helper_list.append('')
                counter += 1
            elif counter == len(finished_dict['Burlington County']):         #To finish, we fill in the very last row of sheet two with our totals (Leaving out the "Totals" key)
                for tonnage in finished_dict['Burlington County'][i]:
                    helper_list.append(tonnage)

        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        helper_list.insert(258, helper_list.pop(279))   #Again for page two

        for i in helper_list:
            form_two_burl_sublist.append(str(i))


    form_two_camd_sublist = []
    if len(finished_dict['Camden County']) < 24:                                        #If there's 23 rows or less, we can fit all the info on one sheet
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_camd_sublist.append('Camden County')

        helper_list = []
        counter = 1
        for i in finished_dict['Camden County']:
            if counter < len(finished_dict['Camden County']):
                helper_list.append(i)
                for tonnage in finished_dict['Camden County'][i]:
                    helper_list.append(tonnage)
            counter += 1
        while len(helper_list) < 242:
            helper_list.append('')
        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        for i in finished_dict['Camden County']['Totals']:
            helper_list.append(i)

        for i in helper_list:
            form_two_camd_sublist.append(str(i))

    else:                                                                               #If there's more than 23 rows, we'll need to play with the list to fit the info on two pages
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_camd_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_camd_sublist.append('Camden County')

        helper_list = []
        counter = 1
        for i in finished_dict['Camden County']:
            if counter < 23:                    #fill up the first page
                helper_list.append(i)
                for tonnage in finished_dict['Camden County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == 23:                 #bottom "total" row of sheet one is left blank to avoid confusion
                for i in range(10):
                    helper_list.append('')
                helper_list.append(str(finished_dict['Header Info']['Facility Name']))      #We also need to fill in another header for the second sheet
                helper_list.append(str(finished_dict['Header Info']['Facility Location']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Month']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Year']))
                helper_list.append('Camden County')
                counter += 1
            elif counter > 23 and counter < (len(finished_dict['Camden County']) - 1):    #Now we fill in the rows normally until...
                helper_list.append(i)
                for tonnage in finished_dict['Camden County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == len(finished_dict['Camden County']) - 1:     #After the second to last row, we fill in the rest of sheet two with blank cells
                helper_list.append(i)
                for tonnage in finished_dict['Camden County'][i]:
                    helper_list.append(tonnage)
                while len(helper_list) < 499:
                    helper_list.append('')
                counter += 1
            elif counter == len(finished_dict['Camden County']):         #To finish, we fill in the very last row of sheet two with our totals (Leaving out the "Totals" key)
                for tonnage in finished_dict['Camden County'][i]:
                    helper_list.append(tonnage)

        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        helper_list.insert(258, helper_list.pop(279))   #Again for page two

        for i in helper_list:
            form_two_camd_sublist.append(str(i))


    form_two_glou_sublist = []
    if len(finished_dict['Gloucester County']) < 24:                                        #If there's 23 rows or less, we can fit all the info on one sheet
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_glou_sublist.append('Gloucester County')

        helper_list = []
        counter = 1
        for i in finished_dict['Gloucester County']:
            if counter < len(finished_dict['Gloucester County']):
                helper_list.append(i)
                for tonnage in finished_dict['Gloucester County'][i]:
                    helper_list.append(tonnage)
            counter += 1
        while len(helper_list) < 242:
            helper_list.append('')
        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        for i in finished_dict['Gloucester County']['Totals']:
            helper_list.append(i)

        for i in helper_list:
            form_two_glou_sublist.append(str(i))

    else:                                                                               #If there's more than 23 rows, we'll need to play with the list to fit the info on two pages
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Facility Name']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Facility Location']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Reporting Period Month']))
        form_two_glou_sublist.append(str(finished_dict['Header Info']['Reporting Period Year']))
        form_two_glou_sublist.append('Gloucester County')

        helper_list = []
        counter = 1
        for i in finished_dict['Gloucester County']:
            if counter < 23:                    #fill up the first page
                helper_list.append(i)
                for tonnage in finished_dict['Gloucester County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == 23:                 #bottom "total" row of sheet one is left blank to avoid confusion
                for i in range(10):
                    helper_list.append('')
                helper_list.append(str(finished_dict['Header Info']['Facility Name']))      #We also need to fill in another header for the second sheet
                helper_list.append(str(finished_dict['Header Info']['Facility Location']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Month']))
                helper_list.append(str(finished_dict['Header Info']['Reporting Period Year']))
                helper_list.append('Gloucester County')
                counter += 1
            elif counter > 23 and counter < (len(finished_dict['Gloucester County']) - 1):    #Now we fill in the rows normally until...
                helper_list.append(i)
                for tonnage in finished_dict['Gloucester County'][i]:
                    helper_list.append(tonnage)
                counter += 1
            elif counter == len(finished_dict['Gloucester County']) - 1:     #After the second to last row, we fill in the rest of sheet two with blank cells
                helper_list.append(i)
                for tonnage in finished_dict['Gloucester County'][i]:
                    helper_list.append(tonnage)
                while len(helper_list) < 499:
                    helper_list.append('')
                counter += 1
            elif counter == len(finished_dict['Gloucester County']):         #To finish, we fill in the very last row of sheet two with our totals (Leaving out the "Totals" key)
                for tonnage in finished_dict['Gloucester County'][i]:
                    helper_list.append(tonnage)

        helper_list.insert(1,helper_list.pop(22))       #Move Muni 3 name to index 1 because of weird csv order
        helper_list.insert(258, helper_list.pop(279))   #Again for page two

        for i in helper_list:
            form_two_glou_sublist.append(str(i))

    template_name = finished_dict['Header Info']['Template Name'] + ".pdf"

    finished_input_list = []
    finished_input_list.append(template_name)
    for i in form_one_input_list:
        finished_input_list.append(str(i))
    for i in form_two_oos_sublist:
        finished_input_list.append(i)
    for i in form_two_burl_sublist:
        finished_input_list.append(i)
    for i in form_two_camd_sublist:
        finished_input_list.append(i)
    for i in form_two_glou_sublist:
        finished_input_list.append(i)
    while len(finished_input_list) < 2076:      #After putting in info, leave any unused Form 2a and 2b sheets blank
        finished_input_list.append('')
    finished_input_list.append(finished_dict['Header Info']['Facility Name'])
    finished_input_list.append(finished_dict['Header Info']['Facility Location'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Month'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Year'])
    while len(finished_input_list) < 2310:      #After filling in header info for form 2c, leave rest blank
        finished_input_list.append('')
    finished_input_list.append(finished_dict['Header Info']['Facility Name'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Month'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Year'])
    while len(finished_input_list) < 2543:      #Same for 3a
        finished_input_list.append('')
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Month'])
    finished_input_list.append(finished_dict['Header Info']['Facility Name'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Year'])
    while len(finished_input_list) < 2775:      #Same for 3b
        finished_input_list.append('')
    finished_input_list.append(finished_dict['Header Info']['Facility Name'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Month'])
    finished_input_list.append(finished_dict['Header Info']['Reporting Period Year'])
    while len(finished_input_list) < 3007:
        finished_input_list.append('')
        

    with open(finished_dict['Header Info']['Template Location'] + finished_dict['Header Info']['Template Name'] + ".csv", 'r', encoding = 'utf-8') as file:
        template_csv = csv.reader(file)
        row_counter = 1
        fields_list = []
        for row in template_csv:
            if row_counter == 1:
                for i in row:
                    fields_list.append(str(i))
            row_counter += 1


    print('''
    
Would you like to inspect the input list before I fill the form out? (Y/N)

''')
    answer = input()
    if answer == 'Y' or answer == 'y' or answer == 'Yes' or answer == 'yes' or answer == 'YES':
        print(len(form_one_input_list))
        print('')
        print(form_one_input_list)
        print('')
        print(len(form_two_oos_sublist))
        print('')
        print(form_two_oos_sublist)
        print('')
        print(len(form_two_burl_sublist))
        print('')
        print(form_two_burl_sublist)
        print('')
        print(len(form_two_camd_sublist))
        print('')
        print(form_two_camd_sublist)
        print('')
        print(len(form_two_glou_sublist))
        print('')
        print(form_two_glou_sublist)
        print('')
        print(len(finished_input_list))
        print('')
        print(finished_input_list)


    print('''

        

        
    
When you're ready, hit enter to finish. You'll see the finished csv file
in the folder you specified earlier. Use foxit reader to import the data
from the csv into the blank report pdf and you're done.

''')
    answer = input()
    while True:
        if answer == '':
            break

    filename = finished_dict['Header Info']['Finished Product Path'] + finished_dict['Header Info']['Finished Product Name'] + ".csv"
    with open(filename, 'w', newline='') as csvfile:  
        csvwriter = csv.writer(csvfile)  
        csvwriter.writerow(fields_list)  
        csvwriter.writerow(finished_input_list) 
    


    

    
   




intro_sequence()

parse_files()

populate_form()