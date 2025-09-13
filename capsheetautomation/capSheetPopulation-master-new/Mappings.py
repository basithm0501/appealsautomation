

moveCells_OrgMaintenance = {'B24': 'AC', 'C24': 'AD', 'B25': 'AE', 'C25': 'AF', 'B26': 'AG', 'C26': 'AH', 'B27': 'AI',
                            'C27': 'AJ', 'B28': 'AK', 'C28': 'AL', 'B30': 'AM',
                            'C30': 'AN', 'B31': 'AO', 'C31': 'AP', 'B32': 'AQ', 'C32': 'AR', 'B33': 'AS', 'C33': 'AT',
                            'B35': 'AU', 'C35': 'AV', 'B37': 'AW'
                            }

# Programming columns
# Program Name, Event Description, Date, Attendance, Location, Admission, Room Rent and Equip., Advertising, Food, Supplies/Decorations, Duplications, Contracts(not included yet), Other,

moveCells_Programming = {'A2': ['BX', 'FU', 'JR'], 'A3': ['BY', 'FV', 'JS'], 'B3': ['BZ', 'FW', 'JT'],
                         'C3': ['CA', 'FX', 'JU'],
                         'D3': ['CB', 'FZ', 'JW'], 'E3': ['CC', 'GA', 'JX'], 'B5': ['CF', 'GC', 'JZ'],
                         'C5': ['CG', 'GD', 'JZ'], 'B6': ['CH', 'GE', 'KB'],
                         'C6': ['CI', 'GF', 'KC'], 'B7': ['CJ', 'GG', 'KD'], 'C7': ['CK', 'GH', 'KE'],
                         'B8': ['CL', 'GI', 'KF'], 'C8': ['CM', 'GJ', 'KG'],
                         'B9': ['CN', 'GK', 'KH'], 'C9': ['CV', 'GS', 'KP'], 'B10': ['CO', 'GM', 'KJ'],
                         'B11': ['CP', 'GN', 'KK'], 'B12': ['CS', 'GO', 'KL'],
                         'B13': ['CT', 'GP', 'KM'], 'B14': ['CR', 'GQ', 'KN'], 'B15': ['CS', 'GR', 'KO'],
                         'B17': ['CW', 'GT', 'KQ'], 'C17': ['CX', 'GU', 'KR'],
                         'B20': ['CY', 'GV', 'KS']
                         }


moveCells_Programming2 = {'G2': ['BX', 'FU', 'JR'], 'G3': ['BY', 'FV', 'JS'], 'H3': ['BZ', 'FW', 'JT'],
                          'I3': ['CA', 'FX', 'JU'], 'J3': ['CB', 'FZ', 'JW'], 'K3': ['CC', 'GA', 'JX'],
                          'H5': ['CF', 'GC', 'JZ'], 'I5': ['CG', 'GD', 'KA'], 'H6': ['CH', 'GE', 'KB'],
                          'I6': ['CI', 'GF', 'KC'], 'H7': ['CJ', 'GG', 'KD'], 'I7': ['CK', 'GH', 'KE'],
                          'H8': ['CK', 'GI', 'KF'], 'I8': ['CL', 'GJ', 'KG'], 'H9': ['CM', 'GK', 'KH'],
                          'I9': ['CV', 'GS', 'KP'], 'H10': ['CN', 'GM', 'KJ'], 'H11': ['CO', 'GN', 'KK'],
                          'H12': ['CP', 'GO', 'KL'], 'H13': ['CQ', 'GP', 'KM'], 'H14': ['CR', 'GP', 'KN'],
                          'H15': ['CS', 'GR', 'KO'], 'H17': ['CW', 'GT', 'KQ'], 'I17': ['CX', 'GU', 'KR'],
                          'H20': ['CY', 'GV', 'KS']
                          }

# Series Programming Column

moveCells_seriesProgramming = {'H21': ['DA', 'GX', 'KU'],  # Name
                               'H22': ['DB', 'GY', 'KV'],  # Description
                               'L21': ['DC', 'GZ', 'KW'],  # Installments
                               'L22': ['DD', 'HA', 'KX'],  # Date
                               'K22': ['DE', 'HB', 'KY'],  # Attendance
                               'J21': ['DF', 'HD', 'LA'],  # Location
                               'K21': ['DG', 'HE', 'LB'],  # admissions fee

                               'H24': ['DJ', 'HG', 'LD'],  # RRE
                               'I24': ['DK', 'HH', 'LE'],
                               'H25': ['DL', 'HI', 'LF'],  # adv
                               'I25': ['DM', 'HJ', 'LG'],
                               'H26': ['DN', 'HK', 'LH'],  # food
                               'I26': ['DO', 'HL', 'LI'],
                               'H27': ['DP', 'HM', 'LJ'],  # supply/deco
                               'I27': ['DQ', 'HN', 'LK'],
                               'H28': ['DR', 'HO', 'LL'],
                               'G29': [['DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ'], ['HM', 'HN', 'HO', 'HP', 'HQ', 'HR'],
                                       ['LI', 'LJ', 'LK', 'LL', 'LM', 'LN']],
                               # The concatenated columns, shows all the types of contracts requested
                               'I29': ['DZ', 'HW', 'LT'],  # Description of Contract
                               'H32': ['EA', 'HU', 'LU'],  # Other
                               'I32': ['EB', 'HV', 'LV'],
                               }

# Trip Competition/Conference

moveCells_tripsCC1 = {'B53': ['EE', 'IB', 'LY'],  # NAME
                      'C54': ['EF', 'IC', 'LZ'],  # SERIES
                      'B55': ['EG', 'ID', 'MA'],  # LOCATION
                      'C55': ['EH', 'IE', 'MB'],  # DESCRIPTION
                      'D55': ['EI', 'IF', 'MC'],  # ATTENDANCE
                      'F55': ['EJ', 'IG', 'MD'],  # DATE

                      'B57': ['EL', 'II', 'MF'],  # TRANSPORT
                      'C57': ['EM', 'IJ', 'MG'],
                      'B58': ['EN', 'IK', 'MH'],  # PARKING
                      'C58': ['EO', 'IL', 'MI'],
                      'B59': ['EP', 'IM', 'MJ'],  # FOOD
                      'C59': ['EQ', 'IN', 'MK'],
                      'B60': ['ER', 'IO', 'ML'],  # LODGING
                      'C60': ['ES', 'IP', 'MM'],
                      'B61': ['ET', 'IQ', 'MN'],  # REGISTRATION
                      'C61': ['EU', 'IR', 'MO'],
                      'B62': ['EV', 'IS', 'MP'],  # OTHER
                      'C62': ['EW', 'IT', 'MQ'],
                      }


moveCells_TripsCC2 = {'H47': ['EE', 'IB', 'LY'],  # NAME
                      'I48': ['EF', 'IC', 'LZ'],  # SERIES
                      'H49': ['EG', 'ID', 'MA'],  # LOCATION
                      'I49': ['EH', 'IE', 'MB'],  # DESCRIPTION
                      'J49': ['EI', 'IF', 'MC'],  # ATTENDANCE
                      'L49': ['EJ', 'IG', 'MD'],  # DATE

                      'H51': ['EL', 'II', 'MF'],  # TRANSPORT
                      'I51': ['EM', 'IJ', 'MG'],
                      'H52': ['EN', 'IK', 'MH'],  # PARKING
                      'I52': ['EO', 'IL', 'MI'],
                      'H53': ['EP', 'IM', 'MJ'],  # FOOD
                      'I53': ['EQ', 'IN', 'MG'],
                      'H54': ['ER', 'IO', 'MK'],  # LODGING
                      'I54': ['ES', 'IP', 'ML'],
                      'H55': ['ET', 'IQ', 'MM'],  # REGISTRATION
                      'I55': ['EU', 'IR', 'MN'],
                      'H56': ['EV', 'IS', 'MP'],  # OTHER
                      'I56': ['EW', 'IT', 'MQ'],
                      }

# Other Trip
moveCells_otherTrip = {'G36': ['EZ', 'IW', 'MT'],  # Name
                       'H36': ['FA', 'IX', 'MU'],  # Series or nah
                       'K36': ['FB', 'IY', 'MV'],  # Location
                       'I37': ['FC', 'IZ', 'MW'],  # desc
                       'H37': ['FD', 'JA', 'MX'],  # atten
                       'J37': ['FE', 'JB', 'MY'],  # date

                       'H39': ['FI', 'JD', 'NA'],  # adv
                       'I39': ['FJ', 'JE', 'NB'],
                       'H40': ['FG', 'JF', 'NC'],  # transport
                       'I40': ['FH', 'JG', 'ND'],
                       'H41': ['FK', 'JH', 'NE'],  # admin
                       'I41': ['FL', 'JI', 'NF'],
                       'H42': ['FM', 'JJ', 'NG'],  # food
                       'I42': ['FN', 'JK', 'NH'],
                       'H43': ['FO', 'JL', 'NI'],  # lodging
                       'I43': ['FP', 'JM', 'NJ'],
                       'H44': ['FQ', 'JN', 'NK'],  # other
                       'I44': ['FR', 'JO', 'NL'],
                       }