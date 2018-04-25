from openpyxl import Workbook
from openpyxl import load_workbook
import requests
import pprint
import os

r = requests.get('http://www.gencon.com/downloads/events.xlsx')
with open('/Users/jvais/Downloads/events-temp.xlsx', 'wb') as f:
    f.write(r.content)
wb = load_workbook('/Users/jvais/Downloads/events-temp.xlsx')

track = {"oneshot":{'total':0,
                    'titles':[],
                    'query':{'column':'B','value':'ONE SHOT Podcast Network'}
                    },
        "trinity":{'total':0,
                    'titles':[],
                    'query':{'column':'G','value':'Trinity'}
                    },
        "genesys":{'total':0,
                    'titles':[],
                    'query':{'column':'G','value':'Genesys'}
                    },
        "genesys":{'total':0,
                    'titles':[],
                    'query':{'column':'G','value':'Genesys'}
        }
}

ws1 = wb.active
for row in ws1:
    for cell in row:
        for e in track.keys():
            if cell.column == track[e]['query']['column']:
                if cell.value == track[e]['query']['value']:
                    track[e]['total'] = track[e]['total'] + 1
                    track[e]['titles'].append("%s - %s" % (row[0].value, row[2].value))

pp = pprint.PrettyPrinter(indent=4)
pp.pprint(track)

os.remove('/Users/jvais/Downloads/events-temp.xlsx')