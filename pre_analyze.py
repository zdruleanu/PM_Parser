import pandas as pd
import os
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.dates as mdates
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Pt
import itertools
import yaml

class graficCollection:
    '''
    This class is used to contain the data from one single graph(figure)
    '''
    def __init__(self, figura):
        self.figura = figura
        self.grafice = {}

    def extrage_gafice(self, ExcelFile, FilterParams, TargetColumns):
        # The function returns a list of pandas series, where each series will be ploted on the same plot
        # ExcelFile is the imported DataFrame (already indexed by time)
        # FilterParams is a dictionary with keys being column names and values being a tuple, where the first item is
        #   the operator, the second being a list of values that each column should be filtered upon,
        # TargetColumns is a list of column names that we want in the graph.
        # when we plot it will always plot based on the ExcelFie indexes on x axis
        #
        # FilterParams = [['Slot', '==',[0,1,4]], ['altceva', '==',['d','d']], ['al3', '==',[7,938,54]]]


        FilterParamNumber = FilterParams.__len__()
        AllColValuesLists = [i[2] for i in FilterParams]
        AllColValuesListsCombinations = list(itertools.product(*AllColValuesLists))

        while True:
            try:
                comb = AllColValuesListsCombinations.pop(0)
            except:
                print("exited while")
                break
            FilterExpression = '('
            for i in range(0,FilterParamNumber):
                FilterExpression += "(ExcelFile['" + FilterParams[i][0] + "'] " + FilterParams[i][1] + " " + str(comb[i]) + ")"
                if i < FilterParamNumber-1:
                    FilterExpression += " & "
                else:
                    FilterExpression += ")"
                print("FiltereExpression is: ", FilterExpression)
            filterResult = ExcelFile[eval(FilterExpression)]
            PrefixNumeGrafic = ''
            for FilterParam in comb:
                PrefixNumeGrafic += str(FilterParam) + "_"
            for TargetColumn in TargetColumns:
                if filterResult[TargetColumn].empty == False:
                    self.grafice[PrefixNumeGrafic + TargetColumn] = (filterResult[TargetColumn])


def find_csv_filenames(path_to_dir, suffix=""):
    filenames = os.listdir(path_to_dir)
    return [path_to_dir+"/"+filename for filename in filenames if filename.endswith(suffix)]

def initializeConfigFiles(path_to_dir):
    filenames = find_csv_filenames("./data")
    for filename in filenames:
        if os.path.splitext(filename)[1] == '.yaml':
            continue

        configFileName = os.path.splitext(filename)[0]+".yaml"
        if configFileName in filenames:
            print(filename + " already has the corresponding config file: " + configFileName)
        else:
            configFile = open(configFileName, "w")
            configFile.close()
            print("Created " + configFileName + " config file for " + filename)
    print("Finished config files initialization")

initializeConfigFiles("./data")
fileNames = find_csv_filenames("./data", '.xls')

#prepare presentation helpers
prs = Presentation()
outputPath = './output/'
tmpPicturePath = './output/tmp/'
tmpPictureName = 'pic1.png'
pptName = 'Report_' + mdates.datetime.date.today().strftime('%d-%m-%Y') + '.pptx'

for fileName in fileNames:
    print(fileName)
    configFileName = os.path.splitext(fileName)[0] + ".yaml"

    #check if the config file is empty. If, so skip this excel file
    if os.path.getsize(configFileName) != 0:
        configFile = open(configFileName, "r")
        configDict = yaml.load(configFile)
    else:
        print(fileName + " has an empty config file")
        continue

    excelfile = pd.read_excel(fileName)
    # change Reported Time type from objet to datetime and set it as index
    excelfile[configDict['indexul']] = pd.to_datetime(excelfile[configDict['indexul']])
    excelfile.set_index(configDict['indexul'], inplace=True)
    FilterParams = configDict['FilterParams']
    TargetColumns = configDict['TargetColumns']

    #reverse the oreder from oldest to newest
    excelfile = excelfile[::-1]
    figura = plt.figure();
    ax = figura.add_subplot(111)
    colectieGraphs = graficCollection(figura)
    colectieGraphs.extrage_gafice(excelfile, FilterParams, TargetColumns)
    for label, serie in colectieGraphs.grafice.items():
        x = serie.index
        y = serie.values
        ax.plot(x, y, label=label)

    #adjustments to the figure
    #---------

    colectieGraphs.figura.axes[0].set_xlim(excelfile.index[0], excelfile.index[-1])
    dateTimeFmt = mdates.DateFormatter('%D %H:%M')
    colectieGraphs.figura.axes[0].xaxis.set_major_locator(plt.MaxNLocator(45))
    colectieGraphs.figura.axes[0].xaxis.set_major_formatter(dateTimeFmt)
    colectieGraphs.figura.axes[0].xaxis.set_tick_params(rotation = 90)
    colectieGraphs.figura.legend(loc=9, mode='expand', fontsize='small', ncol=5)
    colectieGraphs.figura.set_size_inches(9.99, 6.7)
    colectieGraphs.figura.tight_layout()
    colectieGraphs.figura.subplots_adjust(top = 0.900)

    #----------

    #create a slide and add it to the ppt
    title_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_slide_layout)


    plt.savefig(tmpPicturePath + tmpPictureName)
    titlu = slide.placeholders[0]
    titlu.text = configDict['Title']
    titlu.top = 0
    titlu.left = 0
    titlu.height = Inches(0.5)
    titlu.width = Inches(10)
    titlu.text_frame.paragraphs[0].font.size = Pt(14)
    titlu.text_frame.paragraphs[0].font.bold = True

    pic = slide.shapes.add_picture(tmpPicturePath + tmpPictureName, Inches(0.01), Inches(0.5))



prs.save(outputPath + pptName)



#plt.close()