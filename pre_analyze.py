import pandas as pd
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pptx import Presentation
from pptx.util import Inches
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

    def extrage_gafice(self, ExcelFile, FilterParams, TargetColumns, UseTargetColumnsInLegend):
        # The function returns a list of pandas series, where each series will be ploted on the same plot
        # ExcelFile is the imported DataFrame (already indexed by time)
        # FilterParams is a dictionary with keys being column names and values being a tuple, where the first item is
        #   the operator, the second being a list of values that each column should be filtered upon,
        # TargetColumns is a list of column names that we want in the graph.
        # when we plot it will always plot based on the ExcelFie indexes on x axis
        #
        # FilterParams = [['Slot', '==',[0,1,4]], ['altceva', '==',['d','d']], ['al3', '<=',[7]]]


        FilterParamNumber = FilterParams.__len__()
        AllColValuesLists = [i[2] for i in FilterParams]
        AllColValuesListsCombinations = list(itertools.product(*AllColValuesLists))

        while True:
            try:
                comb = AllColValuesListsCombinations.pop(0)
            except:
                print('Finished graphs data generation')
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
            # for FilterParam in comb:
            #    PrefixNumeGrafic += str(FilterParam) + "_"
            # we consider the first filter parameter relevnat for the name of the grapsh as it will be shown in the legend. For example it will represent the Slot number
            PrefixNumeGrafic = str(comb[0])

            for TargetColumn in TargetColumns:
                if not filterResult[TargetColumn].empty:
                    if UseTargetColumnsInLegend:
                        self.grafice[PrefixNumeGrafic + "_" + TargetColumn] = (filterResult[TargetColumn])
                    else:
                        # Filtering parameters might lead to a result where the PrefixNumeGrafic is not unique.
                        # If so, exit and print Error
                        if PrefixNumeGrafic in self.grafice.keys():
                            print("ERROR: graph names are not unique. Please set UseTargetColumnsInLegend to True in the config file")
                            exit()
                        else:
                            self.grafice[PrefixNumeGrafic] = (filterResult[TargetColumn])

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

# prepare presentation helpers
prs = Presentation()
outputPath = './output/'
tmpPicturePath = './output/tmp/'
pptName = 'Report_' + mdates.datetime.date.today().strftime('%d-%m-%Y') + '.pptx'
slidesOrder = {}

for fileName in fileNames:
    print(fileName)
    configFileName = os.path.splitext(fileName)[0] + ".yaml"

    # check if the config file is empty. If, so skip this excel file
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
    UseTargetColumnsInLegend = configDict['UseTargetColumnsInLegend']

    # reverse the oreder from oldest to newest
    excelfile = excelfile[::-1]
    figura = plt.figure();
    ax = figura.add_subplot(111)
    colectieGraphs = graficCollection(figura)
    colectieGraphs.extrage_gafice(excelfile, FilterParams, TargetColumns, UseTargetColumnsInLegend)
    for label, serie in colectieGraphs.grafice.items():
        x = serie.index
        y = serie.values
        ax.plot(x, y, label=label)

    # adjustments to the figure
    # ---------

    colectieGraphs.figura.axes[0].set_xlim(excelfile.index[0], excelfile.index[-1])
    dateTimeFmt = mdates.DateFormatter('%D %H:%M')
    colectieGraphs.figura.axes[0].xaxis.set_major_locator(plt.MaxNLocator(45))
    colectieGraphs.figura.axes[0].xaxis.set_major_formatter(dateTimeFmt)
    colectieGraphs.figura.axes[0].xaxis.set_tick_params(rotation = 90)
    colectieGraphs.figura.legend(loc=9, mode='none', fontsize='small', ncol=5, labelspacing=0.05)
    colectieGraphs.figura.set_size_inches(9.99, 6.7)
    colectieGraphs.figura.tight_layout()
    colectieGraphs.figura.subplots_adjust(top = 0.900)

    # ----------



    tmpPictureName = configDict['Title'] + '.png'
    plt.savefig(tmpPicturePath + tmpPictureName)


    # Create a dictionary where the key is the slide number and the value is the picture name
    # But first check if the order configured is correct. If not, terminate the script.
    # note that the tmptPictureName will also be used for the Title of the slide
    if configDict['slideNumber'] in slidesOrder.keys():
        print("the slide number already exists. Please check the configs. Current config file: " + configFileName)
        exit()
    slidesOrder[configDict['slideNumber']] = tmpPictureName

# We sort the slidesOrder so the outputed ppt would be in the requred
sortedKeysAsInt = sorted([int(k) for k in slidesOrder.keys()])
for slideNumber in sortedKeysAsInt:
    # create a slide and add it to the ppt
    title_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_slide_layout)
    titlu = slide.placeholders[0]
    # As the value stored for the key is the tmpPictureName, we use it for the title by removing the extension
    titlu.text = os.path.splitext(slidesOrder[str(slideNumber)])[0]
    titlu.top = 0
    titlu.left = 0
    titlu.height = Inches(0.5)
    titlu.width = Inches(10)
    titlu.text_frame.paragraphs[0].font.size = Pt(14)
    titlu.text_frame.paragraphs[0].font.bold = True
    pic = slide.shapes.add_picture(tmpPicturePath + slidesOrder[str(slideNumber)], Inches(0.01), Inches(0.5))



prs.save(outputPath + pptName)



#plt.close()