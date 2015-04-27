import traceback
import os
import sys
import collections
import xlsxwriter
from m.common import CompilationError

def time2excel(timestamp):
    import datetime
    dt = datetime.datetime.utcfromtimestamp(timestamp)
    dt1900 = datetime.datetime(1900,1,1)
    delta = dt - dt1900

    # Excel handles time in number of days since 1, January 1900
    # Where 1, Jan 1900 itself is 1
    # Additional +1 is fix to bug in Excel that treats 1900 as a leap year
    value = delta.days + 2 + delta.seconds/24./3600.
    return value

def identity(val):
    return val

def cellName(row, col):
    colName = chr(ord('A')+col)
    return "%s%d" % (colName, row+1)

def absCellName(row, col):
    colName = chr(ord('A')+col)
    return "$%s$%d" % (colName, row+1)

kKB = 1024
kMB = 1024*kKB
kGB = 1024*kMB
kDay = 24*3600

class oExcel:
    ourWorkbooks = {}
    
    class ChartInfo(object):
        def __init__(self):
            self.topRow = 0
            self.leftCol = 0 
            self.chartType = None
            self.chartTitle = None
            self.chartWidth = 12
            self.chartHeight = 15
            self.chartAlign = "top"
            self.chartStyle = 16
            self.chartX = None
            self.chartY = []
        def __str__(self):
            return "width=%d height=%d" % (self.chartWidth, self.chartHeight)

    formatMapping = { ",": ("#,##0", identity),
                      ".": ("#,##0.00", identity),
                      "e": ("0.00E+00", identity),
                      "%": ("0.0%", identity),
                      "K": ('#,##0.0,"K"', identity),
                      "M": ('#,##0.00,,"M"', identity),
                      "G": ('#,##0.00,,,"G"', identity),
                      "KB": ('#,##0"KB"', lambda x: float(x)/kKB),
                      "MB": ('#,##0.0"MB"', lambda x: float(x)/kMB),
                      "GB": ('#,##0.00"GB"', lambda x: float(x)/kGB),
                      "n": ('[<200000]#,##0;[<500000000]0.00,,"M";0.00,,,"G"', identity),
                      "N": ('[<500000000]0.00,,"M";[<500000000000]0.00,,,"G";0.00,,,,"T"', identity),
                      "T": ('yyyy-mm-dd hh:mm:ss', time2excel),
                      "Tm": ('yyyy-mm-dd hh:mm:ss.000', time2excel),
                      "t": ('[m]:ss', lambda x: float(x)/kDay),
                      "mt": ('[m]:ss.000', lambda x: float(x)/1000/kDay),
                    }
    def __init__(self, fileName, variableNames, **moreParams):
        self.myFileName = fileName
        self.mySheetName = moreParams.get("sheetName", None)
        self.myContinue = moreParams.get("continue", False)
        self.myVars = variableNames
        self.lastDataColumn = len(self.myVars) - 1
        self.myNumColumns = len(variableNames)
        self.chartInfoDict = collections.defaultdict(oExcel.ChartInfo)
        self.formattedColumns = {}
        self.columnTitles = {}
        self.conversionFunctions = [identity] * len(self.myVars)
        self.columnFormats = [None] * len(self.myVars)
        for (param, value) in moreParams.iteritems():
            if param.endswith('_format'):
                value = oExcel.formatMapping.get(value, value)
                self.formattedColumns[param[:-7]] = value[0]
            elif param.endswith('_title'):
                self.columnTitles[param[:-6]] = value
            elif param.startswith("chartType"):
                self.chartInfoDict[param[9:]].chartType = value
            elif param.startswith("chartTitle"):
                self.chartInfoDict[param[10:]].chartTitle = value
            elif param.startswith("chartWidth"):
                self.chartInfoDict[param[10:]].chartWidth = value
            elif param.startswith("chartHeight"):
                self.chartInfoDict[param[11:]].chartHeight = value
            elif param.startswith("chartAlign"):
                self.chartInfoDict[param[10:]].chartAlign = value
            elif param.startswith("chartStyle"):
                self.chartInfoDict[param[10:]].chartStyle = value
            elif param.startswith("chartX"):
                try:
                    xColumn = int(value)
                except:
                    try:
                        xColumn = self.myVars.index(value)
                    except ValueError:
                        raise CompilationError("chartX variable %s is unknown" % value)
                #print "firstColumn", firstColumn
                self.chartInfoDict[param[11:]].chartX = xColumn
            elif param.startswith("chartY"):
                varNames = value.split(",")
                for var in varNames:
                    try:
                        yColumn = int(var)
                    except:
                        try:
                            yColumn = self.myVars.index(var)
                        except ValueError:
                            raise CompilationError("chartY variable %s is unknown" % var)
                    self.chartInfoDict[param[10:]].chartY.append(yColumn)
            elif param in self.myVars:
                # consider this as format
                value = oExcel.formatMapping.get(value, value)
                self.formattedColumns[param] = value
        # calculate top tows for all charts that appear at the top
        topRow = 0
        for chartId, chartInfo in sorted(self.chartInfoDict.iteritems()):
            if not chartInfo.chartType:
                raise CompilationError( "No chart type appears for %s chart" % (chartId if chartId else "default"))
            #print "Chart Info", chartInfo
            if chartInfo.chartAlign == "top":
                chartInfo.topRow = topRow
                topRow += chartInfo.chartHeight + 1
        
        if topRow:
            self.myTitleRow = topRow + 1
        else:
            self.myTitleRow = 0
        self.initWorkbook()
        for param, value in self.formattedColumns.iteritems():
            self.setConversion(param, value)
        titleFormat = self.myWorkbook.add_format({'bold': True, 'bottom': 2})
        for i in range(len(variableNames)):
            title = self.columnTitles.get(variableNames[i], variableNames[i])
            self.myDataSheet.write(self.myTitleRow, i, title, titleFormat)
        self.myNextRow = self.myTitleRow + 1
        
            
    def save(self, record):
        for i in range(len(record)):
            val = self.conversionFunctions[i](record[i])
            numFormat = self.columnFormats[i]
            if numFormat:
                self.myDataSheet.write(self.myNextRow, i, val, numFormat)
            else:
                self.myDataSheet.write(self.myNextRow, i, val)
            
        self.myNextRow += 1
    def close(self):
        topRowForBottomCharts = self.myNextRow+1
        topRowForLeftCharts = self.myTitleRow+1
        for chartId,chartInfo in sorted(self.chartInfoDict.iteritems()):
            if not chartInfo.chartType:
                continue
            if chartInfo.chartAlign == "bottom":
                chartInfo.topRow = topRowForBottomCharts
                topRowForBottomCharts += chartInfo.chartHeight+1
            elif chartInfo.chartAlign == "left":
                chartInfo.topRow = topRowForLeftCharts
                topRowForLeftCharts += chartInfo.chartHeight+1
                chartInfo.leftCol = self.myNumColumns + 1
            self.createChart(chartInfo)
        if not self.myContinue:
            self.myWorkbook.close()
    def createChart(self, chartInfo):
        left = 0.5
        right = left + chartInfo.chartWidth
        top = chartInfo.topRow
        bottom = top + chartInfo.chartHeight
        if chartInfo.chartType in ["column", "bar", "stackedColumn", "relativeColumn", "stackedBar", "relativeBar"]:
            chart = self.createColumnChart(chartInfo)
        elif chartInfo.chartType in ["pie", "doughnut"]:
            chart = self.createPieChart(chartInfo)
        elif chartInfo.chartType in ["line", "area", "stackedArea", "relativeArea"]:
            chart = self.createLineChart(chartInfo)
        if chartInfo.chartTitle:
            chart.set_title({'name': chartInfo.chartTitle})
        chart.set_style(chartInfo.chartStyle)
        self.myDataSheet.insert_chart(cellName(chartInfo.topRow, chartInfo.leftCol), chart)
        
    def getNumDataRows(self):
        return self.myNextRow-self.myTitleRow-1
    def getSeriesFormula(self, col):
        formula = "'%s'!%s:%s" % (self.mySheetName, \
                                absCellName(self.myTitleRow+1, col), absCellName(self.myNextRow-1, col))
        #print "categoryFormula = " + formula
        return formula
    def getTitleFormula(self, col):
        formula = "='%s'!%s" % (self.mySheetName, absCellName(self.myTitleRow, col))
        #print "categoryFormula = " + formula
        return formula
    def createDefaultChart(self, chart, chartInfo, seriesParams={}):
        for colNum, col in enumerate(chartInfo.chartY):
            series = dict(seriesParams.items())
            series["values"] = self.getSeriesFormula(col)
            series["name"] = self.getTitleFormula(col)
            if colNum == 0 and chartInfo.chartX is not None:
                series["categories"] = self.getSeriesFormula(chartInfo.chartX)
                chart.set_x_axis({
                                  'name':  self.getTitleFormula(chartInfo.chartX),
                                  'name_font':  {'bold': True },
                                  })
            chart.add_series(series)
    def createColumnChart(self, chartInfo):
        if chartInfo.chartType.startswith("relative"):
            type = chartInfo.chartType[8:].lower()
            subtype = "percent_stacked"
        elif chartInfo.chartType.startswith("stacked"):
            type = chartInfo.chartType[7:].lower()
            subtype = "stacked"
        else:
            type = chartInfo.chartType
            subtype = None
        chartAttr = {"type": type}
        if subtype:
            chartAttr["subtype"] = subtype
        chart = self.myWorkbook.add_chart(chartAttr)
        self.createDefaultChart(chart, chartInfo)
        return chart
    def createPieChart(self, chartInfo):
        chart = self.myWorkbook.add_chart({"type": chartInfo.chartType})
        seriesParams={'data_labels': {'percentage': True}}
        if chartInfo.chartType == "doughnut":
            seriesParams['data_labels']['value'] = True
        self.createDefaultChart(chart, chartInfo, seriesParams=seriesParams)
        return chart

    def createLineChart(self, chartInfo):
        if chartInfo.chartType == "stackedArea":
            chartAttr = {"type": "area", "subtype": "stacked"}
        elif chartInfo.chartType == "relativeArea":
            chartAttr = {"type": "area", "subtype": "percent_stacked"}
        else:
            chartAttr = {"type": chartInfo.chartType}
        chart = self.myWorkbook.add_chart(chartAttr)
        self.createDefaultChart(chart, chartInfo)
        return chart

    def createScatterChart(self, chart, chartInfo):
        chart = self.myWorkbook.add_chart({"type": chartInfo.chartType})
        self.createDefaultChart(chart, chartInfo)
        return chart
           
    def setConversion(self, columnName, format):
        try:
            col = self.myVars.index(columnName)
        except:
            return
        if isinstance(format, tuple):
            self.conversionFunctions[col] = format[1]
            self.columnFormats[col] = self.myWorkbook.add_format({"num_format": format[0]})
    def initWorkbook(self):
        self.myWorkbook = oExcel.ourWorkbooks.get(self.myFileName)
        if not self.myWorkbook: 
            self.myWorkbook = xlsxwriter.Workbook(self.myFileName, {'constant_memory': True})
            if self.myContinue:
                oExcel.ourWorkbooks[self.myFileName] = self.myWorkbook
        else:
            if not self.myContinue:
                del oExcel.ourWorkbooks[self.myFileName]
            
        self.mySheetId = 1
        if self.mySheetName:
            self.myDataSheet = self.myWorkbook.add_worksheet(self.mySheetName)
        else:
            self.mySheetName = "Sheet1"
            self.myDataSheet = self.myWorkbook.add_worksheet()
    @staticmethod
    def closeWorkbook(fileName):
        workbook = oExcel.ourWorkbooks.get(fileName)
        if workbook:
            workbook.close()
            del oExcel.ourWorkbooks[fileName]

def closeExcel(fileName):
    """closeExcel(fileName)
    closes open Excel workbook
"""
    oExcel.closeWorkbook(fileName)

def _outputExcelFromJson(jsonFileName):
    import json
    f = open(jsonFileName, "rb")
    inputData = json.load(f)
    f.close()
    excel = oExcel(inputData['fileName'], inputData['variableNames'], **inputData['moreParams'])
    for record in inputData['data']:
        excel.save(tuple(record))
    excel.close()
    return True

if __name__=="__main__":
    _outputExcelFromJson(sys.argv[1])
    