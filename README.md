from gooey import Gooey, GooeyParser
import pandas as pd
import os
import xlwings as xw

@Gooey(program_name="Combine BCS log for validation", program_description="@author:kdn0mqp,Alan\nUsing DF", default_size=(800, 600))
def parse_args():
    parser = GooeyParser()
    parser.add_argument("fileFolderPath", metavar="File Folder Path",
                        help=" you can drag folder to below \n or you can choose folder by clicking> Browser",
                        widget="DirChooser")
    parser.add_argument("validationofBCS", metavar="BCS Validation File",
                        help="please choose your BCS Validation File\n", widget="FileChooser")
    args = parser.parse_args()
    return args

args = parse_args()

fileFolderPath = args.fileFolderPath
trackerList = os.listdir(fileFolderPath)
validationofBCS = args.validationofBCS
rosterDF=pd.DataFrame()
callDF=pd.DataFrame()
ahtDF=pd.DataFrame()
otherDF=pd.DataFrame()

for ph in trackerList:
    path=os.path.join(fileFolderPath,ph)
    print(path)

    detailRoster=pd.read_excel(path,sheet_name="Roster",header=2)
    detailRoster=detailRoster.loc[detailRoster["Name"].notnull(),:]
    detailRoster = detailRoster.iloc[:, 1:8]
    rosterDF=rosterDF.append(detailRoster)
    rosterDF=rosterDF.fillna("")
    rosterDF=rosterDF.astype("str")
    print(rosterDF)

    detailCall=pd.read_excel(path,sheet_name="Call",header=3)
    detailCall=detailCall.loc[detailCall["Date"].notnull(),:]
    detailCall=detailCall.iloc[:,12:29]
    callDF=callDF.append(detailCall)
    callDF=callDF.fillna("")
    callDF=callDF.astype("str")
    print(callDF)

    detailAHT=pd.read_excel(path,sheet_name="AHT",header=3)
    detailAHT=detailAHT.loc[detailAHT["Date"].notnull(),:]
    detailAHT=detailAHT.iloc[:,3:14]
    ahtDF=ahtDF.append(detailAHT)
    ahtDF=ahtDF.fillna("")
    ahtDF=ahtDF.astype("str")
    print(ahtDF)

    detailOther=pd.read_excel(path,sheet_name="Other",header=3)
    detailOther=detailOther.loc[detailOther["Date"].notnull(),:]
    detailOther=detailOther.iloc[:,8:14]
    otherDF=otherDF.append(detailOther)
    otherDF=otherDF.fillna("")
    otherDF=otherDF.astype("str")
    print(otherDF)

app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
wb1=app.books.open(validationofBCS)
wb1.sheets["Roster"].range("b5:h500").clear_contents()
wb1.sheets["Call"].range("m6:ac5000").clear_contents()
wb1.sheets["AHT"].range("d6:n500").clear_contents()
wb1.sheets["Other"].range("i6:n500").clear_contents()

wb1.sheets["Roster"].range("b5").value=rosterDF.values
wb1.sheets["Call"].range("m6").value=callDF.values
wb1.sheets["AHT"].range("d6").value=ahtDF.values
wb1.sheets["Other"].range("i6").value=otherDF.values
wb1.save()


















