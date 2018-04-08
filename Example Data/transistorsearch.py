print('Compiling Data Summaries from "parameters.csv" files...')
import os
import pandas as pd
import statistics

r=os.getcwd()
i=[]
for path, sundirs, files in os.walk(r):
    for name in files:
        i.append(os.path.join(path, name))

paramlist=[]
index=0
while index < len(i):
    paramsearch=i[index][-14:]
    if paramsearch == 'parameters.csv':
        paramlist.append(i[index])
    index = index + 1
    continue

ntotalparam=[]
ptotalparam=[]
j=0
while j < len(paramlist):
    parampath=paramlist[j]
    params=pd.read_csv(parampath, delimiter=",", error_bad_lines=False)
    mobilityn=[]
    vthn=[]
    ionoffn=[]
    mobilityp=[]
    vthp=[]
    ionoffp=[]
    pcount=0
    ncount=0
    k=0
    while k < len(params):
        polarity=params.values[k][0][0]
        mobility=params.values[k][1]
        vth=params.values[k][2]
        ionoff=params.values[k][4]
        if polarity=='N':
            mobilityn.append(mobility)
            vthn.append(vth)
            ionoffn.append(ionoff)
            ncount=1
        else:
            mobilityp.append(mobility)
            vthp.append(vth)
            ionoffp.append(ionoff)
            pcount=1
        k=k+1
        if k == len(params):
                if ncount==1:
                    nmobavg=sum(mobilityn)/len(mobilityn)
                    nmobstdev=statistics.stdev(mobilityn)
                    nmobmax=max(mobilityn)
                    bestindex=mobilityn.index(nmobmax)
                    nvthavg=sum(vthn)/len(vthn)
                    nvthstdev=statistics.stdev(vthn)
                    nvthmax=vthn[bestindex]
                    nionavg=sum(ionoffn)/len(ionoffn)
                    nionmax=ionoffn[bestindex]
                if pcount==1:
                    pmobavg=sum(mobilityp)/len(mobilityp)
                    pmobstdev=statistics.stdev(mobilityp)
                    pmobmax=max(mobilityp)
                    bestindex=mobilityp.index(pmobmax)
                    pvthavg=sum(vthp)/len(vthp)
                    pvthstdev=statistics.stdev(vthp)
                    pvthmax=vthp[bestindex]
                    pionavg=sum(ionoffp)/len(ionoffp)
                    pionmax=ionoffp[bestindex]
        continue
    comp=parampath.split("\\",7)[5]
    temp=parampath.split("\\",7)[6]
    if pcount==1:
        pparamsavg=(comp,temp,pmobmax,pmobavg,pmobstdev,pionmax,pionavg,pvthmax,pvthavg,pvthstdev)
        ptotalparam.append(pparamsavg)
    if ncount==1:
        nparamsavg=(comp,temp,nmobmax,nmobavg,nmobstdev,nionmax,nionavg,nvthmax,nvthavg,nvthstdev)
        ntotalparam.append(nparamsavg)
    j=j+1
    continue

#-----------------------------------------------N-TYPE EXCEL EXPORT-----------------------------------------------------
if ntotalparam != []:
    # Create a Pandas dataframe title for data.
    df0 = pd.DataFrame({'Data': ['Sample','Annealing','Best Mobility', 'Avg Mobility','Best Ion/Ioff (from Best Mob Sample)','Avg Ion/Ioff','Best Vth (from Best Mob Sample)','Avg Vth']}).T

    writer = pd.ExcelWriter('N-Type Data Summary.xlsx', engine='xlsxwriter')

    #Loop through data from total export list
    j=0
    index=1
    while j < len(ntotalparam):
        compound=ntotalparam[j][0]
        temp=ntotalparam[j][1]
        nmobmax=ntotalparam[j][2]
        nmobavg=ntotalparam[j][3]
        nmobavg=round(nmobavg,7)
        nmobstdev=ntotalparam[j][4]
        nmobstdev=round(nmobstdev,7)
        nmobtotavg=str(nmobavg) + chr(177) + " " + str(nmobstdev)
        nionmax=ntotalparam[j][5]
        nionavg=ntotalparam[j][6]
        nvthmax=ntotalparam[j][7]
        nvthavg=ntotalparam[j][8]
        nvthavg=round(nvthavg,7)
        nvthstdev=ntotalparam[j][9]
        nvthstdev=round(nvthstdev,7)
        nvthtotavg=str(nvthavg) + chr(177) + " " + str(nvthstdev)
        ndata=(compound,temp,nmobmax,nmobtotavg,nionmax,nionavg,nvthmax,nvthtotavg)
        df = pd.DataFrame({'invisible header': ndata}).T
        df.to_excel(writer, sheet_name='Sheet1', startrow=index, index=False,header =False)
        j=j+1
        index=index+1
        continue

    # Convert the dataframe to an XlsxWriter Excel object.
    df0.to_excel(writer, sheet_name='Sheet1', header = False, index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # Sort the columns to make it easier to paste into word tables
    xl = pd.ExcelFile("N-Type Data Summary.xlsx")
    df = xl.parse("Sheet1")

    df['anneal_cat'] = pd.Categorical(
        df['Annealing'], 
        categories=['RT','rt','50C','50c','100C','100c','150C','150c','200C','200c'], 
        ordered=True
    )
    df=df.sort_values(['Sample','anneal_cat'],ascending =True)

    del df['anneal_cat']
     
    writer = pd.ExcelWriter('N-Type Data Summary.xlsx')
    df.to_excel(writer,sheet_name='Sheet1',index=False)
    writer.save()
    print('N-Type Data Summary exported ---->')

#-----------------------------------------------P-TYPE EXCEL EXPORT-----------------------------------------------------
if ptotalparam != []:
    # Create a Pandas dataframe title for data.
    df0 = pd.DataFrame({'Data': ['Sample','Annealing','Best Mobility', 'Avg Mobility','Best Ion/Ioff (from Best Mob Sample)','Avg Ion/Ioff','Best Vth (from Best Mob Sample)','Avg Vth']}).T

    writer = pd.ExcelWriter('P-Type Data Summary.xlsx', engine='xlsxwriter')

    #Loop through data from total export list
    j=0
    index=1
    while j < len(ptotalparam):
        compound=ptotalparam[j][0]
        temp=ptotalparam[j][1]
        pmobmax=ptotalparam[j][2]
        pmobavg=ptotalparam[j][3]
        pmobavg=round(pmobavg,7)
        pmobstdev=ptotalparam[j][4]
        pmobstdev=round(pmobstdev,7)
        pmobtotavg=str(pmobavg) + chr(177) + " " +  str(pmobstdev)
        pionmax=ptotalparam[j][5]
        pionavg=ptotalparam[j][6]
        pvthmax=ptotalparam[j][7]
        pvthavg=ptotalparam[j][8]
        pvthavg=round(pvthavg,7)
        pvthstdev=ptotalparam[j][9]
        pvthstdev=round(pvthstdev,7)
        pvthtotavg=str(pvthavg) + chr(177) + " " +  str(pvthstdev)
        pdata=(compound,temp,pmobmax,pmobtotavg,pionmax,pionavg,pvthmax,pvthtotavg)
        df = pd.DataFrame({'invisible header': pdata}).T
        df.to_excel(writer, sheet_name='Sheet1', startrow=index, index=False,header =False)
        j=j+1
        index=index+1
        continue

    # Convert the dataframe to an XlsxWriter Excel object.
    df0.to_excel(writer, sheet_name='Sheet1', header = False, index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # Sort the columns to make it easier to paste into word tables
    xl = pd.ExcelFile("P-Type Data Summary.xlsx")
    df = xl.parse("Sheet1")

    df['anneal_cat'] = pd.Categorical(
        df['Annealing'], 
        categories=['RT','rt','50C','50c','100C','100c','150C','150c','200C','200c'], 
        ordered=True
    )
    df=df.sort_values(['Sample','anneal_cat'],ascending =True)

    del df['anneal_cat']
     
    writer = pd.ExcelWriter('P-Type Data Summary.xlsx')
    df.to_excel(writer,sheet_name='Sheet1',index=False)
    writer.save()
    print('P-Type Data Summary exported ---->')
