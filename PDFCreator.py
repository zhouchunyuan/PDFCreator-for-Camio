from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer,Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.lib.units import inch,cm
from reportlab.lib.pagesizes import A4

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))

import sys,datetime,re
import Tkinter as tk
import tkMessageBox as mbox
import tkFileDialog
######################################
NIKON_TXT_FILE = 'C:\\nikon.txt'
MAIN_RES_FILE = 'D:\\LGE_Report\\Program\\main.res' # will be given as argv from <PDF.vbs>
#NIKON_logo = 'D:\\LGE_Report\\Program\\dependency\\Nikon_logo.png'
#GM_logo = 'D:\\LGE_Report\\Program\\dependency\\Shanghai_GM_logo.png'
NIKON_logo = 'D:\\LGE_Report\\Program\\Nikon.jpg'
GM_logo = 'D:\\LGE_Report\\Program\\SGM.jpg'
OUTPUT_PDF_DIR = 'D:\\LGE_Report'
OUTPUT_PDF_NAME =''

PAGE_WIDTH,PAGE_HEIGHT = A4
styles = getSampleStyleSheet()
#First we import some constructors, some paragraph styles and other conveniences from other modules.
pageinfo = ""
######################################

###################### function showMessage ######### 
# display message dialog
#####################################################
def showMessage(msg):
    window = tk.Tk()
    window.wm_withdraw()
    mbox.showinfo('Info',msg)
###################### function getResFileDir ######### 
# get res path and name
#######################################################
def getResFileDir(pathname):
    options = {}
    options['defaultextension'] = '.res'
    options['filetypes'] = [('res files', '.res')]
    options['initialfile'] = pathname
    options['title'] = 'Please select the RES file'
    root = tk.Tk()
    root.withdraw()
    file_path_string = tkFileDialog.askopenfilename(**options)
    return file_path_string
###################### function getResDir ######### 
# to guess (find) RES file dir
# according to plan_name in <Nikon.txt>
#       pattern = re.compile(r'(\w\w)_([\w ]+)')# to match name like 'LTG MY2016'
#       match = pattern.match(line_Plan_Name)
#       Body_OR_Head=match.group(1)
#       Type_name   =match.group(2)
#####################################################
def getResDir(Body_OR_Head,Type_name):
    resPathName = ''
    if Body_OR_Head.strip()=='CB':
        if Type_name.strip()=='LTG' or Type_name.strip()=='LKW':
            resPathName = 'D:\\LGE\\CB\\'+Type_name.strip()+'\\Report\\'
        elif Type_name.strip()=='LTG MY2016': 
            resPathName = 'D:\\LGE\\CB MY2016\\LTG\\Report\\'
        elif Type_name.strip()=='LD4 MY2016':
            resPathName = 'D:\\LGE\\CB MY2016\\LD4\\Report\\'
    elif Body_OR_Head.strip()=='CH':
        resPathName = 'D:\\LGE\\CH\\Report\\'
    resPathName = resPathName+'main.res'
    return resPathName
###################### function findline ############ 
# This function returns a list if it is a data line
# Return None, if not.
#####################################################
def findLine(line):
    gap = 10
    a=0
    b=20
    if len(line)>20 :
        col0 = line[a:b]
        a=b+1
        b=b+gap
        col2 = line[a:b].strip()
        a=b+1
        b=b+gap
        col3 = line[a:b].strip()
        a=b+1
        b=b+gap
        col4 = line[a:b].strip()
        a=b+1
        b=b+gap
        col5 = line[a:b].strip()
        a=b+1
        b=b+gap
        col6 = line[a:b].strip()
        a=b+1
        b=b+gap
        col7 = line[a:b].strip()
        # begins with letter ^[a-zA-Z],
        # followed by at least one character [a-zA-Z0-9\.\- ]+
        # at least one space [ ]+
        # numbers with or w/o minus sign [\-\d\.]+
        # end with number
        pattern = re.compile(r'(^[a-zA-Z][a-zA-Z0-9\.\- ]+)[ ]+([\-\d\.]+\d\Z) *')
        match = pattern.match(col0)

        if match:
            col0 = match.group(1).strip()
            col1 = match.group(2)
            return [col0,col1,col2,col3,col4,col5,col6,col7]
        else:
            return None
        
###################### function createTableData ############
#
# This function returns the table data
# regex is used to judge data types
#   1) '==...==' or '===...===' means a caption line,
#       but need to exclude '========'
#   2) 'xxx : yyy' means substitution of names
#       if the xxx is 'Circle', replace the next 3rd line name with 'yyy'
#       if other case, replace the next line name with 'yyy'
#   3) in other cases:
#       if it is a caption line,try to combine any caption line below it.
#
############################################################        
def createTableData() :
    data = [['Feature Name','Actual','Nominal','Lo-To1','Hi-To1','Devi','Graphic','Error']]
    f = open(MAIN_RES_FILE)
    line = f.readline()
    last_Was_DATUM_Head = False
    last_Was_Caption = False
    while line !='':
        # step1 find '== ... ==', this should be displayed
        pattern = re.compile(r'={2,3}[^=]+={2,3}')
        if pattern.match(line):#special process for == ==
            if line == '==Part Location and orientation relative to the MCS==\n':
                #skip 2 lines
                f.readline()#-------------------------
                f.readline()#Point:ORIGINAL_POINT_OF_ABC
                line = line+'       '+f.readline()
                line = line+'       '+f.readline()
                data.append([line])
                last_Was_DATUM_Head = False
            else:
                data.append([line])
                last_Was_DATUM_Head = True
        else :# step2 find xxx : yyy, this is for changing name
            #pattern = re.compile(r'(^\w+):(.+\w)(--(^\w+):(.+\w))?')
            pattern = re.compile(r'(^\w+):( |\w+)(--)?(\w+)?(:)?(\w+)?')# 'Circle:H001--Circle:H002
            match = pattern.match(line)
            if match:
                name0 = match.group(1).strip()
                name1 = match.group(2).strip()
                # step3 check if it is 'Circle' with TP2D
                if name0 == 'Circle' and last_Was_DATUM_Head:
                    names = ['X','Y',name1]
                    for i,name in enumerate(names):
                        line = f.readline()# read 3 times
                        oneLine = findLine(line)
                        #print line,oneLine
                        if i==0 and oneLine[0]!='X-axis':# Circle w/o X Y
                            oneLine[0]=name1
                            data.append(oneLine)
                            break
                        else:
                            oneLine[0]=name
                            data.append(oneLine)
                    last_Was_DATUM_Head = False
                elif match.group(3) is not None:
                    line = f.readline()
                    oneLine = findLine(line)
                    oneLine[0]=match.group(2)+match.group(3)+match.group(6)
                    # 'Circle:H001--Circle:H002
                    # (1   )   (2)(3)(4) (5)(6)
                    data.append(oneLine)
                else: # setp 4 general name substitution
                    #print name0,':::',name1
                    if name0 != 'Time':
                        line = f.readline()# read
                        oneLine = findLine(line)
                        oneLine[0]=name1
                        data.append(oneLine)
                    last_Was_DATUM_Head = False

                # above is not caption line
                last_Was_Caption=False            
            else: # step 5 test
                #pattern: word{space|\n} [repeat 1-5 times] word{space|\n}
                pattern = re.compile(r'^([a-zA-Z0-9_/]+[ \n]){1,5}\Z')# find any line begins with letter
                match = pattern.match(line)
                if match:
                    if last_Was_Caption :
                        data[len(data)-1]=[data[len(data)-1][0]+line]
                    else:
                        data.append([line])
                    last_Was_Caption=True

                last_Was_DATUM_Head = False
        
        line = f.readline()
    f.close()
    return data

################## function myFirstPage ###############
# Fixed features of the first page of the document
#######################################################
def myFirstPage(c, doc):
    c.saveState()
    #c = canvas.Canvas(pdf_name, pagesize=A4)
    width, height = A4
    top = 4*cm

    c.setLineWidth(.3)
    #c.setFont('Helvetica-Bold', 22)
    c.setFont('Courier-Oblique', 22)

    c.drawImage(GM_logo,width/3-4/2*cm,height-top,width=3*cm,height=2*cm)
    c.drawImage(NIKON_logo,width/3*2-2.5/2*cm,height-top,width=2.5*cm,height=2.5*cm)
     
    c.drawCentredString(width/2,height-top-1*cm,'Nikon Metrology Camio Software')
    c.setFont('Courier-Oblique', 9)
    c.drawString(1.5*cm, 0.75 * inch, "Page %d %s" % (doc.page, pageinfo))
    c.drawString(1.5*cm, 1 * inch, OUTPUT_PDF_DIR+OUTPUT_PDF_NAME)
    c.restoreState()

######################## function myLaterPages #################################
# For pages after the first to look different from the first
################################################################################
def myLaterPages(c, doc):
    width, height = A4
    c.saveState()
    c.drawImage(NIKON_logo,18.2*cm,height-2.5*cm,width=1.3*cm,height=1.3*cm)
    c.setFont('Courier-Oblique', 20)
    c.drawString(1.5*cm, height-2.5*cm, "LK Camio Studio Reporting")
    c.setFont('Courier-Oblique', 9)
    c.drawString(1.5*cm, 0.75 * inch, "Page %d %s" % (doc.page, pageinfo))
    c.drawString(1.5*cm, 1 * inch, OUTPUT_PDF_DIR+OUTPUT_PDF_NAME)
    c.restoreState()

#################### function getParameters ###############
# create a table heading block
###########################################################
def getParameters(res_filename):
    #read from Nikon.txt
    Body_OR_Head=''
    Type_name=''
    cell=''

    global OUTPUT_PDF_DIR
    global OUTPUT_PDF_NAME
    global MAIN_RES_FILE

    
    f=open(NIKON_TXT_FILE)
    line_Plan_Name           = f.readline()
    line_Part_Description    = f.readline()
    line_Plant               = f.readline()
    line_Department          = f.readline()
    line_Part_No             = f.readline()                 
    line_Operation           = f.readline()
    line_Spindle             = f.readline()
    line_Part_Serial_No      = f.readline()
    line_Change_Level        = f.readline()
    line_Gage                = f.readline()
    line_Machine             = f.readline()
    line_Event               = f.readline()
    line_Module              = f.readline()

    #prepare pdf name part1
    pattern = re.compile(r'(\w\w)_([\w ]+)')# to match name like 'LTG MY2016'
    match = pattern.match(line_Plan_Name)
    if match:
        sub_pattern = re.compile(r'(\w+) (\w+)')# to match 2nd group if like 'LTG MY2016'
        sub_match = sub_pattern.match(match.group(2))
        if sub_match:
            OUTPUT_PDF_DIR = OUTPUT_PDF_DIR +'\\'+ sub_match.group(1)+'\\'+match.group(1)+' '+ sub_match.group(2)# LTG\CB
        else:
            OUTPUT_PDF_DIR = OUTPUT_PDF_DIR +'\\'+ match.group(2)+'\\'+match.group(1)# LTG\CB
        #print OUTPUT_PDF_DIR +" ===test==="
        Body_OR_Head=match.group(1)
        Type_name   =match.group(2)
        MAIN_RES_FILE = getResDir(Body_OR_Head,Type_name)
        #print MAIN_RES_FILE
    else:
        pattern = re.compile(r'(\w+)-(CH)')# to match name like 'LTG-CH'
        match = pattern.match(line_Plan_Name)
        if match:
            OUTPUT_PDF_DIR = OUTPUT_PDF_DIR +'\\'+ match.group(1)+'\\'+match.group(2)# LTG\CH
            #print OUTPUT_PDF_DIR +" ===test==="
            Body_OR_Head=match.group(2)
            Type_name   =match.group(1)
            MAIN_RES_FILE = getResDir(Body_OR_Head,Type_name)
            #print MAIN_RES_FILE
        else:
            showMessage('Can not find plan name: <'+line_Plan_Name+'>')

    #prepare pdf name part2
    pattern = re.compile(r'OP(\d{2,3})?.*')
    match = pattern.match(line_Operation)
    if match:
        num = int(match.group(1))
        if Body_OR_Head == 'CB':#body
            if num >= 10 and num <=40 :
                cell = 'CELL1'
            elif num >= 70 and num <=90:
                cell = 'CELL2'
            elif num == 100:
                cell = 'CELL3'
            elif num >= 110 and num <=130:
                cell = 'CELL4'
            elif num ==140:
                cell = 'CELL5'
            elif num ==150:
                cell = 'CELL6'
        if Body_OR_Head == 'CH':#head
            if num >= 10 and num <=50 :
                cell = 'CELL1'
            elif num >= 90 and num <=110:
                cell = 'CELL2'
            elif num == 140:
                cell = 'CELL3'                    
    OUTPUT_PDF_DIR = OUTPUT_PDF_DIR + '\\'+cell +'\\'
    OUTPUT_PDF_NAME = Type_name+'_'+Body_OR_Head+'_'+cell
    st = datetime.datetime.now().strftime('%Y%m%d%H%M')
    OUTPUT_PDF_NAME = OUTPUT_PDF_NAME+'_'+line_Event.replace('\n','')+'_'+st+'.pdf'
    #PDF file name ready
                    
    Plan_Name           = Paragraph(line_Plan_Name       ,styles['Italic'])
    Part_Description    = Paragraph(line_Part_Description,styles['Italic'])
    Plant               = Paragraph(line_Plant           ,styles['Italic'])
    Department          = Paragraph(line_Department      ,styles['Italic'])
    Part_No             = Paragraph(line_Part_No         ,styles['Italic'])                    
    Operation           = Paragraph(line_Operation       ,styles['Italic'])
    Spindle             = Paragraph(line_Spindle         ,styles['Italic'])
    Part_Serial_No      = Paragraph(line_Part_Serial_No  ,styles['Italic'])
    Change_Level        = Paragraph(line_Change_Level    ,styles['Italic'])
    Gage                = Paragraph(line_Gage            ,styles['Italic'])
    Machine             = Paragraph(line_Machine         ,styles['Italic'])
    Event               = Paragraph(line_Event           ,styles['Italic'])
    Module              = Paragraph(line_Module          ,styles['Italic'])

    f.close

    
    if res_filename is not None:
        #showMessage('inside getParameter, res_filename='+res_filename)
        MAIN_RES_FILE = res_filename
    else:
        MAIN_RES_FILE = getResFileDir(MAIN_RES_FILE)
        
    #read from RES
    f=open(MAIN_RES_FILE)
    
    content = f.read()
    #search items
    pattern = re.compile(r'(.|\n)+?Time:  (\d\d:\d\d:\d\d)')# the ? means lazy mode
    match = pattern.match(content)
    if match:
        Start_Time=Paragraph(match.group(2),styles['Italic'])
    #pattern = re.compile(r'(.|\n)+Duration +((\d+ +mins)? +\d+ +secs)')
    pattern = re.compile(r'(.|\n)+Duration +((\d+ +mins +)?\d+ +secs)')
    #begin with \n, and group(3) is optional, and " +" means many spaces
    match = pattern.match(content)
    if match:
        Run_Time=Paragraph(match.group(2).replace(' ',''),styles['Italic'])
    #make default temperature = 20    
    Temperature=Paragraph('20',styles['Italic'])
    if 'Temperature Compensation:  OFF' not in content :# when temprature off, set to 20
        pattern = re.compile(r'(.|\n)+?COMP,(\d+\.\d+)')
        match = pattern.match(content)
        if match:
            Temperature=Paragraph(match.group(2),styles['Italic'])
    pattern = re.compile(r'-+\n(\d+-.+-\d+)')
    match = pattern.match(content)
    if match:
        Date=match.group(1).decode('gb2312') # does not use Paragraph() due to Chinese char

    #count numbers
    NumberOfGraphic =0
    NumberOfError=0
    lines = content.split('\n')
    for line in lines:
        lineList= findLine(line)
        if lineList is not None:
            if lineList[-2]!='':
                pattern = re.compile(r'[\*+-><]+')
                if pattern.match(lineList[-2]):
                    NumberOfGraphic = NumberOfGraphic+1
            if lineList[-1]!='':
                pattern = re.compile(r'[+-\.\d]+')
                if pattern.match(lineList[-1]):
                    NumberOfError = NumberOfError+1
    f.close
    
    #Operation1 = Paragraph(Operation, styles['Normal']) # this can auto wrap in side a cell
    #Machine1 = Paragraph(Machine,styles['Normal'])      # auto wrap
    
    #data=  [['Plan Name'        , Plan_Name     ,''             ,'Part Description' ,Part_Description   ,''          ],
    #        ['Plant'            , Plant         ,''             ,'Department'       ,Department         ,''          ],
    #        ['Part No.'         , Part_No       ,'Operation'    ,Operation1         ,'Spindle'          ,Spindle     ],
    #        ['Part Serial No.'  , Part_Serial_No,'Machine'      ,Machine            ,'Change Level'     ,Change_Level],
    #        ['Date'             , Date          ,'Start Time'   ,Start_Time         ,'Gage'             ,Gage        ],
    #        ['Event'            , Event         ,'Run Time'     ,Run_Time           ,'Temperature'      ,Temperature ]]
    
    data=  [['Plan Name'        , Plan_Name     ,''             ,'Part Description' ,Part_Description   ,''          ],
            ['Plant'            , Plant         ,''             ,'Department'       ,Department         ,''          ],
            ['Part No.'         , Part_No       ,'Operation'    ,Operation          ,''                 ,''          ],
            ['Part Serial No.'  , Part_Serial_No,'Machine'      ,Machine            ,''                 ,''          ],
            ['Spindle'          , Spindle       ,'Gage'         ,Gage               ,'Change Level'     ,Change_Level],
            ['Date'             , Date          ,'Start Time'   ,Start_Time         ,'Run Time'         ,Run_Time    ],
            ['Event'            , Event         ,'Module'       ,Module             ,'Temperature'      ,Temperature ],
            ['Total : '+ str(NumberOfGraphic) + '    out : '+ str(NumberOfError)]]  
    return data

################### main function #######################
# main function
#########################################################
def main(argv):


    Story = [Spacer(1,2.5*cm)]
    style = styles["Normal"]
    ########## GM powertrain #####
    data=[['','GM POWERTRAIN']]
    t=Table(data,
            colWidths=[12.1*cm,5.8*cm],
            style=[
                ('BOX',(1,0),(1,-1),1.5,colors.black)
            ])
    Story.append(t)
    ################### table 1st block ##############################
    #check command parameter number
    tmpshowmsg = 'arg number'+str(len(argv))
    #showMessage( tmpshowmsg)
    if len(argv)!=2 :
        showMessage('.res file name error!')    
        data= getParameters(None)
    else:
        #showMessage( 'argv[1]')#show ref pathname
        data=getParameters(argv[1])
        
    t=Table(data,
            colWidths=[3.3*cm,3.1*cm,2.5*cm,3.2*cm,2.9*cm,2.9*cm],
            style=[
         ('GRID',(0,0),(-1,-1),0.5,colors.grey),
         ('SPAN',(1,0),(2,0)),
         ('SPAN',(-2,0),(-1,0)),
         ('SPAN',(1,1),(2,1)),
         ('SPAN',(-2,1),(-1,1)),
         ('SPAN',(3,2),(-1,2)),
         ('SPAN',(3,3),(-1,3)),
         ('SPAN',(0,-1),(-1,-1)),
         ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold'),

         ('FONTSIZE', (0, 0), (-1, -1),10), # here it only defines item names. Does not affect Paragraph()
         #('FONT', (1, 0), (1, -1), 'Helvetica'),
         #('FONT', (3, 2), (3, -1), 'Helvetica'),
         #('FONT', (-2, 0), (-2, 1), 'Helvetica'),
         #('FONT', (-1, 4), (-1, -1), 'Helvetica'),
         ('FONT', (1, 5), (1, 5), 'STSong-Light'),
         ('BOX', (0, 0), (-1, -1), 1.5, colors.black),
         ('BOX', (0, 0), (-1, -1), 0.5, colors.white)
         ])
    Story.append(t)
    ################## table 2nd block ##############################
    
    data = createTableData()
    t=Table(data,
            colWidths=[5.7*cm,2.1*cm,2.1*cm,1.6*cm,1.6*cm,1.6*cm,1.6*cm,1.6*cm],
            repeatRows=1)
    ####### Style setting ####
    style=[
         ('GRID',(0,0),(-1,-1),0.5,colors.grey),
         ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
         ('FONT', (0, 0), (-1, 0), 'Times-Bold'),
         ('FONT', (0, 1), (-1, -1), 'Times-Italic'),
         ('FONTSIZE', (0, 0), (-1, -1), 10),
         ('TOPPADDING', (0, 0), (-1, -1), 0),
         ('BOX', (0, 0), (-1, -1), 1.5, colors.black),
         ('BOX', (0, 0), (-1, -1), 0.5, colors.white)
         ]
    for rowIdx,row in enumerate(data):
        if len(row)==1:
            style.append(('SPAN',(0,rowIdx),(-1,rowIdx)))
            style.append(('BOTTOMPADDING',(0,rowIdx),(-1,rowIdx),-8))
    t.setStyle(style)
    ###### Style setting  ####
                         
    Story.append(t)
    #############################################################
    t = Table([['End of table.']],style=[('FONT',(0,0),(-1,-1),'Times-Italic')])
    Story.append(t)

    doc = SimpleDocTemplate(OUTPUT_PDF_DIR+OUTPUT_PDF_NAME)
    #showMessage('output='+OUTPUT_PDF_DIR+OUTPUT_PDF_NAME)
    doc.build(Story, onFirstPage=myFirstPage, onLaterPages=myLaterPages)

# >>>>>>> entry point >>>>>>
main(sys.argv)
