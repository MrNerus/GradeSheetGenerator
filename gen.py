from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook, load_workbook
from universe import *
import os

# functions
def generateHeader(stdName: str, stdClass: str, stdRoll: str, issuedDate: str, examTitle: str) -> None:
    """
    A function to generate a header on an image with student name, class, roll number, issued date, and exam title.
    Parameters:
        section (int): The section number.
        stdName (str): The name of the student.
        stdClass (str): The class of the student.
        stdRoll (str): The roll number of the student.
        issuedDate (str): The date the header is issued.
        examTitle (str): The title of the exam.
    Returns:
        None
    """
    global myFont
    global newImg
    newImg.text((520,500),   stdName,    font=myFont, fill = (0, 0, 0)) # Name
    newImg.text((1290,500),  stdClass,   font=myFont, fill = (0, 0, 0)) # Grade
    newImg.text((350 ,555),  stdRoll,    font=myFont, fill = (0, 0, 0)) # Roll no.
    newImg.text((1420,555),  issuedDate, font=myFont, fill = (0, 0, 0)) # Date
    newImg.text((1240,400),  examTitle,  font=myFont, anchor="mm", fill = (0, 0, 0)) # ExamTitle


# functions
def generateFooter(gpa: str, gp: str, remarks: str) -> None:
    """
    A function to generate a footer on an image with GPA, Grade, and remarks on a new image.
    
    Parameters:
        gpa (str): The GPA to be displayed.
        gp (str): The Grade to be displayed.
        remarks (str): Any additional remarks to be displayed.
    
    Returns:
        None
    """
    global myFont
    global newImg
    newImg.text((700, 1400), gpa,     font=myFont, fill = (0, 0, 0)) # GPA
    newImg.text((520, 1460), gp,      font=myFont, fill = (0, 0, 0)) # Grade
    newImg.text((360, 1510), remarks, font=myFont, fill = (0, 0, 0)) # Remarks

def generateAttendanceReport(workingDays: str, presentDays: str, absentDays: str) -> None:
    """
    A function to generate a footer on an image with working days, present days, and absent days.
    
    Parameters:
        workingDays (str): The number of working days.
        presentDays (str): The number of days present.
        absentDays (str): The number of days absent.
    
    Returns:
        None
    """
    global myFont
    global newImg
    newImg.text((2070, 1400), workingDays, font=myFont, fill = (0, 0, 0)) # working days
    newImg.text((2050, 1460), presentDays, font=myFont, fill = (0, 0, 0)) # present days
    newImg.text((2040, 1510), absentDays,  font=myFont, fill = (0, 0, 0)) # absent days


def generateBody(yCord: int, subName: str, note: str, gpa: str, gp: str) -> None:
    """
    A function to generate a footer on an image with specified parameters.
    Parameters:
        yCord (int): The y-coordinate for the text.
        subName (str): The subject name to display.
        note (str): The note to include.
        gpa (str): The GPA to display.
        gp (str): The GP to include.
    Returns:
        None
    """
    global myFont
    global newImg
    newImg.text((180, yCord),  subName, font=myFont, fill = (0, 0, 0)) # Subject
    newImg.text((1100, yCord), note,    font=myFont, anchor="ma", fill = (0, 0, 0)) # IsStar
    newImg.text((1360, yCord), gpa,     font=myFont, anchor="ma", fill = (0, 0, 0)) # GPA
    newImg.text((1620, yCord), gp,      font=myFont, anchor="ma", fill = (0, 0, 0)) # GP


# variables
alphabets = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ']

# loads GPA ledger
try:
    ledgerGPA = load_workbook(filename=f'./examLedger.xlsx', data_only=True)
    print(f"{COL_GREEN}File found.{COL_RESET}")
except:
    print(f"{COL_RED}File not found.{COL_RESET}")
    exit()
ledgerGPASheet: Workbook= ledgerGPA["RAW"]

examTitle = str(ledgerGPASheet[f"A1"].value)


issuedDate: str = input("Enter Result Issuance Date: ")

img = Image.open('./BCA GradeSheet Format.jpg')
newImg = ImageDraw.Draw(img)
myFont = ImageFont.truetype('./Inconsolata-SemiBold.ttf', 38)


processedGradeSheet = []

#### This is Third version
i = 4 # Start from 4th row
while True:

    # Exit the loop if current cell is empty. Assumption: SN are never empty
    if ledgerGPASheet[f"A{i}"].value == None:
        break
    
    # define new image EVERY SINGLE TIME
    img = Image.open('./BCA GradeSheet Format.jpg')
    newImg = ImageDraw.Draw(img)
    myFont = ImageFont.truetype('./Inconsolata-SemiBold.ttf', 38)

    # General Data we are working with
    exTitle        = str(ledgerGPASheet[f"A1"   ].value) # Exam Title
    stdGrade       = str(ledgerGPASheet[f"A2"   ].value) # Student Grade
    stdReg         = str(ledgerGPASheet[f"C{i}" ].value) # Regestration No.
    stdName        = str(ledgerGPASheet[f"B{i}" ].value) # Name of Student
    stdSubCount    = int(ledgerGPASheet[f"D{i}" ].value) # Subject Count
    stdGPA         = str(ledgerGPASheet[f"BB{i}"].value) # Grade Point Average
    stdGP          = str(ledgerGPASheet[f"BC{i}"].value) # Grade Point in Word
    stdRemarks     = str(ledgerGPASheet[f"BD{i}"].value) # Remarks
    stdWorkingDays = str(ledgerGPASheet[f"BE{i}"].value) # Working Days
    stdPresentDays = str(ledgerGPASheet[f"BF{i}"].value) # Present Days
    stdAbsentDays  = str(ledgerGPASheet[f"BG{i}"].value) # Absent Days

    # Fill those general Data
    generateHeader(stdName, stdGrade, stdReg, "2079/04/31", exTitle)
    generateFooter(stdGPA, stdGP, stdRemarks)
    generateAttendanceReport(stdWorkingDays, stdPresentDays, stdAbsentDays)

    h = round((1330 - 730)/(stdSubCount-1))
    k = 0
    for j in range(4, 4+7*(stdSubCount),7):
        yCord = 730 + h * k
        subName = str(ledgerGPASheet[f"{alphabets[j+0]}{i}"].value)
        subGPA  = str(ledgerGPASheet[f"{alphabets[j+3]}{i}"].value)
        subGP   = str(ledgerGPASheet[f"{alphabets[j+5]}{i}"].value)
        subFlag = str(ledgerGPASheet[f"{alphabets[j+6]}{i}"].value)

        subName = "<Subject Name>" if subName == 'None' else subName
        subGPA  = "<GPA>"          if subGPA  == 'None' else subGPA
        subGP   = "<GP>"           if subGP   == 'None' else subGP
        subFlag = ""               if subFlag == 'None' else subFlag
        generateBody(yCord, subName, subFlag, subGPA, subGP)
        k += 1
    
    print(f"{stdName}")
    img = img.convert('RGB')
    processedGradeSheet.append(img)
    # img.save(f"./buffer/{str(i).zfill(2)}_{stdReg}_{stdName}.png")

    i += 1

print("Saving to PDF...")
# for i in os.listdir("./Buffer/"):
#     #read the image
#     # print(f"Collecting {i}")
#     img = Image.open(f"./Buffer/{i}")
#     processedGradeSheet.append(img)
coverPic = Image.open(f"./white.png")
coverPic = coverPic.convert('RGB')
coverPic.save("./result.pdf", "PDF" , save_all=True, append_images=processedGradeSheet)

