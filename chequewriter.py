import os
import subprocess
try:
    from reportlab.pdfgen import canvas
except ImportError:
    # ReportLab is not installed, install it using pip
    print("Installing ReportLab...")
    subprocess.run(['pip', 'install', 'reportlab'])
    print("ReportLab installed successfully.")
    from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# define constants
CHEQUE_WIDTH = 215.9 * mm
CHEQUE_HEIGHT = 95.3 * mm
TOP_MARGIN = 10 * mm
LEFT_MARGIN = 15 * mm
RIGHT_MARGIN = 15 * mm
BOTTOM_MARGIN = 15 * mm
MICR_FONT_SIZE = 12
MICR_FONT_PATH = 'GnuMICR.ttf'

# function to convert number to words
def convert_number_to_words(num):
    ones = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]
    tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
    teens = ["ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
    words = ""
    
    if num == 0:
        return "zero"
    
    if num < 0:
        words += "minus "
        num = abs(num)
    
    if num >= 1000000:
        words += convert_number_to_words(num // 1000000) + " million "
        num %= 1000000
        
    if num >= 1000:
        words += convert_number_to_words(num // 1000) + " thousand "
        num %= 1000
        
    if num >= 100:
        words += ones[num // 100] + " hundred "
        num %= 100
        
    if num >= 20:
        words += tens[num // 10] + " "
        num %= 10
        
    if num >= 10 and num <= 19:
        words += teens[num - 10] + " "
        num = 0
        
    if num > 0:
        words += ones[num] + " "
    
    return words.capitalize()

# function to create the printable file
def create_cheque():
    # get input from user
    name = input("Enter Your Name: ")
    address1 = input("Enter Your First Address Line(1/2): ")
    address2 = input("Enter Your Second Address Line(2/2): ")
    bank_name = None
    while bank_name not in ["CIBC", "RBC", "TD", "SCOTIABANK", "BMO"]:
        print("Select your bank:")
        print("1. CIBC")
        print("2. RBC")
        print("3. TD")
        print("4. Scotiabank")
        print("5. BMO")
        bank_choice = input("Enter your choice (1-5): ")
        if bank_choice == "1":
            bank_name = "Canadian Imperial Bank of Commerce"
        elif bank_choice == "2":
            bank_name = "Royal Bank of Canada"
        elif bank_choice == "3":
            bank_name = "TD Canada Trust"
        elif bank_choice == "4":
            bank_name = "Scotiabank"
        elif bank_choice == "5":
            bank_name = "Bank of Montreal"
        else:
            print("Invalid choice. Please try again.")
    bank_address = input("Enter Bank Branch Address: ")
    date = input("Enter Date (YYYY-MM-DD): ")
    payee_name = input("Enter Payee Name: ")
    amount = float(input("Enter the amount: "))
    cheque_number = input("Enter Cheque Number (3 digits): ")
    transit_number = input("Enter Transit (Branch) Number: ")
    institution_number = input("Enter Institution Number: ")
    account_number = input("Enter Account Number: ")
    memo = input("Enter Memo: ")
    
    # convert amount to words
    amount_in_words = convert_number_to_words(int(amount))
    
    # format MICR line
    if bank_name.upper() == "Canadian Imperial Bank of Commerce":
        micr_line = "C{}C A{}D{}A {}D{}C".format(
            cheque_number.zfill(3),
            transit_number.zfill(5),
            institution_number.zfill(3),
            account_number[:2],
            account_number[2:]
        )
    elif bank_name.upper() == "Royal Bank of Canada":
        micr_line = "C{}C A{}D{}A {}D{}D{}".format(
            cheque_number.zfill(5),
            transit_number.zfill(5),
            institution_number.zfill(3),
            account_number[:3],
            account_number[3:6],
            account_number[6:]
        )
    elif bank_name.upper() == "TD Canada Trust":
        micr_line = "C{}C A{}D{}A {}D{}C".format(
            cheque_number.zfill(3),
            transit_number.zfill(5),
            institution_number.zfill(3),
            account_number[:4],
            account_number[4:]
        )
    elif bank_name.upper() == "Scotiabank":
        micr_line = "C{}C A{}D{}A {}D{}C".format(
            cheque_number.zfill(3),
            transit_number.zfill(5),
            institution_number.zfill(3),
            account_number[:5],
            account_number[5:]
        )
    elif bank_name.upper() == "Bank of Montreal":
        micr_line = "C{}C A{}D{}A {}D{}C".format(
            cheque_number.zfill(3),
            transit_number.zfill(5),
            institution_number.zfill(3),
            account_number[:4],
            account_number[4:]
        )

    # create file name
    file_name = "{}_{}.pdf".format(payee_name.lower().replace(" ", ""), date.replace("/", ""))
    
    # create PDF file and write data to it
    c = canvas.Canvas(file_name, pagesize=letter)
    
    # load fonts
    pdfmetrics.registerFont(TTFont('GnuMICR', MICR_FONT_PATH))
    c.setFont('GnuMICR', MICR_FONT_SIZE)
    
    # draw MICR line
    micr_line_y = CHEQUE_HEIGHT - TOP_MARGIN - 75 * mm
    c.drawString(LEFT_MARGIN, micr_line_y, micr_line)
    
    # load Arial font for other text
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    pdfmetrics.registerFont(TTFont('Arial_bold', 'arial_bold.ttf'))
    c.setFont('Arial_bold', 12)
    
    # draw cheque elements
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN, name)
    c.setFont('Arial', 10)
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 5 * mm, address1)
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 10 * mm, address2)
    c.setFont('Arial', 12)
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 20 * mm, CHEQUE_HEIGHT - TOP_MARGIN, cheque_number)
    c.setFont('Arial', 10)
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 30 * mm, CHEQUE_HEIGHT - TOP_MARGIN - 10 * mm, "Date: {}".format(date))
    c.setFont('Arial', 12)
    c.drawString(CHEQUE_WIDTH / 3, CHEQUE_HEIGHT - TOP_MARGIN - 20 * mm, bank_name)
    c.setFont('Arial', 10)
    c.drawString(CHEQUE_WIDTH / 3, CHEQUE_HEIGHT - TOP_MARGIN - 25 * mm, bank_address)
    c.setFont('Arial', 12)
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 35 * mm, "PAY {} & {}/100 Dollars".format(amount_in_words, format_amount(amount)))
    c.setFont('Arial_bold', 12)
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 20 * mm, CHEQUE_HEIGHT - TOP_MARGIN - 45 * mm, "$ {:,.2f}".format(float(amount)))
    c.setFont('Arial', 10)
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 45 * mm, "Pay to the")
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 50 * mm, "order of______________________________________")
    c.setFont('Arial', 12)
    c.drawString(LEFT_MARGIN + 20 * mm, CHEQUE_HEIGHT - TOP_MARGIN - 50 * mm, payee_name)
    c.setFont('Arial', 10)
    c.drawString(LEFT_MARGIN, CHEQUE_HEIGHT - TOP_MARGIN - 60 * mm, "Memo: {}".format(memo))
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 45 * mm, CHEQUE_HEIGHT - TOP_MARGIN - 60 * mm, "____________________")
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 45 * mm, CHEQUE_HEIGHT - TOP_MARGIN - 65 * mm, "AUTHORIZED SIGNATURE")
    c.setFont('Arial', 12)
    
    # add blank line for authorized signature
    c.drawString(CHEQUE_WIDTH - RIGHT_MARGIN - 50 * mm, BOTTOM_MARGIN + 10, "")
    
    # save and close PDF file
    c.save()
    
    # print success message
    print("Cheque created successfully! File Name: {}".format(file_name))

def format_amount(amount):
    cents = int(round(amount * 100))
    formatted_amount = "{:02d}".format(cents % 100)
    return formatted_amount

# run the program
create_cheque()