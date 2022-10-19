import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from PIL import Image
import glob
import time
import pandas as pd
from openpyxl import Workbook

wb = Workbook()
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

df = pd.read_excel('/home/sltech/PycharmProjects/fillForm_selenium/Test.xlsx')
# list(df)
# print(df)

# print(df.loc[0])
s = Service('/home/sltech/Downloads/chromedriver')

browser = webdriver.Chrome(options=options, service=s,
                           )
browser.maximize_window()

browser.get('https://dvprogram.state.gov/application.aspx')


"""Entrant Name,Confirmation Number,Year of Birth,Digital Signature"""
f = open("data.txt", "w")
f.write("""Entrant Name,Confirmation Number,Year of Birth,Digital Signature""")


def img_resize():
    img_list = []
    resize_img = []

    for file_name in glob.glob('/home/sltech/PycharmProjects/fillForm_selenium/resize_images/image/*'):
        img = Image.open(file_name)
        img_list.append(img)

    for image in img_list:
        image = image.resize((600, 600))
        resize_img.append(image)

    for i, new in enumerate(resize_img):
        new.save('{}{}{}'.format('/home/sltech/PycharmProjects/fillForm_selenium/resize_images/', i + 1, '.jpg'))


def authentication(data):
    print('authentication')
    flag=True
    while(flag):
        try:
            browser.find_element('name', '_ctl0:ContentPlaceHolder1:txtCodeInput')
            time.sleep(1)
        except:
            flag=False
            print('auth success')
    formFillup(data)


def start(data):
    print('start')
    delay = 7
    try:
        myElem = WebDriverWait(browser, delay)\
            .until(EC.presence_of_element_located(('name',  '_ctl0:ContentPlaceHolder1:txtCodeInput')))
        print('start try')
        authentication(data)
    except FileNotFoundError:
        print("Loading took too much time!")


def formFillup(data):
    print('fillform')
    # for a in df.index:
    a = data
    # print(data)

    last_name = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl0:txtLastName')
    last_name.send_keys(data['Last_Name'])

    first_name = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl0:txtFirstName')
    first_name.send_keys(data['First_Name'])

    middle_name = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl0:txtMiddleName')
    middle_name.send_keys(data['Middle_Name'])
    # print(middle_name)

    if a['Gender'] == 'Male':
        gender_m = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl1_rdoGenderM')
        gender_m.click()
    else:
        gender_f = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl1_rdoGenderF')
        gender_f.click()

    date_1 = data['Date_of_birth']
    # print(date)

    # date_list = str(date).split('/')
    date_list = str(date_1.date()).split('-')

    birth_month = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl2_txtMonthOfBirth', )
    birth_month.send_keys(date_list[1])

    birth_day = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl2:txtDayOfBirth')
    birth_day.send_keys(date_list[2])

    birth_year = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl2:txtYearOfBirth')
    birth_year.send_keys(date_list[0])

    birth_city = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl3:txtBirthCity')
    birth_city.send_keys(data['Birth_city'])

    birth_country = Select(browser.find_element('name',
                                                '_ctl0:ContentPlaceHolder1:formApplicant:_ctl4:drpBirthCountry'))

    birth_country.select_by_visible_text(data['Birth_country'])

    if a['DV_eligibility'] == 'No':
        dv_no = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl5_rblBirthEligibleCountry_1')
        dv_no.click()

        dv_countries = Select(browser.find_element('name',
                                                   '_ctl0:ContentPlaceHolder1:formApplicant:'
                                                   '_ctl5:drpBirthEligibleCountry'))
        time.sleep(1)
        dv_countries.select_by_visible_text(data['DV_eligibility_country'])
    else:
        dv_yes = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl5_rblBirthEligibleCountry_0')
        dv_yes.click()

    img_resize()
    image_file = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl6_inpPhotograph')
    image_file.send_keys(data['Image'])
    image_btn = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl6_btnPhotoBox')
    # image_btn.click()
    image_btn

    address_line1 = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl7:txtAddress1')
    address_line1.send_keys(data['Address'])

    city_town = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl7:txtCity')
    city_town.send_keys(data['City/Town'])

    dist_state = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl7:txtDistrict')
    dist_state.send_keys(data['District/County/Province/State'])

    pin_code = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl7:txtZipCode')
    pin_code.send_keys(int(data['Postal Code/Zip Code']))

    country = Select(browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl7_drpMailingCountry'))
    country.select_by_visible_text(data['Country'])

    country = Select(browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl8_drpCountry'))
    country.select_by_visible_text(data['Current_living_country'])

    email = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl10:txtEmailAddress')
    email.send_keys(data['Email'])

    confirm_email = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl10:txtConfEmailAddress')
    confirm_email.send_keys(data['Email'])

    if a['Heighest_education'] == 'Primary school only':
        ed1 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_0')
        ed1.click()
    elif a['Heighest_education'] == 'High School, no degree':
        ed2 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_1')
        ed2.click()
    elif a['Heighest_education'] == 'High School degree':
        ed3 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_2')
        ed3.click()
    elif a['Heighest_education'] == 'Vocational School':
        ed4 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_3')
        ed4.click()
    elif a['Heighest_education'] == 'Some University Courses':
        ed5 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_4')
        ed5.click()
    elif a['Heighest_education'] == 'University Degree':
        ed6 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_5')
        ed6.click()
    elif a['Heighest_education'] == 'Some Graduate Level Courses':
        ed7 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_6')
        ed7.click()
    elif a['Heighest_education'] == 'Master\'s Degree':
        ed8 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_7')
        ed8.click()
    elif a['Heighest_education'] == 'Some Doctorate Level Courses':
        ed9 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_8')
        ed9.click()
    elif a['Heighest_education'] == 'Doctorate Degree':
        ed10 = browser.find_element('id', 'ContentPlaceHolder1_formApplicant__ctl11_rblEducation_9')
        ed10.click()

    if a['Marital_status'] == 'Unmarried':
        m_status = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_0')
        m_status.click()
    elif a['Marital_status'] == 'Married and my spouse is NOT a U.S.citizen or U.S. Lawful Permanent Resident (LPR)':
        m_status1 = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_1')
        m_status1.click()
    elif a['Marital_status'] == 'Married and my spouse IS a U.S.citizen or U.S. Lawful Permanent Resident (LPR)':
        m_status2 = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_2')
        m_status2.click()
    elif a['Marital_status'] == 'Divorced':
        m_status3 = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_3')
        m_status3.click()
    elif a['Marital_status'] == 'Widowed':
        m_status4 = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_4')
        m_status4.click()
    elif a['Marital_status'] == 'Legally Separated':
        m_status5 = browser.find_element('id', '_ctl0_ContentPlaceHolder1_formApplicant__ctl12_rblMarried_5')
        m_status5.click()

    children = browser.find_element('name', '_ctl0:ContentPlaceHolder1:formApplicant:_ctl13:txtNumChildren')
    children.send_keys(int(data['Number_of_chindren']))

    submit_btn = browser.find_element('xpath', f'/html/body/form/div[6]/input[1]')
    submit_btn.click()

    time.sleep(2)

    submit_btn2 = browser.find_element('id', 'ContentPlaceHolder1_btnContinueP2')
    submit_btn2.click()



    # sheet['A' + str(sheet_count)] = 'Entrant Name'
    # sheet['B' + str(sheet_count)] = 'Confirmation Number'
    # sheet['C' + str(sheet_count)] = 'Year of Birth'
    # sheet['D' + str(sheet_count)] = 'Digital Signature'
    #
    # sheet_count += 1
    #
    # workbook.save(filename="write_data.xlsx")
    # print(data)
    """Entrant Name,Confirmation Number,Year of Birth,Digital Signature"""
    f = open("data.txt", "a+")
    if browser.find_element('xpath', '//*[@id="printdiv"]'):
        f.write("\n" +
                browser.find_element('xpath', '//*[@id="printdiv"]/div[5]').text+","+
                browser.find_element('xpath', '//*[@id="printdiv"]/div[7]').text+","+
                browser.find_element('xpath', '//*[@id="printdiv"]/div[9]').text+","+
                browser.find_element('xpath', '//*[@id="printdiv"]/div[11]').text
                )
    else:
        print("Please enter new entry!")


    time.sleep(1)
    browser.refresh()

    time.sleep(2)
    browser.get('https://dvprogram.state.gov/application.aspx')



for l in df.index:
    data = df.loc[l]
    a = data
    start(data)





browser.close()

# executable_path=r"/home/sltech/Downloads/chromedriver"

# ContentPlaceHolder1_formApplicant__ctl7_drpMailingCountry

