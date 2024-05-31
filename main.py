import pandas as pd
import openpyxl
from datetime import date, datetime


class SetDates:

    columns = ['Division', 'Dept', 'Instructor Name', 'Class#', 'Section', 'Units','Starting Enrollment', 'Census Enrollment', 'Current Enrollment', 'Start Date', 'Perc Drop']
    section_report_df = pd.DataFrame(columns=columns)

    def __init__(self, section_df):
        self.section_df = section_df
        self.section_df['Start Date'] = pd.to_datetime(self.section_df['Start Date'])
        self.section_df['Add Date'] = pd.to_datetime(self.section_df['Add Date'])
        self.section_df['Drop Dt'] = pd.to_datetime(self.section_df['Drop Dt'])

    # def set_census_date(self):
    #     self.section_df['Census Date'] = self.section_df['Start Date'] + pd.Timedelta(days=21)
    def division(self):
        SEM = ['A&P', 'BIOL', 'BOT','CHEM', 'CIS', 'MATH', 'ASTR', 'BTEC', 'ENGR', 'ESCI', 'GEOL', 'GEOG', 'MICR', 'PHYS',
               'PS', 'ZOOL']
        LA = ['AFRS', 'ASL', 'CHIN', 'COMM', 'ENGL', 'ESL', 'FREN', 'GERM', 'KOR', 'READ', 'SPAN', 'JAPN']
        FA = ['PHOT', 'ART', 'MUS', 'TH', 'DANC', 'FILM', 'JOUR', 'JAMS']
        HSS = ['ANTH', 'ECON', 'PSYC', 'POL', 'HIST', 'CS', 'WGS', 'PHIL', 'AJ', 'EDEL', 'INST', 'HUM', 'SOC']
        BE = ['ACCT', 'BA', 'BCOT', 'LAW', 'EDT', 'RE', 'ACCT', 'FIN']
        TECH = ['ENGT', 'COS', 'WMT', 'MTT', 'WELD', 'AUTO', 'AB', 'ARCH', 'PMT', 'ET']
        HO = ['CDIT', 'CDSE', 'CA', 'NRSG', 'DH', 'DA', 'PTA', 'PHAR', 'CD', 'CDEC', 'HO', 'MA', 'SLP']
        AED = ['AED', 'IWAP', '', '', '', '']
        COUN = ['COUN', '', '', '', '', '']
        ACLR = ['ACLR', '', '', '', '', '']
        KIN = ['KIN', 'HED', 'ATH', 'PEX', '', '']
        ACAD = ['LIBR']
        # LA = ['', '', '', '', '', '']

        for i in range(len(self.section_df)):
            if self.section_df.loc[i, 'Subject'] in SEM:
                self.section_df.loc[i, 'DIV'] = 'SEM'
            elif self.section_df.loc[i, 'Subject'] in LA:
                self.section_df.loc[i, 'DIV'] = 'LA'
            elif self.section_df.loc[i, 'Subject'] in FA:
                self.section_df.loc[i, 'DIV'] = 'FAC'
            elif self.section_df.loc[i, 'Subject'] in HSS:
                self.section_df.loc[i, 'DIV'] = 'HSS'
            elif self.section_df.loc[i, 'Subject'] in BE:
                self.section_df.loc[i, 'DIV'] = 'BE'
            elif self.section_df.loc[i, 'Subject'] in HO:
                self.section_df.loc[i, 'DIV'] = 'HO'
            elif self.section_df.loc[i, 'Subject'] in TECH:
                self.section_df.loc[i, 'DIV'] = 'TECH'
            elif self.section_df.loc[i, 'Subject'] in KIN:
                self.section_df.loc[i, 'DIV'] = 'KIN'
            elif self.section_df.loc[i, 'Subject'] in AED:
                self.section_df.loc[i, 'DIV'] = 'AED'
            elif self.section_df.loc[i, 'Subject'] in COUN:
                self.section_df.loc[i, 'DIV'] = 'COUN'
            elif self.section_df.loc[i, 'Subject'] in ACLR:
                self.section_df.loc[i, 'DIV'] = 'SAS'
            elif self.section_df.loc[i, 'Subject'] in ACAD:
                self.section_df.loc[i, 'DIV'] = 'ACAD'



    def start_date_enrollment(self):

        """We do a calculation by first getting the length of the total data frame then the length of the drops and subract one from the other. """
        self.section_df[['Instructor First Name', 'Instructor Last Name']] = self.section_df[['Instructor First Name', 'Instructor Last Name']].fillna('Staff')
        self.section_df['Instructor Last Name'].fillna('Staff')
        # self.section_df['Drop Dt'] = pd.to_datetime(df['Drop Dt']).dt.date
        start_date = self.section_df.loc[0, 'Start Date']
        # drop_date = self.section_df.loc[0, 'Drop Dt']
        # start_date = start_date.to_date()

        # self.section_df['Enrollment Drop Date'] = pd.to_datetime(self.section_df['Enrollment Drop Date'])
        total_section_enrollment = len(self.section_df)
        start_date_enrollment_df = self.section_df[self.section_df['Drop Dt'] < self.section_df['Start Date']] \
                                   # or self.section_df[[self.section_df['Drop Dt'] == 0]].reset_index()
        drops = len(start_date_enrollment_df)
        starting_enrollment = total_section_enrollment - drops
        # print(start_date_enrollment_df)

        return starting_enrollment

    # def census_date_enrollment(self):
    #     census_date = self.section_df['Start Date'] + pd.Timedelta(days=21)
    #     census_date_enrollment_df = self.section_df[self.section_df['Drop Dt'] > census_date]
    #     census_date_enrollment = len(census_date_enrollment_df)
    #     return census_date_enrollment

    def current_date_enrollment(self):
        self.section_df[['Instructor First Name', 'Instructor Last Name']] = self.section_df[
            ['Instructor First Name', 'Instructor Last Name']].fillna('Staff')
        # self.section_df['Instructor Last Name'].fillna('Staff')
        # self.section_df['Drop Dt'] = pd.to_datetime(df['Drop Dt']).dt.date
        start_date = self.section_df.loc[0, 'Start Date']
        # drop_date = self.section_df.loc[0, 'Drop Dt']
        # start_date = start_date.to_date()

        # self.section_df['Enrollment Drop Date'] = pd.to_datetime(self.section_df['Enrollment Drop Date'])
        section_enrollment_df = self.section_df[self.section_df['Enrollment Status'] == 'Enrolled']
        total_section_enrollment = len(section_enrollment_df)
        waitlist_enrollment_df = self.section_df[self.section_df['Enrollment Status'] == 'Waiting']
        print('waitlist', waitlist_enrollment_df)
        total_waitlist = len(waitlist_enrollment_df)
        print('number', total_waitlist)
        current_date_enrollment_df = self.section_df[self.section_df['Drop Dt'] < self.section_df['Today Date']] \
            # or self.section_df[[self.section_df['Drop Dt'] == 0]].reset_index()
        drops = len(current_date_enrollment_df)
        current_enrollment = total_section_enrollment
        # print(current_date_enrollment_df)

        return current_enrollment, total_waitlist

        # current_date = self.section_df.loc[0, 'Today Date']
        # semester_end_date = self.section_df.loc[0, 'Semester End Date']

        # current_enrollment_df = self.section_df[self.section_df['Drop Dt'] < self.section_df.loc[0, 'Today Date']]



    def create_report(self, start_date_enrollment, current_enrollment, waitlist_enrollment):
        length = len(SetDates.section_report_df)
        # print(type(section_df.loc[0, 'Instructor First Name']))
        # print((section_df.loc[0, 'Instructor First Name']))
        SetDates.section_report_df.loc[length, 'Division'] = section_df.loc[0, 'DIV']
        SetDates.section_report_df.loc[length, 'Dept'] = section_df.loc[0, 'Subject']
        SetDates.section_report_df.loc[length, 'Instructor Name'] = section_df.loc[0, 'Instructor First Name'] + " " + section_df.loc[0, 'Instructor Last Name']
        SetDates.section_report_df.loc[length, 'Course'] = section_df.loc[0, 'Subject'] + "" + section_df.loc[0,'Catalog']
        SetDates.section_report_df.loc[length, 'Class#'] = section_df.loc[0, 'Class Nbr']
        SetDates.section_report_df.loc[length, 'Units'] = section_df.loc[0, 'Units']
        SetDates.section_report_df.loc[length, 'Start Date'] = section_df.loc[0, 'Start Date']
        SetDates.section_report_df.loc[length, 'Starting Enrollment'] = start_date_enrollment
        SetDates.section_report_df.loc[length, 'Current Enrollment'] = current_enrollment
        SetDates.section_report_df.loc[length, 'Current Waitlist'] = waitlist_enrollment
        if start_date_enrollment == 0:
            SetDates.section_report_df.loc[length, 'Perc Drop'] = 0
        else:
            SetDates.section_report_df.loc[length, 'Perc Drop'] = (start_date_enrollment - current_enrollment) / start_date_enrollment
        # print(SetDates.section_report_df.to_string())
        return SetDates.section_report_df

# def set start dates
# def set census dates
# def calculate enrollment at start date
# def calculate drop rate
# def calculate FTES
df_not_sorted = pd.read_csv(
    'C:/Users/fmixson/Desktop/Enrollment/Spring/CER_SR_ENROLLMENT_ACADAF_24610.csv', encoding='latin-1')

pd.set_option('display.max_columns', None)
# print('not sorted',df_not_sorted)
df = df_not_sorted.sort_values(by=['Class Nbr'])

df.fillna(0)
# df = df.Class_Nbr.astype(str).str.split('.', expand = True)[0]


# semester_end_date = '2023-05-31'
# df['Semester End Date'] = semester_end_date
# df['Semester End Date'] = pd.to_datetime(df['Semester End Date'])

today = date.today()
df.loc[:, 'Today Date'] = today
df['Today Date'] = pd.to_datetime(df['Today Date'])

# start_date = '2023-01-09'
# df.loc[:, 'Start Date'] = start_date
df['Start Date'] = pd.to_datetime(df['Start Date'])
# semester_end_date = datetime.strptime(semester_end_date, '%Y-%m-%d').date()
df = df[df['Term'] != 'nan']
# df['Enrollment Drop Date'] = datetime.strptime(df['Enrollment Drop Date'], '%Y-%m-%d').date()
df['Drop Dt'] = pd.to_datetime(df['Drop Dt']).dt.date
# print(df.dtypes)
section_numbers = []

for i in range(len(df)):
    if df.loc[i, 'Class Nbr'] not in section_numbers:
        section_numbers.append(df.loc[i, 'Class Nbr'])
cleaned_section_numbers = [x for x in section_numbers if pd.notnull(x)]
# print('The number of sections', len( cleaned_section_numbers))
# print(cleaned_section_numbers)
for section in cleaned_section_numbers:
    section_df = df[df['Class Nbr'] == section].reset_index()
    # section_df = section_df.fillna(semester_end_date)

    dates = SetDates(section_df=section_df)
    dates.division()
    # dates.set_census_date()
    start_date_enrollment = dates.start_date_enrollment()
    # census_date_enrollment = dates.census_date_enrollment()
    current_enrollment, waitlist_enrollment= dates.current_date_enrollment()
    report = dates.create_report(start_date_enrollment=start_date_enrollment, current_enrollment=current_enrollment, waitlist_enrollment=waitlist_enrollment)

report.to_csv('SP_Enrollment_by_Section.csv')
report.to_excel('Test.xlsx')
