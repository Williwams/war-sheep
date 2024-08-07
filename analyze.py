import re
import pikepdf
from tika import parser
import random
import csv
from PIL import Image
import pytesseract
import os
from mailmerge import MailMerge
from pdf2image import convert_from_path
import time
from datetime import datetime
import shutil

# Initialized with global scope
gender_identifiers = {'He': 'They',
                      'She': 'They',
                      'Himself': 'Themself',
                      'Herself': 'Themself',
                      'His': 'Their',
                      'Her': 'Their',
                      'Male': 'GenderRedact',
                      'Female': 'GenderRedact'
                      }
af910_template = 'templates/AF910_TEMPLATE.docx'


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class chk_codes():
    Checked = "R"
    Unchecked = "£"


class UserData():
    def get_user_folder_path(self):
        return os.path.join("user_files", self.folder_name)

    def files(self):
        return os.listdir(self.get_user_folder_path())

    def get_file(self, file):
        return os.path.join("user_files", self.folder_name, file)

    def get_name(self):
        return f"{self.lname}, {self.fname} {self.mname}"

    def __init__(self, sanitized_name, folder_name, lname="", fname="", mname=""):
        self.sanitized_name = sanitized_name
        self.folder_name = folder_name
        self.lname = lname
        self.fname = fname
        self.mname = mname if len(mname) > 0 else "NMN"

    def name_regex(self):
        options = []
        options.append(re.compile(fr'({self.lname} {self.fname} {self.mname})'))
        options.append(re.compile(fr'({self.fname} {self.mname} {self.lname})'))
        options.append(re.compile(fr'({self.fname} {self.mname[0:1]}[.] {self.lname})'))
        options.append(re.compile(fr'(Sergeant|Airman) ({self.lname})'))
        options.append(re.compile(fr'({self.lname} {self.fname})'))
        options.append(re.compile(fr'({self.fname} {self.lname})'))
        options.append(re.compile(fr'({self.lname})'))
        return options

    def output_file(self, file):
        return os.path.join("output_files", self.sanitized_name, file)

    def print_textfile(self, text, filename):
        with open(self.output_file(filename), 'w') as f:
            f.write(text.strip())
            f.close()

    def create_output_folder(self):
        try:
            os.mkdir(os.path.join("output_files", self.sanitized_name), mode=0o777)
        except FileExistsError:
            for old_file in os.listdir(os.path.join("output_files", self.sanitized_name)):
                os.remove(os.path.join("output_files", self.sanitized_name, old_file))

    def standardize_user_files(self):
        for file in os.listdir(self.get_user_folder_path()):
            os.rename(os.path.join(self.get_user_folder_path(), file), os.path.join(self.get_user_folder_path(), file.upper()))

    def convert_pdf_to_image(self, file, orig_id):
        if (file[file.index('.'):] == ".PDF"):
            pages = convert_from_path(file)
            for id, page in enumerate(pages):
                new_file = os.path.join("tmp",f'{file.split("/")[-1][:file.index("."):]}_{orig_id}_Part{id}.PNG')
                page.save(new_file, 'PNG')
                self.process_image(new_file, f'{orig_id}_Part{id}')
                os.remove(new_file)

    def process_image(self, file, id):
        if (file[file.index('.'):] == ".PDF"):
            return self.convert_pdf_to_image(file, id)
        img_start = time.time()
        dec_text = pytesseract.image_to_string(Image.open(file))
        self.print_textfile(dec_text, f"Dec{id}_{self.sanitized_name}.txt")
        print(f"{file} took {time.time()-img_start} seconds to print")

    def get_file_list(self, filetype):
        surf_pattern = re.compile(r'.*SURF.*[.]PDF')
        epr_pattern = re.compile(r'.*EPR.*[.]PDF')
        dec_pattern = re.compile(r'.*[.]TIF|.*[.]PNG|.*DECORATION.*[.]PDF')
        if filetype == "SURF":
            file_list = list(filter(surf_pattern.search, self.files()))
        if filetype == "EPR":
            file_list = list(filter(epr_pattern.search, self.files()))
        if filetype == "DEC":
            file_list = list(filter(dec_pattern.search, self.files()))
        if len(file_list) == 0:
            print(f"{bcolors.WARNING}Warning! User {self.folder_name} does not have a {filetype} that could be recognized{bcolors.ENDC}")
        return file_list

    def map_surf(self, file):
        text = read_pdf(self.get_file(file))
        username = re.search(r'Name:(.*?)SSAN:', text).group(1).strip().split(" ")
        self.lname = username[0]
        self.fname = username[1]
        self.mname = username[2] if len(username) >= 3 else "NMN"
        text = sanitize_data(text, self)
        text = re.sub(r'SEX/RACE/ETH-GR:(.*)', "SEX/RACE/ETH-GR: MASKED", text)
        text = re.sub(r'SSAN:(.*)', "SSAN: SSN-MASKED", text)
        self.print_textfile(text, f"SURF_{self.sanitized_name}.txt")


class EPRInfo():
    Username = ""
    Rank = ""
    DAFSC = ""
    Org = ""
    ReportFrom = ""
    ReportThru = ""
    DaysNonRated = ""
    DaysRated = ""
    ReportReason = ""
    DutyTitle = ""
    KeyDuties = ""
    BulletsTaskKnowledge = ""
    BulletsFollowership = ""
    BulletsWholeAirman = ""
    BulletsAdditionalRater = ""
    BulletsCommander = ""
    FutureRole1 = ""
    FutureRole2 = ""
    FutureRole3 = ""
    PromotionEligible = ""
    RaterName = ""
    RaterDutyTitle = ""
    RaterDate = ""
    AddlRaterName = ""
    AddlRaterDutyTitle = ""
    AddlRaterDate = ""
    Referral = ""
    QualityForceReview = ""
    Remarks = ""
    UnitCCName = ""
    UnitCCDutyTitle = ""
    UnitCCDate = ""

    def get_new_report_name(self, cur_user):
        return cur_user.output_file(f"{self.ReportThru[-4:]}_{self.Rank}_{cur_user.sanitized_name}.docx")

    def chkUpdate(self, text, cur_user):
        check_types = {'NotRatedCheck': 1, 'MetSomeCheck': 2, 'MetAllCheck': 3, 'ExceededSomeCheck': 4, 'ExceededMostCheck': 5,
                       'DoNotPromoteCheck': 1, 'NotReadyCheck': 2, 'PromoteCheck': 3, 'MustPromoteCheck': 4, 'PromoteNowCheck': 5,
                       'NonConcurCheck': 1, 'ConcurCheck': 2}
        chk_list = {'III': re.compile(fr'\n\tIII(.*): 1'),
                    'IV': re.compile(fr'\n\tIV(.*): 1'),
                    'V': re.compile(fr'\n\tIV(.*): 1'),
                    'VI': re.compile(fr'\n\tVI(.*): 1'),
                    'VIIIC': re.compile(r'\n\tVIII(.*ConcurCheck)(.*): 1'),
                    'IXC': re.compile(r'\n\tIX(.*ConcurCheck)(.*): 1'),
                    'IXP': re.compile(r'\n\tIX(?!.*Concur)(.*Check): 1')
                    }
        for section in chk_list:
            check_text = chk_list[section].search(text)
            if (check_text is not None and check_text.group(1) in check_types):
                score = check_types[check_text.group(1)]
            else:
                if not (section == 'IXP' and self.PromotionEligible == 'NO'):
                    print(f"{bcolors.FAIL}Section {section} did not have a valid checkbox checked for user {cur_user.folder_name}{bcolors.ENDC}")
            for x in range(1, 6):
                setattr(self, f"{section}_{x}", chk_codes.Checked if x == score else chk_codes.Unchecked)

    def merge(self, cur_user):
        document = MailMerge(af910_template)
        variables = vars(self)
        document.merge(**variables)
        document.write(self.get_new_report_name(cur_user))


def read_pdf(file):
    encr_pdf = pikepdf.open(file)
    # Place Decrypted file elsewhere to scan/delete
    decrypt_filename = os.path.join("tmp", file.split("/")[-1])
    encr_pdf.save(decrypt_filename)
    parsedPDF = parser.from_file(decrypt_filename)
    pdf = parsedPDF["content"]
    pdf = pdf.replace('nn', 'n')
    pdf = clean_unicode_spaces(pdf)
    os.remove(decrypt_filename)
    return pdf


def sanitize_data(text, userdata):
    text = text.upper()
    # Cleaning any SSNs out
    text = re.sub(r'(?!000|.+0{4})(?:\d{9}|\d{3}-\d{2}-\d{4})', "SSN-MASKED", text)
    # Cleaning name
    for opt in userdata.name_regex():
        text = re.sub(opt, userdata.sanitized_name, text)
    for opt in gender_identifiers.keys():
        text = re.sub(fr'(\s)({opt.upper()})([\s|.])', fr'\1{gender_identifiers[opt].upper()}\3', text)
    return text


def clean_unicode_spaces(text):
    space_chars = "|".join(['\u2001', '\u2002', '\u2003', '\u2004',
                            '\u2005', '\u2006', '\u2007', '\u2008', '\u2009'])
    return re.sub(space_chars, ' ', text)


def process_epr(cur_user, file):
    text = read_pdf(cur_user.get_file(file))
    cur_epr = EPRInfo()
    cur_epr.Username = cur_user.sanitized_name
    cur_epr.Rank = re.search(r'\n\tRank: (.*)', text).group(1).strip()
    cur_epr.DAFSC = re.search(r'\n\tDAFSC: (.*)', text).group(1).strip()
    cur_epr.Org = re.search(r'\n\tOrgCCLocal: (.*)', text).group(1).strip()
    cur_epr.ReportFrom = re.search(r'Enter Report From Date as DD Mmm YYYY: (\d{2} [a-zA-Z]{3} \d{4})', text).group(1).strip()
    cur_epr.ReportThru = re.search(r'Enter Report Thru Date as DD Mmm YYYY: (\d{2} [a-zA-Z]{3} \d{4})', text).group(1).strip()
    cur_epr.DaysNonRated = str(int(float(re.search(r'\n\tDaysNonRated: (.*)', text).group(1).strip())))
    cur_epr.DaysRated = re.search(r'\n\tDaysSupervised: (.*)', text).group(1).strip()
    cur_epr.ReportReason = re.search(r'\n\tReason4Rpt: (.*)', text).group(1).strip()
    cur_epr.DutyTitle = re.search(r'\n\tDutyTitle: (.*)', text).group(1).strip()
    cur_epr.KeyDuties = "\n".join(re.search(r'\n\tKeyDuties: (.*)', text).group(1).strip().splitlines())
    cur_epr.BulletsTaskKnowledge = "\n".join(re.search(r'\n\tIIIComments: (.*)', text).group(1).strip().splitlines())
    cur_epr.BulletsFollowership = "\n".join(re.search(r'\n\tIVComments: (.*)', text).group(1).strip().splitlines())
    cur_epr.BulletsWholeAirman = "\n".join(re.search(r'\n\tVComments: (.*)', text).group(1).strip().splitlines())
    cur_epr.BulletsAdditionalRater = "\n".join(re.search(r'\n\tVIIIComments: (.*)', text).group(1).strip().splitlines())
    cur_epr.BulletsCommander = re.search(r'\n\tIXComments: (.*)', text).group(1).strip()
    cur_epr.FutureRole1 = re.search(r'\n\tFutureRole1: (.*)', text).group(1).strip()
    cur_epr.FutureRole2 = re.search(r'\n\tFutureRole2: (.*)', text).group(1).strip()
    cur_epr.FutureRole3 = re.search(r'\n\tFutureRole3: (.*)', text).group(1).strip()
    cur_epr.PromotionEligible = re.search(r'\n\tPromotion Eligible: (.*)', text).group(1).strip()
    cur_epr.RaterName = "\n".join(re.search(r'\n\tRaterName: (.*)', text).group(1).strip().splitlines())
    cur_epr.RaterDutyTitle = re.search(r'\n\tRaterDutyTitle: (.*)', text).group(1).strip()
    cur_epr.RaterDate = re.search(r'\n\tRaterSign: \n\tThis field will auto populate once digitally signed: (.*)', text).group(1).strip()
    cur_epr.AddlRaterName = "\n".join(re.search(r'\n\tAddRaterName: (.*)', text).group(1).strip().splitlines())
    cur_epr.AddlRaterDutyTitle = re.search(r'\n\tAddRaterDutyTitle: (.*)', text).group(1).strip()
    cur_epr.AddlRaterDate = re.search(r'\n\tAddRaterSign: \n\tThis field will auto populate once digitally signed: (.*)', text).group(1).strip()
    cur_epr.Referral = re.search(r'\n\tIX4DropDown: (.*)', text).group(1).strip()
    cur_epr.QualityForceReview = re.search(r'\n\tQuality Force Review: (.*)', text).group(1).strip()
    cur_epr.Remarks = re.search(r'\n\tXIRemarks: (.*)', text).group(1).strip()
    cur_epr.UnitCCName = "\n".join(re.search(r'\n\tUnitCCName: (.*)', text).group(1).strip().splitlines())
    cur_epr.UnitCCDutyTitle = re.search(r'\n\tUnitCCDutyTitle: (.*)', text).group(1).strip()
    cur_epr.UnitCCDate = re.search(r'\n\tUnitCCSign: \n\tThis field will auto populate once digitally signed: (.*)', text).group(1).strip()
    cur_epr.chkUpdate(text, cur_user)
    cur_epr.merge(cur_user)


def check_total(check):
    total = 0
    for root, dirs, files in os.walk('output_files'):
        if (check == 'files'):
            total += len(files)
        elif (check == 'users'):
            total += len(dirs)
    # Gitignore and the user map
    if (check == 'files'):
        total -= 2
    return total


def main():
    start_time = time.time()
    users = [name for name in os.listdir("user_files") if os.path.isdir(os.path.join("user_files", name))]
    random.shuffle(users)
    with open("output_files/user_map.csv", 'w', newline='') as csv_file:
        header_row = ["Folder Name", "Generic Identifier", "Name"]
        user_writer = csv.writer(csv_file, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        user_writer.writerow(header_row)
        for id, user in enumerate(users):
            cur_user = UserData(sanitized_name=f"User_{id}",
                                folder_name=user)
            cur_user.create_output_folder()
            cur_user.standardize_user_files()
            remaining_files = cur_user.files()
            for surf in cur_user.get_file_list("SURF"):
                cur_user.map_surf(surf)
                remaining_files.remove(surf)
            for epr in cur_user.get_file_list("EPR"):
                process_epr(cur_user, epr)
                remaining_files.remove(epr)
            for id, dec in enumerate(cur_user.get_file_list("DEC")):
                cur_user.process_image(cur_user.get_file(dec), id)
                remaining_files.remove(dec)
            if (len(remaining_files) > 0):
                print(f"{bcolors.WARNING}Finished {cur_user.folder_name} with unused files! Files unchecked {remaining_files}{bcolors.ENDC}")
            else:
                print(f"{bcolors.OKGREEN}Finished {cur_user.folder_name} Using all files!{bcolors.ENDC}")
            user_writer.writerow([user, cur_user.sanitized_name, cur_user.get_name()])
        backup_time = datetime.now().strftime('%Y_%m_%d_%H%M')
        shutil.make_archive(f"archive/ZIP{backup_time}", 'zip', "output_files")
        print(f"{bcolors.HEADER}BACKUP CREATED IN archive/ZIP{backup_time}")
        print(f"{bcolors.HEADER}---Execution took {time.time() - start_time} seconds ---{bcolors.ENDC}")
        print("")
        print(f"{bcolors.OKBLUE}Ingested {bcolors.OKCYAN}{check_total('files')}{bcolors.OKBLUE} files for {bcolors.OKCYAN}{check_total('users')}"
              f"{bcolors.OKBLUE} users. Collect files from output_files folder.{bcolors.ENDC}")
        print(f"{bcolors.OKBLUE}user_map.csv is to be retained by CSS to trace users back after review.{bcolors.ENDC}")
        print(f"{bcolors.OKBLUE}If no errors were displayed, User_x folders can be delivered to all reviewers for boarding.{bcolors.ENDC}")


if __name__ == "__main__":
    main()
