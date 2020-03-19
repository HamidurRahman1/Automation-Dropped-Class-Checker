
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import getpass
import time
import re
import os
import pandas
import platform
import csv
import datetime


class Utility:
    Login_Url = "https://home.cunyfirst.cuny.edu"
    Login_Domain = "@login.cuny.edu"
    Username_Loc = "CUNYfirstUsernameH"
    Password_Loc = "CUNYfirstPassword"
    Login_Submit_Loc = "submit"
    Bookmarked_Login_Url = "https://ssologin.cuny.edu/errorPage/pages/Error.html"
    Password_Error_Url = "https://ssologin.cuny.edu/unrecognized-credentials.html?p_error_code=OAM-5"
    StudentSrvCtr_Url = "https://cssa.cunyfirst.cuny.edu/psp/cnycsprd/EMPLOYEE/CAMP/c/SCC_ADMIN_OVRD_STDNT" \
                               ".SSS_STUDENT_CENTER.GBL?Folder=MYFAVORITES"
    StudentSrvCtr_Frame_Loc = "ptifrmtgtframe"
    StudentEmpl_Loc = "STDNT_SRCH_EMPLID"
    StudentEmplSrch_Loc = "PSPUSHBUTTONTBSEARCH"
    StudentCourseHis_Loc = "DERIVED_SSS_SCL_SSS_MORE_ACADEMICS"
    StudentCourseHisLoader_Loc = "DERIVED_SSS_SCL_SSS_GO_1"
    StudentCourseHisTable_Loc = "PSLEVEL1GRIDWBO"
    DY_Student_CrsName_Loc = "CRSE_NAME$"
    DY_Student_Grade_Loc = "CRSE_GRADE$"
    DY_Student_Term_Loc = "CRSE_TERM$"

    __DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'downloads')

    PATH_TO_READ_FILE = None
    PATH_TO_WRITE_FILE = None

    REPEATABLE_GRADES = {"w", "wa", "wn", "wd", "wu", "r", "u", "f", "inc"}

    TAB = "\t=> "

    @staticmethod
    def get_term():
        try:
            term = str(input("Enter the current term: ")).strip()
            parts = term.lower().split()
            if len(parts) != 2:
                raise AttributeError("Invalid term entered, given="+term)
            now = datetime.datetime.now()
            year = int(parts[0].strip())
            if now.year != year and now.year+1 != year and now.year-1 != year:
                raise AttributeError("Expected year=" + str(now.year-1)
                                     + " or " + str(now.year) + " or " + str(now.year+1) + " but found " + parts[0])
            if str(parts[1]).lower() != "spring" and str(parts[1]).lower() != "fall":
                raise AttributeError("Expected spring or fall, found="+str(parts[1]+" as term name."))
            return parts[0]+" "+parts[1]+" term"
        except ValueError:
            print(Utility.TAB, ("Expected year=" + str(now.year-1)
                                + " or " + str(now.year) + " or " + str(now.year) + " but found " + parts[0]))
            return Utility.get_term()
        except AttributeError as ae:
            print(Utility.TAB+ae.args[0])
            return Utility.get_term()
        except Exception as e:
            print(Utility.TAB+e.args[0])
            return Utility.get_term()

    @staticmethod
    def delete_temp_csv(file_path):
        if os.path.isfile(file_path) and os.path.exists(file_path):
            os.remove(file_path)

    @staticmethod
    def check_dir_and_get_drop_file():
        try:
            os_name = platform.system().lower()
            if os_name == "darwin" or os_name == "linux":
                Utility.PATH_TO_READ_FILE = Utility.__DOWNLOAD_DIR+"/_droplist_/"
                Utility.PATH_TO_WRITE_FILE = Utility.PATH_TO_READ_FILE+"StudentsToBeChecked.csv"
            elif os_name == "windows":
                Utility.PATH_TO_READ_FILE = Utility.__DOWNLOAD_DIR+"\\_droplist_\\"
                Utility.PATH_TO_WRITE_FILE = Utility.PATH_TO_READ_FILE+"StudentsToBeChecked.csv"
            if not os.path.isdir(Utility.PATH_TO_READ_FILE):
                raise IsADirectoryError(Utility.TAB+"Directory does not exists. PLease create an empty folder name "
                                                    "_droplist_ in downloads directory")
            files = list()
            for file in os.listdir(Utility.PATH_TO_READ_FILE):
                if file.endswith(".csv"):
                    files.append(file)
            if len(files) != 1:
                raise FileExistsError(Utility.TAB+Utility.PATH_TO_READ_FILE
                                      + " folder must contain only one file but found " + str(len(files)))
            Utility.PATH_TO_READ_FILE += str(files[0])
            return Utility.PATH_TO_READ_FILE
        except IsADirectoryError as dir:
            print(dir)
            exit(0)
        except FileExistsError as fe:
            print(fe)
            exit(0)

    @staticmethod
    def get_login_cred():
        username = str(input("Enter CUNYfirst Username: "))
        password = getpass.getpass(prompt="Enter CUNYfirst Password: ", stream=None)
        return username, password

    @staticmethod
    def to_csv(file_dir):
        excel = pandas.read_excel(file_dir)
        csv_path = file_dir.replace(".xlsx", ".csv")
        excel.to_csv(csv_path, index=None, header=True)
        return csv_path


class DroppedClassChecker:

    @staticmethod
    def goto_login(driver, username, password):
        try:
            username = username.lower()+Utility.Login_Domain

            driver.get(Utility.Login_Url)
            username_loc = WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.ID, Utility.Username_Loc)))
            username_loc.clear()
            username_loc.send_keys(username)

            WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.ID, Utility.Password_Loc)))\
                .send_keys(password)

            WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.ID, Utility.Login_Submit_Loc))).click()

            if Utility.Bookmarked_Login_Url == driver.current_url:
                return DroppedClassChecker.goto_login(driver, username, password)
            elif driver.current_url == Utility.Password_Error_Url:
                print(Utility.TAB+"Invalid login username or password provided")
                username, password = Utility.get_login_cred()
                return DroppedClassChecker.goto_login(driver, username, password)
            else:
                return driver
        except TimeoutException as te:
            driver.close()
            print(Utility.TAB, te.msg, te.args)
        except Exception as e:
            driver.close()
            print(Utility.TAB, e.args)

    @staticmethod
    def goto_student_service(driver):
        try:
            driver.get(Utility.StudentSrvCtr_Url)
            return driver
        except TimeoutException as t:
            driver.close()
            print(Utility.TAB, t.msg, t.args)
        except Exception as e:
            driver.close()
            print(Utility.TAB, e.args)

    @staticmethod
    def apply_empl_id_course(driver, empl_id, subject_to_check, sub_level_to_check, term_to_check):
        try:
            fr = WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.ID, Utility.StudentSrvCtr_Frame_Loc)))
            driver.switch_to.frame(fr)
            driver.find_element_by_id(Utility.StudentEmpl_Loc).clear()
            driver.find_element_by_id(Utility.StudentEmpl_Loc).send_keys(empl_id)
            driver.find_element_by_class_name(Utility.StudentEmplSrch_Loc).click()

            WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.ID, Utility.StudentCourseHis_Loc)))\
                .find_elements_by_tag_name("option")[1].click()

            driver.find_element_by_id(Utility.StudentCourseHisLoader_Loc).click()

            time.sleep(4)

            classes = WebDriverWait(driver, 10)\
                .until(ec.visibility_of_element_located((By.CLASS_NAME, Utility.StudentCourseHisTable_Loc)))\
                .find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")

            subject_to_check = subject_to_check.lower()
            term_to_check = term_to_check.lower()
            sub_level_to_check = int(next(re.finditer(r'\d+$', sub_level_to_check)).group(0))
            next_trm = DroppedClassChecker.next_term(term_to_check)

            if len(classes) > 0:
                for r in range(0, len(classes)-1):

                    course = str((WebDriverWait(driver, 10).until(ec.visibility_of_element_located
                        ((By.ID, Utility.DY_Student_CrsName_Loc+str(r)))).text)).lower()

                    grade = str((WebDriverWait(driver, 10).until(ec.visibility_of_element_located
                        ((By.ID, Utility.DY_Student_Grade_Loc+str(r)))).text)).lower().strip()

                    term = str((WebDriverWait(driver, 10).until(ec.visibility_of_element_located
                        ((By.ID, Utility.DY_Student_Term_Loc+str(r)))).text)).lower()

                    parts = course.split()
                    if parts[1].isdigit() and course.startswith(subject_to_check) \
                            and (term == term_to_check or term == next_trm) and len(grade) == 0:
                        return DroppedClassChecker.next_course_level_checker(subject_to_check, int(parts[1]))

        except NoSuchElementException as ne:
            driver.close()
            print(Utility.TAB, ne.msg, ne.args)
        except TimeoutException as te:
            driver.close()
            print(Utility.TAB, te.msg, te.args)
        except Exception as e:
            driver.close()
            print(Utility.TAB, e.args)

    @staticmethod
    def next_course_level_checker(subject, next_level):
        if subject == 'mat':
            if next_level in {99, 117, 119, 123}:
                return False
            else:
                return True

    @staticmethod
    def next_term(this_term):
        parts = this_term.lower().split()
        if parts[1].strip() == "fall":
            return str(int(parts[0].strip())+1)+str(" spring term")
        else:
            return parts[0].strip()+" fall term"


if __name__ == "__main__":
    csv_file = None
    driver = None
    file_reader = None
    file_writer = None
    try:
        csv_file = Utility.check_dir_and_get_drop_file()

        term = Utility.get_term()
        username, password = Utility.get_login_cred()

        file_reader = open(csv_file, "r")
        csv_reader = csv.DictReader(file_reader)

        file_writer = open(Utility.PATH_TO_WRITE_FILE, "w")
        csv_writer = csv.DictWriter(file_writer, fieldnames=csv_reader.fieldnames,
                                    delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)

        csv_writer.writeheader()

        driver = webdriver.Chrome()
        driver.maximize_window()
        driver = DroppedClassChecker.goto_login(driver, username, password)

        for student in csv_reader:
            if len(str(student["Grade In"]).strip()) == 0:
                csv_writer.writerow(student)
                continue
            if str(student["Grade In"]).lower() in Utility.REPEATABLE_GRADES:
                if DroppedClassChecker\
                        .apply_empl_id_course(DroppedClassChecker.goto_student_service(driver),
                        student["ID"], student["Subject"], student["Catalog"], term):
                    csv_writer.writerow(student)

    except FileNotFoundError as f:
        print(Utility.TAB, f.args[0])
        exit(0)
    except KeyboardInterrupt as k:
        print(Utility.TAB, k.args[0])
        exit(0)
    except TimeoutException as te:
        print(Utility.TAB, te.msg, te.args, te.screen)
        exit(0)
    except Exception as e:
        print(Utility.TAB, e.args)
        exit(0)
    finally:
        if csv_file is str:
            file_reader.close()
            file_writer.close()
        if driver is not None:
            driver.close()
