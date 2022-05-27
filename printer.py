import subprocess
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import textwrap
from selenium.webdriver.firefox.options import Options


def ad_to_printer(filename: str, ad_group: str, web_address: str, address_book_no: int, enter_password=False, password="7654321") -> None:
    """Simple script to upload an Active Directory group into Canon printers.

    Args:
        filename        (str): Filename for csv. (not really important)
        ad_group        (str): Active Directory group to extract.
        web_address     (str): IP or web address of the Canon printer webinterface (with port number).
        address_book_no (int): html number of the address book.
        enter_password  (boolean): Defaults to False.
        password        (str): Defaults to 7654321
    """
    
    printer_csv = textwrap.dedent(
        """# Canon AddressBook CSV version: 0x0002
# CharSet: UTF-8
# SubAddressBookName: 
# DB Version: 0x010a

objectclass,cn,cnread,cnshort,subdbid,mailaddress,dialdata,uri,url,path,protocol,username,pwd,member,indxid,enablepartial,sub,faxprotocol,ecm,txstartspeed,commode,lineselect,uricommode,uriflag,pwdinputflag,ifaxmode,transsvcstr1,transsvcstr2,ifaxdirectmode,documenttype,bwpapersize,bwcompressiontype,bwpixeltype,bwbitsperpixel,bwresolution,clpapersize,clcompressiontype,clpixeltype,clbitsperpixel,clresolution,accesscode,uuid,cnreadlang,enablesfp,memberobjectuuid,loginusername,logindomainname,usergroupname,personalid,"""
    )

    output = subprocess.check_output(
        f"chcp 65001 | powershell -ExecutionPolicy Unrestricted C:\\ad_to_printer\\getAdGroupMembers {ad_group}",
        shell=True,
        stderr=subprocess.STDOUT,
    )

    workers = []
    workers.append({})
    i = 0
    for line in output.split(b"\r\n"):
        line = line.decode("utf-8")
        line_splitted = line.split(":")
        if len(line_splitted) == 2:
            type, value = line_splitted
            type = type.strip()
            value = value.strip()
            if type == "EmailAddress" or "GivenName" or "Surname":
                workers[i][type.lower()] = value

            if type == "UserPrincipalName":
                i += 1
                workers.append({})

    workers.pop(-1)
    for i, worker in enumerate(workers):
        printer_csv += f'\nemail,"{worker["givenname"]} {worker["surname"]}","{worker["givenname"]}",,{i},"{worker["emailaddress"]}",,,,,smtp,,,,{201 + i},off,,,,,,,,,,,,,,,,,,,,,,,,,0,,"{i+1}",,,,,,,'

    with open(filename, "w", encoding="utf-8") as file:
        file.write(printer_csv)

    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options)

    driver.get(web_address)

    if enter_password:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "USERNAME"))
        ).send_keys("Administrator")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "PASSWORD_T"))
        ).send_keys(password)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div/div/form[1]/div[2]/div/div/div[3]/fieldset/button",
                )
            )
        ).click()

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".Standby.App1_4"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(), 'Datenverwaltung')]")
        )
    ).click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Adresslisten')]"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "InportButton"))
    ).click()

    Select(driver.find_element_by_id("AID")).select_by_value(str(address_book_no))

    Select(driver.find_element_by_id("AMOD")).select_by_value("2")

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ACLS3"))
    ).click()
    file_upload = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "AFILE"))
    )
    file_upload.send_keys(filename)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ENC_MODE_D"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/form[5]/div/div[2]/div[2]/div/div[1]/div/div[4]/fieldset/input",
            )
        )
    ).click()

    driver.switch_to.alert.accept()

    driver.close()
