from printer import ad_to_printer

"""
Args:
    filename        (str): Filename for csv. (not really important)
    ad_group        (str): Active Directory group to extract.
    web_address     (str): IP or web address of the Canon printer webinterface (with port number).
    address_book_no (int): html number of the address book.
    enter_password  (boolean): Defaults to False.
    password        (str): Defaults to 7654321
"""

ad_to_printer("abook.csv", "really-cool-and-exclusive-ad-group", "192.168.0.168:8000", 1, enter_password=True, password="strongest_password_ever")
