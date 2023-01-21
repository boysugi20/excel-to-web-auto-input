import os
import random
import time
import urllib.request

import requests
import stem.connection
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import logging

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException

from stem import Signal
from stem.control import Controller
from stem.util.log import get_logger

# from selenium import webdriver
from seleniumwire import webdriver

import psutil

logger = logging.getLogger("seleniumwire")
logger.setLevel(logging.ERROR)

class Proxy():
    def __init__(self, proxy_list):

        self.proxy_list = []
        for line in proxy_list:
            split = line.split(":")
            self.proxy_list.append(
                split[2] + ":" + split[3] + "@" + split[0] + ":" + split[1]
            )

        self.proxy_status = {proxy: "free" for proxy in self.proxy_list}

    def get_proxies(self):

        return self.proxy_status

    def get_free_proxy(self):

        temp = []
        for k, v in self.proxy_status.items():
            if v == "free":
                temp.append(k)

        if len(temp) == 0:
            time.sleep(5)
            for k, v in self.proxy_status.items():
                if v == "free":
                    temp.append(k)

        proxy = random.choice(temp)

        return proxy

    def set_proxy_busy(self, proxy):

        self.proxy_status[proxy] = "busy"

        return

    def set_proxy_free(self, proxy):

        self.proxy_status[proxy] = "free"

        return

    def set_proxy_busy_all(self):
        for k, v in self.proxy_status.items():
            self.proxy_status[k] = "busy"

        return

    def set_proxy_free_all(self):
        for k, v in self.proxy_status.items():
            self.proxy_status[k] = "free"

        return

    def set_proxy_sleep(self, seconds):
        self.set_proxy_busy_all()
        time.sleep(seconds)
        self.set_proxy_free_all()

        return


def driver_quit(driver):

	# get the process ID of your driver
	p = psutil.Process(driver.service.process.pid)

	# recursively get the process ID of the children
	proc_children = p.children(recursive=True)

	try:
		# send a term signal to the driver itself
		os.kill(driver.service.process.pid, signal.SIGTERM)

		# for each sub process, send a term signal
		for proc in proc_children:
			os.kill(proc.pid, signal.SIGTERM)

		del driver.requests

	except Exception as e:
		pass


# Signal TOR for a new connection
def renew_connection(show_logs=False):
    """
    Renews tor connection (change IP)
    """
    # Silence logs (annoying)
    if not show_logs:
        logger = get_logger()
        logger.propagate = False

    with Controller.from_port(port=9051) as controller:
        controller.authenticate(password=os.getenv("TOR_PASS"))
        controller.signal(Signal.NEWNYM)
        time.sleep(5)
        time.sleep(controller.get_newnym_wait())


def get_tor_session(username=None, password=None):
    # Tor uses the 9050 port as the default socks port
    session = requests.session()

    if username is not None and password is None:
        temp = random.randint(10000, 0x7FFFFFFF)
        session.proxies = {
            "http": "socks5://" + str(username) + ":" + str(temp) + "@127.0.0.1:9050",
            "https": "socks5://" + str(username) + ":" + str(temp) + "@127.0.0.1:9050",
        }
    elif username is None and password is not None:
        temp = random.randint(10000, 0x7FFFFFFF)
        session.proxies = {
            "http": "socks5://" + str(temp) + ":" + str(password) + "@127.0.0.1:9050",
            "https": "socks5://" + str(temp) + ":" + str(password) + "@127.0.0.1:9050",
        }
    elif username is not None and password is not None:
        session.proxies = {
            "http": "socks5://"
            + str(username)
            + ":"
            + str(password)
            + "@127.0.0.1:9050",
            "https": "socks5://"
            + str(username)
            + ":"
            + str(password)
            + "@127.0.0.1:9050",
        }
    else:
        session.proxies = {
            "http": "socks5://127.0.0.1:9050",
            "https": "socks5://127.0.0.1:9050",
        }

    return session


# Generate Service object
def get_chrome_driver_service(path):
    """
    Returns Service object to be used when creating chrome driver instances.

                    Parameters:
                                    path (str): Path to chromedriver

                    Returns:
                                    Service (obj): object for creating driver instances.
    """

    service = Service(path)

    return service


# Generate fake user agent
def get_fake_ua():
    """
    Returns random user agent from fake-useragent library

    Returns: user_agent (str): random user agent.
    """

    user_agent = UserAgent().random

    return user_agent


# Create a driver instance
def get_driver(chromeDriverPath, ipCheckLink=None, user_agent=None, proxy=None, headless=True):
    """
    Returns chrome driver

                    Parameters:
                                chromeDriverPath (str): ChromeDriver path
                                user_agent (obj): user agent
                                proxy (str): proxy to use (url or 'tor')
                                headless (bool): headless or not

                    Returns:
                                    driver (obj): chrome driver instances
                                    driver_info (dict): user_agent, ip
    """

    options = webdriver.ChromeOptions()

    if headless:
        options.add_argument("--headless")

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--incognito")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if proxy is not None:
        sw_options = {
            "proxy": {
                "http": "http://" + str(proxy),
            }
        }
    elif proxy == "tor":
        proxy = "socks5://localhost:9050"
        options.add_argument("--proxy-server=" + str(proxy))
    else:
        sw_options = {"proxy": {"no_proxy": "localhost,127.0.0.1"}}

    if user_agent is not None:
        options.add_argument("user-agent=" + str(user_agent))
    else:
        try:
            user_agent = get_fake_ua()
            options.add_argument("user-agent=" + str(user_agent))
        except Exception as e:
            pass

    # Try passing service, else pass path
    try:
        driver = webdriver.Chrome(
            service=get_chrome_driver_service(chromeDriverPath),
            options=options,
            seleniumwire_options=sw_options,
        )
    except:
        try:
            driver = webdriver.Chrome(
                chromeDriverPath, options=options, seleniumwire_options=sw_options
            )
        except:
            raise Exception("[WARNING] Failed to create chrome session (check version)")

    if ipCheckLink is not None:
        ip = check_ip(driver=driver, ipCheckLink=ipCheckLink)
    else:
        ip = check_ip(driver=driver)

    driver_info = {"user_agent": user_agent, "ip": ip}

    return driver, driver_info


# Check driver IP
def check_ip(driver, ipCheckLink="https://ident.me"):
    """
    Returns ip address used by specified driver instance

                    Parameters:
                                    ipCheckLink (str): URL to check ip Address
                                    driver (obj): Chrome driver

                    Returns:
                                    ip (str): ip address
    """

    driver.get(ipCheckLink)
    check_ip_html = driver.page_source

    soup = BeautifulSoup(check_ip_html, "lxml")

    ip = str(soup.text)

    ip = ip.strip()

    return ip


# Click button
def click_button(driver, wait_time, xpath):

    button = WebDriverWait(driver, wait_time).until(
        ec.element_to_be_clickable((By.XPATH, xpath))
    )
    button.click()

    return


# Click item from dropdown
def click_dropdown_item(driver, xpath, index):

    items = driver.find_elements_by_xpath(xpath)
    action = ActionChains(driver)
    action.move_to_element(items[index]).perform()

    return


# Input text
def input_text(driver, wait_time, xpath, text, clear=False, enter=False):

    text_input = WebDriverWait(driver, wait_time).until(
        ec.visibility_of_element_located((By.XPATH, xpath))
    )

    if clear:
        text_input.send_keys("\ue009" + "\ue003")

    text_input.send_keys(text)

    if enter:
        text_input.send_keys(Keys.ENTER)

    return


# Check if element exist
def check_element(driver, xpath):
    try:
        driver.find_element(by=By.XPATH, value=xpath)
    except NoSuchElementException:
        return False
    return True
