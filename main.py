from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from tkinter import *
import time
from datetime import date
import mmap
from random import randint
import csv
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.font as font
import os
import ntpath

error_occurred = False
filename = None
row_count = 0
total_lines = 0
flipkart_val = 0
amazon_val = 0
directory = None
output_file = None


def UploadAction(event=None):
    global directory
    directory = filedialog.askdirectory()


def mapcount(filename):
    f = open(filename, "r+", encoding='utf8')
    buf = mmap.mmap(f.fileno(), 0)
    lines = 0
    readline = buf.readline
    while readline():
        lines += 1
    return lines


def read_csv(file):
    isbn_list = []
    with open(file, encoding='utf8') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                pass
            else:
                isbn_list.append(row[1])
            line_count += 1
    return isbn_list


def amazon_scrape(filename):
    global row_count
    global total_list
    global output_file
    options = Options()

    found = False
    listed = "Not Listed"
    price_tag = "NA"
    options.headless = False

    binding = "NA"
    per_page_ratio = "NA"
    # options.add_argument('headless')
    options.add_argument('--no-sandbox')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_driver_path = r"chromedriver.exe"
    driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)
    driver.minimize_window()
    original_link = 'https://www.amazon.in/'
    base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
    base_file_path = os.path.join(output_file, base_name)
    writing_file = open(base_file_path, 'a', encoding='utf8', newline='')
    amazon_flip_writer = csv.writer(writing_file, delimiter=',')
    amazon_flip_writer.writerow(
        ["Date", "ISBN13","Prime",'Ranking',"Buybox",'Buybox Seller','Buybox Price',"Repro Listing"])
    writing_file.close()
    csv_file = pd.read_csv(filename)
    isbn_list = csv_file['ISBN13']
    line_count = 1
    for i in range(len(isbn_list)):
        found = False
        listed = "Not Listed"
        # price_tag = "NA"
        # start = time.time()
        # pages = "NA"
        current_isbn = str(isbn_list[i])
        rank = "NA"
        buybox = "No"
        buybox_seller = "NA"
        buybox_price = "NA"
        prime = "Not Prime"
        final_shipping = 0
        ## stack = [original_link]
        # curr_per_page_ratio = "NA"
        # count = 0
        # page = 1
        # time.sleep(randint(1, 5))
        try:
            driver.get('https://www.amazon.in/s?k='+current_isbn)
            # driver.get(stack.pop())
            # time.sleep(3)
            # isbn_field = driver.find_element_by_id('twotabsearchtextbox')
            if True:
            #     isbn_field.send_keys(str(current_isbn))
            #     isbn_field.send_keys(Keys.RETURN)
            #
            #     while driver.execute_script('return document.readyState;') != "complete":
            #         pass

                time.sleep(2)
                try:
                    book_box = driver.find_element_by_css_selector("div[class='sg-col-4-of-12 sg-col-8-of-16 sg-col-12-of-20 sg-col']")
                    prime_status = book_box.find_element_by_css_selector("span[class='aok-relative s-icon-text-medium s-prime']")
                    prime="Prime"
                except:
                    pass
                search_results = driver.find_elements_by_css_selector("a[class='a-size-base a-link-normal']")
                all_reference_links = []
                for search_result in search_results:
                    link = search_result.get_attribute("href")
                    if 'Paperback' in search_result.text or 'Hardcover' in search_result.text:
                        all_reference_links.append(link)
                other_results = driver.find_elements_by_css_selector(
                    "a[class='a-size-base a-link-normal a-text-bold']")
                for other_result in other_results:
                    link = other_result.get_attribute("href")
                    if 'Paperback' in other_result.text or 'Hardcover' in other_result.text:
                        all_reference_links.append(link)
                first_link = True
                while all_reference_links:
                    try:
                        curr_found = False
                        curr_listed = "Not Listed"
                        # price_tag = "NA"
                        # pages = "NA"
                        driver.get(all_reference_links.pop(0))
                        time.sleep(2)
                        if prime != 'Prime':
                            shipping = str(driver.find_element_by_css_selector('span[class="a-color-base buyboxShippingLabel"]').text)
                            temp_shipping = ''
                            for i in shipping:
                                if i.isdigit() or i == '.':
                                    temp_shipping += i
                            final_shipping = float(temp_shipping)
                        the_details_id = driver.find_element_by_id('detailBullets_feature_div')
                        span_elements = the_details_id.find_elements_by_class_name('a-list-item')
                        for span_element in span_elements:
                            if 'Sellers Rank' in span_element.text:
                                temp_rank = str(span_element.text).split('\n')[0].replace('See Top 100', '')
                                rank = ""
                                for i in temp_rank:
                                    if i.isdigit():
                                        rank += i
                                break
                        # if first_link:
                        #     first_link = False
                        #     try:
                        #         the_details_id = driver.find_element_by_id('detailBullets_feature_div')
                        #         span_elements = the_details_id.find_elements_by_class_name('a-list-item')
                        #         for span_element in span_elements:
                        #             other_spans = span_element.find_elements_by_tag_name('span')
                        #             for other_span in other_spans:
                        #                 if 'pages' in other_span.text or 'Pages' in other_span.text:
                        #                     pages = other_span.text
                        #     except NoSuchElementException:
                        #         pass
                        # temp_binding = ""
                        # try:
                        #     binding = driver.find_element_by_id('productSubtitle')
                        #     if 'Hardcover' in binding.text:
                        #         temp_binding = 'Hardcover'
                        #     elif 'Paperback' in binding.text:
                        #         temp_binding = 'Paperback'
                        # except NoSuchElementException:
                        #     pass
                        seller_name = None
                        try:
                            # seller_name = driver.find_element_by_id('sellerProfileTriggerId')
                            delay = 5
                            seller_name = WebDriverWait(driver, delay).until(
                                EC.presence_of_element_located((By.ID, 'sellerProfileTriggerId')))
                        except Exception as e:
                            with open(ntpath.basename(filename)[:-4]+"amazon_error_log.csv", 'a',newline='') as error_file:
                                error_writer = csv.writer(error_file)
                                error_writer.writerow(['Amazon', current_isbn])
                        # try:
                        #     mrp_prices = driver.find_elements_by_class_name('a-list-item')
                        #     for mrp_price in mrp_prices:
                        #         if 'M.R.P.:' in mrp_price.text:
                        #             price_tag = str(mrp_price.text)
                        #             price_tag = price_tag.replace('M.R.P.:','')
                        # except e:
                        #     pass

                        # try:
                        #     selling_price = driver.find_element_by_css_selector(
                        #         "span[class='a-size-medium a-color-price inlineBlock-display offer-price a-text-normal price3P']")
                        # except NoSuchElementException:
                        #     pass
                        # if selling_price and price_tag == "NA":
                        #     price_tag = selling_price.text

                        # pages = pages.replace('pages', '').replace('Pages', '').strip()
                        # price_tag = price_tag.replace('₹', '').strip()
                        # try:
                        #     curr_per_page_ratio = round(float(price_tag) / float(pages), 2)
                        # except:
                        #     pass
                        buybox_seller = str(seller_name.text)
                        buybox_price = str(driver.find_element_by_id('soldByThirdParty').text).replace('₹','').strip()
                        buybox_price = float(buybox_price) + final_shipping
                        if seller_name and 'repro' in str(seller_name.text).lower():
                            curr_found = True
                            found = True
                            buybox = "Yes"
                            if curr_found:
                                curr_listed = 'Listed'
                            # selling_price = None

                            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                            base_file_path = os.path.join(output_file, base_name)
                            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                            amazon_flip_writer.writerow(
                                [date.today(), current_isbn] + [prime,rank,buybox,buybox_seller,buybox_price,curr_listed])
                            writing_file.close()
                            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn,prime,rank,buybox,buybox_seller,buybox_price,curr_listed)
                        else:
                            try:
                                other_sellers = driver.find_element_by_id('mbc-olp-link')
                                a_tag = other_sellers.find_element_by_tag_name("a")
                                # href_link = a_tag.get_attribute("href")
                                a_tag.click()
                                # driver.get(href_link)
                                time.sleep(3)
                                try:
                                    delay = 5
                                    WebDriverWait(driver, delay).until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, "a[class='a-size-small a-link-normal']")))
                                except Exception as e:
                                    with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                              newline='') as error_file:
                                        error_writer = csv.writer(error_file)
                                        error_writer.writerow(['Amazon', current_isbn])
                                more_pages = True
                                while more_pages:
                                    more_pages = False
                                    search_results = driver.find_elements_by_css_selector(
                                        "a[class='a-size-small a-link-normal']")
                                    seller_num = -1
                                    for search_result in search_results:
                                        seller_num += 1
                                        seller_name = search_result.text
                                        if 'repro' in str(seller_name).lower():
                                            curr_found = True
                                            curr_listed = 'Listed'
                                            found = True
                                            break
                                    # selling_prices = driver.find_elements_by_css_selector("span[class='a-size-large a-color-price olpOfferPrice a-text-bold']")
                                    if curr_found:
                                        #     curr_sell_num = -1
                                        #     for selling_price in selling_prices:
                                        #         curr_sell_num += 1
                                        #         if curr_sell_num == seller_num:
                                        #             price_tag = selling_price.text
                                        #             break
                                        #     pages = pages.replace('pages', '').replace('Pages', '').strip()
                                        #     price_tag = price_tag.replace('₹', '').strip()
                                        #     try:
                                        #         curr_per_page_ratio = round(float(price_tag) / float(pages),2)
                                        #     except:
                                        #         pass
                                        base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                                        base_file_path = os.path.join(output_file, base_name)
                                        writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                                        amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                                        amazon_flip_writer.writerow(
                                            [date.today(), current_isbn] + [prime, rank, buybox, buybox_seller,
                                                                            buybox_price, curr_listed])
                                        writing_file.close()
                                        print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn, prime,
                                              rank, buybox, buybox_seller, buybox_price,curr_listed)
                                    try:
                                        next_button = driver.find_element_by_css_selector('li[class="a-last"]')
                                        next_button_link = next_button.find_element_by_tag_name('a')
                                        href_tag = next_button_link.get_attribute('href')
                                        more_pages = True
                                        driver.get(href_tag)
                                    except NoSuchElementException:
                                        pass

                            except NoSuchElementException:
                                pass
                            except Exception as e:
                                with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                          newline='') as error_file:
                                    error_writer = csv.writer(error_file)
                                    error_writer.writerow(['Amazon', current_isbn])
                            if curr_found == False:
                                try:
                                    new_sellers = driver.find_element_by_id("buybox-see-all-buying-choices-announce")
                                    new_sellers_link = new_sellers.get_attribute("href")
                                    driver.get(new_sellers_link)
                                    time.sleep(2)
                                    try:
                                        delay = 5
                                        WebDriverWait(driver, delay).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, "span[class='a-size-medium a-text-bold']")))
                                    except Exception as e:
                                        with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                                  newline='') as error_file:
                                            error_writer = csv.writer(error_file)
                                            error_writer.writerow(['Amazon', current_isbn])
                                    more_pages = True
                                    while more_pages:
                                        more_pages = False
                                        search_results = driver.find_elements_by_css_selector(
                                            "span[class='a-size-medium a-text-bold']")
                                        seller_num = -1
                                        for search_result in search_results:
                                            seller_num += 1
                                            seller_name = search_result.text
                                            if 'repro' in str(seller_name).lower():
                                                curr_found = True
                                                curr_listed = 'Listed'
                                                found = True
                                                break
                                        # selling_prices = driver.find_elements_by_css_selector("span[class='a-size-large a-color-price olpOfferPrice a-text-bold']")
                                        if curr_found:
                                            #     curr_sell_num = -1
                                            #     for selling_price in selling_prices:
                                            #         curr_sell_num += 1
                                            #         if curr_sell_num == seller_num:
                                            #             price_tag = selling_price.text
                                            #             break
                                            #     pages = pages.replace('pages', '').replace('Pages', '').strip()
                                            #     price_tag = price_tag.replace('₹', '').strip()
                                            #     try:
                                            #         curr_per_page_ratio = round(float(price_tag) / float(pages),2)
                                            #     except:
                                            #         pass
                                            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                                            base_file_path = os.path.join(output_file, base_name)
                                            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                                            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                                            amazon_flip_writer.writerow(
                                                [date.today(), current_isbn] + [prime, rank, buybox, buybox_seller,
                                                                                buybox_price, curr_listed])
                                            writing_file.close()
                                            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn,
                                                  prime, rank, buybox, buybox_seller, buybox_price,curr_listed)
                                        try:
                                            next_button = driver.find_element_by_css_selector('li[class="a-last"]')
                                            next_button_link = next_button.find_element_by_tag_name('a')
                                            href_tag = next_button_link.get_attribute('href')
                                            more_pages = True
                                            driver.get(href_tag)
                                        except NoSuchElementException:
                                            pass
                                except:
                                    pass
                    except NoSuchElementException:
                        pass
                    except Exception as e:
                        with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a', newline='') as error_file:
                            error_writer = csv.writer(error_file)
                            error_writer.writerow(['Amazon', current_isbn])
        except Exception as e:
            with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a', newline='') as error_file:
                error_writer = csv.writer(error_file)
                error_writer.writerow(['Amazon', current_isbn])
        if not found:
            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
            base_file_path = os.path.join(output_file, base_name)
            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
            amazon_flip_writer.writerow(
                [date.today(), current_isbn] + [prime, rank, buybox, buybox_seller, buybox_price, curr_listed])
            writing_file.close()
            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn, prime, rank, buybox, buybox_seller,
                  buybox_price,"Not Listed")
        line_count += 1
    driver.close()
    driver.quit()
    return base_file_path


def setup_ui():
    global row_count
    global output_file
    global total_lines
    global flipkart_val
    global amazon_val
    root = tk.Tk()
    root.geometry("600x200")
    root.title("Select Your Folder To Check Whether The Titles are Listed in Amazon and Flipkart")
    select = tk.Label(root)
    select.pack()
    myFont = font.Font(size=20)
    button = tk.Button(root, bg='yellow', text='Select Folder', command=UploadAction, font=myFont)
    button.configure(width=600, height=2)
    button.pack()
    label = tk.Label(root)
    label.pack()

    def task():
        if directory:
            select.config(text="You have successfully selected the folder you can now close this window")
            label.config(text="Current Selected Folder Path: " + str(directory))
        root.after(1000, task)

    root.after(0, task)
    root.mainloop()
    if directory:
        output_file = os.path.join(directory, 'Output')
        try:
            os.mkdir(output_file)
        except FileExistsError:
            pass
        for name in os.listdir(directory):
            if name[-5:] == '.xlsx' or name[-4:] == '.xls':
                global total_records

                filename = os.path.join(directory, name)
                excel_reader = pd.read_excel(filename)
                filename = os.path.join(directory, name[:-5] + '.csv')
                excel_reader.to_csv(filename, index=None, header=True)
                total_lines = mapcount(filename) - 1
                print("Processing the file now")
                file_path = amazon_scrape(filename)
                read_file = pd.read_csv(file_path)
                if os.path.exists(file_path[:-4] + '.xlsx'):
                    existing_file = pd.read_excel(file_path[:-4] + ".xlsx")
                    read_file = pd.concat([existing_file, read_file], ignore_index=True)
                    with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                        read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                else:
                    with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                        read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                os.remove(file_path)

                os.remove(filename)
                print("All Files Processed")
        for name in os.listdir(os.path.dirname(sys.argv[0])):
            if name[-4:] == '.csv':
                convert_to_excel = pd.read_csv(name)
                with pd.ExcelWriter(name[:-4] + '.xlsx', mode='w') as writer:
                    convert_to_excel.to_excel(writer, sheet_name='Sheet1', index=None, header=True)




setup_ui()
