import time
import csv
import uuid
import os
import tempfile
import pandas as pd
from urllib3.exceptions import ReadTimeoutError
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Configure Edge options
edge_options = Options()
edge_options.add_argument("start-maximized")
edge_options.add_experimental_option("detach", True)
# Add a unique user-data-dir so each session uses a new profile
temp_profile = tempfile.mkdtemp()
edge_options.add_argument(f"--user-data-dir={temp_profile}")

# Launch Edge browser
driver = webdriver.Edge(options=edge_options)
driver.set_page_load_timeout(180)

# List of product URLs to scrape
product_urls = [
  "https://www.olliix.com/product/data/3932/FPF18-0160?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/13088/MP130-0945?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/14175/MP130-1036?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/13685/II130-0406?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/3734/MP130-0373?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/15180/MP130-1207?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/14013/MPS130-0306?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/15690/CH130-1011?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/15702/CH130-1012?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/4023/FPF18-0419?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/15701/CH130-1001?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/13053/MP130-0929?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/15700/CH130-1000?ia=0&c=21&o=3&i=1",
  "https://www.olliix.com/product/data/7816/MP130-0156?ia=0&c=21&o=3&i=1"
]

# Set cookies (if needed for all pages)
cookies = [
    {"name": ".AspNetCore.Antiforgery.6jtYfgA0IZU", 
     "value": "CfDJ8Owt0b9-lC5Evmhcb8Y8-dsjNx6l2z_7Ab5oG95b_HlHkq0TSrSvYasDzTy7848AsnfrP48CLMZhiEaHFuY4KF4QwUtzPLijhxXS6CwOOQAvI3VldG3HtA2pEwlIFlTGyvPodsaXte4tJOutPjyLji4", 
     "domain": ".olliix.com"},
    
    {"name": ".AspNetCore.Cookies", 
     "value": "CfDJ8Owt0b9-lC5Evmhcb8Y8-dsoq1J7HbEPU0WdbpIzXFTyMIdZHORqIEwcZ4QHYPHYkGNHXmHrIKR86SoUSYRLad4FR46WC1K6QyHhWE3xtx19cGToQ-p7cDmo9VePfgmvYqmMCYCE0vLu1l0Up4QreQmYQ8IG5VwMZdfmhQStp36QC0L2k0V3sqbFpbdD48X8h20bIblyTWGgtiCjPtNlw5J8m56W76uk80N90eUkQ5VlNodPbHgu1ic0XCjF4nTHnQq5hGZOgHUDMIkpva9IBHfp00ZPrrV3Hk8e5XNAUEpzsbn7V3OYhZiM1xQDHLEBXJ2KQwAVr2umWiXOQzQOc8srfCzt5jGGjgVdLxfZwPnyEIAD25E7a59Fq1ofFeQBZYbvD5oloaQn5XYpt7T3IDxD9INKWRZWd5m007uHsvmf3eWIv17ojKAtvzLVlZT0wGk-qDEc6p1CtLRuwktb_WMM1--U5GZE2xhrwVcCH-jRfSgdkq-PB33kDEPc_SqJHpTShrNtjdhCzXQDFeOXVA3tGjzxV3tBtcI51oWsx0sCE4ScIAxX3PhOgC0CNlYG9VdRtYfBprIrzWmUuoCwCd_GXPf3PNQqPe3P1mvApwQCiuIV-siTc1ppgsS9ky7NJKll3IpT5pqSgI1C3NoEDKwkXYIVkPqVhFBFwrRpxoFYwOSl6Icjju_wDvWnTQIuN94WWmrEZ7bQX6Jz1ssqmWcfee7q1gK7dsEyZL9TJtM5kKac---SW_m9A8cR5K51ObI-LN_5BRLh3l-RvaZQbpE", 
     "domain": ".olliix.com"},
    
    {"name": ".AspNetCore.Session", 
     "value": "CfDJ8Owt0b9%2BlC5Evmhcb8Y8%2Bdttb6F62Jl1VQ7qEAatzFoWozbkygpxUhwneiah6b%2B0hhVMvHvbElx5zrDB%2FpztdnLRMUIjHQ43liLOwoQDFzyiATE1iPw3lbfb%2FTFTDDds6xNat8gtafssQBl34USICCjKrCXzgfYRoWkK5FBTlcp5", 
     "domain": ".olliix.com"},
    
    {"name": "atuh.cookie2", 
     "value": "CfDJ8Owt0b9-lC5Evmhcb8Y8-dsH9KBKaMYzClrPf2orxqOQMLdJNcbw3JLYcHUkI-kw4LkqAqyjwj4kZr0HlDk-35Wx3SgPZTSLUhYzFie5ZIhPxhoeZuxGiV-Or-FeFLb9GEABrc0xptFulw3-qwbOxN_voXib45_N8Z4iRoLxVIZZJZoSEVP5bDz183N671UwkCm5exSQlhBGKtHXPiMWZx7qE0ARZOz_1AB9L22kqpgZwTdJpBkbIun-xKLDLJGKo17pDDbhyBScA0ZFrgBtOyIf8tnBa2jh6oK9rDmovQd3Z9Gq4d6yVaDDw1eI9ZzbjkDaVCaClvP3Jg6KTyYK5Nhmjr-Mh5A08F1PTmtg1o1f-i5W8p52ue41NoEKtTBLPJj5HqJ45EMU_pYSCr1oX5avl5Vlf3cmbTjIQjIs6eeYXc8A60cvivB0wRGBtXllOv2xmPtoWVvprBRjrmYU0dshrYRPaMNo4y_mlcwF7lwXT0gW0FbPMy3Adv2I_9b2y9NoC3S0ArUqNCRE7oqeTckQr268FP52GgoNul26OknGrP-wFC8_w8Lom_OOIsnPfBUvxV7AmmR654u4Dps5oOSaqeFla05YgQFWdLfYPN1jWoO5NC9oLVfHwJ1rOFt7_go9P2DTvUbxAfSy7gy8tfsgQ2oxcbIIdV6D3kDrJVo3npmSHGcAC-9FFx25cZFX3YcvOHQY7H1TeXhygRzHaboFDog7BCICn6Op4k6Sm_FtXcDkSbrIQeDXcpRVWGQM56p6N51TrIsWWhIYR-QTCLjS45oAnfw03LZCB2VnXuZoZchIAtLr4YbS_5p9UdaIgYcqldZ8MAHjXEti6PpTWvJmSyzdE-JLJsHHtnqH3n3i826wSX0vko_l2D03SbD_lTe--w7jmK88tngMvM-rbyGr8ONaFpSQfU8WcrSyHiTs1YqLUqotYLEW0LzcPVpc5Dsmq4L7gDANPjB0MEfugIiMC6pimjN38NcEYmoTDPF---462K8eufVeXAQ5PxT6btjkJDhW3vksEMDypWn49lhO6oHFY6PS_rmwoUfa3t2UO9gATGf5HJ204vjNSLSP9eg579h1IbAxn5RMzRzumMQSvKMu6nZJkCJzR9fk554X4uFePvNlvMBElcF0JC0vTnOIJSY1GwhMAqvyEEJriCPp_I2gyJP5CyF2dGfMRTQUrDH4JmdbcgJbVrCgJENWGtQzeBx2QAwWbs28z8OC_KyQHMBEMa2WI6iQ08lbifb7OQf2D_8ePybxNM_w_9djdiRsVC279FJVQSG8_iqozGAtVyJEc2FCqrjFqwEf-8gXPBFd8EAAKdlCHUsa0OKvvXPMbT4ORJgFaf7q7v2A7kZgaYu0mseI6caNZAkL9ZQDnEV7beFEFj0Yot7D2rjIKYG5gt5k7Wbwlr_Mr0CW1SqeHI_rzfMTYV2ejSE5eKYvJ1G9tzRR4cAkKMGxbT--gVKrvV_Jv3KpZotdQ3G0cxkH1Aijs3N8SgLPRwFN6cJ0xfzbSpCMbjJ6ARM3T79T6k_dQmMcsr1qABC2YLbl1XxoOn4OaV997a-MOFQXdyxjKva8sz2nZ_14Ha4LzI7hLNmkgbCGkqBgR8KmkTxt5UwNg2mbd1CpP7OrUK60kl6PGdPtEaD1q9GAQeLYPznTRKOpFA0yclWiIojtTTfOB-NrylwANKOY-Iw687bG1ChWYy5qt4RktpbHaXgVQaBjZVuhR0GNc_N_qxu7PT0oxBL__qELHxtpei7NEtYRCuRopD802vhsaQA0mh8OeIYmHOAuxkDhDxrgxv695KMhO1EbYPgqlrird5lKcgcIciAl9tFBOqsemTVr9Xs5XvrcaZk26bng-6j_ZK22hFrtTHYfh_xZNSOPhj9PJXj7lD1M9QY9GUxfh2DJJCDkRgXqOv8vAguB_oJplPwSIeJzxl20wMDku8RbVXEkXisIrNOoHi9FMJRNufE2Sy_-gmLL5K6zvlvvQUsGM1ubFEzmBNR2aelSC-vGyE7MYMq_6VP6dvTFEH2sjdoyPL8xX0jf5HUO6uaafVE5ZM-jSKUpSTBSaWrEyCZHc4YfOfAt-x1dAkbODHdemrT9AMQmmsWkBA4XkbDv4TNJvyjWjPQroB3vaUdkbWMbmDy5Z2NgocKJcAXkKdOPbyiUD0n5CKhdNcLWthfZubYcbM1DqFav1eo"}
]

driver.get("https://www.olliix.com")
for cookie in cookies:
    driver.add_cookie(cookie)
    print("Cookie set:", cookie["name"])

# Refresh once to apply cookies (if necessary)
driver.refresh()
time.sleep(5)  # Wait for login/session activation

# Global variables to control size output for each product
current_size_present = False
current_size_name = ""

# Open CSV file once for writing output
csv_filename = "olliix_products.csv"
with open(csv_filename, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    header = [
        "Handle", "Title", "Variation", "Vendor", "Tags", 
        "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value",
        "Variant SKU", "Variant Inventory Tracker", "Variant Grams", 
        "Variant Inventory Qty", "Variant Inventory Policy", 
        "Variant Fulfillment Service", "Variant Price", 
        "Variant Requires Shipping", "Variant Taxable", 
        "Variant Weight Unit", "Gift Card", 
        "Google Shopping / Condition", "Status", 
        "Image Src", "Image Position", "Variant Image", 
        "Variant item no.", "Variant UPC", "variant pack qty", 
        "Variant shipping method", "Variant Description", 
        "variant product detail", "variant features", 
        "variant material", "variant care instruction", 
        "variant shipping dimension", "variant assembly", 
        "variant breadcrumb"
    ]
    writer.writerow(header)

    # Define the scrape_variant() function
    def scrape_variant():
        try:
            # Wait for the active variation name
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//li[contains(@class, 'actived')]//span[@data-bind='text: Value']"))
            )
            variation_name_element = variation_lists[1].find_element(
                By.XPATH, ".//li[contains(@class, 'actived')]//span[@data-bind='text: Value']"
            )
            variation_name = variation_name_element.text.strip()
            print(f"Scraping variation: {variation_name}")
            time.sleep(3)  # Allow UI updates

            # Get the active image element (variant image)
            try:
                active_image_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img"))
                )
                variant_image = active_image_element.get_attribute("src")
                if variant_image:
                    global_variant_images.append(variant_image)
            except Exception:
                variant_image = ""
            print(f"Active image: {variant_image}")

            try:
                price_element = driver.find_element(By.XPATH, "//span[@data-bind='price:SuggestedRetail']")
                price = price_element.text.strip()
            except Exception:
                price = ""
            print(f"Price: {price}")

            try:
                shipping_element = driver.find_element(
                    By.XPATH, "//div[@class='property-con' and contains(@data-bind, 'OnlyByLTL')]"
                )
                shipping_method = shipping_element.text.strip()
            except Exception:
                shipping_method = ""
            print(f"Shipping: {shipping_method}")

            # Scrape all images from the DownBox
            try:
                image_elements = driver.find_elements(By.XPATH, "//div[@id='DownBox']//img")
                image_urls = [img.get_attribute("src") for img in image_elements if img.get_attribute("src")]
                global_all_images.extend(image_urls)
                images = ", ".join(image_urls)
            except Exception:
                images = ""
            print(f"Images: {images}")

            try:
                variant_upc_element = driver.find_element(
                    By.XPATH, "//div[@class='property-con' and @data-bind='text: DisplayUPC']"
                )
                variant_upc = str(variant_upc_element.text.strip())  # Always convert to string

                # Convert possible scientific notation to a normal string
                variant_upc = str(int(float(variant_upc))) if "E" in variant_upc or "e" in variant_upc else variant_upc

                # Ensure Excel treats it as a string
                variant_upc = f"'{variant_upc}"
            except Exception:
                variant_upc = ""
            print(f"UPC: {variant_upc}")

            try:
                pack_qty_element = driver.find_element(
                    By.XPATH, "//div[@class='property-con' and @data-bind='text: PackQty']"
                )
                variant_pack_qty = pack_qty_element.text.strip()
            except Exception:
                variant_pack_qty = ""
            print(f"Pack Qty: {variant_pack_qty}")

            try:
                item_no_element = driver.find_element(
                    By.XPATH, "//div[@class='property-con' and @data-bind='text: DisplayItemNo']"
                )
                variant_item_no = item_no_element.text.strip()
            except Exception:
                variant_item_no = ""
            print(f"Item No: {variant_item_no}")

            variant_sku = str(uuid.uuid4())[:8]
            print(f"SKU: {variant_sku}")

            # One-time expansions of accordion sections
            nonlocal_variant_flags = {
                "variant_details_opened": scrape_variant.variant_details_opened,
                "features_opened": scrape_variant.features_opened,
                "material_opened": scrape_variant.material_opened,
                "care_opened": scrape_variant.care_opened,
                "shipping_opened": scrape_variant.shipping_opened,
                "assembly_opened": scrape_variant.assembly_opened
            }

            if not nonlocal_variant_flags["variant_details_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 2:
                        variant_detail_section = accordion_sections[1]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.variant_details_opened = True
                except Exception as e:
                    print(f"Error clicking variant details: {e}")

            if not nonlocal_variant_flags["features_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 3:
                        variant_detail_section = accordion_sections[2]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.features_opened = True
                except Exception as e:
                    print(f"Error clicking features: {e}")

            if not nonlocal_variant_flags["material_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 4:
                        variant_detail_section = accordion_sections[3]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.material_opened = True
                except Exception as e:
                    print(f"Error clicking material: {e}")

            if not nonlocal_variant_flags["care_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 5:
                        variant_detail_section = accordion_sections[4]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.care_opened = True
                except Exception as e:
                    print(f"Error clicking care: {e}")

            if not nonlocal_variant_flags["shipping_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 6:
                        variant_detail_section = accordion_sections[5]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.shipping_opened = True
                except Exception as e:
                    print(f"Error clicking shipping: {e}")

            if not nonlocal_variant_flags["assembly_opened"]:
                try:
                    accordion_sections = driver.find_elements(By.CLASS_NAME, "block.col-12.accordion-module")
                    if len(accordion_sections) >= 7:
                        variant_detail_section = accordion_sections[6]
                        toggle_button = variant_detail_section.find_element(By.TAG_NAME, "h2")
                        driver.execute_script("arguments[0].scrollIntoView();", toggle_button)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", toggle_button)
                        time.sleep(3)
                        scrape_variant.assembly_opened = True
                except Exception as e:
                    print(f"Error clicking assembly: {e}")

            # Grab all relevant HTML
            try:
                description_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "(//div[@class='accordion-content'])[1]"))
                )
                description_html = description_section.get_attribute("outerHTML").strip()
            except Exception:
                description_html = ""
            print(f"Description: {description_html}")
            time.sleep(3)

            try:
                variant_details_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content'][.//p[@data-bind='text: SizeName'] or .//p[contains(@data-bind, 'GetTextFront')]]"))
                )
                variant_details_html = variant_details_section.get_attribute("outerHTML").strip()
            except Exception:
                variant_details_html = ""
            print(f"Variant Details: {variant_details_html}")

            try:
                features_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content' and .//pre[@data-bind='html: Features']]"))
                )
                features_html = features_section.get_attribute("outerHTML").strip()
            except Exception:
                features_html = ""
            print(f"Features: {features_html}")

            try:
                material_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content'][.//p[@data-bind='html: Material']]"))
                )
                material_html = material_section.get_attribute("outerHTML").strip()
            except Exception:
                material_html = ""
            print(f"Material: {material_html}")

            try:
                care_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content'][.//p[@data-bind='text: CareInstructions']]"))
                )
                care_html = care_section.get_attribute("outerHTML").strip()
            except Exception:
                care_html = ""
            print(f"Care: {care_html}")

            try:
                shipping_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content'][.//table[@class='shipping-dimension-table'] or .//p[contains(text(),'Shipping')]]"))
                )
                shipping_html = shipping_section.get_attribute("outerHTML").strip()
            except Exception:
                shipping_html = ""
            print(f"Shipping: {shipping_html}")

            try:
                assembly_section = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='accordion-content' and @data-bind='foreach: AssemblyInstructions']"))
                )
                assembly_links = assembly_section.find_elements(By.TAG_NAME, "a")
                assembly_html = "; ".join([link.get_attribute("outerHTML") for link in assembly_links]) if assembly_links else ""
            except Exception:
                assembly_html = ""
            print(f"Assembly: {assembly_html}")

            option2_value = current_size_name if current_size_present else ""
            option2_name = "Size" if option2_value != "" else ""

            # Write a row for this variation
            writer.writerow([
                Handle,                           # Handle
                title,                            # Title
                variation_name,                   # Variation
                Vendor,                           # Vendor
                tags,                             # Tags
                "Color",                          # Option1 Name
                variation_name,                   # Option1 Value
                option2_name,                     # Option2 Name (empty if sizes missing)
                option2_value,                    # Option2 Value (empty if sizes missing)
                variant_sku,                      # Variant SKU
                "shopify",                        # Variant Inventory Tracker (placeholder)
                "",                               # Variant Grams
                "50000",                          # Variant Inventory Qty
                "deny",                           # Variant Inventory Policy
                "manual",                         # Variant Fulfillment Service
                price,                            # Variant Price
                "True",                           # Variant Requires Shipping
                "True",                           # Variant Taxable
                "",                               # Variant Weight Unit
                "False",                          # Gift Card
                "",                               # Google Shopping / Condition
                "active",                         # Status
                "",                               # Image Src
                "",                               # Image Position
                variant_image,                    # Variant Image
                variant_item_no,                  # Variant item no.
                variant_upc,                      # Variant UPC
                variant_pack_qty,                 # variant pack qty
                shipping_method,                  # Variant shipping method
                description_html,                 # Variant Description
                variant_details_html,             # variant product detail
                features_html,                    # variant features
                material_html,                    # variant material
                care_html,                        # variant care instruction
                shipping_html,                    # variant shipping dimension
                assembly_html,                    # variant assembly
                tags                              # variant breadcrumb
            ])

        except Exception as e:
            print(f"Error scraping variation: {e}")

    # Initialize one-time expansion flags for the accordion sections
    scrape_variant.variant_details_opened = False
    scrape_variant.features_opened = False
    scrape_variant.material_opened = False
    scrape_variant.care_opened = False
    scrape_variant.shipping_opened = False
    scrape_variant.assembly_opened = False

    # Loop over each URL to scrape products
    for product_url in product_urls:
        print(f"\n---\nScraping product URL: {product_url}")
        # Retry logic for loading the page
        max_attempts = 3
        attempt = 0
        while attempt < max_attempts:
            try:
                driver.get(product_url)
                break  # Break on success
            except ReadTimeoutError as e:
                attempt += 1
                print(f"Read timeout on attempt {attempt} for {product_url}. Retrying...")
                if attempt == max_attempts:
                    print(f"Max attempts reached for {product_url}. Skipping.")
                    continue  # Skip to the next URL

        # Refresh the page if needed and wait a moment
        driver.refresh()
        time.sleep(5)

        # Reset global image lists for this product
        global_all_images = []
        global_variant_images = []
        # Reset the variant expansion flags for this product
        scrape_variant.variant_details_opened = False
        scrape_variant.features_opened = False
        scrape_variant.material_opened = False
        scrape_variant.care_opened = False
        scrape_variant.shipping_opened = False
        scrape_variant.assembly_opened = False

        # Reset current size globals
        current_size_present = False
        current_size_name = ""

        try:
            # Wait for product title
            title_element = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, "product-title"))
            )
            title = title_element.text.strip()
            Handle = title

            # Wait for brand name
            brand_element = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//a[@aria-label='brand']"))
            )
            Vendor = brand_element.text.strip()

            # Get tags from breadcrumb
            breadcrumb_links = driver.find_elements(By.CSS_SELECTOR, "#breadcrumb a")
            if breadcrumb_links:
                tags = breadcrumb_links[-1].text.strip()
                print("Last breadcrumb:", tags)
            else:
                print("No breadcrumb links found.")

            # Check if sizes are present and if there is a valid size value.
            try:
                size_list = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//ul[@class='product-detail__list' and @data-bind='foreach: AllSizes']"))
                )
                size_elements = size_list.find_elements(By.TAG_NAME, "li")
                valid_size_elements = []
                for elem in size_elements:
                    size_val = elem.find_element(By.XPATH, ".//span[@data-bind='text: Value']").text.strip()
                    if size_val.lower() != "see below":
                        valid_size_elements.append(elem)
                if valid_size_elements:
                    current_size_present = True
                    print("Valid sizes found.")
                else:
                    current_size_present = False
                    print("Only placeholder sizes found ('See below'); skipping size option.")
            except Exception:
                current_size_present = False
                print("Sizes are missing. Proceeding with variations only.")

            # Wait for variation list
            variation_lists = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.XPATH, "//ul[@class='product-detail__list']"))
            )
            if len(variation_lists) < 2:
                raise Exception("Could not find the second variation list.")

            variation_elements = variation_lists[1].find_elements(By.TAG_NAME, "li")

            # Loop through sizes and variations if sizes exist
            if current_size_present:
                for size_index in range(len(valid_size_elements)):
                    try:
                        # Refresh valid size list for each iteration
                        valid_size_elements = size_list.find_elements(By.TAG_NAME, "li")
                        # Ensure we pick the valid ones
                        filtered_sizes = []
                        for elem in valid_size_elements:
                            size_val = elem.find_element(By.XPATH, ".//span[@data-bind='text: Value']").text.strip()
                            if size_val.lower() != "see below":
                                filtered_sizes.append(elem)
                        if size_index >= len(filtered_sizes):
                            break
                        size = filtered_sizes[size_index]
                        previous_size = driver.find_element(
                            By.XPATH, ".//li[contains(@class, 'actived')]//span[@data-bind='text: Value']"
                        ).text.strip()
                        size_name = size.find_element(By.XPATH, ".//span[@data-bind='text: Value']").text.strip()
                        current_size_name = size_name  # update the global size name
                        if "actived" in size.get_attribute("class"):
                            print(f"Size {size_index + 1} ({size_name}) already selected.")
                        else:
                            driver.execute_script("arguments[0].scrollIntoView();", size)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", size)
                            print(f"Clicked on size {size_index + 1} ({size_name}); waiting for change...")
                            WebDriverWait(driver, 10).until(
                                lambda d: d.find_element(By.XPATH, ".//li[contains(@class, 'actived')]//span[@data-bind='text: Value']").text.strip() != previous_size
                            )
                            time.sleep(3)

                        # Loop through variations for this size
                        for variation_index in range(len(variation_elements)):
                            try:
                                variation_elements = variation_lists[1].find_elements(By.TAG_NAME, "li")
                                variation = variation_elements[variation_index]
                                previous_variant = driver.find_element(
                                    By.XPATH, ".//li[contains(@class, 'actived')]//span[@data-bind='text: Value']"
                                ).text.strip()
                                previous_image = driver.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src")
                                if variation_index == 0:
                                    if "actived" not in variation.get_attribute("class"):
                                        driver.execute_script("arguments[0].scrollIntoView();", variation)
                                        time.sleep(1)
                                        driver.execute_script("arguments[0].click();", variation)
                                        print(f"Clicked on variation {variation_index + 1}; waiting for change...")
                                        WebDriverWait(driver, 10).until(
                                            lambda d: d.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src") != previous_image
                                        )
                                        time.sleep(3)
                                else:
                                    if "actived" in variation.get_attribute("class"):
                                        print(f"Skipping variation {variation_index + 1} (already selected).")
                                        continue
                                    driver.execute_script("arguments[0].scrollIntoView();", variation)
                                    time.sleep(1)
                                    driver.execute_script("arguments[0].click();", variation)
                                    print(f"Clicked on variation {variation_index + 1}; waiting for change...")
                                    WebDriverWait(driver, 10).until(
                                        lambda d: d.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src") != previous_image
                                    )
                                    time.sleep(3)

                                driver.execute_script("window.scrollBy(0, 500);")
                                time.sleep(1)
                                driver.execute_script("window.scrollBy(0, -500);")
                                time.sleep(1)
                                scrape_variant()
                            except Exception as e:
                                print(f"Error processing variation {variation_index + 1}: {e}")
                    except Exception as e:
                        print(f"Error processing size {size_index + 1}: {e}")

            # If sizes are missing, loop directly through variations
            else:
                for variation_index in range(len(variation_elements)):
                    try:
                        variation_elements = variation_lists[1].find_elements(By.TAG_NAME, "li")
                        variation = variation_elements[variation_index]
                        previous_variant = driver.find_element(
                            By.XPATH, ".//li[contains(@class, 'actived')]//span[@data-bind='text: Value']"
                        ).text.strip()
                        previous_image = driver.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src")
                        if variation_index == 0:
                            if "actived" not in variation.get_attribute("class"):
                                driver.execute_script("arguments[0].scrollIntoView();", variation)
                                time.sleep(1)
                                driver.execute_script("arguments[0].click();", variation)
                                print(f"Clicked on variation {variation_index + 1}; waiting for change...")
                                WebDriverWait(driver, 10).until(
                                    lambda d: d.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src") != previous_image
                                )
                                time.sleep(3)
                        else:
                            if "actived" in variation.get_attribute("class"):
                                print(f"Skipping variation {variation_index + 1} (already selected).")
                                continue
                            driver.execute_script("arguments[0].scrollIntoView();", variation)
                            time.sleep(1)
                            driver.execute_script("arguments[0].click();", variation)
                            print(f"Clicked on variation {variation_index + 1}; waiting for change...")
                            WebDriverWait(driver, 10).until(
                                lambda d: d.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//img").get_attribute("src") != previous_image
                            )
                            time.sleep(3)

                        driver.execute_script("window.scrollBy(0, 500);")
                        time.sleep(1)
                        driver.execute_script("window.scrollBy(0, -500);")
                        time.sleep(1)
                        scrape_variant()
                    except Exception as e:
                        print(f"Error processing variation {variation_index + 1}: {e}")

        except Exception as e:
            print(f"Error scraping {product_url}: {e}")

        # Process extra images (if any) after variations
        try:
            unique_images = []
            for img in global_all_images:
                if img not in unique_images:
                    unique_images.append(img)
            final_images = [
                img for img in unique_images 
                if img not in global_variant_images and not img.startswith("https://img.youtube.com")
            ]

            print(f"Final extra images for {Handle}: {final_images}")
            max_extra_images = 12 - len(global_variant_images)
            if len(final_images) > max_extra_images:
                print(f"Trimming extra images from {len(final_images)} to {max_extra_images} to meet the limit.")
                final_images = final_images[:max_extra_images]
            for pos, img in enumerate(final_images, start=1):
                row = [
                    Handle, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                    "", "", img, pos, "", "", "", "", "", "", "", "", "", ""
                ]
                writer.writerow(row)
        except Exception as e:
            print(f"Error writing extra image rows for {Handle}: {e}")

driver.quit()
print(f"Data successfully saved in {csv_filename}")

# ---- New code to convert CSV output to Excel for GitHub artifacts ----
output_dir = "output"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

excel_filename = os.path.join(output_dir, "olliix_products.xlsx")
try:
    df = pd.read_csv(csv_filename)
    df.to_excel(excel_filename, index=False)
    print(f"Excel file saved as {excel_filename}")
except Exception as e:
    print(f"Error converting CSV to Excel: {e}")
