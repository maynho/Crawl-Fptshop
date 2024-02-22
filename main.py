import os
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException


class FPT:
    def __init__(self):
        pass

    def UpdateExcel(self, data_list, excel_file='danh_sach_laptop2.xlsx', output_directory='output', sheet_name='Sheet1', start_stt=None):
            if not os.path.exists(output_directory):
                os.makedirs(output_directory)

            excel_file_path = os.path.join(output_directory, excel_file)

            try:
                df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

                if start_stt is None:
                    start_stt = df['STT'].max() + 1 if 'STT' in df else 1

            except FileNotFoundError:
                df = pd.DataFrame()
                start_stt = 1

            for data in data_list:
                data['STT'] = start_stt
                start_stt += 1

            new_data = pd.DataFrame(data_list, columns=['STT', 'Tên sản phẩm', 'Giá gốc', 'Giá khuyến mãi', 'Màn hình', 'Cpu', 'Ram', 'Ổ cứng', 'Card đồ họa', 'Trọng lượng'])

            df = pd.concat([df, new_data], ignore_index=True)

            df.to_excel(excel_file_path, index=False, sheet_name=sheet_name)

            print(f'Đã cập nhật dữ liệu vào {excel_file_path}.')

    def GetContent(self, driver, xpath):
        js_code = '''
            var xpathExpression = "'''+xpath+'''";
            var result = document.evaluate(xpathExpression, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            var screenElement = result.singleNodeValue;

            if (screenElement) {
                var screenText = screenElement.innerText;
                console.log(screenText);
                return screenText;
            } else {
                console.log("Element not found with xpath.");
                return "False";
            }
        '''
        content = driver.execute_script(js_code)
        return content

    def Chrome(self):
        driver = webdriver.Chrome()

        driver.get("https://fptshop.com.vn/may-tinh-xach-tay?sort=ban-chay-nhat&trang=50")

        product_containers = driver.find_elements(By.XPATH, '//div[@class="cdt-product__info"]') #cdt-product__info
        
        data_update_excel = []

        i = 0
        graphics_card_count = 0
        for product_container in product_containers:
            i += 1
            product_name = product_container.find_element(By.XPATH, './/h3') # Tên sản phẩm
            print(f'Đang crawl : {product_name.text}')
            try:
                product_price = product_container.find_element(By.XPATH, './/div[@class="strike-price"] //strike') # Giá gốc
            except:
                product_price = product_container.find_element(By.XPATH, './/div[@class="price"]') # Giá gốc

            try:
                promotion_product_price = product_container.find_element(By.XPATH, './/div[@class="progress pdiscount2"]') # Giá khuyến mãi 
            except:
                try:
                    promotion_product_price = product_container.find_element(By.XPATH, './/div[@class="progress"]') # Giá khuyến mãi
                except:
                    promotion_product_price = product_container.find_element(By.XPATH, './/div[@class="price"]') # Giá gốc

            xpath_screen = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='Màn hình'])[{i}]"
            screen = self.GetContent(driver, xpath_screen)

            xpath_cpu = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='CPU'])[{i}]"
            cpu = self.GetContent(driver, xpath_cpu)

            xpath_ram = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='RAM'])[{i}]"
            ram = self.GetContent(driver, xpath_ram)

            xpath_hard_driver = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='Ổ cứng'])[{i}]"
            hard_driver = self.GetContent(driver, xpath_hard_driver)

            has_graphics_card_info = False
            try:
                product_container.find_element(By.XPATH, '(.//div[@class="cdt-product__config list-layout"] //span[@data-title="Đồ họa"])')
                has_graphics_card_info = True
                graphics_card_count += 1
            except NoSuchElementException:
                pass

            if has_graphics_card_info:
                xpath_graphics_card = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='Đồ họa'])[{graphics_card_count}]"
                graphics_card = self.GetContent(driver, xpath_graphics_card)
            else:
                graphics_card = "Không có thông tin"


            xpath_weight = f"(//div[@class='cdt-product__config list-layout'] //span[@data-title='Trọng lượng'])[{i}]"
            weight = self.GetContent(driver, xpath_weight)

            dict_data = {}

            dict_data['Tên sản phẩm'] = product_name.text
            dict_data['Giá gốc'] = product_price.text
            dict_data['Giá khuyến mãi'] = promotion_product_price.text
            dict_data['Màn hình'] = screen
            dict_data['Cpu'] = cpu
            dict_data['Ram'] = ram
            dict_data['Ổ cứng'] = hard_driver
            dict_data['Card đồ họa'] = graphics_card
            dict_data['Trọng lượng'] = weight
            
            data_update_excel.append(dict_data)

        self.UpdateExcel(data_update_excel)
        driver.quit()

if __name__ == "__main__":
    FPT().Chrome()