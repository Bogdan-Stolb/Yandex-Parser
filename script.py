import time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains

class YandexMapsParser:
    def __init__(self, headless=False):
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.geolocation": 1
        })
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 15)
        self.actions = ActionChains(self.driver)
        
    def search_companies(self, company_type, max_companies=None):
        print("Открываем Яндекс.Карты")
        self.driver.get("https://yandex.ru/maps/")
        time.sleep(0.2)

        print(f"Ищем: {company_type}")
        search_input = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.input__control"))
        )
        
        search_input.clear()
        search_input.send_keys(company_type)
        time.sleep(0.2)
        
        search_input.send_keys(Keys.ENTER)
        print("Выполняется поиск")
        time.sleep(0.5)
        
        companies_data = []
        seen_companies = set()
        scroll_count = 0
        max_scrolls = 50
        
        print(f"\nНачинаем сбор компаний")
        
        body = self.driver.find_element(By.TAG_NAME, "body")
        
        results_container = None
        try:
            results_container = self.driver.find_element(By.CSS_SELECTOR, "div.search-list-view, ul.search-list-view__list, div.scroll__container")
        except:
            pass
        
        last_count = 0
        no_change_count = 0
        
        while scroll_count < max_scrolls:
            cards = self.driver.find_elements(By.CSS_SELECTOR, "li.search-snippet-view")
            print(f"Найдено карточек: {len(cards)}, собрано: {len(companies_data)}")
            
            for card in cards:
                try:
                    name_element = card.find_element(By.CSS_SELECTOR, "a.link-overlay")
                    name = name_element.get_attribute('aria-label') or name_element.text
                    
                    if name in seen_companies:
                        continue
                    seen_companies.add(name)
                    
                    category = "N/A"
                    try:
                        category_element = card.find_element(By.CSS_SELECTOR, "div.search-business-snippet-view__categories a")
                        category = category_element.text
                    except:
                        pass
                    
                    address = "N/A"
                    try:
                        address_element = card.find_element(By.CSS_SELECTOR, "a.search-business-snippet-view__address")
                        address = address_element.text
                    except:
                        pass
                    
                    company_data = {
                        'Название': name,
                        'Категория': category,
                        'Адрес': address,
                        'Телефоны': '',
                        'Сайт': ''
                    }
                    
                    try:
                        self.driver.execute_script("arguments[0].click();", name_element)
                        time.sleep(0.13)
                        
                        self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.business-card-view"))
                        )
                        
                        phones = []
                        try:
                            phone_blocks = self.driver.find_elements(By.CSS_SELECTOR, "div.card-phones-view")
                            
                            for phone_block in phone_blocks:
                                try:
                                    expand_button = phone_block.find_element(By.CSS_SELECTOR, "div.card-feature-view._interactive")
                                    if expand_button:
                                        self.driver.execute_script("arguments[0].click();", expand_button)
                                        time.sleep(0.08)
                                except:
                                    pass
                                
                                phone_elements = phone_block.find_elements(By.CSS_SELECTOR, "div.card-phones-view__phone-number")
                                for phone_elem in phone_elements:
                                    phone_text = phone_elem.text.strip()
                                    if phone_text and phone_text != "Показать телефон":
                                        import re
                                        clean_phone = re.sub(r'[^\d\+\-\(\)\s]', '', phone_text)
                                        if clean_phone:
                                            phones.append(clean_phone)
                        except:
                            pass
                        
                        company_data['Телефоны'] = "; ".join(phones) if phones else "Не указан"
                        
                        try:
                            website_element = self.driver.find_element(By.CSS_SELECTOR, "a.business-urls-view__link")
                            company_data['Сайт'] = website_element.get_attribute('href') or website_element.text
                        except:
                            company_data['Сайт'] = "Не указан"
                        
                        self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                        time.sleep(0.08)
                        
                    except Exception as e:
                        print(f"Ошибка деталей для {name}: {str(e)[:50]}")
                        try:
                            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                        except:
                            pass
                    
                    companies_data.append(company_data)
                    print(f"{len(companies_data)}. ✓ {name}")
                    
                    if max_companies and len(companies_data) >= max_companies:
                        print(f"\nДостигнут лимит в {max_companies} компаний")
                        return companies_data
                    
                except Exception as e:
                    continue
            
            current_count = len(companies_data)
            if current_count == last_count:
                no_change_count += 1
                print(f"Новых компаний не появилось ({no_change_count}/3)")
                if no_change_count >= 3:
                    print(f"\nНовые компании не найдены. Завершаем.")
                    break
            else:
                no_change_count = 0
                last_count = current_count
            
            print(f"  📜 Прокрутка {scroll_count + 1}...")
            
            body.send_keys(Keys.PAGE_DOWN)
            time.sleep(0.08)
            
            body.send_keys(Keys.PAGE_DOWN)
            time.sleep(0.08)
            
            try:
                if results_container:
                    self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", results_container)
                else:
                    panels = self.driver.find_elements(By.CSS_SELECTOR, "div.search-list-view, div.scroll__container, div._scrollable")
                    for panel in panels:
                        self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", panel)
            except:
                pass
            
            try:
                if results_container:
                    self.actions.move_to_element(results_container).scroll(0, 500).perform()
                else:
                    self.actions.scroll(0, 500).perform()
                time.sleep(0.08)
            except:
                pass
            
            self.driver.execute_script("window.scrollBy(0, 800);")
            time.sleep(0.2)
            
            try:
                show_more = self.driver.find_element(By.CSS_SELECTOR, "button.more-button, div.more-button, a.more-button, span.more-button")
                if show_more.is_displayed():
                    self.driver.execute_script("arguments[0].click();", show_more)
                    print("Нажата кнопка 'Показать еще'")
                    time.sleep(0.2)
            except:
                pass
            
            scroll_count += 1
        
        print(f"\nВсего найдено компаний: {len(companies_data)}")
        return companies_data
    
    def export_to_excel(self, companies, filename):
        if not companies:
            print("Нет данных для экспорта")
            return
        
        wb = Workbook()
        ws_companies = wb.active
        ws_companies.title = "Компании"
        
        headers = ['№', 'Название', 'Категория', 'Адрес', 'Телефоны', 'Сайт']
        ws_companies.append(headers)
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        for cell in ws_companies[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for idx, company in enumerate(companies, 1):
            ws_companies.append([
                idx,
                company['Название'],
                company['Категория'],
                company['Адрес'],
                company['Телефоны'],
                company['Сайт']
            ])
        
        # Статистика
        ws_stats = wb.create_sheet("Статистика")
        ws_stats['A1'] = f"СТАТИСТИКА ПОИСКА"
        ws_stats['A1'].font = Font(bold=True, size=14)
        
        with_phones = sum(1 for c in companies if c['Телефоны'] != 'Не указан')
        with_sites = sum(1 for c in companies if c['Сайт'] != 'Не указан')
        
        stats_data = [
            ["Показатель", "Значение"],
            ["Всего компаний", len(companies)],
            ["С телефонами", with_phones],
            ["С сайтом", with_sites],
        ]
        
        for row in stats_data:
            ws_stats.append(row)
        
        for sheet in [ws_companies, ws_stats]:
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(filename)
        print(f"\nФайл Excel сохранен: {filename}")
        print(f"Всего записей: {len(companies)}")
    
    def close(self):
        self.driver.quit()


def main():
    parser = YandexMapsParser(headless=False)
    
    try:
        company_type = input("Введите тип компании (например: IT-компания, ресторан, аптека): ")
        limit_input = input("Введите максимальное количество компаний (Enter = без лимита): ")
        max_companies = int(limit_input) if limit_input.strip() else None
        
        companies = parser.search_companies(company_type, max_companies)
        
        if companies:
            filename = f"{company_type.replace(' ', '_').replace('-', '_')}_companies.xlsx"
            parser.export_to_excel(companies, filename)
            print(f"\nГотово! Файл сохранен: {filename}")
        else:
            print("Компании не найдены")
        
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        parser.close()


if __name__ == "__main__":
    main()