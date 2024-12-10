import os
import subprocess
import hashlib
import platform
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
import selenium
import glob
import shutil
import warnings
import time
import logging
from pathlib import Path
import sys
warnings.filterwarnings('ignore')

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException

def setup_logging():
    """Log dosyaları, çıkarılabilir."""
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f'download_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

def get_firefox_path():
    """İşletim sistemine göre Firefox yolunu bulma"""
    system = platform.system()
    
    if system == "Windows":
        import winreg
        try:
            # 64-bit Firefox
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Mozilla\Mozilla Firefox")
            path = winreg.QueryValue(key, None)
            return os.path.join(path, 'firefox.exe')
        except:
            try:
                # 32-bit Firefox
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Mozilla\Mozilla Firefox")
                path = winreg.QueryValue(key, None)
                return os.path.join(path, 'firefox.exe')
            except:
                return None
    
    elif system == "Darwin":  # macOS
        paths = [
            "/Applications/Firefox.app/Contents/MacOS/firefox",
            os.path.expanduser("~/Applications/Firefox.app/Contents/MacOS/firefox")
        ]
        for path in paths:
            if os.path.exists(path):
                return path
    
    elif system == "Linux":
        paths = [
            "/usr/bin/firefox",
            "/usr/lib/firefox/firefox"
        ]
        for path in paths:
            if os.path.exists(path):
                return path
    
    return None

def setup_directories():
    """İndirme ve log klasörlerini oluşturma"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    download_dir = os.path.join(script_dir, 'downloads')
    os.makedirs(download_dir, exist_ok=True)
    return download_dir

def setup_firefox_options(download_dir):
    """Firefox ayarlarını yapılandırma"""
    options = Options()
    
    firefox_path = get_firefox_path()
    if firefox_path and os.path.exists(firefox_path):
        options.binary_location = firefox_path
    
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", download_dir)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", 
                         "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    return options

def get_geckodriver_name():
    """İşletim sistemine göre geckodriver adını belirleme"""
    system = platform.system()
    return "geckodriver.exe" if system == "Windows" else "geckodriver"

def setup_webdriver(options):
    """Webdriver'ı yapılandırma"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    geckodriver_name = get_geckodriver_name()
    geckodriver_path = os.path.join(script_dir, geckodriver_name)
    
    if not os.path.exists(geckodriver_path):
        logging.error(f"Geckodriver bulunamadı: {geckodriver_path}")
        logging.error(f"Lütfen {geckodriver_name} dosyasını script ile aynı klasöre koyun")
        sys.exit(1)
    
    service = Service(executable_path=geckodriver_path)
    return webdriver.Firefox(service=service, options=options)

def find_correct_row_index(df, key_text):
    """Girdiğimiz sitede, tabloda doğru satırı bulma"""
    for i, val in enumerate(df['Unnamed: 0']):
        val_str = str(val)
        if val_str.strip() == '-' or 'güncellenmemektedir' in val_str:
            continue
        if 'İstatistiksel Tablolar' in val_str and key_text in val_str:
            return i
    return None

def find_download_button(browser, key):
    """İndirme butonunu bulma"""
    xpaths_to_try = [
        f"//*[@id='istatistikselTable']/tbody/tr[{key+1}]/td[3]/a/img",
        f"//*[@id='istatistikselTable']/tbody/tr[{key+1}]/td[3]/a",
        f"//table[@id='istatistikselTable']//tr[{key+1}]//a[.//img[contains(@src, 'excel.svg')]]",
        f"//table[@id='istatistikselTable']//tr[{key+1}]//td[3]//a[.//img]"
    ]
    
    for xpath in xpaths_to_try:
        try:
            logging.info(f"XPath deneniyor: {xpath}")
            elements = browser.find_elements(By.XPATH, xpath)
            if elements:
                logging.info(f"XPath ile {len(elements)} element bulundu: {xpath}")
                element = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                if 'word.svg' not in element.get_attribute('outerHTML'):
                    return element
                else:
                    logging.info("Word linki bulundu, atlanıyor")
            else:
                logging.info(f"XPath ile element bulunamadı: {xpath}")
        except Exception as e:
            logging.info(f"XPath hatası {xpath}: {e}")
            continue
    
    try:
        row = browser.find_element(By.XPATH, f"//*[@id='istatistikselTable']/tbody/tr[{key+1}]")
        logging.info(f"Satır {key+1} içeriği: {row.get_attribute('outerHTML')}")
    except:
        logging.error(f"Satır bulunamadı: {key+1}")
    
    return None

def handle_popup(browser):
    """Popup'ı kapatma"""
    try:
        popup_button = WebDriverWait(browser, 3).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/div/div[2]/div/p/button[1]"))
        )
        browser.execute_script("arguments[0].click();", popup_button)
        logging.info("Popup kapatıldı")
        time.sleep(1)
        return True
    except TimeoutException:
        return True
    except Exception as e:
        try:
            browser.execute_script("""
                var elements = document.getElementsByClassName('modal');
                for(var i=0; i<elements.length; i++) {
                    elements[i].remove();
                }
                var elements = document.getElementsByClassName('modal-backdrop');
                for(var i=0; i<elements.length; i++) {
                    elements[i].remove();
                }
            """)
            logging.info("Popup JavaScript ile kaldırıldı")
            time.sleep(1)
            return True
        except:
            logging.warning(f"Popup kapatılamadı: {e}")
            return False

def process_link(browser, link, key_text, note_text, file_mapping, download_dir):
    """Linki işle ve dosyayı indir"""
    try:
        browser.get(link)
        time.sleep(4)
        
        logging.info(f"İşleniyor: {key_text}")
        handle_popup(browser)
        
        # İstatistik sekmesine tıkla
        max_retries = 3
        for retry in range(max_retries):
            try:
                if retry > 0:
                    browser.refresh()
                    time.sleep(4)
                    handle_popup(browser)
                
                stats_tab = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/nav/div/a[2]"))
                )
                browser.execute_script("arguments[0].click();", stats_tab)
                logging.info("İstatistik sekmesi tıklandı")
                break
            except Exception as e:
                if retry == max_retries - 1:
                    logging.error(f"İstatistik sekmesi tıklanamadı: {e}")
                    return False
                logging.warning(f"Yeniden deneme {retry + 1}: {e}")
                time.sleep(2)
        
        time.sleep(4)
        
        # Tabloyu bul ve işle
        try:
            table_element = WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.ID, "istatistikselTable"))
            )
            
            table_html = table_element.get_attribute('outerHTML')
            logging.info(f"Tablo HTML uzunluğu: {len(table_html)}")
            
            df = pd.read_html(table_html)[0]
            logging.info(f"Tablo boyutu: {df.shape}")
            
            for i, row in df['Unnamed: 0'].items():
                logging.info(f"Satır {i}: {row}")
            
            key = find_correct_row_index(df, key_text)
            
            if key is None:
                logging.warning(f"Anahtar metin bulunamadı: {key_text}")
                return False
            
            logging.info(f"Eşleşme bulundu: {key}")
            
            # İndirme butonunu bul ve tıkla
            download_button = find_download_button(browser, key)
            
            if download_button is None:
                logging.info("Sayfa yenileniyor...")
                browser.refresh()
                time.sleep(4)
                handle_popup(browser)
                stats_tab = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/nav/div/a[2]"))
                )
                browser.execute_script("arguments[0].click();", stats_tab)
                time.sleep(4)
                download_button = find_download_button(browser, key)
                
                if download_button is None:
                    logging.error("İndirme butonu bulunamadı")
                    return False
            
            try:
                if download_button.tag_name.lower() == 'img':
                    download_button = download_button.find_element(By.XPATH, '..')
                
                href = download_button.get_attribute('href')
                logging.info(f"İndirme linki: {href}")
                
                browser.execute_script("arguments[0].click();", download_button)
                logging.info("İndirme butonu tıklandı")
            except Exception as e:
                logging.error(f"İndirme butonu tıklanamadı: {e}")
                return False
            
            time.sleep(5)
            
            # Dosyayı yeniden adlandır
            list_of_files = glob.glob(os.path.join(download_dir, '*'))
            if not list_of_files:
                logging.error("Dosya indirilemedi")
                return False
            
            latest_file = max(list_of_files, key=os.path.getctime)
            file_extension = Path(latest_file).suffix
            
            file_hash = hashlib.md5(f"{key_text}{datetime.now().strftime('%Y%m%d%H%M%S')}".encode()).hexdigest()[:12]
            hashed_filename = f"{file_hash}{file_extension}"
            new_filepath = os.path.join(download_dir, hashed_filename)
            
            try:
                os.rename(latest_file, new_filepath)
                file_mapping.append({
                    'Orijinal_İsim': key_text,
                    'Hash_İsim': hashed_filename,
                    'İndirme_Tarihi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Link': link,
                    'Not': note_text if note_text else ''
                })
                logging.info(f"Dosya indirildi ve yeniden adlandırıldı: {hashed_filename}")
                return True
            except Exception as e:
                logging.error(f"Dosya yeniden adlandırılamadı: {e}")
                return False
            
        except Exception as e:
            logging.error(f"Tablo işlenirken hata: {e}")
            return False
            
    except Exception as e:
        logging.error(f"Link işlenirken hata {link}: {e}")
        return False

def main():
    """Ana fonksiyon"""
    try:
        setup_logging()
        download_dir = setup_directories()
        logging.info(f"İndirme klasörü: {download_dir}")
        
        options = setup_firefox_options(download_dir)
        browser = setup_webdriver(options)
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(script_dir, 'tuik.xlsx')
        
        if not os.path.exists(excel_path):
            logging.error("tuik.xlsx bulunamadı")
            logging.error("Lütfen tuik.xlsx dosyasını script ile aynı klasöre koyun")
            sys.exit(1)
        
        df = pd.read_excel(excel_path, sheet_name='Hazır')
        file_mapping = []
        
        for index, row in df.iterrows():
            link = row['Link']
            key_text = row['Kelime']
            note_text = row['Not'] if pd.notna(row['Not']) else None
            
            if pd.isna(link) or pd.isna(key_text):
                logging.warning(f"Satır atlanıyor {index}: Link veya kelime eksik")
                continue
            
            logging.info(f"İşleniyor {index + 1}/{len(df)}: {key_text}")
            success = process_link(browser, link, key_text, note_text, file_mapping, download_dir)
            
            if not success:
                logging.warning(f"İşlenemedi: {key_text}")
            
            time.sleep(2)
        
        mapping_file = os.path.join(download_dir, 'file_mapping.xlsx')
        pd.DataFrame(file_mapping).to_excel(mapping_file, index=False)
        logging.info(f"Eşleştirme dosyası kaydedildi: {mapping_file}")
        
        browser.quit()
        logging.info("Tüm indirmeler tamamlandı")
        
    except Exception as e:
        logging.error(f"Ana süreç hatası: {e}")
        if 'browser' in locals():
            browser.quit()

if __name__ == "__main__":
    main() 