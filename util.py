import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

class Handywrapper:

    def __init__(self,driver):
        self.driver = driver


    def find_element(self,By_type, locator=""):
        element = ""
        try:
            element = self.driver.find_element(By_type, locator)
            return element
        except:
            return element


    def get_attribute(self,By_type, locator="", att="value"):
        value = ""
        try:
            element = self.find_element(By_type, locator)
            value = element.get_attribute(att)
            return value
        except:
            return value


    def find_elements(self,By_type, locator=""):
        time.sleep(0.5)
        element_list = []
        try:
            element_list = self.driver.find_elements(By_type, locator)
            return element_list
        except:
            return element_list


    def get_list_of_attributes(self,By_type, locator="", att="value"):
        value_list = []
        try:
            element_list = self.find_elements(By_type, locator)
            for element in element_list:
                value = element.get_attribute(att)
                value_list.append(value)
            return value_list
        except:
            return value_list


    def get_element_tag(self, element=None, tag=""):
        try:
            element = element.find_element(By.TAG_NAME, tag)
            return element
        except:
            return ""


    def find_element_text(self,By_type, locator=""):
        try:
            element = self.find_element(By_type, locator)
            element_text = element.text
            return element_text
        except:
            return ""


    def find_elements_list_of_text(self,By_type, locator=""):
        elements_list_of_text = []
        try:
            elements = self.find_elements(By_type, locator)
            for element in elements:
                elements_list_of_text.append(element.text)
            return elements_list_of_text
        except:
            return []


    def Click_element(self,By_type="", locator="", element=None ):
        try:
            time.sleep(0.5)
            if element is not None:
                # self.wait_explicitly(By_type, locator)
                # self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                time.sleep(0.5)
                element.click()
            else:
                self.wait_explicitly(By_type, locator)
                element = self.find_element(By_type, locator)
                self.wait_explicitly(By_type, locator)
                
                time.sleep(0.5)
                element.click()
            return element
        except Exception as e:
            print("cannot click the element")
            return element

    def scroll_to_element(self, By_type="", locator="", element=None):
        try:
            if element is None:
                self.wait_explicitly(By_type, locator)
                element = self.find_element(By_type, locator)
            if element:
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                time.sleep(0.5)
                return element
            return ""
        except Exception:
            return ""
        
    def hover(self, By_type="", locator="", element=None, pause=0.5):

        try:
            if element is None:
                self.wait_explicitly(By_type, locator)
                element = self.find_element(By_type, locator)
            if element:
                action = ActionChains(self.driver)
                action.move_to_element(element).perform()
                time.sleep(pause)
                return element
            return ""
        except Exception:
            return ""


    def is_element_present(self,By_type, locator=""):
        try:
            element = self.find_element(By_type, locator)
            if element is not None:
                return True
            else:
                return False
        except:
            return False

    def wait_explicitly(self, By_type, locator=""):
        try:
            WebDriverWait(self.driver, 50).until(
                EC.presence_of_element_located((By_type, locator)))
        except:
            pass


    def hover_over_element_by_id(self, driver, element_id, timeout=10, pause=0.5):
        """Wait for an element by ID and hover over it using ActionChains.

        Returns the WebElement if found, else raises on timeout.
        """
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, element_id)))
        element = driver.find_element(By.ID, element_id)
        ActionChains(driver).move_to_element(element).perform()
        time.sleep(pause)
        return element

    def click_in_shadow(self, selectors):
        js = """
        return (function(selectors){
        let el = document.querySelector(selectors[0]);
        if(!el) return null;
        for(let i=1;i<selectors.length;i++){
            if(!el) return null;
            el = el.shadowRoot ? el.shadowRoot.querySelector(selectors[i]) : el.querySelector(selectors[i]);
        }
        if(el){ el.click(); return true; }
        return null;
        })(arguments[0]);
        """
        return self.driver.execute_script(js, selectors)
    