from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    TimeoutException,
    WebDriverException,
)
from logging import Logger
def HandleException(func):
    def wrapper(self, *args, **kwargs):
        try:
            return func(self,*args,**kwargs)
        except (
            StaleElementReferenceException,
            ElementNotInteractableException,
            ElementClickInterceptedException,
            TimeoutException,
            WebDriverException
        ):
            logger = None
            if hasattr(self,"logger"):
                logger:Logger = self.logger
            if logger:
                logger.info(f"Retry: {func.__name__}")
            return func(self,*args,**kwargs)
        except Exception as e:
            self.logger.error(f"{func.__name__}: {e}")
    return wrapper