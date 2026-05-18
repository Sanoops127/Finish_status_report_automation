import time
from functools import wraps
from src.utils.logger import logger

def with_retry(max_retries=3, base_delay=2, exceptions=(Exception,)):
    """
    Retry decorator with exponential backoff.
    
    :param max_retries: Maximum number of retries before giving up.
    :param base_delay: Base delay in seconds. It multiplies by 2 on each retry.
    :param exceptions: Tuple of exceptions that trigger a retry.
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = base_delay
            last_exception = None
            
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt < max_retries:
                        logger.warning(
                            f"Attempt {attempt + 1}/{max_retries} failed for {func.__name__} "
                            f"with error: {e}. Retrying in {delay} seconds..."
                        )
                        time.sleep(delay)
                        delay *= 2  # Exponential backoff
                    else:
                        logger.error(
                            f"All {max_retries} retries failed for {func.__name__}. "
                            f"Final error: {e}"
                        )
            
            # If we reach here, it means all retries failed
            raise last_exception
            
        return wrapper
    return decorator
