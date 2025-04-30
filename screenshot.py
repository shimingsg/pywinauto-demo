# -*- coding: utf-8 -*-
from PIL import ImageGrab
import os
from datetime import datetime

# os.environ.get('ENV_VARIABLE_NAME')
defaut_screen_path = r'.\results'  

# https://pillow.readthedocs.io/en/stable/reference/ImageGrab.html
# image_path = os.path.join(defaut_screen_path, f'screenshot_{datetime.now().strftime("%Y-%m-%d_%H_%M_%S_%f")}.{image_extension}')


def capture_screenshot(image_path:str, image_extension:str = '.png', all_screens:bool = True, bbox:tuple = None) -> str:
    """
    Capture a screenshot and save it to the specified path.

    Args:
        image_path (str): The path where the screenshot will be saved.
        image_extension (str): The file extension for the screenshot. Default is '.png'.
        all_screens (bool): Whether to capture all screens or just the primary one. Default is True.
        bbox (tuple): A 4-tuple defining the left, upper, right, and lower pixel coordinates of the box to capture.

    Returns:
        str: The path where the screenshot was saved.
    """
    os.makedirs(defaut_screen_path, exist_ok=True)
    # image_name = f'screenshot_{datetime.now().strftime("%Y-%m-%d_%H_%M_%S_%f")}.{image_extension}'
    image = ImageGrab.grab(bbox=bbox, all_screens=all_screens).save(os.path.join(defaut_screen_path, image_path), image_extension)   
    return os.path.join(defaut_screen_path, image_path)