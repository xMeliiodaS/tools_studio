import os
import json
import shutil

from src.config.constants import APP_DATA_FOLDER_NAME, CONFIG_FILE_NAME
from src.config.logging_config import get_logger

logger = get_logger(__name__)

class ConfigProvider:
    """
    Class for loading configuration from a per-user JSON file in AppData,
    or a custom path if provided.
    """

    @staticmethod
    def load_config_json(path=None):
        # """
        # Load config JSON from an explicit path when provided, otherwise fall
        # back to repository-local config discovery.
        # """
        # if path is None:
        #     path = os.path.join(os.getcwd(), "config.json")
        #     if not os.path.exists(path):
        #         path = os.path.join(os.path.dirname(os.getcwd()), "config.json")
        #
        # try:
        #     with open(path, 'r', encoding='utf-8') as f:
        #         data = json.load(f)
        #         return data
        # except Exception as e:
        #     logger.exception("Error reading %s from '%s': %s", CONFIG_FILE_NAME, path, e)
        #     return {}

        """
        Load the configuration from the user's JSON file or a specified path.

        :param path: Optional full path to config.json
        :return: The loaded configuration as a dictionary.
        """
        if path is None:
            # --- Determine per-user AppData path ---
            appdata = os.getenv('APPDATA') or os.path.expanduser('~\\AppData\\Roaming')
            user_config_folder = os.path.join(appdata, APP_DATA_FOLDER_NAME)
            os.makedirs(user_config_folder, exist_ok=True)
            path = os.path.join(user_config_folder, CONFIG_FILE_NAME)

            # --- If config is missing, optionally copy default from exe folder ---
            if not os.path.exists(path):
                default_config_path = os.path.join(os.path.dirname(__file__), CONFIG_FILE_NAME)
                if os.path.exists(default_config_path):
                    shutil.copy(default_config_path, path)
                else:
                    logger.info("Default %s not found. Creating empty config.", CONFIG_FILE_NAME)
                    with open(path, 'w', encoding='utf-8') as f:
                        json.dump({}, f)

        # --- Load and return JSON ---
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            logger.info("Loaded config from: %s", path)
            return data
        except Exception as e:
            logger.exception("Error reading %s: %s", CONFIG_FILE_NAME, e)
            return {}
