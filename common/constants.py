import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

# Settings files
KYORUGI_SETTINGS_FILE = os.path.join(PROJECT_ROOT, "kyorugi_settings.json")
POOMSAE_SETTINGS_FILE = os.path.join(PROJECT_ROOT, "poomsae_settings.json")

# Template paths
TEMPLATES_DIR = os.path.join(PROJECT_ROOT, "templates")
GAME_NUMBER_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "경기번호_계산기_양식.xlsx")
KYORUGI_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "겨루기_경기시간_계산기_양식.xlsx")
POOMSAE_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "품새_경기시간_계산기_양식.xlsx")
