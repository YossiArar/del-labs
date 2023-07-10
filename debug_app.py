import os
import sys

PATH = os.path.dirname(__file__)
ENV_PATH = f"{PATH}/the-gallery-system-env/lib/site-packages"
sys.path.append(ENV_PATH)

import platform
import streamlit.web.cli as stcli

if 'Windows' in platform.platform():
    PATH = os.path.dirname(__file__)
else:
    PATH = os.environ.get("PWD")

APP_PATH = f"{PATH}/app"


def streamlit_run():
    sys.argv = ["streamlit", "run", f"{APP_PATH}/main_app.py", "--global.developmentMode=False"]
    sys.exit(stcli.main())


if __name__ == '__main__':
    streamlit_run()
