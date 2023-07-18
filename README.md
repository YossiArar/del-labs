# Environment and Installations

### Python workspace:
- Install python 3.10
- Install Latest Pycharm IDE and open it
- Open project directory and Create Virtual Environment

# install project requirements
$ pip install -r requirements.txt

# or 
def build():
    build_commands = ["python setup.py develop",
                      "python -m pip install --upgrade pip",
                      "python -m pip install --upgrade build", "python -m build", "python setup.py sdist bdist_wheel"]
    for command in build_commands:
        os.system(command)


# open the gs app
$ streamlit run start_app.py