from setuptools import setup, find_packages
setup(
    name="Weleda Oogstlijst Manager",
    version="0.1",
    packages=find_packages(),
    scripts=['wol/views.py'],


    install_requires=['pyqt5',
                      "pyforms==2.1.0",
                      "pysettings==1.0.0",
                      "logging-bootstrap==1.0.0",
                      'xlrd',
                      'xlsxwriter'
                     ],
    dependency_links = [
        'git+https://github.com/dominic-dev/pyforms.git#egg=pyforms-2.1.0 ',
        'git+https://github.com/UmSenhorQualquer/pysettings.git#egg=pysettings-1.0.0',
        'git+https://bitbucket.org/fchampalimaud/logging-bootstrap.git#egg=logging-bootstrap-1.0.0',
    ],
    package_data={
        'data': ['*.xlsx'],
    },
)
