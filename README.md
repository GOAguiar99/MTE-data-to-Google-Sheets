<h1 align="center">
    MTE-data-to-Google-Sheets
    <br>
</h1>

<h4 align="center">Python software scrapes MTE website data with Selenium.</h4>

<p align="center">
    <a href="https://github.com/GOAguiar99/MTE-data-to-Google-Sheets">
        <img src="https://img.shields.io/badge/status-development-green?style=for-the-badge">
    </a>
    <a href="https://github.com/GOAguiar99/MTE-data-to-Google-Sheets/releases">
        <img alt="GitHub commits since latest release (by date)" src="https://img.shields.io/github/commits-since/GOAguiar99/MTE-data-to-Google-Sheets/latest?style=for-the-badge">
    </a>
    <a href="https://github.com/GOAguiar99/MTE-data-to-Google-Sheets/blob/master/LICENSE">
        <img src="https://img.shields.io/badge/license-MIT-yellow?style=for-the-badge">
    </a>
</p>

<p align="center">
    <a href="#overview">Overview</a> •
    <a href="#dependencies">Dependencies</a> •
    <a href="#license">License</a>
</p>

## Overview

This Python software scrapes MTE website data with Selenium, processes it, and uploads it to Google Sheets API. It simplifies data collection and analysis from MTE for frequent users. Selenium and Google Sheets API combination make it efficient and easily automated.

The MTE website is the online platform for Brazil's Ministry of Labor and Employment (Ministério do Trabalho e Emprego). It provides information and services related to labor regulations, worker's rights, job opportunities, and more.

One important feature of the MTE website is the CAGED (Cadastro Geral de Empregados e Desempregados) system, which is used to collect data on employment and unemployment in Brazil. Employers are required to report hiring, firing, and other employment changes through the CAGED system on a monthly basis.

The CCTS (Cadastro de Conformidade de Trabalho Seguro) is another system available on the MTE website. It is used to assess and certify the compliance of companies with Brazilian labor laws related to health and safety in the workplace.

Overall, the MTE website is a crucial resource for workers, employers, and anyone interested in labor regulations and job market information in Brazil.


## Dependencies

* altgraph==0.17.2
* async-generator==1.10
* attrs==21.4.0
* cachetools==5.0.0
* certifi==2021.10.8
* cffi==1.15.0
* charset-normalizer==2.0.12
* cryptography==36.0.2
* future==0.18.2
* google-api-core==2.7.1
* google-api-python-client==2.42.0
* google-auth==2.6.2
* googleapis-common-protos==1.56.0
* httplib2==0.20.4
* idna==3.3
* oauthlib==3.2.0
* outcome==1.1.0
* pefile==2022.5.30
* protobuf==3.19.4
* pyasn1==0.4.8
* pyasn1-modules==0.2.8
* pycparser==2.21
* pygsheets==2.0.5
* pyinstaller==5.2
* pyinstaller-hooks-contrib==2022.3
* pyOpenSSL==22.0.0
* pyparsing==3.0.7
* PySocks==1.7.1
* pywin32-ctypes==0.2.0
* requests==2.27.1
* requests-oauthlib==1.3.1
* rsa==4.8
* selenium==4.1.3
* sniffio==1.2.0
* sortedcontainers==2.4.0
* trio==0.20.0
* trio-websocket==0.9.2
* uritemplate==4.1.1
* urllib3==1.26.9
* wsproto==1.1.0

### Installation
0. For using GeckoDriver you need firefox installed in your computer!
1. Clone this repository
2. Create a virtual environment using 'python -m venv venv'
3. Activate your venv using 'venv\Scripts\activate'
4. Run `pip install -r requirements.txt`
5. Download your client_secret.json and token.json to use google api and put in the same folder (https://developers.google.com/drive/api/quickstart/python?hl=pt-br)
6. Run `python main.py`


## License
This project is licensed under MIT License.
