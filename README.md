# Active Directory Audit Report
_Transform a PingCastle HTML or XML report into a branded DOCX and PDF report._

## How it works

- Add DOCX templates with a custom header, footer and styles:
    - assets/templates/MyFirstTemplate/MyFirstTemplate.docx
    - assets/templates/MyFirstTemplate/fonts/MyFirstFont/font.ttf (installed)
- Add the output report of a PingCastle run:
    - input/ad_hc_<domain>.com.(xml or html)
- Specify the template to use in the main program:
    - FILE_STYLE_TEMPLATE = path.join(PATH_ASSETS, "templates", "template1.docx")
- Run the python script and get the DOCX and PDF reports:
    - output/ActiveDirectoryAuditReport.(docx and pdf)

## Installation
_Requires python3 and pip to run._

#### Install the dependencies and build the report.

```sh
py -m pip install -r .\requirements.txt
py audit.py
```

## Compatibilities

| Python | Windows 11 |
|--------|------------|
| 3.11.4 | OK         |

## License

MIT

**Free Software, Hell Yeah!**