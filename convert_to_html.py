#! /usr/bin/env python3

from bs4 import BeautifulSoup
import sys
import subprocess

SHEETS = [ "quali_schedule", "quali_score", "knockout_score" ]

if len(sys.argv) != 2:
    print(f'Argument error\n  Usage: {sys.argv[0]} <xlsx_file>')
    exit(1)

xlsx_file = sys.argv[1]
html_file = xlsx_file.replace('.xlsx', '.html')

results = subprocess.run(f'libreoffice --invisible --convert-to html {xlsx_file}', shell=True, capture_output=True)
if results.returncode:
    print(results.stdout.decode())
    print(results.stderr.decode())
    exit(1)

for sheet in SHEETS:

    with open(html_file) as f_html:
        soup = BeautifulSoup(f_html, 'html.parser')

        title = soup.find('title')
        title.string = sheet

        overview = soup.find('p')
        overview.decompose()

        hrs = soup.find_all('hr')
        for hr in hrs:
            hr.decompose()

        tables = soup.find_all('table')
        headings = []
        for table in tables:
            heading = table.find_previous_sibling('a')
            if not sheet in heading.text:
                heading.decompose()
                table.decompose()
            else:
                h1 = soup.find('h1')
                h1.string = sheet

        file_out = f'{sheet}.html'

        with open(file_out, "w") as file:
            file.write(str(soup))

