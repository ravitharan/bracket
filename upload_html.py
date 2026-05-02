#! /usr/bin/env python3

import sys
import subprocess
from pathlib import Path

WEB_STEM_URL = 'https://storageareaincloud.s3.eu-west-1.amazonaws.com/badminton/'
if len(sys.argv) != 2:
    print(f'Argument error\n  Usage: {sys.argv[0]} <html_file>')
    exit(1)

html_file = Path(sys.argv[1])


results = subprocess.run(f'aws s3 cp {html_file} s3://storageareaincloud/badminton/', shell=True, capture_output=True)
print(results.stdout.decode())

if results.returncode:
    print(results.stderr)
else:
    print(f'Uploaded {html_file}')
    print(f'{WEB_STEM_URL}{html_file.name}')
