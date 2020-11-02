#!/usr/bin/env python3
import os
import argparse
import zipfile
import re
from enum import Enum

from xml.dom import minidom


class ExecMode(Enum):
    Unprotect = 'unprot'
    Reprotect = 'reprot'


sheet_xml_regex = re.compile(r'^xl\/worksheets\/[^\/]+\.xml$')
unprot_token = '.unprot'


def unprotect(target_file):
    path_parts = os.path.splitext(target_file)
    output_file = path_parts[0] + unprot_token + path_parts[1]
    if os.path.exists(output_file):
        os.remove(output_file)
    with zipfile.ZipFile(target_file, mode='r') as target_file_r, \
            zipfile.ZipFile(output_file, mode='x', compression=zipfile.ZIP_DEFLATED) as output_file_x:
        for f in target_file_r.namelist():
            is_sheet = sheet_xml_regex.match(f) is not None
            content = target_file_r.read(f)
            if is_sheet:
                doc = minidom.parseString(content)
                prots = doc.getElementsByTagName('sheetProtection')
                assert not prots or len(prots) == 1
                if prots:
                    prot = prots[0]
                    prot.parentNode.removeChild(prot)
                content = doc.toxml()
            output_file_x.writestr(f, content)


prot_token = '.prot'


def reprotect(target_file):
    assert (unprot_token + '.') in target_file
    orig_file = target_file.replace(unprot_token, '')
    path_parts = os.path.splitext(orig_file)
    output_file = path_parts[0] + prot_token + path_parts[1]
    if os.path.exists(output_file):
        os.remove(output_file)
    with zipfile.ZipFile(orig_file, mode='r') as orig_file_r, \
        zipfile.ZipFile(target_file, mode='r') as target_file_r, \
            zipfile.ZipFile(output_file, mode='x', compression=zipfile.ZIP_DEFLATED) as output_file_x:
        # Read new content from target_file
        for f in target_file_r.namelist():
            is_sheet = sheet_xml_regex.match(f) is not None
            content = target_file_r.read(f)
            if is_sheet:
                # Read protection from orig_file
                orig_doc = minidom.parseString(orig_file_r.read(f))
                prots = orig_doc.getElementsByTagName('sheetProtection')
                assert not prots or len(prots) == 1
                # Insert into new sheet
                if prots:
                    doc = minidom.parseString(content)
                    prot = prots[0]
                    # Insert after sheetData node. Apparently order matters..
                    sd_sibling = doc.getElementsByTagName('sheetData')[
                        0].nextSibling
                    if sd_sibling:
                        sd_sibling.parentNode.insertBefore(prot, sd_sibling)
                    else:
                        # If sheetData is the last child of worksheet, then append to the end
                        ws = doc.getElementsByTagName('worksheet')[0]
                        ws.appendChild(prot)
                content = doc.toxml()
            # Write to output_file
            output_file_x.writestr(f, content)


def main():
    ap = argparse.ArgumentParser(
        'Unprotect Excel worksheet protection, and re-protect it after modification.'
    )
    ap.add_argument('mode', type=ExecMode,
                    help=f'Execution mode. Possible values: {{{", ".join((m.value for m in ExecMode))}}}')
    ap.add_argument('file', type=str,
                    help='Target Excel file.')
    args = ap.parse_args()
    if args.mode == ExecMode.Unprotect:
        unprotect(args.file)
    elif args.mode == ExecMode.Reprotect:
        reprotect(args.file)


if __name__ == "__main__":
    main()
