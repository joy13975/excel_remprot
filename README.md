# remprot.py
Remove + restore Excel worksheet protection.

## Disclaimer

This program is only for educational purposes and the author takes no responsibility in any harm caused by the use of this code. It is not extensively tested. Only use at your own risk and always backup files before use.

Microsoft Excel is a product developed and sold by Microsoft. This repository is in no way associated with Microsoft and the author does not intend to infringe Microsoft's trademark.

## Known Issues

1. Result file size varies slightly w.r.t. original file size.


## Before use
Download `remprot.py` one way or another.

Requires python3.

`chmod +x remprot.py` for skipping the preceeding `python3` command (optional). 

## Usage

`./remprot.py -h`

```
usage: Unprotect Excel worksheet protection, and re-protect it after modification. [-h] mode file

positional arguments:
  mode        Execution mode. Possible values: {unprot, reprot}
  file        Target Excel file.

optional arguments:
  -h, --help  show this help message and exit
  ```

## Example

1. Remove password protection (`unprot` mode)

`./remprot.py unprot example.xlsx`

This creates a `example.unprot.xlsx` with unlocked sheets and does not alter the orignal file.

2. Modify the target sheets and save (i.e. to `example.unprot.xlsx`).
3. Restore the password protection (`reprot` mode):

`./remprot.py reprot exmple.unprot.xslx`

Note that this requires `example.xslx`, the original file to be present in the same directory to source original password protection data from.

This creates a `example.prot.xlsx` with password protected sheets and does not alter the intermediate file ending in `.unprot.xslx`.