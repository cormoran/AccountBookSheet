# AccountBookSheet

AccountBookSheet is collection of google app script to generate account book.
The account book supports importing csv from money forward via Google Drive.

## Features

- Support setting self-pay ratio for expense sharing.
- Support virtual installments payment for smoothing.

## Usage

1. Install [clasp](https://github.com/google/clasp)
2. Copy `secret.ts.template` to `secret.ts` and edit the contents.
3. Managed to setup project and do `clasp push`.
4. Download money forward csv files with `bash download_moneyforward_csv.sh`.

- The script supports mac os only for now.

5. Execute apps script.

- `mainImportFromGoogleDrive` to import csv files from google drive. The directory is specified in `secret.ts`.
- `mainRebuildSummarySheets` to creates sheets which collects and summarizes imported contents.

## Sheets

- `0_ImportState` maintains imported files and last imported time for efficient/faster diff sync.
- `Import_year_month` contains imported contents with additional editable columns (`擬似分割払い回数`, `自己負担率`).
- `Z_AllRepeated` contains all items and its pay month considering `擬似分割払い回数` & `自己負担率`.

Z_AllRepeated からがんばっていい感じのグラフを作る（がんばって 😉）
