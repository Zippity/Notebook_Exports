# OneNote Desktop Local Notebook Exporter

Uses OneNote Win32 COM API to export all locally stored notebooks into a folder called `/Backups` as `.onepkg` files. 
A `hierarchy.xml` file containing local notebook metadata will also be generated.

Notebook names are automatically sanitized for Windows filesystem naming schemes. Notebook files use the name of a given notebook, not its nickname.

Make sure that there aren't any sections or pages with syncing errors, or the export will break in odd ways.

If the `/Backups` folder has pre-existing notebook exports, those notebooks will be skipped during the export process.

## Why did I make this?
I just graduated from university, and I stored all of my class notes from junior year onward in individual OneNote notebooks. As I didn't want to manually export each and every notebook myself, I wrote up a quick script that would do it for me. You'll need to load every notebook you want to export into OneNote Desktop, but that's a small price to pay for fully automated batch exports.


---

## Prerequisites
- Windows
- OneNote Desktop (Office 2016/2019/365, not the Windows 10/UWP version)
- Python 3.7+
- OneNote must be installed and **registered for COM automation** - *see this [Stack Overflow](https://stackoverflow.com/a/22098588) post for more information*

## Installation
1. Clone or download this repository.
2. Install dependencies:
   ```sh
   pip install lxml tqdm pywin32
   ```

## Usage
1. Make sure OneNote Desktop is installed and you are signed in.
2. Run the script:
   ```sh
   python local_notebook_export.py
   ```
3. The script will export all local notebooks to the `/Backups` folder as `.onepkg` files.

**Note:**
- Run the script as administrator if you encounter permission errors.
- Large notebooks may take several minutes to export.
- The script waits for each export to finish before starting the next.

## Limitations & Notes
- Only local notebooks are exported (not cloud-only or UWP/Store notebooks).
- Notebooks already present in `/Backups` are skipped.
- Export may fail if there are syncing errors in OneNote.
- The script uses the OneNote 2013 schema for validation.

## Credits
- OneNote 2013 schema from [OneMore](https://github.com/stevencohn/OneMore/blob/main/Reference/0336.OneNoteApplication_2013.xsd)
