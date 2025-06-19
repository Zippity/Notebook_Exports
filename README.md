## OneNote Desktop Local Notebook Exporter

Uses OneNote Win32 COM API to export all locally stored notebooks into a folder called `/Backups` as `.onepkg` files. 
A `hierarchy.xml` file containing all local notebooks will also be generated.

Notebook names are automatically sanitized for windows filesystem naming schemes. Notebook files use the name of a given notebook, not its nickname.

Make sure that there aren't any sections or pages with syncing errors, or the export will break in odd ways.

If the `/Backups` folder has pre-existing notebook exports, those notebooks will be skipped during the export process.

OneNote 2013 schema copied from [OneMore](https://github.com/stevencohn/OneMore/blob/main/Reference/0336.OneNoteApplication_2013.xsd)
