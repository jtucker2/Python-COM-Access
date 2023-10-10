# Incomplete features and known bugs
- Navigation pane groups can be imported but don't have objects assigned to them.
- The program will sometimes (seemingly after a Windows update) stop working with a `has no attribute 'CLSIDToClassMap'` error. To fix this delete the folder `C:\Users\<my username>\AppData\Local\Temp\gen_py` (https://stackoverflow.com/questions/33267002/why-am-i-suddenly-getting-a-no-attribute-clsidtopackagemap-error-with-win32com)

# Unsolved problems and mysteries
## Navigation pane groups
Navigation pane categories and groups can be exported into an XML file. However the link between objects and groups doesn't appear to be straight forward. MSysNavPaneObjectIDs point to MSysNavPaneGroupToObjects which point to MSysNavPaneGroups. These pointers use IDs which don't appear to be unique, so can't be followed without manual checking through the possible paths to find the right one.

## Changes in exports after multiple exports/reimports
When comparing the exports after multiple exports/import cycles there are large numbers of differences between the files, however no important information seems to have changed. SQL queries have minor changes that don't affect their meaning, such as the order of AND segments of the queries being swapped. Forms have GUIDs of objects changed and sections of large numbers changed. However, nothing seems to change in terms of using the forms in Access itself after however many exports/imports.

## ProjectVariants form not exporting
ProjectVariants is the only form that doesn't get exported due to an error: "The record source 'ProjectVariants' specified on this form or report does not exist". This error also happens when trying to open the form in Access so it seems this is a problem with the Access project.