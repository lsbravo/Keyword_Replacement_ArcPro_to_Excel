# Keyword_Replacement_ArcPro_to_Excel
Arcpy script for custom tool that allows mapping intersect results (Points against Polygons) to keywords in an Excel Spreadsheet.

This script is for situations where you need to run multiple intersect queries on a single point to populate a report in Word. For example, if you regularly find yourself needing to fill out values like zip codes, counties, or districts on a specific location. The tool allows you to add as many layers for intersect queries as you need but will only return one result per feature class, which is why it's best for situations without overlapping polygons within the same feature class.

Tested in ArcGIS Pro 3.1.0.

# What You Need to Get Started

1. An Excel spreadsheet populated with keywords to be replaced by your intersect results. Keywords need to be unique and cannot contain spaces or semicolons. One way to guarantee no accidental replacements are made is to surround your unique keyword with <> or XXX's. See screenshot below for an example.
2. A shapefile with your single point location.
3. Shapefiles on the polygons to be intersected (states, regions, counties). You will also need to know how the attribute table is formatted to map the appropriate value to the keyword in Word.

# How It Works
All parameters are required. Make sure the Workbook Filepath includes the xlsx extension. The LayersToCheck portion can accept several layers. These must be accompanied by the name of the attribute that will be populated in Excel and the keyword that will be replaced.


When the tool is run, you'll see a new Excel Spreadsheet in the same location as your original document. It will have "UPDATE" appended to the name. Open it, and you should see the keywords now mapped to the results of the intersection queries.

#Before Getting Started
Because the script interacts with Excel, you'll need the OPENPYXL module. If you haven't already, you can open a Jupyter Notbook and import the module before getting started.

# How To Add It To A Custom Toolbox
In the Catalog panel, right click Toolboxes, add a new one, and rename it.*

![Add_Toolbox](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/b4471a70-c984-435d-8247-3897dae775ae)

*If you already have a custom toolbox, you can skip this step.

Then right click the new tookbox, under New, click Script.

![Add_Script](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/00c6fbdf-2d71-4244-8694-f5092cd392c1)

It'll prompt you for a name and label.

![General_Excel](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/38edd4ab-3cf2-46ed-9cfa-7853534a49a2)

Once you're done with that, move to the parameters tab and set these three parameters up exactly as displayed in the screenshot.

![Excel_Parameters](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/e69a86c3-5491-45d1-873f-a1fbdd702ef1)

The LayersToCheck Value Table will prompt for column names and types. Set those to these values.

![Parameter_Data_Type](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/fb6ddacb-43bd-422a-b451-c29981805357)

Last, you'll go to the Execution tab, paste the Python script that's in this repository, and then hit OK to finish.

![Execution_Excel](https://github.com/lsbravo/Keyword_Replacement_ArcPro_to_Excel/assets/121823541/07038975-fa34-4c6b-8870-e24636865397)

# Other Notes
Made with assistance from Chat GPT.
There's a repository that has this exact same script compatible with Microsoft Word documents.
