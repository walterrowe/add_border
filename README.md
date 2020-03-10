## add-border

macOS AppleScript to add borders to image files. This works as a droplet with macOS Finder or with Capture One Pro when added to the Open With field in Output Recipes.

| Role | Name | Year |
| ---: | :--- | ---: |
| Original | Kim Aldis | 2016 |
| Modified | Walter Rowe | 2019 |

### To create an app from this script

1. Open add_border.scpt in ScriptEditor
2. File > Export and save in a place where you can reference it
	* File Format: Application

### To use inside Capture One

1. Go to `Open With` field in a Capture One Output Recipe
2. Choose `Other` from the `Open With` drop-down menu
3. Navigate to and select your add_border droplet
4. Select Process Recipe
5. Select images to process
6. Process images (`CMD-d`)

### macOS Catalina Considerations

macOS Catalina introduces new levels of security to protect your system from unwanted access to your files. This change requires you to grant explicit access. This droplet uses the Image Event core service to modify the border of the selected image files. We therefore need to grant Full Disk access to the Image Event service.

This fix was found in [this article](https://darjeelingsteve.com/articles/Fixing-%22Image-Events%22-AppleScripts-Broken-in-macOS-10.15-Catalina.html).

1. Open System Preferences
2. Go to Security & Privacy
3. Click the Lock icon to unlock the panel
4. Click the Privacy tab at the top
5. Scroll down to and select Full Disk Access
6. In the right side click the "+" button
7. In the navigator popup, select the following:
	* /System/Library/CoreServices/Image Events

### To use directly in macOS Finder

1. Select image files in Finder
2. Drag-n-drop onto add_border droplet

### Origins

This script originates from Apple’s Recursive Image File Processing Droplet template. You can read more about it in the [Mac Automation Scripting Guide to Process Dropped Files and Folders](https://developer.apple.com/library/content/documentation/LanguagesUtilities/Conceptual/MacAutomationScriptingGuide/ProcessDroppedFilesandFolders.html). It formats and executes terminal `sips` command to edit the selected image files.
1. Open Apple ScriptEditor
2. Navigate to menu option File > New from Template > Droplets > Recursive Image File Processing Droplet
