## Applescript to search a Capture One Catalog for variants not in a User Collection
## Version 1.12.06     NOTICE Best effort Support, no warranty
## Copyright 2020 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
## 

-- ***To Setup
-- Start Script Editor, open a new (blank) file, copy and paste both parts into one Script Editor Document, compile (hammer symbol) and save.
-- Best if you make "Scripts" folder somewhere in your Documents or Desktop
-- This file is suitable to use as an application in Capture One Pro's Script Menu

-- *** Operation in Script Editor
-- Open  the compiled and saved file
-- Open the Script Editor log window, and select the messages tab
-- Select only a small number of variants, from any collection

-- Run the script
-- Results can appear in Notifications, the Script Editor Log and the Text Edit document when the search for each variant is completed, and in the Clipboard and Display Dialog Window when all searching is completed
--
-- The user may elect to set defaults for enabling or disabling results in Notifications, TextEdit and Script Editor by setting the "enable" variables at beginning of the script
-- The user may change the default amount of reporting by setting the "debugLogLevel" and "ResultsFileMaxDebug" variables at beginning of the script
-- If you are having some issues, then set debugLogLevel to 3 and send me the results from Script Editors log window, or Text Edit.

## Values in this section are safe to change, within limits indicated. Support is likely but no commitment


use AppleScript version "2.5"
use scripting additions

set debugLogLevel to 0 --                 1...6   Set to 1 to allow changing other settings. Increasing Values result in increasing amounts of debug data that takes longer to report
set enableResultsFile to true --          Enables Results in Textedit file (true/false)
set enableResultsByDialog to false --      (true/false)
set enableResultsByClipboard to false --   (true/false)
set enableNotifications to false --         (true/false)
set ResultsFileMaxDebug to 3 --         1...6  
set maxProgressLevel to 3 --         Levels shown in Script Editor's progress bar  1...6
set maxSearchlevel to 100 --         Reduce if you only want to search top level collections or if you start getting stack overflow messages  Range 1....... Verified to 100
set enableFastGUI to true --             (true/false)


## ***** Not safe to change stuff below this line, unless you have some background in SW development. 
## I generally won't help much if you change stuff below this line. I may explain the design intent.

tell application "Finder" to get count of windows
tell application "TextEdit" to get count of windows
tell application "System Events" to set parent_name to name of current application

set script_path to (path to me) as text
set astid to AppleScript's text item delimiters
try
   set AppleScript's text item delimiters to ":"
   if (script_path ends with ":") then
      set Script_Title to text item -2 of script_path
   else
      set Script_Title to text item -1 of script_path
   end if
end try
set AppleScript's text item delimiters to astid
--set Script_Title to (get name of me)

set SE_Parent to (parent_name = "Script Editor")
set SE_Logging to SE_Parent as boolean
set COPisParent to (get parent_name begins with "Capture One") -- version independent
set Result_DocName to "CO_Image_Search.txt"
set Result_ProjectName to "ScriptSearchResults"


set ResultMethod to my InitializeLoqqing3(Result_DocName, Script_Title)
loq_Results2(0, false, ("Started from: " & parent_name & "  Action: Find Images not in a User Collection"))

set minCOPversion to "12"
set maxCOPversion to "21"
validateCOP2(minCOPversion, maxCOPversion)
validateCOPdoc3({"catalog", "session"})

tell application "Capture One 20" to set thisDocRef to get current document

set notFoundNameList to searchDocument(thisDocRef)
set countNotFound to (count of notFoundNameList)

if 0 < countNotFound then
   set displayString to (get countNotFound as text) & " Images not found in a User Collection: " & joinListToString(notFoundNameList, ", ")
else
   set displayString to "All Images found in a User Collection"
end if

loq_Results2(0, true, displayString)

return

## Script Specific Handlers ########

on searchDocument(thisDocRef)
   ## Copyright 2019 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## Seach each user collection, and scheck the results
   global theImageListSize, theImageList, debugLogLevel
   local theCollRefList, countImages, theCollKindList, theCollCollRefList, thisCollCollList, theDocName
   
   ## bulk collection of information improves speed
   
   tell application "Capture One 20" to tell thisDocRef to tell (every collection whose user is true) to ¬
      set {theCollRefList, theCollCollRefList, theCollKindList} to ¬
         {it, its collection, my convertKindList(its kind)}
   
   ## The image list is created, but nothing gets changed here
   tell application "Capture One 20" to tell thisDocRef to tell collection "All Images" to set countImages to (get count of images)
   set {theImageList, theImageListSize} to {{}, 0} -- initialise and clear the Image list
   makeImageList(countImages)
   
   tell application "Capture One 20" to tell thisDocRef to set theDocName to its name
   display notification "Finding Images not in a user collection: " & countImages & " images of " & theDocName
   
   display notification "Checking each collection"
   
   local thisCollRef, thisCollImageIDList, collCtr
   repeat with collCtr from 1 to (count of theCollRefList)
      set thisCollRef to (get contents of theCollRefList's item collCtr) -- a CO reference
      tell application "Capture One 20" to tell thisDocRef to set thisCollImageIDList to get id of images of thisCollRef
      indexCollection(thisCollRef, (theCollKindList's item collCtr), (theCollCollRefList's item collCtr), thisCollImageIDList)
   end repeat
   
   display notification "Done Checking Collections, Now the analysis .."
   
   local everyImageIDList, notFoundIdList, notFoundIndexList, notFoundNameList, vctr, theImageID, theImageNameList, anIndex
   tell application "Capture One 20" to tell thisDocRef to tell collection "All Images" to set {everyImageIDList, everyImageNameList} to {id of images, name of images}
   set notFoundNameList to {}
   repeat with vctr from 1 to countImages
      set thisImageID to (get (everyImageIDList's item vctr) as integer)
      if thisImageID > theImageListSize then
         copy (everyImageNameList's item vctr) to end of notFoundNameList
      else
         if not (get theImageList's item thisImageID) then ¬
            copy (everyImageNameList's item vctr) to end of notFoundNameList
      end if
   end repeat
   
   set {theImageList, everyImageNameList, everyImageIDList} to {null, null, null} -- clear large lists to control AppleScript's memory utilization
   
   return notFoundNameList
end searchDocument

on indexCollection(ThisRef, theCollKind, theCollRefList, theIDStringList)
   ## Copyright 2019 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   global theImageListSize, theImageList, debugLogLevel
   
   ## index the images if the collection holds images
   local anIDString, theImageID
   tell application "Capture One 20" to tell ThisRef to if (0 < (get count of images)) then
      repeat with anIDString in theIDStringList
         set theImageID to (get anIDString as integer)
         if theImageID > theImageListSize then my makeImageList(theImageID)
         set theImageList's item theImageID to true
      end repeat
   end if
   
   if (0 = (get count of theCollRefList)) or ("project" = theCollKind) then return -- don't search below this collection if it is a project, or has no subcollections
   ## This handler only handles user collections, no need to check subcollections
   tell application "Capture One 20" to tell ThisRef to set {theCollCollRefList, theCollImageIDList, theCollKindList} to ¬
      {every collection's collection, every collection's image's id, my convertKindList(every collection's kind)}
   
   local collCtr
   repeat with collCtr from 1 to (get count of theCollRefList)
      indexCollection((theCollRefList's item collCtr), (theCollKindList's item collCtr), (theCollCollRefList's item collCtr), (theCollImageIDList's item collCtr))
   end repeat
end indexCollection

on makeImageList(newImageListSize)
   global theImageListSize, theImageList, debugLogLevel
   ## make some extras to avoid calling this repeatedly
   
   set newImageListSize to newImageListSize + 24
   if 0 < theImageListSize then
      if newImageListSize > theImageListSize then
         set theImageList to theImageList & makeList3((newImageListSize - theImageListSize), false)
         if debugLogLevel ≥ 2 then log "Increased image list from " & theImageListSize & " to " & newImageListSize
      end if
   else
      set theImageList to makeList3(newImageListSize, false)
   end if
   set theImageListSize to get count of theImageList
end makeImageList

on finalCleanup()
   global theImageList, everyImageNameList, everyImageIDList
   set {theImageList, everyImageNameList, everyImageIDList} to {null, null, null}
end finalCleanup

###########################################################################################################################
## Capture One General Handlers  Version 2018/12/16

on validateCOP2(minCOPversionstr, maxCOPversionstr)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract and check basic information about the Capture One application
   
   global debugLogLevel, theAppName, copVersion, copDetailedVersion, enableNotifications
   tell application "System Events"
      set COPProcList to every process whose name begins with "Capture One" and background only is false
      if debugLogLevel ≥ 2 then
         set COPProcNameList to name of every process whose name begins with "Capture One" and background only is false
         my loq_Results2(2, false, ("COP Processes:" & COPProcNameList))
      end if
   end tell
   if (count of COPProcList) = 0 then my loqqed_Error_Halt3(true, "COP is not running")
   if (count of COPProcList) ≠ 1 then my loqqed_Error_Halt3(true, "Unexpected: >1 COP instances")
   set theAppRef to item 1 of COPProcList
   tell application "System Events" to set theAppName to ((get name of theAppRef) as text)
   tell application "System Events" to set copDetailedVersion to get version of my application theAppName
   
   tell application "Capture One 20" to set copVersion to (get app version)
   
   if debugLogLevel ≥ 2 then
      tell application "System Events"
         my loq_Results2(2, false, ("All Processes: " & (get my joinListToString((get name of every process whose background only is false), ", "))))
      end tell
      loq_Results2(2, false, ("theAppName: " & theAppName))
      loq_Results2(2, false, ("COP Version: " & copVersion))
      loq_Results2(2, false, ("COP Detailed Version: " & copDetailedVersion))
   end if
   
   set numCOPversion to (splitstringtolist((word -1 of copVersion), "."))
   set minCOPversion to (splitstringtolist(minCOPversionstr, "."))
   set maxCOPversion to (splitstringtolist(maxCOPversionstr, "."))
   
   set digit_mult to 1000000
   set Version_digit to 0
   repeat with dig_ctr from 1 to count of numCOPversion
      set digit_mult to digit_mult / 100
      set Version_digit to Version_digit + (get item dig_ctr of numCOPversion as integer) * digit_mult
   end repeat
   
   set digit_mult to 1000000
   set min_digit to 0
   repeat with dig_ctr from 1 to count of minCOPversion
      set digit_mult to digit_mult / 100
      set min_digit to min_digit + (get item dig_ctr of minCOPversion as integer) * digit_mult
   end repeat
   
   set digit_mult to 1000000
   set max_digit to 0
   set digit_count to count of maxCOPversion
   repeat with dig_ctr from 1 to digit_count
      set digit_mult to digit_mult / 100
      set max_digit to max_digit + (get item dig_ctr of maxCOPversion as integer) * digit_mult
      if (dig_ctr = digit_count) then set max_digit to max_digit + digit_mult
   end repeat
   
   if (Version_digit < min_digit) or (Version_digit ≥ max_digit) then
      if enableNotifications then display notification "COP Version is unsupported"
      my loqqed_Error_Halt3(true, "This COP Version is " & copDetailedVersion & " - the supported COP versions are from " & minCOPversionstr & " to " & maxCOPversionstr)
   end if
   
   tell application "System Events" to set frontmost of process theAppName to true
   loq_Results2(1, false, ("Capture One version: " & copDetailedVersion))
end validateCOP2

on validateCOPdoc3(COP_kind_list)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract and check basic information about the current document
   
   global debugLogLevel, COPDocName, COPDocKind_s, COPDocRef, theAppName
   
   try
      tell application "Capture One 20" to set COPDocName to get name of current document
   on error
      loqqed_Error_Halt3(true, "The Script could not retrieve the Capture One document - Perhaps a Capture One Dialog window is open?")
   end try
   
   tell application "Capture One 20"
      set current_doc_kind_p to (get kind of current document)
      set current_doc_ref_list to (get every document whose name is COPDocName and kind is current_doc_kind_p)
      set number_of_hits to count of current_doc_ref_list
   end tell
   set COPDocKind_s to convertKindList(current_doc_kind_p)
   if COPDocKind_s = "session" then set COPDocName to text 1 thru ((offset of "." in COPDocName) - 1) of COPDocName
   
   loq_Results2(2, false, ("Found Documents: " & number_of_hits))
   
   if COP_kind_list does not contain COPDocKind_s then loq_Results2(0, false, (COPDocName & " is a " & COPDocKind_s & " -- unsupported type of document"))
   
   if number_of_hits = 0 then
      loqqed_Error_Halt3(false, "Could not find find " & COPDocKind_s & COPDocName)
      error "Could not find find " & COPDocKind_s & COPDocName
   else if number_of_hits > 1 then
      loqqed_Error_Halt3(false, "Found more than one " & COPDocKind_s & " with the name " & COPDocName)
      error "Found more than one " & COPDocKind_s & " with the name " & COPDocName
   else
      tell application "Capture One 20" to set COPDocRef to item 1 of current_doc_ref_list
   end if
   
   loq_Results2(1, false, ("CO Document: " & COPDocKind_s & " " & COPDocName))
end validateCOPdoc3

on validateCOPcollections3()
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract basic information regarding the current collection, and thhe top level collections
   global debugLogLevel, COPDocName, COPDocKind_s, COPDocRef, enableNotifications
   global everyTopCollection, namesTopCollections, kindsTopCollections_s, countTopCollections, selectedCollectionRef, selectedCollectionIndex, kindSelectedCollection_s, nameSelectedCollection
   global selectedCollectionMirroredAtTopLast, selectedCollectionIsUser, bottomUserCollectionIndex, topUserCollectionIndex
   -- selectedCollectionMirroredAtTopLast replaces selectedCollectionAtTopEnd
   -- bottomUserCollectionIndex, topUserCollectionIndex replaces indexInCatalog
   
   tell application "Capture One 20" to tell COPDocRef
      set selectedCollectionRef to get current collection
      if (missing value = selectedCollectionRef) then
         set current collection to collection 1
         set selectedCollectionRef to get current collection
      end if
      set nameSelectedCollection to name of selectedCollectionRef
      set kindSelectedCollection_s to my convertKindList(kind of selectedCollectionRef)
      set {everyTopCollection, namesTopCollections} to {it, name} of every collection
      set kindsTopCollections_s to my convertKindList(kind of every collection)
   end tell
   set countTopCollections to count of namesTopCollections
   
   repeat with collectionCounter from 1 to countTopCollections
      if (nameSelectedCollection = item collectionCounter of namesTopCollections) and ¬
         (kindSelectedCollection_s = item collectionCounter of kindsTopCollections_s) then
         set selectedCollectionIndex to collectionCounter
         exit repeat
      end if
   end repeat
   
   if COPDocKind_s = "catalog" then
      repeat with collectionCounter from countTopCollections to 1 by -1
         if ("in Catalog" = item collectionCounter of namesTopCollections) and ¬
            ("smart album" = item collectionCounter of kindsTopCollections_s) then
            set topUserCollectionIndex to collectionCounter - 1
            exit repeat
         end if
      end repeat
      repeat with collectionCounter from 1 to countTopCollections
         if ("Trash" = item collectionCounter of namesTopCollections) and ¬
            ("smart album" = item collectionCounter of kindsTopCollections_s) then
            set bottomUserCollectionIndex to collectionCounter + 1
            exit repeat
         end if
      end repeat
      set selectedCollectionMirroredAtTopLast to ¬
         (selectedCollectionIndex = countTopCollections) and ¬
         ({"catalog folder", "favorite"} does not contain last item of kindsTopCollections_s)
      
      set selectedCollectionIsUser to ¬
         (selectedCollectionMirroredAtTopLast or ((selectedCollectionIndex ≥ bottomUserCollectionIndex) and ¬
            (selectedCollectionIndex ≤ topUserCollectionIndex)))
      
   else if COPDocKind_s = "session" then
      repeat with collectionCounter from countTopCollections to 1 by -1
         if ("favorite" ≠ item collectionCounter of kindsTopCollections_s) then
            set topUserCollectionIndex to collectionCounter
            exit repeat
         end if
      end repeat
      repeat with collectionCounter from 1 to countTopCollections
         if ("Trash" = item collectionCounter of namesTopCollections) and ¬
            ("favorite" = item collectionCounter of kindsTopCollections_s) then
            set bottomUserCollectionIndex to collectionCounter + 1
            exit repeat
         end if
      end repeat
      set selectedCollectionMirroredAtTopLast to false
      set selectedCollectionIsUser to ({"project", "album", "group", "smart album"} contains kindSelectedCollection_s)
   end if
end validateCOPcollections3

on convertKindList(kind_list)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Many releases of Capture One return the chevron form of the property "kind" when AppleScript is run as an Application
   ## This script converts a list of kind enums to text, handling the chevron form correctly
   ## Assume that the list either contains the plain text enums or the chevron form enums, but not both
   ## Assume that the list contains the kind enums from the same class
   
   if "list" = (get (class of kind_list) as text) then
      set input_is_list to true
   else
      set kind_list to {kind_list}
      set input_is_list to false
   end if
   set kind_s1 to (item 1 of kind_list) as text
   
   set kind_s_list to {}
   set fail_flag to false
   
   if "«" ≠ (get text 1 of kind_s1) then
      repeat with Kind_item in kind_list
         set the end of kind_s_list to (get Kind_item as text)
      end repeat
   else
      ## minimum 6 characters for valid logic. Normally 19
      if 6 > length of kind_s1 then loqqed_Error_Halt3(true, "convertKind received an unexpected Kind string: " & kind_s1)
      set code_start to (get length of kind_s1) - 4
      set kind_type to get (text code_start thru (code_start + 1) of kind_s1)
      
      repeat with Kind_item in kind_list
         set kind_s to Kind_item as text
         set kind_code to get (text code_start thru (code_start + 3) of kind_s)
         
         if kind_type = "CC" then ## Collection Kinds
            if kind_code = "CCpj" then
               set the end of kind_s_list to "project"
            else if kind_code = "CCgp" then
               set the end of kind_s_list to "group"
            else if kind_code = "CCal" then
               set the end of kind_s_list to "album"
            else if kind_code = "CCsm" then
               set the end of kind_s_list to "smart album"
            else if kind_code = "CCfv" then
               set the end of kind_s_list to "favorite"
            else if kind_code = "CCff" then
               set the end of kind_s_list to "catalog folder"
            else
               set fail_flag to true
            end if
            
         else if kind_type = "CL" then ## Layer Kinds
            if kind_code = "CLbg" then
               set the end of kind_s_list to "background"
            else if kind_code = "CLnm" then
               set the end of kind_s_list to "adjustment"
            else if kind_code = "CLcl" then
               set the end of kind_s_list to "clone"
            else if kind_code = "CLhl" then
               set the end of kind_s_list to "heal"
            else
               set fail_flag to true
            end if
            
         else if kind_type = "CR" then ## Watermark Kinds
            if kind_code = "CRWn" then
               set the end of kind_s_list to "none"
            else if kind_code = "CRWt" then
               set the end of kind_s_list to "textual"
            else if kind_code = "CRWi" then
               set the end of kind_s_list to "imagery"
            else
               set fail_flag to true
            end if
            
         else if kind_type = "CO" then ## Document Kinds
            if kind_code = "COct" then
               set the end of kind_s_list to "catalog"
            else if kind_code = "COsd" then
               set the end of kind_s_list to "session"
            else
               set fail_flag to true
            end if
         else
            set fail_flag to true
         end if
         
         if fail_flag then loqqed_Error_Halt3(true, "convertKindList received an unexpected Kind string: " & kind_s)
      end repeat
      
   end if
   
   if input_is_list then
      return kind_s_list
   else
      return item 1 of kind_s_list
   end if
   
end convertKindList

on InitializeResultsCollection(nameResultProject, nameResultAlbumRoot, Coll_Init_Text)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Sets up a project and albums for collecting images
   
   global debugLogLevel, COPDocRef, Ref2ResultAlbum, enableNotifications
   
   tell application "Capture One 20" to tell COPDocRef
      if not (exists collection named (get nameResultProject)) then
         set ref2ResultProject to make new collection with properties {kind:project, name:nameResultProject}
      else
         if ("project" = my convertKindList(kind of (get collection named nameResultProject))) then
            set ref2ResultProject to collection named nameResultProject
         else
            my loqqed_Error_Halt3(true, ("A collection named \"" & nameResultProject & "\" already exists, and it is not a project."))
         end if
      end if
   end tell
   
   set coll_ctr to 1
   set nameResultAlbum to nameResultAlbumRoot & "_" & (get short date string of (get current date)) & "_"
   repeat
      tell application "Capture One 20" to tell ref2ResultProject
         if not (exists collection named (get nameResultAlbum & coll_ctr)) then
            set nameResultAlbum to (get nameResultAlbum & coll_ctr)
            set Ref2ResultAlbum to make new collection with properties {kind:album, name:nameResultAlbum}
            exit repeat
         else
            set coll_ctr to coll_ctr + 1
         end if
      end tell
   end repeat
   
   if enableNotifications then display notification (Coll_Init_Text & " " & nameResultProject & ">" & nameResultAlbum)
   loq_Results2(1, false, (Coll_Init_Text & "  " & nameResultProject & ">" & nameResultAlbum))
   
end InitializeResultsCollection

###########################################################################################################################
## General Handlers  Version 2019/11/6

on InitializeLoqqing3(DocName_Ext, sourceTitle)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose handler to set up Text Editor document for logging results
   
   global debugLogLevel, Script_Title, Result_Doc_ref, SE_Logging, enableResultsFile, enableResultsByDialog, enableResultsByClipboard, DialogTextList, enableNotifications
   global initEnableResultsByDialog, initEnableResultsByClipboard
   
   tell current application to set date_string to (current date) as text
   set LogMethods to {}
   set LogHeader to (sourceTitle & " results on " & date_string)
   
   if enableResultsFile then
      set end of LogMethods to DocName_Ext
      set targetFileWasCreated to false
      set ResultDocIsOpen to false
      
      ## Check if TextEdit is already open and has the document open
      tell application "System Events" to set TextEditlist to get background only of every application process whose name is "TextEdit"
      if (0 < (count of TextEditlist)) and not item 1 of TextEditlist then
         if (DocName_Ext is in (get name of documents of application "TextEdit")) then
            tell application "TextEdit" to set Result_Doc_ref to document DocName_Ext
            set ResultDocIsOpen to true
         end if
      end if
      
      if not ResultDocIsOpen then
         -- create the document and the folder if necessary
         -- Do not use finder to test for the file existence because it has a bug that ignores leading 0's 
         -- https://www.macscripter.net/viewtopic.php?id=45178
         
         set target_folder_parent_a to alias (get path to desktop folder as text)
         set target_folder_parent_p to get POSIX path of target_folder_parent_a
         set target_folder_name to "ScriptReports"
         set target_folder_p to (target_folder_parent_p & target_folder_name)
         set Result_Doc_Path_p to target_folder_p & "/" & DocName_Ext
         
         try
            set Result_Doc_Path_a to (get alias POSIX file Result_Doc_Path_p)
         on error
            try
               set target_folder_a to (get alias POSIX file target_folder_p) --x1
            on error
               tell application "Finder" to set newFolder to make new folder at target_folder_parent_a with properties {name:target_folder_name}
               set target_folder_a to newFolder as alias
            end try
            tell application "Finder" to set newFile to make new file at target_folder_a with properties {name:DocName_Ext}
            set Result_Doc_Path_a to newFile as alias
            set targetFileWasCreated to true
         end try
         
         set First_line to ("Created by " & Script_Title & " on " & date_string)
         tell application "TextEdit" -- open the document and add the first line if empty
            activate
            set Result_Doc_ref to open Result_Doc_Path_a
            set ResultDocIsOpen to true -- For consistency
            tell text of Result_Doc_ref
               if targetFileWasCreated then
                  set paragraph 1 to First_line & return & return
                  tell me to if 2 ≤ debugLogLevel then log Result_Doc_Path_p & ": " & First_line
               else
                  try
                     if (0 = (count of paragraphs)) then set paragraph 1 to First_line & return & return
                  on error
                     set paragraph 1 to First_line & return & return
                  end try
               end if
            end tell
         end tell
      end if
      
      tell application "TextEdit" to tell text of Result_Doc_ref to ¬
         set paragraph (1 + (count paragraphs)) to return & LogHeader & return
      
   end if
   
   if enableResultsByDialog then
      set end of LogMethods to "Display Dialog"
      try
         initEnableResultsByDialog
         set DialogTextList to DialogTextList & ""
         set initEnableResultsByDialog to false
      on error
         set DialogTextList to {LogHeader}
         set initEnableResultsByDialog to true
      end try
   end if
   
   if enableResultsByClipboard then
      set end of LogMethods to "Clipboard"
      try
         initEnableResultsByClipboard
         set initEnableResultsByClipboard to false
      on error
         set the clipboard to LogHeader
         set initEnableResultsByClipboard to true
      end try
   end if
   
   if SE_Logging then -- if Script Editor logging, then open the Log History window
      try
         tell application "System Events" to tell application process "Script Editor"
            if (get name of windows) does not contain "log History" then
               click menu item "Log History" of menu "Window" of menu bar 1
            end if
         end tell
      end try
      set end of LogMethods to " Script Editor Log"
   end if
   
   set LogMethods_S to joinListToString(LogMethods, ", ")
   loq_Results2(2, false, ("Results by " & LogMethods_S))
   
   return LogMethods_S
end InitializeLoqqing3

on loq_Results2(thisLogDebugLevel, MakeFront, log_Text)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose handler for logging results
   ## log results if the debug level of the message is below the the threshold set by debugLogLevel
   ## log the results by whatever mechanism is ebabled - {Script Editor Log, Text Editor Log, Display Dialog}
   
   global Result_Doc_ref, debugLogLevel, SE_Logging, parent_name, ResultsFileMaxDebug, enableResultsFile, enableResultsByDialog, enableResultsByClipboard, DialogTextList
   if thisLogDebugLevel > debugLogLevel then return
   
   set log_Text_S to joinListToString(log_Text, ", ")
   
   if enableResultsFile and ((thisLogDebugLevel ≤ ResultsFileMaxDebug) or not SE_Logging) then
      tell application "TextEdit" to tell text of Result_Doc_ref to ¬
         set paragraph (1 + (count paragraphs)) to ((log_Text_S as text) & return)
      tell application "System Events" to if MakeFront then set frontmost of process "TextEdit" to true
   end if
   
   if enableResultsByDialog and (1 ≥ thisLogDebugLevel) then
      set DialogTextList to DialogTextList & log_Text_S
      tell application "System Events" to set frontmost of process parent_name to true
      display dialog joinListToString(DialogTextList, return)
   end if
   
   if enableResultsByClipboard and (1 ≥ thisLogDebugLevel) then set the clipboard to ((get the clipboard) & return & (log_Text_S as text))
   if SE_Logging then log (log_Text_S as text)
   
end loq_Results2

on loqqed_Error_Halt3(createErrorHere, error_text)
   ## General purpose handler for logging during script termination
   ##
   ## found an error somewhere, so now we exit in a controlled fashion
   ## set createError to "false" to create a local error  instead of here
   global debugLogLevel, Script_Title, enableNotifications
   tell current application to set date_string to (current date) as text
   finalCleanup()
   if enableNotifications then display notification error_text
   if createErrorHere then
      loq_Results2(0, true, ("Script \"" & Script_Title & "\" has halted at " & date_string & return & "Reason: " & error_text & return & return))
      error error_text
   else
      loq_Results2(0, true, ("Script \"" & Script_Title & "\" is exitting at " & date_string & "Reason: " & error_text & return))
   end if
end loqqed_Error_Halt3

on splitstringtolist(theString, theDelim)
   ## Public Domain
   set astid to AppleScript's text item delimiters
   try
      set AppleScript's text item delimiters to theDelim
      set theList to text items of theString
   on error
      set AppleScript's text item delimiters to astid
   end try
   set AppleScript's text item delimiters to astid
   return theList
end splitstringtolist

to joinListToString(theList, theDelim)
   ## Public Domain
   set theString to ""
   set astid to AppleScript's text item delimiters
   try
      set AppleScript's text item delimiters to theDelim
      set theString to theList as string
   on error
      set AppleScript's text item delimiters to astid
   end try
   set AppleScript's text item delimiters to astid
   return theString
end joinListToString

on makeList3(listLength, theElement)
   -- Note that the theElement can even be a List
   if listLength = 0 then return {}
   if listLength = 1 then return {theElement}
   
   set theList to {theElement}
   repeat while (count of theList) < listLength / 2
      copy contents of theList to ListB
      copy theList & ListB to theList
   end repeat
   copy contents of theList to ListB
   return (theList & items 1 thru (listLength - (count of ListB)) of ListB)
end makeList3