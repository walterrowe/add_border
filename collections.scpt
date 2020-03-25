## Applescript to search a COP 11 Catalog for Collections containing selected variants
## Version 1.29 !! NO GUARANTEE OF SUPPORT !!  Best effort
## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
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
-- There is a GUI which allows the user to modify settings
-- If you are having some issues, then set debugLogLevel to 3 and send me the results from Script Editors log window, or Text Edit.

## Values in this section are safe to change, within limits indicated. Support is likely but no commitment

use AppleScript version "2.5"
use scripting additions


set debugLogLevel to 0 --                 1...6   Set to 1 to allow changing other settings. Increasing Values result in increasing amounts of debug data that takes longer to report
set enableResultsFile to true --          Enables Results in Textedit file (true/false)
set enableResultsByDialog to false --      (true/false)
set enableResultsByClipboard to true --   (true/false)
set enableNotifications to false --         (true/false)
set ResultsFileMaxDebug to 3 --         1...6  
set maxProgressLevel to 3 --         Levels shown in Script Editor's progress bar  1...6
set maxSearchlevel to 100 --         Reduce if you only want to search top level collections or if you start getting stack overflow messages  Range 1....... Verified to 100
property searchTimePerVariant : 0 --      The search time in seconds. The value from the previous run is used
set WarnMaxVariants to 6 --          1...n (number of variants) If you make this large you will get no warning when the script may lock up CO for a long time
set WarnSearchTime to 30 --         1...n (seconds) If you make this large you will get no warning when the script may lock up CO for a long time
set SearchDiskFileSystem to false --      untested probably safe (true/false)
set SearchRecentCaptures to false --      (true/false)
set SearchTrash to false --            untested probably safe (true/false)
set ExcludedSubCollectionNames to {} -- The collections in this list will not be searched
set StopSearchAtProject to false --      (true/false)
set enableFastGUI to true --             (true/false)


## ***** Not safe to change stuff below this line, unless you have some background in SW development. 
## I generally won't help much if you change stuff below this line. I may explain the design intent.
set SearchRecentImports to false --      doesn't work in CO 11 (true/false)

tell application "System Events" to set parent_name to name of current application

tell application "Finder"
   set script_path to path to me
   set Script_Title to name of file script_path as text
end tell

set SE_Parent to (parent_name = "Script Editor")
set SE_Logging to SE_Parent as boolean
set COPisParent to (parent_name = "Capture One 11")
set Result_DocName to "COP_Image_Search.txt"
set Result_ProjectName to "ScriptSearchResults"


set ResultMethod to my InitializeLoqqing3(Result_DocName, Script_Title)
loq_Results2(0, false, ("Started from: " & parent_name & "  Action: Find collections where variants are located"))

set minCOPversion to "11.2"
validateCOP(minCOPversion)
validateCOPdoc3({"catalog", "session"})
validateCOPcollections3()

tell application "Capture One 11" to set countSelectedVariants to get count of selected variants
loq_Results2(3, false, ("In this window " & countSelectedVariants & " Selected Variants"))

if countSelectedVariants = 0 then
   display alert "No images selected -  select one or more images"
   my loqqed_Error_Halt3(true, "No images selected")
end if

CO_settingHandler()

set ExcludedTopCollectionNames to {"In Catalog", "Catalog", "All Images"}
if SearchDiskFileSystem then
   set ExcludedTopCollectionKinds to {""}
else
   set ExcludedTopCollectionKinds to {"catalog folder"}
end if
if not SearchRecentImports then set ExcludedTopCollectionNames to ExcludedTopCollectionNames & "Recent Imports"
if not SearchRecentCaptures then set ExcludedTopCollectionNames to ExcludedTopCollectionNames & "Recent Captures"
if not SearchTrash then set ExcludedTopCollectionNames to ExcludedTopCollectionNames & "Trash"

if not SearchDiskFileSystem then
   set max_search_coll_index to topUserCollectionIndex
else
   if selectedCollectionMirroredAtTopLast then
      set max_search_coll_index to countTopCollections - 1
   else
      set max_search_coll_index to countTopCollections
   end if
end if

set Mark1 to GetTick_Now()

set estimatedDuration to countSelectedVariants * searchTimePerVariant

tell application "System Events" to set frontmost of process parent_name to true

if (estimatedDuration > WarnSearchTime) then
   try
      display dialog "Are you sure you want to search for " & countSelectedVariants & " variants? Estimated time: " & estimatedDuration & " seconds" with icon caution
      set Mark1 to GetTick_Now()
   on error errmess
      my loqqed_Error_Halt3(true, "User Cancelled the Script")
   end try
else if (searchTimePerVariant = 0) and (countSelectedVariants > WarnMaxVariants) then
   try
      display dialog "Are you sure you want to search for " & countSelectedVariants & " variants? " with icon caution
      set Mark1 to GetTick_Now()
   on error errmess
      my loqqed_Error_Halt3(true, "User Cancelled the Script")
   end try
end if

if SE_Logging then set progress description to "Setting up"

tell application "Capture One 11"
   set selectedVariantsNameList to get name of variants whose selected is true
   set selectedVariantsIDList to get id of variants whose selected is true
   set selectedVariantsList to (get selected variants)
end tell

loq_Results2(4, false, (selectedVariantsNameList))

set countFoundVariants to 0
set countProcessedVariants to 0
set Mark2 to GetTick_Now()

set search_executed to false
if SE_Logging then
   set progress total steps to countSelectedVariants
   set progress completed steps to 0
   set progress description to "Searching for Variants ..."
end if
repeat with v_Counter from 1 to countSelectedVariants
   
   set thisVariant to item v_Counter of selectedVariantsList
   set searchedVariantName to item v_Counter of selectedVariantsNameList
   set searchedVariantID to item v_Counter of selectedVariantsIDList
   
   if SE_Logging then
      set progress completed steps to v_Counter
      set progress description to "Searching for Variant " & searchedVariantName
   end if
   
   loq_Results2(2, false, (return & "Searching for " & searchedVariantName))
   
   set isFound to false
   set foundPaths to ""
   set pathSeparator to ""
   set searchLevel to 0
   set nextSearchLevel to searchLevel + 1
   
   if nextSearchLevel ≤ maxSearchlevel then
      repeat with c_Counter from 1 to max_search_coll_index
         set thisTopCollection to item c_Counter of everyTopCollection
         set topCollName to item c_Counter of namesTopCollections
         set topCollKind_s to convertKind(item c_Counter of kindsTopCollections_p)
         
         if (ExcludedTopCollectionNames does not contain topCollName) and ¬
            (ExcludedTopCollectionKinds does not contain topCollKind_s) then
            set search_executed to true
            loq_Results2(3, false, ("Search: " & topCollKind_s & " " & topCollName))
            
            set resultPaths to my search_collection(thisTopCollection, nextSearchLevel, topCollName, topCollKind_s) -- The search
            set resultCount to count of resultPaths
            
            loq_Results2(3, false, ("Result of search in " & topCollName & " (" & topCollKind_s & ") is: " & resultPaths))
            
            if resultCount > 0 then
               set isFound to true
               repeat with p_counter from 1 to count of resultPaths
                  set foundPaths to foundPaths & pathSeparator & item p_counter of resultPaths
                  set pathSeparator to return -- after the first path is added the separator become the newline character /r
               end repeat
            end if
         else
            loq_Results2(3, false, ("Reject: " & topCollKind_s & " " & topCollName))
            
         end if
      end repeat
   end if
   
   if not search_executed then -- there are no top level collections to search - don't keep searching
      if enableNotifications then display notification "No Top Level Collections to Search"
      my loqqed_Error_Halt3(true, "No Top Level Collections to Search")
   end if
   
   if isFound then
      loq_Results2(0, false, (return & searchedVariantName & " was found in:" & return & foundPaths))
      set countFoundVariants to countFoundVariants + 1
   else
      loq_Results2((return & searchedVariantName & " was not found"), false, 0)
      if enableNotifications then display notification searchedVariantName & " was not found"
   end if
   set countProcessedVariants to countProcessedVariants + 1
end repeat

set Mark3 to GetTick_Now()

tell application "System Events" to set frontmost of process "Capture One 11" to true
tell current application to set date_string to (current date) as text

set elapsedTime1s to roundDecimals(Mark2 - Mark1, 3)
set elapsedTime2s to roundDecimals(Mark3 - Mark2, 3)

if countSelectedVariants > 0 then
   set searchTimePerVariant to get ((Mark3 - Mark2) / countSelectedVariants) -- update the property to provide a time estimate for the next run
   set searchTimePerVariantDisplay to roundDecimals(searchTimePerVariant, 3)
else
   set searchTimePerVariantDisplay to "--"
end if

loq_Results2(1, false, (return & "SetupTime: " & elapsedTime1s & "sec   Process time: " & elapsedTime2s & "sec   per Variant: " & searchTimePerVariantDisplay & "sec"))

loq_Results2(0, true, (return & "*** Done on " & date_string & " ***  " & countProcessedVariants & " of " & countFoundVariants & " variants found" & return & return))

finalCleanup() -- cleanup large arrays to avoid the situation where the script cannot be saved

## Arrange the windows to show results on top
tell application "System Events" to set frontmost of process "Capture One 11" to true

if enableResultsFile then
   tell application "System Events" to set frontmost of process "TextEdit" to true
else
   tell application "System Events" to set frontmost of process parent_name to true
end if

## Script Specific Handlers ########

on search_collection(thisCollection, searchLevel, thisCollName, thisCollKind)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   -- recursive handler to search a collection and it's subcollections
   -- if successful, returns a list of paths each as a text string
   
   global debugLogLevel, SE_Logging, maxSearchlevel, searchedVariantID, searchedVariantName, COPDocName, StopSearchAtProject, ExcludedSubCollectionNames, maxProgressLevel, enableNotifications
   global selectedCollectionRef, selectedCollectionIndex, kindSelectedCollection_s, nameSelectedCollection, selectedCollectionAtTopEnd, selectedCollectionIsUser
   
   set found_paths to {} -- initialise the list
   
   if ExcludedSubCollectionNames contains thisCollName then return found_paths
   
   if SE_Logging and (searchLevel ≤ maxProgressLevel) then set progress additional description to thisCollKind & " " & thisCollName
   
   tell application "Capture One 11" to tell document COPDocName to tell thisCollection to ¬
      set searchVariantIDhits to get count of (get every variant whose id is searchedVariantID)
   
   if searchVariantIDhits > 0 then
      set found_paths to found_paths & (thisCollName & "(" & thisCollKind & ")")
      if enableNotifications or (debugLogLevel ≥ 2) then display notification searchedVariantName & " found in: " & thisCollKind & " '" & thisCollName & "'"
      loq_Results2(4, false, ("Found " & searchVariantIDhits & " hits in " & thisCollKind & " " & thisCollName))
   end if
   
   set countSubcollsthisColl to 0
   set namesSubcollsthisColl to {}
   set kindsSubcollsthisColl_p to {}
   
   if {"album", "smart album"} does not contain thisCollKind then -- albums and smart albums don't contain collections
      tell application "Capture One 11" to tell document COPDocName to tell thisCollection
         set listSubcollsthisColl to every collection
         tell me to set countSubcollsthisColl to count of listSubcollsthisColl
         if countSubcollsthisColl > 0 then set {namesSubcollsthisColl, kindsSubcollsthisColl_p} to {name, kind} of every collection
      end tell
   end if
   
   if debugLogLevel ≥ 5 then
      tell application "Capture One 11" to tell document COPDocName to tell thisCollection
         set countImages to count of every image
         set countVariants to count of every variant
      end tell
      loq_Results2(debugLogLevel, false, ("In " & thisCollKind & " " & thisCollName & " at level " & searchLevel & "   found " & countImages & " images  ," & countVariants & " variants  ," & countSubcollsthisColl & " collections"))
   end if
   
   set nextSearchLevel to searchLevel + 1
   if (nextSearchLevel ≤ maxSearchlevel) and ¬
      (not (StopSearchAtProject and thisCollKind = "project")) then
      
      loq_Results2(5, false, ("Searching Collections in " & thisCollKind & " " & thisCollName & " at level " & searchLevel & "   found " & countSubcollsthisColl & " collections"))
      
      repeat with c_Counter from 1 to countSubcollsthisColl
         
         set searchColl to item c_Counter of listSubcollsthisColl
         set searchCollName to item c_Counter of namesSubcollsthisColl
         set searchCollKind_s to convertKind(get item c_Counter of kindsSubcollsthisColl_p)
         
         if selectedCollectionIsUser then
            if (searchCollName = nameSelectedCollection) and (searchCollKind_s = kindSelectedCollection_s) then
               ## Check the collection number to trap the error caused by the selected subcollection returning an invalid reference
               try
                  || of {searchColl} -- generate an error to capture error text containing the collection reference
               on error CO11_errtxt number CO11_errnbr
               end try
               
               loq_Results2(4, false, ("***Error Message " & CO11_errtxt & "; number " & CO11_errnbr & " Collection: " & c_Counter & " of " & thisCollKind & " " & thisCollName))
               
               ## extract the collection number from the error text
               if CO11_errnbr = -1728 then
                  set CollNum to (get last word of item 2 of splitstringtolist(CO11_errtxt, "of")) as integer
               else
                  loq_Results2(0, false, ("*****" & CO11_errtxt & "; number " & CO11_errnbr))
                  loqqed_Error_Halt3(true, "Unexpected error when checking collection integrity")
               end if
               
               if CollNum ≠ c_Counter then
                  ## this is the selected user subcollection, it must be accessed as the last top level collection in the document
                  if selectedCollectionIndex ≠ CollNum then
                     ## the collection reference is wrong in some other unexpected way
                     loq_Results2(4, false, ("*****" & CO11_errtxt & "; number " & CO11_errnbr))
                     loqqed_Error_Halt3(false, "Unexpected collection index when checking collection integrity: " & CollNum)
                     error "Unexpected collection index when checking collection integrity: " & CollNum
                  end if
                  ## Set the collection reference to the last top level collection
                  set searchColl to selectedCollectionRef
                  loq_Results2(3, false, ("Top Level Collection number:" & CollNum & " replaces Sub collection " & c_Counter & " of " & thisCollKind & " " & thisCollName))
               end if
               
               loq_Results2(4, false, ("Verified sub-Collection " & c_Counter & " of " & thisCollKind & " " & thisCollName))
            end if
         end if
         
         loq_Results2(4, false, ("Starting search in " & searchCollName & " (" & searchCollKind_s & ") "))
         
         set resultPaths to my search_collection(searchColl, nextSearchLevel, searchCollName, searchCollKind_s) -- the search
         set countResultPaths to count of resultPaths
         
         loq_Results2(3, false, ("Result of search in " & searchCollName & " (" & searchCollKind_s & ") is: " & resultPaths))
         
         if countResultPaths > 0 then
            repeat with p_counter from 1 to countResultPaths
               set found_paths to found_paths & (thisCollName & ">>" & item p_counter of resultPaths)
            end repeat
         end if
      end repeat
   end if
   return found_paths
end search_collection

on finalCleanup()
   ## clean up the large arrays to avoid a large stack that may prevent AppleScript from saving the script
   global selectedVariantsList, selectedVariantsNameList, selectedVariantsIDList, enableNotifications
   set selectedVariantsList to {}
   set selectedVariantsNameList to {}
   set selectedVariantsIDList to {}
end finalCleanup

on CO_settingHandler()
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## Initialisation Handler for scripts using Capture One Pro
   ## Collects and sets up information about the Script settings for the GUI handler
   
   global parent_name, Script_Title, enableFastGUI, debugLogLevel, ResultsFileMaxDebug, maxProgressLevel, countSelectedVariants
   global theAppName, copVersion, COPDocName, COPDocKind_s, COPDocRef
   global namesTopCollections, bottomUserCollectionIndex, topUserCollectionIndex, selectedCollectionRef, nameSelectedCollection, selectedCollectionIsUser, kindSelectedCollection_s
   global SearchDiskFileSystem, SearchRecentImports, SearchRecentCaptures, SearchTrash, ExcludedSubCollectionNames, StopSearchAtProject, maxSearchlevel
   global SE_Logging, SE_Parent, enableResultsFile, enableResultsByDialog, enableResultsByClipboard, enableNotifications, ResultMethod, Result_DocName
   
   set basic_setting_list to {}
   set end of basic_setting_list to {SettingName:(get copVersion & " in " & COPDocKind_s & " \"" & COPDocName & "\"")}
   set end of basic_setting_list to {SettingName:(get kindSelectedCollection_s & " \"" & nameSelectedCollection & "\" with " & countSelectedVariants & " selected variants")}
   set end of basic_setting_list to {SettingName:"Debug Level", SettingValue:debugLogLevel}
   set end of basic_setting_list to {SettingName:"Results by " & ResultMethod}
   if settingGUI_Bypass(basic_setting_list) then
      loq_Results2(0, false, joinListToString(basic_setting_list, return))
      return
   end if
   
   set helpdebugLogLevel to "the amount of debug info reported"
   set helpmaxProgressLevel to "lowest subcollection level reported in the progess bar"
   set helpmaxSearchlevel to "lowest subcollection level which is searched"
   set helpSearchDiskFileSystem to "search the OSX filesystem"
   set helpSearchTrash to "search CaptureOne Trash"
   set helpStopSearchAtProject to "do not search Albums within a Project"
   set helpExcludedCollection to "collections with these names will not be searched"
   set helpFastGUI to "enable faster GUI response, with less help text"
   
   set nameUserCollections to {nameSelectedCollection} & (get items bottomUserCollectionIndex thru topUserCollectionIndex of namesTopCollections)
   if selectedCollectionIsUser then tell application "Capture One 11" to tell selectedCollectionRef to ¬
      set nameUserCollections to nameUserCollections & (get name of every collection)
   
   set initialLogging to {(get enableNotifications as boolean), (get enableResultsByDialog as boolean), (get enableResultsFile as boolean), (get enableResultsByClipboard as boolean), (get SE_Logging as boolean)}
   
   set setting_list to {}
   set end of setting_list to {SettingID:1, SettingName:"Parent", SettingValue:parent_name, UserSet:false}
   set end of setting_list to {SettingID:2, SettingName:"Document", SettingValue:(COPDocKind_s & " " & COPDocName), UserSet:false}
   set end of setting_list to {SettingID:3, SettingName:"Collection", SettingValue:(kindSelectedCollection_s & " " & nameSelectedCollection & "    with " & countSelectedVariants & " selected variants"), UserSet:false}
   set end of setting_list to {SettingID:4, SettingName:"Search Managed Files", SettingHelp:"", SettingValue:false, UserSet:false}
   set end of setting_list to {SettingID:5, SettingName:"Search Referenced Files", SettingHelp:SearchDiskFileSystem, SettingValue:(a reference to SearchDiskFileSystem), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:7, SettingName:"Search Recent Captures", SettingValue:SearchRecentCaptures, UserSet:false}
   set end of setting_list to {SettingID:8, SettingName:"Search All Images", SettingValue:false, UserSet:false}
   set end of setting_list to {SettingID:9, SettingName:"Search COP Trash", SettingHelp:helpSearchTrash, SettingValue:(a reference to SearchTrash), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:11, SettingName:"Collections not searched", SettingHelp:"", SettingValue:(a reference to ExcludedSubCollectionNames), UserSet:true, SettingClass:"List_Text", SettingLimited:"List&Free", SettingLimit_L:nameUserCollections}
   set end of setting_list to {SettingID:12, SettingName:"Do Not Search below Projects", SettingHelp:helpStopSearchAtProject, SettingValue:(a reference to StopSearchAtProject), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:13, SettingName:"Debug Level", SettingHelp:helpdebugLogLevel, SettingValue:(a reference to debugLogLevel), UserSet:true, SettingClass:"Integer", SettingLimited:"Min_Max", SettingLimit_L:{0, 6}}
   set end of setting_list to {SettingID:14, SettingName:"Highest Debug Level in Result File", SettingHelp:"", SettingValue:(a reference to ResultsFileMaxDebug), UserSet:true, SettingClass:"Integer", SettingLimited:"Min_Max", SettingLimit_L:{1, 6}}
   set end of setting_list to {SettingID:15, SettingName:"Maximum Search Level", SettingHelp:helpmaxSearchlevel, SettingValue:(a reference to maxSearchlevel), UserSet:true, SettingClass:"Integer", SettingLimited:"Min", SettingLimit_L:{2}}
   set end of setting_list to {SettingID:16, SettingName:"Enable Results in Text Edit", SettingHelp:"", SettingValue:(a reference to enableResultsFile), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:17, SettingName:"Enable Results in Dialog", SettingHelp:"", SettingValue:(a reference to enableResultsByDialog), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:18, SettingName:"Enable Results in Clipboard", SettingHelp:"", SettingValue:(a reference to enableResultsByClipboard), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:19, SettingName:"Enable Results in Notifications", SettingHelp:"", SettingValue:(a reference to enableNotifications), UserSet:true, SettingClass:"Boolean"}
   set end of setting_list to {SettingID:20, SettingName:"Results Logged by Script Editor", SettingValue:(a reference to SE_Logging), UserSet:SE_Parent, SettingClass:"Boolean"}
   if SE_Logging then set end of setting_list to {SettingID:21, SettingName:"Maximum Level in Progress Bar", SettingHelp:helpmaxProgressLevel, SettingValue:(a reference to maxProgressLevel), UserSet:true, SettingClass:"Integer", SettingLimited:"Min_Max", SettingLimit_L:{1, 10}}
   set end of setting_list to {SettingID:23, SettingName:"Enable Fast GUI Response", SettingHelp:helpFastGUI, SettingValue:(a reference to enableFastGUI), UserSet:true, SettingClass:"Boolean"}
   
   set settings_display_string to settingGUI(setting_list)
   
   if ((not (get enableNotifications as boolean)) and (not (get enableResultsByDialog as boolean)) and (not (get enableResultsFile as boolean)) and (not (get SE_Logging as boolean))) then
      display alert "You have turned off all results notification!! " message "Turning results by Dialog Display on" as critical giving up after 30
      set enableResultsByDialog to true
   end if
   if (initialLogging ≠ {(get enableNotifications as boolean), (get enableResultsByDialog as boolean), (get enableResultsFile as boolean), (get enableResultsByClipboard as boolean), (get SE_Logging as boolean)}) then
      set ResultMethod to my InitializeLoqqing3(Result_DocName, Script_Title)
   end if
   set settings_display_string to settings_display_string & "Result Logging Methods: " & ResultMethod & return
   
   loq_Results2(1, false, (return & "Settings:" & return & settings_display_string))
   
end CO_settingHandler

## Insert the second Part Here ######

## Capture One General Handlers  Version 2018/08/28  #######

on validateCOP(minCOPversionstr)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract and check basic information about the Capture One application
   
   global debugLogLevel, theAppName, copVersion, copDetailedVersion, enableNotifications
   tell application "System Events"
      set COPProcList to every process whose name contains "Capture One" and background only is false
      if debugLogLevel ≥ 2 then
         set COPProcNameList to name of every process whose name contains "Capture One" and background only is false
         my loq_Results2(2, false, ("COP Processes:" & COPProcNameList))
      end if
   end tell
   if (count of COPProcList) = 0 then my loqqed_Error_Halt3(true, "COP is not running")
   if (count of COPProcList) ≠ 1 then my loqqed_Error_Halt3(true, "Unexpected: >1 COP instances")
   set theAppRef to item 1 of COPProcList
   tell application "System Events" to set theAppName to ((get name of theAppRef) as text)
   tell application "System Events" to set copDetailedVersion to get version of my application theAppName
   
   tell application "Capture One 11" to set copVersion to (get app version)
   
   if debugLogLevel ≥ 2 then
      --properties of application "Capture One 11"
      tell application "System Events"
         my loq_Results2(2, false, ("All Processes: " & (get my joinListToString((get name of every process whose background only is false), ", "))))
      end tell
      loq_Results2(2, false, ("theAppName: " & theAppName))
      loq_Results2(2, false, ("COP Version: " & copVersion))
      loq_Results2(2, false, ("COP Detailed Version: " & copDetailedVersion))
   end if
   
   if the theAppName ≠ "Capture One 11" then
      if enableNotifications then display notification "Wrong COP Application"
      my loqqed_Error_Halt3(true, "Found COP Application " & theAppName & " The only supported COP application is Capture One 11")
   end if
   
   set numCOPversion to (splitStringToList((word -1 of copVersion), "."))
   set minCOPversion to (splitStringToList(minCOPversionstr, "."))
   
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
   
   if Version_digit < min_digit then
      if enableNotifications then display notification "COP Version is too low"
      my loqqed_Error_Halt3(true, "This COP Version is " & (word -1 of copVersion) & " - the minimum supported COP version is " & minCOPversionstr)
   end if
   
   tell application "System Events" to set frontmost of process theAppName to true
   loq_Results2(1, false, ("Capture One version: " & copDetailedVersion))
end validateCOP

on validateCOPdoc3(COP_kind_list)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract and check basic information about the current document
   global debugLogLevel, COPDocName, COPDocKind_s, COPDocRef, enableNotifications
   
   try
      tell application "Capture One 11" to set COPDocName to get name of current document
   on error
      tell application "System Events" to tell process "Capture One 11" to ¬
         if (1 = (count of windows)) and ("" = name of window 1) then ¬
            tell window 1 to tell group 1 to tell button "Cancel" to click
      tell application "Capture One 11" to set COPDocName to get name of current document
   end try
   
   tell application "Capture One 11"
      set current_doc_kind_p to (get kind of current document)
      set current_doc_ref_list to (get every document whose name is COPDocName and kind is current_doc_kind_p)
      set number_of_hits to count of current_doc_ref_list
   end tell
   set COPDocKind_s to convertKind(current_doc_kind_p)
   if COPDocKind_s = "session" then set COPDocName to text 1 thru ((offset of "." in COPDocName) - 1) of COPDocName
   
   
   loq_Results2(2, false, ("Is: " & COPDocKind_s & "   was: " & (get current_doc_kind_p as text)))
   loq_Results2(2, false, ("Found Documents: " & number_of_hits))
   
   if COP_kind_list does not contain COPDocKind_s then loq_Results2(0, false, (COPDocName & " is a " & COPDocKind_s & " -- unsupported type of document"))
   
   if number_of_hits = 0 then
      loqqed_Error_Halt3(false, "Could not find find " & COPDocKind_s & COPDocName)
      error "Could not find find " & COPDocKind_s & COPDocName
   else if number_of_hits > 1 then
      loqqed_Error_Halt3(false, "Found more than one " & COPDocKind_s & " with the name " & COPDocName)
      error "Found more than one " & COPDocKind_s & " with the name " & COPDocName
   else
      tell application "Capture One 11" to set COPDocRef to item 1 of current_doc_ref_list
   end if
   
   loq_Results2(1, false, ("CO Document: " & COPDocKind_s & " " & COPDocName))
end validateCOPdoc3

on validateCOPcollections3()
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose initialisation handler for scripts using Capture One Pro
   ## Extract basic information regarding the current collection, and thhe top level collections
   global debugLogLevel, COPDocName, COPDocKind_s, COPDocRef, enableNotifications
   global everyTopCollection, namesTopCollections, kindsTopCollections_p, countTopCollections, selectedCollectionRef, selectedCollectionIndex, kindSelectedCollection_s, nameSelectedCollection
   global selectedCollectionMirroredAtTopLast, selectedCollectionIsUser, bottomUserCollectionIndex, topUserCollectionIndex
   -- selectedCollectionMirroredAtTopLast replaces selectedCollectionAtTopEnd
   -- bottomUserCollectionIndex, topUserCollectionIndex replaces indexInCatalog
   
   tell application "Capture One 11" to tell COPDocRef
      try
         set selectedCollectionRef to get current collection
         set {nameSelectedCollection, kindSelectedCollection_p} to {name, kind} of selectedCollectionRef
      on error
         my loqqed_Error_Halt3(true, "There is no collection selected in " & COPDocKind_s & " " & COPDocName)
      end try
      set everyTopCollection to get every collection
      set {namesTopCollections, kindsTopCollections_p} to {name, kind} of every collection
   end tell
   set kindSelectedCollection_s to convertKind(kindSelectedCollection_p)
   set countTopCollections to count of namesTopCollections
   
   repeat with collectionCounter from 1 to countTopCollections
      if (nameSelectedCollection = item collectionCounter of namesTopCollections) and (kindSelectedCollection_s = convertKind(item collectionCounter of kindsTopCollections_p)) then
         set selectedCollectionIndex to collectionCounter
         exit repeat
      end if
   end repeat
   
   if COPDocKind_s = "catalog" then
      repeat with collectionCounter from countTopCollections to 1 by -1
         if ("in Catalog" = item collectionCounter of namesTopCollections) and ("smart album" = convertKind(item collectionCounter of kindsTopCollections_p)) then
            set topUserCollectionIndex to collectionCounter - 1
            exit repeat
         end if
      end repeat
      repeat with collectionCounter from 1 to countTopCollections
         if ("Trash" = item collectionCounter of namesTopCollections) and ("smart album" = convertKind(item collectionCounter of kindsTopCollections_p)) then
            set bottomUserCollectionIndex to collectionCounter + 1
            exit repeat
         end if
      end repeat
      set selectedCollectionMirroredAtTopLast to (selectedCollectionIndex = countTopCollections) and ({"catalog folder", "favorite"} does not contain convertKind(last item of kindsTopCollections_p))
      set selectedCollectionIsUser to (selectedCollectionMirroredAtTopLast or ((selectedCollectionIndex ≥ bottomUserCollectionIndex) and (selectedCollectionIndex ≤ topUserCollectionIndex)))
      
   else if COPDocKind_s = "session" then
      repeat with collectionCounter from countTopCollections to 1 by -1
         if ("favorite" ≠ convertKind(item collectionCounter of kindsTopCollections_p)) then
            set topUserCollectionIndex to collectionCounter
            exit repeat
         end if
      end repeat
      repeat with collectionCounter from 1 to countTopCollections
         if ("Trash" = item collectionCounter of namesTopCollections) and ("favorite" = convertKind(item collectionCounter of kindsTopCollections_p)) then
            set bottomUserCollectionIndex to collectionCounter + 1
            exit repeat
         end if
      end repeat
      set selectedCollectionMirroredAtTopLast to false
      set selectedCollectionIsUser to ({"project", "album", "group", "smart album"} contains kindSelectedCollection_s)
   end if
end validateCOPcollections3

on validateCOPvariant(theVariant, theVariantID)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Capture One identifies a variant as "variant n of collrction m of ..."
   ## If the user changes the order of variant of a collection, "variant n" refers to a different variant, and a running script will behave unpredicatably
   ## The usefulness of this handler depends on some other part of the script obtaining a list of the variant IDs when the script starts
   ## The handler is provided with the variant reference and the original variant ID.
   ## This handler obtains the variant's ID from Capture One, and compares that to the original variant ID. If there is a difference, the script is halted
   ## This handler should be called when processing of a variant starts, and just before any operation that change something of the variant 
   
   global COPisParent, enableNotifications
   if COPisParent then return -- when Capture One is running a script, the GUI is locked ... no need to check
   
   tell application "Capture One 11" to tell theVariant to if theVariantID = id then return
   loqqed_Error_Halt3(true, "Variant ID Mismatch, most likely due to manipulation of Capture One by the user while the script is running")
   
end validateCOPvariant

on convertKind(kind_p)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Release 11 of Capture One returns the chevron form of the property "kind" when AppleScript is run as an Application
   ## This script converts the chevron form into the expected string
   ## Code looks ugly but executes in under 0.25 msec
   
   set kind_s to kind_p as text
   if (get text 1 of kind_s) ≠ "«" then return kind_s -- Check if the first character is a Chevron. If not, return the same string
   
   ## minimum 6 characters for valid logic. Normally 19
   if 6 > length of kind_s then loqqed_Error_Halt3(true, "convertKind received an unexpected Kind string: " & kind_s)
   
   set code_start to (get length of kind_s) - 4
   set kind_code to get (text code_start thru (code_start + 3) of kind_s)
   set kind_type to get text 1 thru 2 of kind_code
   set fail_flag to false
   
   if kind_type = "CC" then ## Collection Kinds
      if kind_code = "CCpj" then
         set kind_res to "project"
      else if kind_code = "CCgp" then
         set kind_res to "group"
      else if kind_code = "CCal" then
         set kind_res to "album"
      else if kind_code = "CCsm" then
         set kind_res to "smart album"
      else if kind_code = "CCfv" then
         set kind_res to "favorite"
      else if kind_code = "CCff" then
         set kind_res to "catalog folder"
      else
         set fail_flag to true
      end if
      
   else if kind_type = "CL" then ## Layer Kinds
      if kind_code = "CLbg" then
         set kind_res to "background"
      else if kind_code = "CLnm" then
         set kind_res to "adjustment"
      else if kind_code = "CLcl" then
         set kind_res to "clone"
      else if kind_code = "CLhl" then
         set kind_res to "heal"
      else
         set fail_flag to true
      end if
      
   else if kind_type = "CR" then ## Watermark Kinds
      if kind_code = "CRWn" then
         set kind_res to "none"
      else if kind_code = "CRWt" then
         set kind_res to "textual"
      else if kind_code = "CRWi" then
         set kind_res to "imagery"
      else
         set fail_flag to true
      end if
      
   else if kind_type = "CO" then ## Document Kinds
      if kind_code = "COct" then
         set kind_res to "catalog"
      else if kind_code = "COsd" then
         set kind_res to "session"
      else
         set fail_flag to true
      end if
   else
      set fail_flag to true
   end if
   
   if fail_flag then loqqed_Error_Halt3(true, "convertKind received an unexpected Kind string: " & kind_s)
   
   return kind_res
end convertKind

on InitializeResultsCollection(nameResultProject, nameResultAlbumRoot, Coll_Init_Text)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Sets up a project and albums for collecting images
   
   global debugLogLevel, COPDocRef, ref2ResultAlbum, enableNotifications
   
   tell application "Capture One 11" to tell COPDocRef
      if not (exists collection named (get nameResultProject)) then
         set ref2ResultProject to make new collection with properties {kind:project, name:nameResultProject}
      else
         if ("project" = my convertKind(kind of (get collection named nameResultProject))) then
            set ref2ResultProject to collection named nameResultProject
         else
            my loqqed_Error_Halt3(true, ("A collection named \"" & nameResultProject & "\" already exists, and it is not a project."))
         end if
      end if
   end tell
   
   set coll_ctr to 1
   set nameResultAlbum to nameResultAlbumRoot & "_" & (get short date string of (get current date)) & "_"
   repeat
      tell application "Capture One 11" to tell ref2ResultProject
         if not (exists collection named (get nameResultAlbum & coll_ctr)) then
            set nameResultAlbum to (get nameResultAlbum & coll_ctr)
            set ref2ResultAlbum to make new collection with properties {kind:album, name:nameResultAlbum}
            exit repeat
         else
            set coll_ctr to coll_ctr + 1
         end if
      end tell
   end repeat
   
   if enableNotifications then display notification (Coll_Init_Text & " " & nameResultProject & ">" & nameResultAlbum)
   loq_Results2(1, false, (Coll_Init_Text & "  " & nameResultProject & ">" & nameResultAlbum))
   
end InitializeResultsCollection

## General Handlers  Version 2018/08/26 #######

on settingGUI_Bypass(theSettingList)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose handler that provides a bypass to settingGUI()
   
   ####
   ## The Setting record
   ## SettingName        Name of the Setting (text)
   ## SettingValue         Value of the Setting (text)
   
   global debugLogLevel, parent_name, Script_Title
   tell application "System Events" to set frontmost of process parent_name to true
   
   set settings_display_string to ""
   repeat with theSetting_r in theSettingList
      set settings_display_string to settings_display_string & theSetting_r's SettingName
      try
         set theValueString to joinListToString(theSetting_r's SettingValue, ", ")
         if 0 < length of theValueString then set settings_display_string to settings_display_string & ": " & theValueString
      end try
      set settings_display_string to settings_display_string & return
   end repeat
   
   set debugButtonName to "Configure"
   set legendString to return & "Click \"" & debugButtonName & "\" to configure settings; \"OK\" to continue"
   try
      set dialog_result to display dialog settings_display_string & legendString with title "Settings for " & Script_Title ¬
         buttons {"Cancel", "OK", debugButtonName} default button "OK" cancel button "Cancel"
   on error errmess
      my loqqed_Error_Halt3(true, "User Cancelled the Script")
   end try
   
   return ("OK" = (get button returned of dialog_result))
   
end settingGUI_Bypass

on settingGUI(theSettingList)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose handler that provides a user interface to control AppleScript settings
   
   #############
   ## The Setting record
   ## SettingID             ID of the setting (integer)
   ## SettingName        Name of the Setting (text)
   ## SettingHelp          Help Text for the Setting  
   ## SettingValue         Value or  Reference to the Global Variable which holds the option value (reference)
   ## UserSet                True if the user can set this variable (boolean)
   ## SettingClass         Class of the data representing the option  {Boolean, Text, List_Text, Integer, Real} (text)
   ## SettingLimited      How are the option values constrained {false, List&Free, List, Min_Max, Min, Max}
   ## SettingLimit_L      The list of permissible values of the option (List of Text, List of Integer, List of Real, List of [Min],[Max])
   
   #############
   
   global debugLogLevel, parent_name, Script_Title, enableFastGUI
   
   copy {} to settings_ID_List
   repeat with theSetting_r in theSettingList -- Duplicate setting IDs cause subtle and bizarre errors - catch this problem
      set theSettingID to (get (theSetting_r's SettingID) as text)
      if settings_ID_List contains theSettingID then error "Duplicate Setting ID!! " & theSettingID
      set the end of settings_ID_List to theSettingID
   end repeat
   
   repeat
      tell application "System Events" to set frontmost of process parent_name to true
      
      set settings_display_string to ""
      repeat with theSetting_r in theSettingList
         set settings_display_string to settings_display_string & (theSetting_r's SettingName) & ": " & joinListToString((get contents of theSetting_r's SettingValue), ";")
         if theSetting_r's UserSet then set settings_display_string to settings_display_string & "  *"
         set settings_display_string to settings_display_string & return
      end repeat
      
      set legendString to return & "Click \"Edit\" to edit items with an *asterisk*" & return & "Click \"OK\" to continue without editing settings"
      
      try
         set dialog_result to display dialog settings_display_string & legendString with title "Settings for " & Script_Title ¬
            buttons {"Cancel", "OK", "Edit"} default button "OK" cancel button "Cancel"
      on error errmess
         my loqqed_Error_Halt3(true, "User Cancelled the Script")
      end try
      
      if "OK" = (get button returned of dialog_result) then exit repeat -- Exit 3 of 3 from the handler
      
      set userCancelled to false
      repeat while userCancelled is false
         
         copy {} to settings_choose_List
         repeat with theSetting_r in theSettingList
            ##!!!## Do NOT screw with this line, or the list of text editing starts losing information  ##!!!## 
            if theSetting_r's UserSet then ¬
               copy (get ((theSetting_r's SettingID as text) & ": " & (theSetting_r's SettingName as text) & ": " & joinListToString((contents of theSetting_r's SettingValue), "; "))) to ¬
                  the end of settings_choose_List
         end repeat
         
         set ChooseResult to (get choose from list settings_choose_List with title "Settings for " & Script_Title with prompt "Select an Item to Edit" cancel button name "Done" OK button name "Edit")
         
         if ChooseResult = false then exit repeat --main exit from this repeat loop
         
         set selectedSetting to item 1 of ChooseResult
         set selectedSettingID to get ((item 1 of (splitStringToList(selectedSetting, ":"))) as integer)
         
         set settingFound to false
         repeat with selectedSetting_r in theSettingList
            if selectedSettingID = (get selectedSetting_r's SettingID) then
               set settingFound to true
               exit repeat
            end if
         end repeat
         
         if not settingFound then
            loqqed_Error_Halt3(false, "Unexpected value found while finding setting: " & selectedSettingID)
            error "SW Error 1"
         end if
         
         set selectedSettingName to selectedSetting_r's SettingName as text
         set setting_Class to selectedSetting_r's SettingClass as text
         
         set hasFreeInput to true
         set hasSettingLimitType to false
         set hasSettingMin to false
         set hasSettingMax to false
         set hasSettingLimitList to false
         set theSettingMin to missing value
         set theSettingMax to missing value
         set theSettingLimitList to missing value
         
         try
            set SettingLimitType to get selectedSetting_r's SettingLimited as text -- this works even if the class is boolean
            set hasSettingLimitType to true
            if false = (get SettingLimitType as boolean) then set hasSettingLimitType to false
         end try
         
         if hasSettingLimitType then
            
            if ("Min_Max" = SettingLimitType) then
               set hasFreeInput to true
               set hasSettingMin to true
               set hasSettingMax to true
               set theSettingMin to get item 1 of (get selectedSetting_r's SettingLimit_L as list)
               set theSettingMax to get item 2 of (get selectedSetting_r's SettingLimit_L as list)
               
            else if ("Min" = SettingLimitType) then
               set hasFreeInput to true
               set hasSettingMin to true
               set theSettingMin to get item 1 of (get selectedSetting_r's SettingLimit_L as list)
               
            else if ("Max" = SettingLimitType) then
               set hasFreeInput to true
               set hasSettingMax to true
               set theSettingMax to get item 1 of (get selectedSetting_r's SettingLimit_L as list)
               
            else if ("List" = SettingLimitType) then
               set hasFreeInput to false
               set hasSettingLimitList to true
               set theSettingLimitList to (get selectedSetting_r's SettingLimit_L) as list
               
            else if ("List&Free" = SettingLimitType) then
               set hasFreeInput to (not enableFastGUI)
               set hasSettingLimitList to true
               set theSettingLimitList to (get selectedSetting_r's SettingLimit_L) as list
               
            else
               loqqed_Error_Halt3(false, "Unexpected value found while evaluating Setting Limit. Type: " & SettingLimitType & "   Length: " & (length of selectedSetting_r's SettingLimit_L))
               error "SW Error 4"
            end if
         end if
         
         set settingAddTitle to "Enter New Value"
         set settingChooseTitle to "Select New Item"
         set settingDeleteTitle to "Select Item to be removed"
         set settingEditPrompt to "for setting #" & selectedSettingID & ": " & selectedSettingName
         set settingDeletePrompt to "from setting #" & selectedSettingID & ": " & selectedSettingName
         if (not enableFastGUI) and (0 < (count of selectedSetting_r's SettingHelp)) then set settingEditPrompt to ¬
            settingEditPrompt & return & " (" & (get (selectedSetting_r's SettingHelp) as text) & ")"
         
         
         set userCancelled to false
         set hasNewSettingValue to false
         
         if setting_Class = "Boolean" then
            
            set selectedSettingValue to (selectedSetting_r's SettingValue) as boolean
            
            if enableFastGUI then
               set the contents of selectedSetting_r's SettingValue to not selectedSettingValue
            else
               set ChooseResult to (get choose from list {"True", "False"} with prompt settingEditPrompt with title settingChooseTitle OK button name "Select" default items {(get selectedSettingValue as text)})
               if ChooseResult = false then
                  set userCancelled to true
               else
                  set newSettingValue to item 1 of ChooseResult
                  
                  if "True" = (get newSettingValue) then
                     set the contents of selectedSetting_r's SettingValue to true
                     set hasNewSettingValue to true
                  else if "False" = (get newSettingValue) then
                     set the contents of selectedSetting_r's SettingValue to false
                     set hasNewSettingValue to true
                  else
                     loqqed_Error_Halt3(false, "Unexpected value found while editing Boolean: " & newSettingValue)
                     error "SW Error 2"
                  end if
               end if
            end if
            
         else if setting_Class = "Integer" then
            
            set selectedSettingValue to (get selectedSetting_r's SettingValue) as integer
            
            if hasSettingLimitList then
               set ChooseResult to (get choose from list theSettingLimitList with prompt settingEditPrompt with title settingChooseTitle OK button name "Select")
               if ChooseResult = false then
                  set userCancelled to true
               else
                  if (0 ≤ (count of ChooseResult)) then set contents of selectedSetting_r's SettingValue to (get item 1 of ChooseResult) as integer
                  set hasNewSettingValue to true
                  set selectedSettingValue to (get item 1 of ChooseResult) as integer
               end if
            end if
            
            if hasFreeInput then
               set settingEditPrompt to settingEditPrompt & return
               if hasSettingMin then set settingEditPrompt to settingEditPrompt & "Minimum: " & (get theSettingMin as integer)
               if hasSettingMax then
                  if hasSettingMin then set settingEditPrompt to settingEditPrompt & "      "
                  set settingEditPrompt to settingEditPrompt & "Maximum: " & (get theSettingMax as integer)
               end if
               
               repeat while not hasNewSettingValue
                  try
                     set dialog_result to display dialog settingEditPrompt with title settingAddTitle default answer selectedSettingValue
                  on error
                     set userCancelled to true
                     exit repeat
                  end try
                  try
                     if 0 < (length of text returned of dialog_result) then
                        set newSettingValue to get (text returned of dialog_result) as integer
                     else
                        copy (get selectedSettingValue as integer) to newSettingValue
                     end if
                     set hasNewSettingValue to true
                  on error
                     display dialog (settingEditPrompt & return & "Unable to convert " & (get text returned of dialog_result) & "  to an integer") with title "Press OK to try again" buttons {"OK"} default button "OK"
                  end try
                  
                  if hasNewSettingValue then
                     if (hasSettingMin and (newSettingValue < theSettingMin)) or (hasSettingMax and (newSettingValue > theSettingMax)) then
                        set hasNewSettingValue to false
                        display dialog (settingEditPrompt & return & "The value " & newSettingValue & " is below the Min or above the Max") with title "Press OK to try again" buttons {"OK"} default button "OK"
                     else
                        set contents of selectedSetting_r's SettingValue to newSettingValue
                     end if
                  end if
               end repeat
            end if
            
         else if setting_Class = "Real" then
            
            set selectedSettingValue to (get selectedSetting_r's SettingValue) as real
            
            if hasSettingLimitList then
               set ChooseResult to (get choose from list theSettingLimitList with prompt settingEditPrompt with title settingChooseTitle OK button name "Select")
               if ChooseResult = false then
                  set userCancelled to true
               else
                  if (0 ≤ (count of ChooseResult)) then set contents of selectedSetting_r's SettingValue to (get item 1 of ChooseResult) as real
                  set hasNewSettingValue to true
                  set selectedSettingValue to (get item 1 of ChooseResult) as real
               end if
            end if
            
            if hasFreeInput then
               set settingEditPrompt to settingEditPrompt & return
               if hasSettingMin then set settingEditPrompt to settingEditPrompt & "Minimum: " & (get theSettingMin as real)
               if hasSettingMax then
                  if hasSettingMin then set settingEditPrompt to settingEditPrompt & "      "
                  set settingEditPrompt to settingEditPrompt & "Maximum: " & (get theSettingMax as real)
               end if
               
               repeat while not hasNewSettingValue
                  try
                     set dialog_result to display dialog settingEditPrompt with title settingAddTitle default answer selectedSettingValue
                  on error
                     set userCancelled to true
                     exit repeat
                  end try
                  try
                     if 0 < (length of text returned of dialog_result) then
                        set newSettingValue to get (text returned of dialog_result) as real
                     else
                        copy (get selectedSettingValue as real) to newSettingValue
                     end if
                     set hasNewSettingValue to true
                  on error
                     display dialog (settingEditPrompt & return & "Unable to convert " & (get text returned of dialog_result) & "  to an real") with title "Press OK to try again" buttons {"OK"} default button "OK"
                  end try
                  
                  if hasNewSettingValue then
                     if (hasSettingMin and (newSettingValue < theSettingMin)) or (hasSettingMax and (newSettingValue > theSettingMax)) then
                        set hasNewSettingValue to false
                        display dialog (settingEditPrompt & return & "The value " & newSettingValue & " is below the Min or above the Max") with title "Press OK to try again" buttons {"OK"} default button "OK"
                     else
                        set contents of selectedSetting_r's SettingValue to newSettingValue
                     end if
                  end if
               end repeat
            end if
            
         else if setting_Class = "Text" then
            
            set selectedSettingValue to (get selectedSetting_r's SettingValue) as text
            
            if hasSettingLimitList then
               set ChooseResult to (get choose from list theSettingLimitList with prompt settingEditPrompt with title settingChooseTitle OK button name "Select")
               if ChooseResult = false then
                  set userCancelled to true
               else
                  if (0 ≤ (count of ChooseResult)) then
                     set contents of selectedSetting_r's SettingValue to (get item 1 of ChooseResult)
                     set hasNewSettingValue to true
                     set selectedSettingValue to (get item 1 of ChooseResult) as text
                  end if
               end if
            end if
            
            if hasFreeInput then
               try
                  set dialog_result to display dialog settingEditPrompt with title settingAddTitle default answer selectedSettingValue
                  set contents of selectedSetting_r's SettingValue to get (text returned of dialog_result)
                  set hasNewSettingValue to true
               on error
                  set userCancelled to true
               end try
            end if
            
         else if setting_Class = "List_Text" then
            
            copy ((get contents of selectedSetting_r's SettingValue) as list) to selectedSettingValueList
            
            if (count of selectedSettingValueList) = 0 then
               set deleteList to {}
            else
               set deleteList to (get choose from list selectedSettingValueList ¬
                  with prompt (settingDeletePrompt & return & "(empty selection ok)") ¬
                  with title settingDeleteTitle OK button name "Remove" cancel button name "Skip" with multiple selections allowed and empty selection allowed)
               if (deleteList = false) then
                  set deleteList to {}
               else
                  if (count of deleteList) > 0 then set hasNewSettingValue to true
               end if
            end if
            
            set newSettingValueList to {}
            repeat with theValueString in selectedSettingValueList
               if deleteList does not contain theValueString then copy (get theValueString as text) to the end of newSettingValueList
            end repeat
            
            set addedSettingValueList to {}
            if hasSettingLimitList then
               set ChooseResult to (get choose from list theSettingLimitList with prompt settingEditPrompt ¬
                  with title settingChooseTitle OK button name "Add" with empty selection allowed and multiple selections allowed)
               if ChooseResult = false then
                  set userCancelled to true
               else
                  set addedSettingValueList to (get ChooseResult as list)
               end if
            end if
            
            if hasFreeInput then
               try
                  set dialog_result to display dialog settingEditPrompt with title settingAddTitle default answer ""
                  set addedSettingValueList to addedSettingValueList & splitStringToList((get text returned of dialog_result), ";")
               on error
                  set userCancelled to true
               end try
            end if
            
            if userCancelled then
               set hasNewSettingValue to false
            else
               repeat with theValueString in addedSettingValueList
                  set theValueString to removeLeadingTrailingSpaces(theValueString)
                  ## Don't add zero length strings and strings already in the list
                  if (0 < (count of theValueString)) and ¬
                     newSettingValueList does not contain theValueString then
                     copy (get theValueString as text) to the end of newSettingValueList
                     set hasNewSettingValue to true
                  end if
               end repeat
               
               if hasNewSettingValue and (newSettingValueList ≠ selectedSettingValueList) then
                  set the contents of selectedSetting_r's SettingValue to {}
                  copy (get newSettingValueList as list) to contents of selectedSetting_r's SettingValue
               end if
            end if
            
         else
            loqqed_Error_Halt3(false, "Unexpected class found while editing Boolean: " & setting_Class)
            error "SW Error 3"
         end if
      end repeat
      if enableFastGUI then exit repeat
   end repeat
   
   
   set settings_display_string to ""
   repeat with theSetting_r in theSettingList
      set settings_display_string to settings_display_string & (get SettingName of theSetting_r) & ": " & (get contents of (SettingValue of theSetting_r)) & return
   end repeat
   
   return settings_display_string ##  Exit point
end settingGUI

on InitializeLoqqing3(DocName_Ext, sourceTitle)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose handler to set up Text Editor document for logging results
   
   global debugLogLevel, Script_Title, Result_Doc_ref, SE_Logging, enableResultsFile, enableResultsByDialog, enableResultsByClipboard, DialogTextList, enableNotifications
   global initEnableResultsByDialog, initEnableResultsByClipboard, initEnableNotifications
   
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
      tell application "System Events" to tell application process "Script Editor"
         if (get name of windows) does not contain "log History" then
            click menu item "Log History" of menu "Window" of menu bar 1
         end if
      end tell
      set end of LogMethods to " Script Editor Log"
   end if
   
   if enableNotifications then set end of LogMethods to "Notifications"
   
   set LogMethods_S to joinListToString(LogMethods, ", ")
   loq_Results2(2, false, ("Results by " & LogMethods_S))
   
   return LogMethods_S
end InitializeLoqqing3

on loq_Results2(thisLogDebugLevel, MakeFront, log_Text)
   ## Copyright 2018 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General purpose handler for logging results
   ## log results if the debug level of the message is below the the threshold set by debugLogLevel
   ## log the results by whatever mechanism is ebabled - {Script Editor Log, Text Editor Log, Display Dialog}
   
   global Result_Doc_ref, debugLogLevel, SE_Logging, parent_name, ResultsFileMaxDebug, enableResultsFile, enableResultsByDialog, enableResultsByClipboard, DialogTextList, enableNotifications
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

on splitStringToList(theString, theDelim)
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
end splitStringToList

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

on removeLeadingTrailingSpaces(theString)
   ## Public Domain, modified
   repeat while theString begins with space
      -- When the string is only 1 character long, then it is exactly 1 space, and the next operation willl crash. So return ""
      if 1 ≥ (count of theString) then return ""
      set theString to text 2 thru -1 of theString
   end repeat
   repeat while theString ends with space
      set theString to text 1 thru -2 of theString
   end repeat
   return theString
end removeLeadingTrailingSpaces

on getIndexOf(theItem, theList)
   ## credits Emmanuel Levy
   set astid to AppleScript's text item delimiters
   set AppleScript's text item delimiters to return
   set theList to return & theList & return
   set AppleScript's text item delimiters to astid
   try
      -1 + (count (paragraphs of (text 1 thru (offset of (return & theItem & return) in theList) of theList)))
   on error
      0
   end try
end getIndexOf

on roundToQuantum(thisValue, quantum)
   ## Public domain author unknown
   return (round (thisValue / quantum) rounding to nearest) * quantum
end roundToQuantum

on roundDecimals(n, numDecimals)
   ## Nigel Garvey, Macscripter
   set x to 10 ^ numDecimals
   tell n * x to return (it div 0.5 - it div 1) / x
end roundDecimals

on MSduration(firstTicks, lastTicks)
   ## Public domain
   ## returns duration in ms
   ## inputs are durations, in seconds, from GetTick's Now()
   return (round (10000 * (lastTicks - firstTicks)) rounding to nearest) / 10
end MSduration

on GetTick_Now()
   ## From MacScripter Author "Jean.O.matiC"
   ## returns duration in seconds since since 00:00 January 2nd, 2000 GMT, calculated using computer ticks
   script GetTick
      property parent : a reference to current application
      use framework "Foundation" --> for more precise timing calculations
      on Now()
         return (current application's NSDate's timeIntervalSinceReferenceDate) as real
      end Now
   end script
   
   return GetTick's Now()
end GetTick_Now