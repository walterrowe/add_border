use AppleScript version "2.5"
use scripting additions
set debugEnable to false

tell application "Capture One 20" to set thisDocRef to get current document

set notFoundNameList to searchDocument(thisDocRef)
set countNotFound to (count of notFoundNameList)

if 0 < countNotFound then
   set displayString to (get countNotFound as text) & " Images not found in a User Collection: " & joinListToString(notFoundNameList, ",")
else
   set displayString to "All Images found in a User Collection"
end if

set the clipboard to displayString
display notification displayString

return


on searchDocument(thisDocRef)
   ## Copyright 2019 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## Seach each user collection, and scheck the results
   global theImageListSize, theImageList, debugEnable
   local theCollRefList, countImages, theCollKindList, theCollUserList, theCollCollRefList, thisCollCollList, theDocName
   
   ## bulk collection of information improves speed
   ## Do not collect Image ID lists as this will include All Images, Recent Imports and Catalog Folders (every image 3x or 4x)
   tell application "Capture One 20" to tell thisDocRef to ¬
      set {theCollRefList, theCollUserList, theCollCollRefList, theCollKindList} to ¬
         {every collection, every collection's user, every collection's collection, my convertKindList(every collection's kind)}
   
   ## The image list is created, but nothing gets changed here
   tell application "Capture One 20" to tell thisDocRef to tell collection "All Images" to set countImages to (get count of images)
   set {theImageList, theImageListSize} to {{}, 0} -- initialise and clear the Image list
   makeImageList(countImages)
   
   tell application "Capture One 20" to tell thisDocRef to set theDocName to its name
   display notification "Finding Images not in a user collection: " & countImages & " images of " & theDocName
   
   local thisCollRef, thisCollImageIDList, collCtr
   repeat with collCtr from 1 to (count of theCollUserList)
      if (get contents of theCollUserList's item collCtr) then
         set thisCollRef to (get contents of theCollRefList's item collCtr) -- a CO reference
         tell application "Capture One 20" to tell thisDocRef to set thisCollImageIDList to get id of images of thisCollRef
         searchCollection(thisCollRef, (theCollKindList's item collCtr), (theCollCollRefList's item collCtr), thisCollImageIDList)
      end if
   end repeat
   
   display notification "Half done  ..."
   
   local everyImageIDList, notFoundIdList, notFoundIndexList, notFoundNameList, vctr, theImageID, theImageNameList, anIndex
   tell application "Capture One 20" to tell thisDocRef to tell collection "All Images" to set {everyImageIDList, everyImageNameList} to {id of images, name of images}
   set notFoundNameList to {}
   repeat with vctr from 1 to countImages
      if not (get theImageList's item (get (everyImageIDList's item vctr) as integer)) then ¬
         copy (everyImageNameList's item vctr) to end of notFoundNameList
   end repeat
   
   set {theImageList, theImageListSize} to {{}, 0} -- clear large lists to control AppleScript's memory utilization
   
   return notFoundNameList
end searchDocument

on searchCollection(ThisRef, theCollKind, theCollRefList, theIDStringList)
   ## Copyright 2019 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   global theImageListSize, theImageList, debugEnable
   
   ## index the images if the collection holds images
   local anIDString, theImageID
   tell application "Capture One 20" to tell ThisRef to if (0 < (get count of images)) then
      repeat with anIDString in theIDStringList
         set theImageID to (get anIDString as integer)
         if theImageID > theImageListSize then my makeImageList(theImageID)
         set theImageList's item theImageID to true
      end repeat
   end if
   
   if (0 = (get count of theCollRefList)) or ("project" = theCollKind) then return -- don't search below this collection if it is a project, or has no collections
   ## bulk collection of information improves speed
   ## This handler only handles user collections, so every subcollection is also a user collection
   tell application "Capture One 20" to tell ThisRef to set {theCollCollRefList, theCollImageIDList, theCollKindList} to ¬
      {every collection's collection, every collection's image's id, my convertKindList(every collection's kind)}
   
   local collCtr
   repeat with collCtr from 1 to (get count of theCollRefList)
      searchCollection((theCollRefList's item collCtr), (theCollKindList's item collCtr), (theCollCollRefList's item collCtr), (theCollImageIDList's item collCtr))
   end repeat
end searchCollection

on makeImageList(newImageListSize)
   global theImageListSize, theImageList, debugEnable
   ## make some extras to avoid calling this repeatedly
   set newImageListSize to newImageListSize + 24
   if 0 < theImageListSize then
      if newImageListSize > theImageListSize then set theImageList to theImageList & makeList3((newImageListSize - theImageListSize), false)
   else
      set theImageList to makeList3(newImageListSize, false)
   end if
   set theImageListSize to get count of theImageList
end makeImageList

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

on joinListToString(theList, theDelim)
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

on convertKindList(kind_list)
   ## Copyright 2019 Eric Valk, Ottawa, Canada   Creative Commons License CC BY-SA    No Warranty.
   ## General Purpose Handler for scripts using Capture One Pro
   ## Capture One returns the chevron form of the "kind" property when AppleScript is run as an Application
   ## Unless care is taken to avoid text conversion of this property, this bug breaks script decisions based on "kind"
   ## This script converts text strings with the chevron form to strings with the expected text form
   ## The input may be a single string, a single enum, a list of strings or a list of enums
   ## The code is not compact but runs very fast, between 60us and 210us per item
   
   set kind_s_list to {}
   set input_is_list to ("list" = (get (class of kind_list) as text))
   if input_is_list then
      if ("text" = (get (class of item 1 of kind_list) as text)) and ¬
         ("«" ≠ (get text 1 of item 1 of kind_list)) then return kind_list -- quick pass through if first item is OK
      repeat with theItem in kind_list
         tell application "Capture One 20" to set the end of kind_s_list to (get theItem as text)
      end repeat
      if ("«" ≠ (get text 1 of item 1 of kind_s_list)) then return kind_s_list
   else
      if ("text" = (get (class of kind_list) as text)) and ¬
         ("«" ≠ (get text 1 of kind_list)) then return kind_list -- quick pass through if first item is OK
      tell application "Capture One 20" to set kind_s1 to (get kind_list as text)
      if "«" ≠ (get text 1 of kind_s1) then return kind_s1 -- quick pass through if input is OK
      set kind_s_list to {kind_s1}
   end if
   
   set fail_flag to false
   set code_start to -5
   
   set kind_list to {}
   repeat with Kind_s_item in kind_s_list
      if "«" ≠ (get text 1 of Kind_s_item) then loqqed_Error_Halt3(true, "convertKindList received an unexpected Kind string: " & Kind_s_item)
      set kind_code to get (text code_start thru (code_start + 3) of Kind_s_item)
      set kind_type to get (text code_start thru (code_start + 1) of Kind_s_item)
      
      if kind_type = "CC" then ## Collection Kinds
         if kind_code = "CCpj" then
            set the end of kind_list to "project"
         else if kind_code = "CCgp" then
            set the end of kind_list to "group"
         else if kind_code = "CCal" then
            set the end of kind_list to "album"
         else if kind_code = "CCsm" then
            set the end of kind_list to "smart album"
         else if kind_code = "CCfv" then
            set the end of kind_list to "favorite"
         else if kind_code = "CCff" then
            set the end of kind_list to "catalog folder"
         else
            set fail_flag to true
         end if
         
      else if kind_type = "CL" then ## Layer Kinds
         if kind_code = "CLbg" then
            set the end of kind_list to "background"
         else if kind_code = "CLnm" then
            set the end of kind_list to "adjustment"
         else if kind_code = "CLcl" then
            set the end of kind_list to "clone"
         else if kind_code = "CLhl" then
            set the end of kind_list to "heal"
         else
            set fail_flag to true
         end if
         
      else if kind_type = "CR" then ## Watermark Kinds
         if kind_code = "CRWn" then
            set the end of kind_list to "none"
         else if kind_code = "CRWt" then
            set the end of kind_list to "textual"
         else if kind_code = "CRWi" then
            set the end of kind_list to "imagery"
         else
            set fail_flag to true
         end if
         
      else if kind_type = "CO" then ## Document Kinds
         if kind_code = "COct" then
            set the end of kind_list to "catalog"
         else if kind_code = "COsd" then
            set the end of kind_list to "session"
         else
            set fail_flag to true
         end if
      else
         set fail_flag to true
      end if
      
      if fail_flag then loqqed_Error_Halt3(true, "convertKindList received an unexpected Kind string: " & Kind_s_item)
   end repeat
   
   if input_is_list then
      return kind_list
   else
      return item 1 of kind_list
   end if
   
end convertKindList

