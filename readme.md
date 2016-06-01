# Secret Parser ACT Plugin
Advanced Combat Tracker Parser plugin for The Secret World.

# Changelog

## 1.0.7.0001 pre 1.0.7.1
* changed HTML export to allow serving over https as well as http

## 1.0.7.0
* restored healing-skill breakdown in html-output
* fixed self-hits with aegis

##### English
* fix for Keziah Mason's Ring

## 1.0.6.9
* adapted to the updated ACT API (this allows bigger damage numbers to be parsed correctly)

##### English
* fix for Protector's Wrath

## 1.0.6.8
* incoming dps & aegis dps
* enhanced Autoupdater, searching on github after ACT internal update check
* Autoupdater now in a seperate thread and respects checkbox-setting
* hps using heal-duration instead of damage-duration now
* fix for incoming hits larger then the maximum value of a 32bit integer

##### English
* fix for critically self hits

##### German
* correct DamageType when "von"-Skills hitting Aegis
* removed "-Kraft" from DamageType on some Aegis-healing to self

## 1.0.6.7
+ merged with v1.0.6.6 from RedEyeEagle
+ added checkbox for colored output, rearranged some checkboxes
+ new section "Tank" in chat export, sorting highest to lowest incoming damage
+ incoming aegis-damage
* incoming damage displayed consistently in the chat script
* fixed missing players due to bug in tanking stats
* added option to export rounded dps values to the chat script
* prevent exception during parsing
+ new option to exports damage, heal and tank as separate pop-ups in /actchat if the script would be shortened otherwise
* fixed empty scripts for encounters with quotes in their name

##### German
* Workaround for combat-log problems where the attacker is missing, counting the damage to the player with act running
* fix for critically self hits

##### French
* /acttell script now works properly for French clients
* workaround for regex-compiling times

## 1.0.6.6 - by Inkraja
* incoming damage with block, evade and glance percentages to chat script added
* TSW script has aegis damage/healing colored

## 1.0.6.5
* act export shows more respect for chat limit (open world events)
* TSW script displays all numbers in a consistent manner
+ new columns for aegis dps & hps
+ new option to split self damage by player
+ new option to show the legend in the TSW script

##### German
* added "von" skill 'Sturm von Niflheim'
* support for aegis healing

##### French (with help from Jakfass & Eafalas):
* included some "de" skills
* removed "de" on some damage-types
* support for aegis healing (might be incomplete)

## 1.0.6.4
aegis-healing fully only for english (german is broken; currently no data for french)

* some changes for export script, new option for hiding the legend
* Aegis-heal /w export to Html and script
* new column "Healed" for Zone View
* new column "AegisHeal" for Zone and Encounter Views

## 1.0.6.3
* German "Runter von meinem Land" Skill
* Minor fixes

## 1.0.6.2
*  Minor fixes

## 1.0.6.1 - Moar Aegis
##### General:
* (English/German) fixed some cases with displaying wrong skills/attacker
* fixed sorting for some columns within act's tables

##### Plugin Options:
* New Option : Limit Playernames to 7 characters (default on = old behavior)
* New Option : Reduce Aegis on Names (default off) - Removes Aegis-information from enemy-names

##### New Columns for View Options (Options/Main Table/Encounters):
+ AegisHits : Count skills hitting Aegis
+ AegisDmg : Damage done to Aegis
+ AegisMismatch : Count skills hitting Aegis with wrong controller
+ AegisMismatch% : percentage of AegisMismatch / AegisHits

##### Export:
* Added /actchat export field "AegisMismatch%=a" (only exported if encounter has aegis-dmg)
* Added /act (html) columns AegisDmg & AegisMismatch
* Added total damage and [aegis damage] of encounter
* New /acttell <playername> script, which is basically a /whisper of /actchat

## 1.0.6.0
DamageType für Aegis setzen auf "Aegis" statt "None", Option?

## 1.0.5.9
Liste von Eafalas für Französische Version eingefügt

## 1.0.5.8
Neues Gadget eingefügt

## 1.0.5.7
* resorted tabstop for elements on Secret Parsing Options page
* /actchat abbreviation-line only lists elements which are actually exported (checked "Pen%" -> "p = pen" and so on)

##### German:
* fixed healing from "von"-skills

## 1.0.5.6
* renaming enemies with Aegis to "Enemy - xxx AEGIS" for all languages for better sorting.

##### German:
* damage type from other-chars to aegis now correctly set to "none" instead of "Schaden zu".
* fixed some missing incoming healing lines

##### French: ("de"-skill-list might be incomplete):
* added evading line
* added Aegis support
