Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDrawMap
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'050602 drk: enable or disable menu items based on sector context
	'120503 rjk: added new designatioin info. to distribution tab,
	'            increased the designation label on the distribution tab to show all of the name,
	'            added a check to prevent runtime error when doing ctrl-T at sea or unknown sector
	'
	'130503 rjk: revamped the tab functions for Sector Panel.  Only displays vaild tabs.
	'130503 drk,rjk: Added information for sea
	'140503 rjk: If the Sector Panel can not show the current tab then go back to the first tab.
	'110803 rjk: Added Airport Code to FillGrid and SortGrid so the SortGrid can sort the Range
	'            column when formated with an Airport Code
	'130803 efj: removed dead variables and procedures
	'140903 rjk: Added Center Map menu item to attack, sector(under misc) and sea right clicks,
	'            Removed Sect menu item from Sector->Misc menu, Fixed Warload from code cleanup,
	'            Added more checks to Sector menus
	'150903 rjk: Fixed Fuel from code cleanup, added some units types for LOTR II
	'160903 rjk: Added Sea title for Sea Frame, reposition the rest of the text.
	'            Fixed runtime error deleting Owners with "???" for enemy units.
	'170903 rjk: Changed the Unit Filters to be a combo box instead of buttons.
	'210903 rjk: Removed Path and Route menu items from Sector->Deliver menu
	'230903 rjk: Removed Plane Cargo menu item
	'280903 rjk: General reformatting.  Moved most of form logic to individual forms.
	'            Switched type to an enum.  Added SetUnitDisplay to control the unit grid.
	'            Made Display info private.  Removed FillSubSetGrid and automated.
	'            Added enum for type of unit grid display.
	'            Switched to SelX and  SelY for top level menu.
	'            Fixed to show Unit grid shows when InRange selected.
	'            Added ndump to Unit database refreshes.
	'            Added follow menu click to call frmPromptFollow.
	'            Fixed so Cancel from CenterMap does not cause an error.
	'            Set up automatic load for Plane and Land similar to Ship for single
	'            unit entry.  Position the fields for Consider cmd to appear correctly.
	'            Removed hard coding for units and switched to UnitCharacteristicCheck.
	'            Fixed frmPromptFleet to work for Planes and Land units.
	'            Fixed map(s) command (removed & from command string).
	'            Combined substype and fleet code into single argument for AddSubBox.
	'            Moved enableOrDisableMenu call for top level menu.
	'            Change "Fleet" to contain the Fleet name as well, allows separation and
	'            identification for Fleet A from Wing A.
	'            Created SetIndexSubCombo so the code could be called from more then
	'            one location. Added a ClearSubCombo as well.
	'            Removed ChangeUnitDisplay and combined with Toolbar1_ButtonClick.
	'            Eliminate ShowPlaneInRange, ShowMissileInRange and ShowAttackInRange,
	'            functionality now in SetUnitRange.  Added Type for "Train" units.
	'011003 rjk: Added map, bmap, report(tech), relations, nation, orders, and missions to update all option
	'            Clear Markers in when switching territories.
	'021003 rjk: Changed to Anti menu item to not have to be occuppied but must have mil. and mob.
	'            In the deliver menu item, the default des. for selecting shells was an "s" and now is an "i".
	'            Moved the field logic for frmPromptMove to form.
	'031003 rjk: For ltend switched to UnitPrompt ship from land so initial field shows ships.
	'061003 rjk: Removed query on rsEnemyUnit, does not work this.  Need to investigate more later.
	'            Also removed rs.Index = "Primary Key" as it caused an error on Unit tables.  Need to investigate more later.
	'            Updated udENEMY so it produced output in FillUnitList.
	'051003 rjk: Removed timestamp ndump
	'061003 rjk: Added nuke information to bottom status line.
	'071003 rjk: Nation list now in relations only so do not need "report *"
	'081003 rjk: Added CStr for nation number in FillGrid when doing AddSubBox
	'091003 rjk: Split Reject menu into two menus, modify frmPromptOffer to work for the new Reject/Accept menu items
	'            Rename Accept to AcceptReport so as to not be confused with Accept from the Reject command
	'091003 rjk: Moved the field logic to the form for frmPromptOffer
	'141003 rjk: Moved "show sect s" here from FormLoad frmDrawMap to RefreshDatabase to be inside the database refresh batch
	'141003 rjk: Moved BuildType to be private variable of FillGrid
	'            Change mnuPlaneScutle_Click() to mnuPlaneScuttle_Click(), fixed spelling
	'            Added timestamp for lost when doing update all changes
	'151003 rjk: Switch Oil and Fert fields so they display correctly on the Sea sector panel
	'161003 rjk: Removed dump request from spy as none of our sectors should change.
	'161003 rjk: Added functionality for Buy, Sell, Reset and enhanced Market
	'171003 rjk: Removed the show sect s from the timer as this information is static
	'            and is checked upon startup, added to Update All
	'171003 rjk: Added Strength fields to Sector database
	'171003 rjk: Added Export Nation to the file menu
	'181003 rjk: Fixed that if enemy unit is at x=-1 or y=-1, it does not display ???
	'            Added Route menu to Maps, it called frmPromptCom
	'181003 rjk: Added multiple selection to MouseDown, builds sector list with '\' separating sectors.
	'181003 rjk: Modified Distribute to support multiple sectors select.
	'191003 rjk: Modified Threshold/Move/Deliver to support multiple sector selection.
	'231003 rjk: Added three territory fields to the sector display, moved the strength fields to the right.
	'251003 rjk: Added Shift F9 to forward in the command list
	'            No Repeat Panel exists, so some code cleanup was done in sb1_PanelClick
	'261003 rjk: Added Spy to report to allow section of multiple sectors.
	'271003 rjk: Switched Spy report to use frmPromptRadar instead of frmPromptCensus.
	'301003 rjk: Added Keys for moving about on the map (only active from the map)
	'            ExportIntelligence has its own form for selecting parameters.
	'            For active menu determination, switched to SelX and SelY from MxPos and MyPos.
	'311003 rjk: For frmPromptMove - allow Destination selection while not in the destination field
	'031103 rjk: Added relation command to Form load for when the AutoUpdate is off to build the basic relation info.
	'061103 rjk: Modified the Key pattern for moving around the map using the Keyboard
	'131103 rjk: Added World wraparound checks to AddtoPath (frmPromptPath) to fixed problem doing a recon over a boundary.
	'161103 rjk: Added txtSector to sector selection using the mouse for frmPromptSector
	'181103 rjk: Created one copy of the DisplayPath code and moved here.
	'191103 rjk: Changed MoveAll to accomodate the multiple sector selection, added Commodity Total Report
	'201103 rjk: Added a new option for controlling strength updates
	'            Suppress Sector menu when on sea sector
	'211103 rjk: Resolved nav. sector visibility run-time error by moving the mnuSectCmd and mnuUnits menu determination
	'            to enableOrDisableMenus
	'241103 rjk: Code Cleanup, changed muuUnitsMarkNukes to mnuUnitsMarkNukes
	'            If the strength information is unknown display "???" (UpdateSectors, deletes the records, an suppress option on)
	'251103 rjk: Fixed cmbSub to show all when All is selected.
	'            Moved mnuAttCenter to mnuAttAttack to save a control. Moved the General into an array to save controls
	'            Moved the ShipMisc/PlaneOthers/PlaneMissile/LandMisc commands into an array to save controls.
	'            Added the ability to center the map to the first unit in the list.
	'271103 rjk: Changed to show Attack and Unit menus together on enemy sectors.
	'281103 rjk: Moved the pictures to EmpUnitView and added initialize for unitView object and removed ViewUnits
	'021203 rjk: Changed from mnuViewOptions to mnuOptions to fixed status bar option.
	'            Added Start and Stop to Reports Menu to allow multiple sector section.
	'            Moved options to frmOptions and switched to boolean options.
	'031203 rjk: display enemy mines in red "?".  Added F3 shortcut to UnitView menu.
	'071203 rjk: Changed hldindex to hldIndex.
	'121203 rjk: Switched to Items and Item classes and objects.
	'121203 rjk: Added closing bracket to display of newdes in Deity mode.
	'211203 rjk: Separate Unit number as separate column in the unit grid.
	'291203 rjk: Added Local Help
	'260104 rjk: Modify production double click to product info (St@r W@rs, y changes at the same time)
	'040204 rjk: Fixed Help menu.
	'050204 rjk: Fixed Unit List with 'All', and unit filter, moved boxes to fix overlap.
	'070204 rjk: Update NavMarkers to use Collection to do an exhaustive path search
	'260204 rjk: Added show sect c/b to RefreshAll ensure the database is fill, as it starts empty
	'070304 rjk: Switch the Enemy timestamp to UTC
	'080304 rjk: Fixed the grid selection to pass unit id to form when the row is not actually selected
	'110304 rjk: Switched to use IsNull check for EnemySector designation, display Unknown designation for null entry
	'            if you can not find the information in the rsBmap on the sector panel. Show ??? for unknown efficiency.
	'150304 rjk: Added Estimate Tool from Jim's original code
	'270304 rjk: Switched over to DeleteAllRecords for clearing a table
	'010404 rjk: Added actual Firing Range column to Unit grid.
	'030404 rjk: Added for Copy Game Database form, only visible for Deity
	'160404 rjk: Fixed the sort column for id to be numeric instead of text.
	'200404 rjk: Fixed the redrawing of EventMarkers.
	'250404 rjk: Fixed run-time error drawing NavMarkers, color variable should have been a long.
	'250404 rjk: Added the ability to include Update mobility in NavMarkers
	'250404 rjk: Added the remaining fields from the rangetool to calculated fields in the grid.
	'250404 rjk: Fixed the column width for enemy timestamp.
	'250404 rjk: Added the ability to save Telegram/Server window output to log or file.
	'250404 rjk: Added the ability to right click on the Telegram/Server window to center map if
	'            cursor is on a sector address.
	'260404 rjk: Fixed run-time crash if deleting unit enemy by selecting the whole grid or including the total line.
	'290404 rjk: Added the ability to include Update efficicency in NavMarkers
	'110604 rjk: Unload the chat/flash windows when closing.
	'110604 rjk: Updated the menu for plague marking.
	'200604 rjk: Updated Tooltip for Che Marker.  Added Satellite Path Calculation and Nuke Damage Report
	'010704 rjk: Filled the default origin for Nuke Damage (Range Tool).
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'230704 rjk: Added Radius for delivery for StarWars.
	'010704 rjk: Added "!" as an end delimiter for sector check for Telegram Window.
	'250804 rjk: Enlist does not require civilians in 2K4.
	'290804 rjk: Fixed read (quick) to work for deity mode.
	'250904 rjk: Added Simple Territory Form.
	'270904 rjk: Added the saving of display mode (DD_CURRENT, DD_NEW and DD_BOTH).
	'151104 rjk: Added the ability to start with your Capital instead of 0,0
	'011204 rjk: Added new PrintMap function.
	'190105 rjk: Set UTF8 on the server based on the option on startup.  Fixed spelling.
	'            Postpone the update commands by 10 seconds to prevent the command from being aborted.
	'220305 rjk: Added SubmitTelegramRead to single function for requesting telegrams.
	'220305 rjk: Added SendTelegramRead to single function for requesting telegrams.
	'            Added version check to deal syntax change in 4.2.21.
	'260305 rjk: Fixed NavMarkers to deal with the extra mobility costs with ships with
	'            SWEEP capabilities.
	'060405 rjk: Changed the cmbSub store Nations instead of Fleets for deity mode.
	'            Fixed the unit filter for enemy to work with country #s greater then 10.
	'100405 rjk: UTF8 encapsulated in server version check for 4.2.21 or newer.
	'120405 rjk: Added Update combo box for NavMarkers
	'200405 rjk: Added checkbox for NavMarkers to control whether to limit
	'            mobility to max. mobility or not.
	'240405 rjk: Added Three custom script buttons.
	'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
	'            ships with SWEEP capabilities.  Removed correction factor from NavMarkers.
	'140505 rjk: Added @sector parsing in the telegram frame in main form.
	'260505 rjk: Switched UTF8 from an registry option to login option.
	'290505 rjk: Added ability show Commodity Ratio and an option to control it.
	'160705 rjk: Added GetOilContent, added Clear Intelligence for sector.
	'230705 rjk: Moved Exit to the bottom of the menu.
	'080805 rjk: Added 'n' to the build menu check.
	'010106 rjk: Change Hidden command to Peek to reflect the server change in 4.2.22.
	'280106 rjk: Added mission type and radius to plane grid.
	'180206 rjk: Replace nation with GetNation.  Replace sdump with GetShips.
	'            Replace ldump with GetLandUnits.  Replace ndump with GetNukes.
	'            Replace pdump with GetPlanes.  Replace lost with GetLost.
	'            Replace relation with GetCountryList.
	'            Use xdump to replace mission for 4.3.0.  Use xdump to replace orders for 4.3.0.
	'120306 rjk: Add a call to Compute Unit Counts for ships for 4.3.0.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'210306 rjk: Moved the RequestMetaTables to DrawMap form loading.
	'210306 rjk: Added the peeper account to the visitor account.
	'230406 rjk: Added the Unit Counts for Land Units for 4.3.X servers.
	'060506 rjk: Added Nuke grid.  Remove Nuke warning label.
	'140506 rjk: Fixed mobility and effeciency buttons from grid changes.
	'            Fixed the distribution panel disappearing from grid changes.
	'200506 rjk: Enable menus for radar and improve all the time for SP: Atlantis.
	'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
	'            Also SP: Atlantis will have some the xdump nat and coun changes.
	'250506 rjk: Added Four SP: Atlantis infrastructures to the Sector Panel,
	'            use runway -> min, radar -> gld, fort -> fert nad navigate -> oil.
	'270506 rjk: Increased the size of civilians to 5 digits and exceed civilians to
	'            4 digits.
	'            Do not reset the database you on startup for SP: Atlantis.
	'300506 rjk: Added shortcuts for Jump points.
	'100606 rjk: Added double click to centre map for the unit grid.
	'            Fixed a problem where the Unit Text List in Unit grid.
	'            Not able to requery against a table, only works for the queries.
	'110606 rjk: Enhanced script buttons to support jump points.
	'            Fixed radar and fire menus to check sector fields for SP:Atlantis.
	'            Change toolbar button for bmap not to erase all the map information.
	'            for SP:Atlantis.
	'180606 rjk: Added support for canal flag to ship characteristics.
	'250606 rjk: Added start/stop unit support for 4.3.6 servers.
	'            Set the id to red if stopped.  Added menu for start/stop for units.
	'030706 rjk: Update the grid when transporting.
	'230806 rjk: Added Import/Export Sea Information
	'181006 rjk: Added Import Offset for Intelligence
	'311206 rjk: Added an option to show short names in the Unit grid.
	'090607 rjk: Display whether production is material limited or now.
	'090907 rjk: Added code for using xdump game and xdump update to get
	'            next update for server 4.3.10 and newer.  Use new functions
	'            GetUpdate, UpdateCommands() and StartUpdateTimer().
	'221007 rjk: Use for xdump ver for 4.3.10 servers and newer for version button
	'311007 rjk: Mark sectors that are now fully ours in yellow
	'031107 rjk: Filter the load/unload/lload/lunload commands to relevant units only.
	'251207 rjk: Changed to cargo to start cargo and added end cargo to mission string.
	'010108 rjk: Switch to GetOrders function.
	'100408 rjk: Fixed Version to ensure it shows in a menu when selected from the menu
	'            or toolbar (broken 221007).
	'190708 rjk: Fixed Escape Key to sent to 'aborted' instead of 'ctld', fixes Escape
	'            cause disconnect from the session.
	'250309 rjk: Changed to use RefreshAll for UpdateAll.
	'110410 rjk: Probe changed to n n n n instead of 0 0 0 0 (attack using y/n instead of 1/0)
	
	Public ShutDown As Boolean
	Public PositionHelp As Boolean
	
	Public PromptForm As System.Windows.Forms.Form 'Hold form currently on prompt
	Public PromptUp As Boolean
	Public DrawingPath As Boolean
	Public DisplayingPath As Boolean
	Public MarkingTerritory As Boolean
	Public WorldBuilding As Boolean
	Public MapValid As Boolean
	Public NavMarkerShip As String
	Public bNavMarkerShipIncludeUpdateMob As Boolean
	Public bNavMarkerShipIncludeUpdateEff As Boolean
	Public iNavMarkerShipUpdates As Short
	Public bNavMarkerShipLimitMobility As Boolean
	
	'Track which control activated the pop-up menu
	Dim PopUpMenuSource As Short
	Const DMAP_PUMS_MAP As Short = 0
	Const DMAP_PUMS_GRID As Short = 1
	
	' Selected Sectors
	Public SelX As Short
	Public SelY As Short
	Private SelSectType As Short
	
	Const SEL_SECT_OWN As Short = 1
	Const SEL_SECT_ENEMY As Short = 2
	Const SEL_SECT_UNKNOWN As Short = 3
	Const SEL_SECT_SEA As Short = 4 'Added rjk 05/13/03: Sea Sector that is not a bridge span or bridge tower
	
	' Const MAPBORDER = 10   removed efj 8/2003
	
	'This static var contains the last retreat string so that it can
	'be pulled up in the Nav Control
	Public DoRetreat As Boolean
	Public LastRetreatString As String
	Public LastRetreatUnits As String
	Public LastRetreatType As String
	
	'This holds the cursor position on the command line
	'and if it currently has the focus
	Dim CommandCursorPos As Short
	Dim CommandCursorFocus As Boolean
	
	'This holds the pending telegrams
	Public MsgQ As New Collection
	
	'These hold the timer variables
	Public SecondsToUpdate As Integer
	Public TimerAtUpdate As Integer
	
	Dim MxPos As Short
	Dim MyPos As Short
	Public Magnification As Single
	Dim strLandid As String
	Dim strShipid As String
	Dim strPlaneid As String
	Dim strNukeid As String
	Dim strEnemyid As String
	Dim ViewShipWake As Boolean
	Dim GridSelection As Boolean
	Dim strTextBox As String
	
	
	Dim bShips As Boolean
	Dim bLands As Boolean
	Dim bPlanes As Boolean
	Dim benemys As Boolean
	Dim bNukes As Boolean
	
	'Private mintCurFrame As Integer ' Current Frame visible rjk 05/13/03: removed,
	' DisplaySelectPanel determines which frame should be visible.
	
	' Dim Grid() As Integer   removed efj 8/2003
	' Dim Gridxy() As Integer   removed efj 8/2003
	Dim WorkingPath() As Short
	
	'unit box vars
	' 0901903 rjk: must match to button index in Toolbar1,
	'              also FillUnitList depends on the numbers and order
	Public Enum enumUnitDisplay
		udUNKNOWN = 0
		udSHIP = 1
		udLAND = 2
		udPLANE = 3
		udNUKE = 4
		udENEMY = 5
		udList = 6
	End Enum
	Private Enum enumSubset
		ssALL
		ssSECT
		ssFLEET
		ssPLANE_RANGE
		ssMISSILE_RANGE
		ssATTACK_RANGE
	End Enum
	Public Enum enumUnitType
		TYPE_START 'must be the first type
		TYPE_ALL
		'qualifiers
		TYPE_LAND_UNLOADED
		TYPE_LAND_LOADED
		TYPE_SHIP_UNLOADED
		TYPE_SHIP_LOADED
		'ship capabilities/abilities
		TYPE_S_FISH
		TYPE_S_TORP
		TYPE_S_DCHRG
		TYPE_S_PLANE
		TYPE_S_MISS
		TYPE_S_OIL
		TYPE_S_OILER
		TYPE_S_SONAR
		TYPE_S_MINE
		TYPE_S_SWEEP
		TYPE_S_SUB
		TYPE_S_SPY
		TYPE_S_LAND
		TYPE_S_SEMI_LAND
		TYPE_S_SUB_TORP
		TYPE_S_TRADE
		TYPE_S_CANAL
		TYPE_S_ANTI_MISSILE
		'ship commands
		TYPE_SC_FIRE
		'plane capabilities/abilities
		TYPE_P_TACTICAL
		TYPE_P_BOMBER
		TYPE_P_INTERCEPT
		TYPE_P_CARGO
		TYPE_P_VTOL
		TYPE_P_MISSILE
		TYPE_P_LIGHT
		TYPE_P_SPY
		TYPE_P_IMAGE
		TYPE_P_SATELLITE
		TYPE_P_STEALTH
		TYPE_P_SDI
		TYPE_P_HALF_STEALTH
		TYPE_P_X_LIGHT
		TYPE_P_CHOPPER
		TYPE_P_ANTISUB
		TYPE_P_PARA
		TYPE_P_ESCORT
		TYPE_P_MINE
		TYPE_P_SWEEP
		TYPE_P_MARINE
		'plane groups and commands
		TYPE_PG_ESCORTS
		'plane commands
		TYPE_PC_BOMB
		TYPE_PC_LAUNCH
		TYPE_PC_DROP
		'land units capabilities/abilities
		TYPE_L_XLIGHT
		TYPE_L_ENGINEER
		TYPE_L_SUPPLY
		TYPE_L_SECURITY
		TYPE_L_LIGHT
		TYPE_L_MARINE
		TYPE_L_RECON
		TYPE_L_RADAR
		TYPE_L_ASSAULT
		TYPE_L_TRAIN
		'land unit commands
		TYPE_LC_FIRE
		TYPE_END 'Must be the last type
	End Enum
	Private DisplaySelect As enumUnitDisplay
	Private LastUnitDisplay As enumUnitDisplay
	Private DisplaySubset As enumSubset '092503 rjk: Switched to enum from Integer
	'Public secx As Integer
	'Public secy As Integer
	Public Fleet As String
	'Public BuildType As String 101403 rjk: Moved to be private variable of FillGrid
	Public SubType As enumUnitType '092503 rjk: Switched to enum from Integer
	
	'Public FillSubSetFlag As Boolean 092503 rjk: Removed, logic in FillGrid.
	Private strSubSect As String 'used to keep a list of Sectors in subCombo
	Private strSubFleet As String 'used to keep a list of Fleets in subCombo
	'Private strSel As String moved to MouseUp function, only place used
	
	'Sort Variables
	Dim SortCol As Short
	Dim SortDecending As Boolean
	
	'Selection Variables
	Private NeedMob As Boolean
	Private Needeff As Boolean
	
	Private Const SPLITTER_HEIGHT As Short = 40
	
	' The percentage occupied by the PictureBox.
	Private PercentageY As Single
	Private PercentageX1 As Single
	Private PercentageX2 As Single
	
	'True when we are dragging the splitter.
	Private Dragging As Boolean
	
	'Which Splitter are we draging
	Private DragY As Boolean
	Private DragX1 As Boolean
	
	'Map Maximized Vars
	Private MapMax As Boolean
	Private MMPercentageY As Single
	Private MMPercentageX1 As Single
	Private MMPercentageX2 As Single
	
	'Unit Maximized Vars
	Private UnitMax As Boolean
	
	'Telegram Center Map
	Dim iTelegramSelX As Short
	Dim iTelegramSelY As Short
	
	'UPGRADE_WARNING: Lower bound of array strType was changed from TYPE_START to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private strType(enumUnitType.TYPE_END) As String
	'UPGRADE_WARNING: Lower bound of array strTypeTitle was changed from TYPE_START to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private strTypeTitle(enumUnitType.TYPE_END) As String
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function HtmlHelp Lib "HHCtrl.ocx"  Alias "HtmlHelpA"(ByVal hWndCaller As Integer, ByVal pszFile As String, ByVal uCommand As Integer, ByRef dwData As Any) As Integer
	
	Const HH_DISPLAY_TOPIC As Integer = 0
	Const HH_DISPLAY_TOC As Integer = &H1
	Const HH_DISPLAY_INDEX As Integer = &H2
	Const HH_HELP_CONTEXT As Integer = &HF
	
	'UPGRADE_WARNING: Event cmbUnitFilter.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbUnitFilter_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbUnitFilter.SelectedIndexChanged
		SubType = VB6.GetItemData(cmbUnitFilter, cmbUnitFilter.SelectedIndex)
		FillGrid()
	End Sub
	
	Private Sub cmdMaxMap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMaxMap.Click
		'Handle the Max Maximize option
		If MapMax Then
			MapMax = False
			PercentageY = MMPercentageY
			PercentageX1 = MMPercentageX1
			PercentageX2 = MMPercentageX2
			tbMain.Visible = frmOptions.bToolbar
			sb1.Visible = frmOptions.bStatusBar
		Else
			MapMax = True
			MMPercentageY = PercentageY
			MMPercentageX1 = PercentageX1
			MMPercentageX2 = PercentageX2
			PercentageY = 1
			PercentageX1 = 1
			tbMain.Visible = False
			sb1.Visible = False
		End If
		ResizePanels()
		ArrangeControls()
		
	End Sub
	Private Sub cmdMaxUnit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMaxUnit.Click
		'Handle the Max Maximize option
		If UnitMax Then
			UnitMax = False
			PercentageY = MMPercentageY
			PercentageX1 = MMPercentageX1
			PercentageX2 = MMPercentageX2
			tbMain.Visible = frmOptions.bToolbar
			sb1.Visible = frmOptions.bStatusBar
		Else
			UnitMax = True
			MMPercentageY = PercentageY
			MMPercentageX1 = PercentageX1
			MMPercentageX2 = PercentageX2
			PercentageY = 0
			PercentageX2 = 1
			DisplayUnitBox() 'rjk 05/13/03: which to DisplayUnitBox from DisplaySectorPanel 3
			tbMain.Visible = False
			sb1.Visible = False
		End If
		ResizePanels()
		ArrangeControls()
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDrawMap.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDrawMap_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'030404 rjk: Added for Copy Game Database form, only visible for Deity
		mnuFileOption(11).Visible = bDeity
		mnuFileOption(12).Visible = bDeity
	End Sub
	
	Private Sub frmDrawMap_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim iCapX As Short
		Dim iCapY As Short
		Dim m As Short
		
		frmOptions.LoadOptions() '120203 rjk: Moved to frmOptions
		
		MapValid = True
		MarkingTerritory = False
		DrawingPath = False
		PositionHelp = False
		NavMarkerShip = vbNullString
		SecondsToUpdate = 0
		TimerAtUpdate = 0
		' LowResGraphics = False    8/2003 efj  removed
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: PictureBox property picMap.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.ScaleMode = vbPixels
		
		'Refresh Database from the server data
		If frmEmpCmd.ConnectedtoHost Then
			RequestMetaTables()
			
			'Set the map to be invalid
			MapValid = False
			
			'Skip database refill if you are visitor or guest
			If Not (LCase(frmEmpCmd.CountryName) = "visitor" Or LCase(frmEmpCmd.CountryName) = "guest" Or LCase(frmEmpCmd.CountryName) = "peeper") Then
				Cursor = System.Windows.Forms.Cursors.WaitCursor
				'set AutoRead
				UpdateAutoRead() '120203 rjk: Switched to frmOptions and boolean options
				
				UpdateAutoUpdate() '120203 rjk: Switched to frmOptions and boolean options
				StartUpdateTimer()
				'fill the database
				If frmEmpCmd.bAutoUpdate Then
					frmEmpCmd.ForceUpdates = True
					'clear database
					If Not frmOptions.bSPAtlantis Then
						frmEmpCmd.SubmitEmpireCommand("dbr", False)
					End If
					
					' thomas lullier - fix show sect s here
					' used to calculate the max safe civ at connection
					' when autoupdate is on
					'101403 rjk: Moved to RefreshDatabase
					'frmEmpCmd.SubmitEmpireCommand "bf1", False 'drk 2/26/03: keep the show sect s window from popping up...
					'frmEmpCmd.SubmitEmpireCommand "show sect s", False
					'frmEmpCmd.SubmitEmpireCommand "bf2", False
					
					RefreshDataBase() ' Fill it back up
				Else
					'submit the db2 to get the initial map update
					frmEmpCmd.SubmitEmpireCommand("db1", False)
					GetCountryList()
					frmEmpCmd.SubmitEmpireCommand("db2", False)
				End If
			Else
				Label4.Text = "Visitor"
			End If
		End If
		
		'get layout info from registry '060304 rjk: Added Val to deal with regional settings.
		Magnification = Val(GetSetting(APPNAME, "Layout", "Magnification", CStr(1)))
		PercentageY = Val(GetSetting(APPNAME, "Layout", "Percentage Y", CStr(0)))
		PercentageX1 = Val(GetSetting(APPNAME, "Layout", "Percentage X1", CStr(0)))
		PercentageX2 = Val(GetSetting(APPNAME, "Layout", "Percentage X2", CStr(0)))
		
		'Check for error - can happen when country setting changes.
		'Make sure values are reasonable
		If Magnification > 15 Then
			Magnification = 1
		ElseIf Magnification < 0.3 Then 
			Magnification = 1
		End If
		
		'Same here, make sure settings are in range
		If PercentageY < 0.05 Or PercentageY > 1 Then
			PercentageY = 0.65
		End If
		If PercentageX1 < 0.05 Or PercentageX1 > 1 Then
			PercentageX1 = 0.75
		End If
		If PercentageX2 < 0.05 Or PercentageX2 > 1 Then
			PercentageX2 = 0.55
		End If
		
		'Now load as normal
		SetHexSideLength(picMap, (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * Magnification)
		DrawingPath = False
		
		'MsgBox "Loading Pictures"
		frmUnitView.unitView.Load() '112803 rjk: Moved the pictures to EmpUnitView, and initialize the unitView object and removed ViewUnits
		LoadPictures() '112803 rjk: Moved the pictures to EmpUnitView
		
		'MsgBox "Loading Descriptions"
		LoadSectorDesc()
		
		'Set Tab box
		'mintCurFrame = 1 rjk 05/13/03: not used more, done in DisplaySectorPanel
		Frame1(0).Visible = False
		Frame1(1).Visible = False
		Frame1(1).Height = TabStrip1.Height
		Frame1(2).Visible = False
		Frame1(3).Visible = False 'rjk 05/13/03: ensure Frame1(3) is initialized as well
		Frame1(4).Visible = False
		'rjk 05/13/03: Initialize Sector Panel to location 0,0
		If frmOptions.bCapital And ParseSectors(iCapX, iCapY, CapSect) Then
			FillSectorBox(iCapX, iCapY)
		Else
			FillSectorBox(0, 0)
		End If
		DisplayFirstSectorPanel()
		
		'Set Origins
		OriginX = -6
		OriginY = -6
		
		'Set map size from version info - may need to change later
		If rsVersion.BOF And rsVersion.EOF Then
			'If there is no version record, use 96x48 as default
			MapSizeX = 48
			MapSizeY = 48
		Else
			rsVersion.MoveFirst()
			MapSizeX = rsVersion.Fields("World X").Value
			MapSizeY = rsVersion.Fields("World Y").Value
			If MapSizeX = 0 Then MapSizeX = 48
			If MapSizeY = 0 Then MapSizeY = 48
		End If
		
		'Set ScrollBar Properties
		Dim bhld As Boolean
		bhld = MapValid
		MapValid = False
		vsMap.Minimum = -1.15 * (MapSizeY / 2)
		vsMap.SmallChange = 2
		vsMap.Value = OriginY
		
		hsMap.Minimum = -1.15 * (MapSizeX / 2)
		hsMap.SmallChange = 2
		hsMap.Value = OriginX
		MapValid = bhld
		
		'Set Window Caption
		'092603 rjk: Added Offline mode
		If frmEmpireLogin.bOffline Then
			Text = My.Application.Info.Title & " - " & GameName & " (Offline)"
		Else
			Text = My.Application.Info.Title & " - " & GameName
		End If
		
		'Max to fill the full screen
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		
		'Set status bar panels
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.key was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(1).key = "Timer"
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(1).Text = "Unknown"
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(1).ToolTipText = "Time Remaining until Update"
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(1).MinWidth = 1
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(2, "BTU"))
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(2).Text = " BTU: 640 "
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(2).ToolTipText = "Bureaucratic Time Units Remaining"
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(2).MinWidth = 1
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(3, "Lights"))
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(3).Image = VB6.ImageToIPictureDisp(picGreenLight)
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(3).MinWidth = 1
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(3).ToolTipText = "Empire Server Available"
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(4, "Mail"))
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(4).Text = vbNullString
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(4).MinWidth = 600
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(4).ToolTipText = "Telegram Indicator"
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(5, "Mail Count"))
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(5).MinWidth = 300
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(5).ToolTipText = "Number of pending telegrams"
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(6, "Anno"))
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(6).Text = vbNullString
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(6).ToolTipText = "Announcement Indicator"
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_ISSUE: MSComctlLib.IPanel property sb1.Panels.Item.MinWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		sb1.Items.Item(6).MinWidth = 600
		
		For m = 1 To 6
			'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sb1.Items.Item(m).AutoSize = True
			'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sb1.Items.Item(m).TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
			sb1.Items.Item(m).TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Next 
		
		'UPGRADE_WARNING: MSComctlLib.Panels method sb1.Panels.Add has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		sb1.Items.Add(New System.Windows.Forms.ToolStripStatusLabel(7, "Hex"))
		sb1.Items.Item(7).Spring = True
		'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sb1.Items.Item(7).AutoSize = True
		'set rtb font
		rtbTelegram.Font = Me.Font
		'120203 rjk: Moved loading the options to frmOptions
		
		LastShip = -1
		
		'FillSubSetFlag = True 092503 rjk: Eliminated
		'092103 rjk: Fill the strType dynamically
		BuildUnitFilters()
	End Sub
	
	Public Sub SetMapSize(ByRef sx As Short, ByRef sy As Short)
		'error check
		If sx <= 0 Or sy <= 0 Then
			Exit Sub
		End If
		
		MapSizeX = sx
		MapSizeY = sy
		vsMap.Minimum = -1.15 * CShort(MapSizeY / 2)
		hsMap.Minimum = -1.15 * CShort(MapSizeX / 2)
		DrawHexes()
	End Sub
	
	'Resize the Panels of on the screen
	Private Sub ResizePanels()
		Dim hgt11 As Single
		Dim hgt12 As Single
		Dim hgt21 As Single
		Dim hgt22 As Single
		Dim wdt1 As Single
		Dim wdt2 As Single
		Dim ohgt As Single
		Dim ftop As Single
		
		'take into account toolbar and status bar height
		ohgt = 0
		ftop = 0
		If tbMain.Visible Then
			ohgt = VB6.PixelsToTwipsY(tbMain.Height)
			ftop = VB6.PixelsToTwipsY(tbMain.Height)
		End If
		
		If sb1.Visible Then
			ohgt = ohgt + VB6.PixelsToTwipsY(sb1.Height)
		End If
		
		
		' Don't bother if we're iconized.
		If WindowState = System.Windows.Forms.FormWindowState.Minimized Then Exit Sub
		
		wdt1 = (VB6.PixelsToTwipsX(ClientRectangle.Width) - SPLITTER_HEIGHT) * PercentageY
		wdt2 = (VB6.PixelsToTwipsX(ClientRectangle.Width) - SPLITTER_HEIGHT) - wdt1
		hgt11 = (VB6.PixelsToTwipsY(ClientRectangle.Height) - ohgt - SPLITTER_HEIGHT) * PercentageX1
		hgt12 = (VB6.PixelsToTwipsY(ClientRectangle.Height) - ohgt - SPLITTER_HEIGHT) - hgt11
		hgt21 = (VB6.PixelsToTwipsY(ClientRectangle.Height) - ohgt - SPLITTER_HEIGHT) * PercentageX2
		hgt22 = (VB6.PixelsToTwipsY(ClientRectangle.Height) - ohgt - SPLITTER_HEIGHT) - hgt21
		
		Picture1.SetBounds(0, VB6.TwipsToPixelsY(ftop), VB6.TwipsToPixelsX(wdt1), VB6.TwipsToPixelsY(hgt11))
		Picture2.SetBounds(0, VB6.TwipsToPixelsY(ftop + hgt11 + SPLITTER_HEIGHT), VB6.TwipsToPixelsX(wdt1), VB6.TwipsToPixelsY(hgt12))
		Picture3.SetBounds(VB6.TwipsToPixelsX(wdt1 + SPLITTER_HEIGHT), VB6.TwipsToPixelsY(ftop), VB6.TwipsToPixelsX(wdt2), VB6.TwipsToPixelsY(hgt21))
		Picture4.SetBounds(VB6.TwipsToPixelsX(wdt1 + SPLITTER_HEIGHT), VB6.TwipsToPixelsY(ftop + hgt21 + SPLITTER_HEIGHT), VB6.TwipsToPixelsX(wdt2), VB6.TwipsToPixelsY(hgt22))
	End Sub
	' Arrange the controls on the form.
	Private Sub ArrangeControls()
		
		On Error Resume Next
		
		'Picture Box1 - scroll bars and map
		vsMap.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture1.ClientRectangle.Width) - VB6.PixelsToTwipsX(vsMap.Width))
		vsMap.Top = 0
		
		hsMap.Left = 0
		hsMap.Width = vsMap.Left
		hsMap.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture1.ClientRectangle.Height) - VB6.PixelsToTwipsY(hsMap.Height))
		vsMap.Height = hsMap.Top
		
		'Set up map picture box
		picMap.SetBounds(0, 0, vsMap.Left, hsMap.Top)
		
		'Set the map Max Button
		cmdMaxMap.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(vsMap.Left) - VB6.PixelsToTwipsX(cmdMaxMap.Width)), 0, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		'Picture Box2 - Command Display and Text Entry
		txtEntry.SetBounds(vsMap.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture2.ClientRectangle.Height) - VB6.PixelsToTwipsY(txtEntry.Height)), Picture2.ClientRectangle.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		lstCmdResult.Visible = (VB6.PixelsToTwipsY(txtEntry.Height) < VB6.PixelsToTwipsY(Picture2.ClientRectangle.Height))
		If lstCmdResult.Visible Then
			lstCmdResult.SetBounds(vsMap.Width, 0, hsMap.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture2.ClientRectangle.Height) - VB6.PixelsToTwipsY(txtEntry.Height)))
		End If
		
		'Picture Box 3 - Sector Marker
		lblSect.SetBounds(0, 0, VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture3.ClientRectangle.Width) - VB6.PixelsToTwipsX(cmdMaxUnit.Width)), lblSect.Height)
		TabStrip1.SetBounds(0, lblSect.Height, Picture3.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Picture3.ClientRectangle.Height) - VB6.PixelsToTwipsY(lblSect.Height)))
		
		'Put the Frames on the tabstrip
		Dim X As Short
		With TabStrip1
			For X = 0 To 3
				Frame1(X).SetBounds(VB6.TwipsToPixelsX(.ClientLeft), VB6.TwipsToPixelsY(.ClientTop), VB6.TwipsToPixelsX(.ClientWidth), VB6.TwipsToPixelsY(.ClientHeight))
			Next X
			grid1.Width = VB6.TwipsToPixelsX(.ClientWidth)
			grid1.Height = VB6.TwipsToPixelsY(.ClientHeight - VB6.PixelsToTwipsY(grid1.Top))
			rtbUnitList.SetBounds(grid1.Left, grid1.Top, grid1.Width, grid1.Height)
		End With
		cmdMaxUnit.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Picture3.Width) - VB6.PixelsToTwipsX(cmdMaxUnit.Width)), 0, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		'Picture Box 4 - Telegram Box
		'Update buttons
		
		'Set telegram text box
		rtbTelegram.SetBounds(0, 0, Picture4.ClientRectangle.Width, Picture4.ClientRectangle.Height)
		
		DrawHexes()
	End Sub
	
	Private Sub frmDrawMap_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim wdt1 As Single
		
		wdt1 = (VB6.PixelsToTwipsX(ClientRectangle.Width) - SPLITTER_HEIGHT) * PercentageY
		Dragging = True
		DragY = (X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT)
		DragX1 = (X < wdt1)
		
	End Sub
	
	Private Sub frmDrawMap_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		Dim wdt1 As Single
		
		If MapMax Then Exit Sub
		
		If Not Dragging Then
			wdt1 = (VB6.PixelsToTwipsX(ClientRectangle.Width) - SPLITTER_HEIGHT) * PercentageY
			If X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT Then
				Me.Cursor = System.Windows.Forms.Cursors.SizeWE
			Else
				Me.Cursor = System.Windows.Forms.Cursors.SizeNS
			End If
		Else
			If DragY Then
				PercentageY = X / VB6.PixelsToTwipsX(ClientRectangle.Width)
				If PercentageY < 0 Then PercentageY = 0
				If PercentageY > 1 Then PercentageY = 1
			ElseIf DragX1 Then 
				PercentageX1 = Y / VB6.PixelsToTwipsY(ClientRectangle.Height)
				If PercentageX1 < 0 Then PercentageX1 = 0
				If PercentageX1 > 1 Then PercentageX1 = 1
			Else
				PercentageX2 = Y / VB6.PixelsToTwipsY(ClientRectangle.Height)
				If PercentageX2 < 0 Then PercentageX2 = 0
				If PercentageX2 > 1 Then PercentageX2 = 1
			End If
			
			If PercentageY > 0.98 Then PercentageY = 0.98
			If PercentageY < 0.02 Then PercentageY = 0.02
			If PercentageX1 > 0.98 Then PercentageX1 = 0.98
			If PercentageX1 < 0.02 Then PercentageX1 = 0.02
			If PercentageX2 > 0.98 Then PercentageX2 = 0.98
			If PercentageX2 < 0.02 Then PercentageX2 = 0.02
			
			ResizePanels()
		End If
	End Sub
	
	Private Sub frmDrawMap_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dragging = False
		ArrangeControls()
	End Sub
	
	'UPGRADE_WARNING: Event frmDrawMap.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDrawMap_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		' Dim x As Integer   removed efj 8/2003
		
		'Don't try to redraw when form is minimized
		If WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			If PromptUp Then
				PromptForm.WindowState = System.Windows.Forms.FormWindowState.Minimized
			End If
			frmEmpCmd.WindowState = System.Windows.Forms.FormWindowState.Minimized
			frmTelegram.WindowState = System.Windows.Forms.FormWindowState.Minimized
			Exit Sub
		End If
		
		'Check to make sure its at least the minimum width
		If VB6.PixelsToTwipsX(Me.Width) < (VB6.PixelsToTwipsX(TabStrip1.Width) + VB6.PixelsToTwipsX(vsMap.Width) * 5) Then
			Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(TabStrip1.Width) + VB6.PixelsToTwipsX(vsMap.Width) * 5)
		End If
		
		'Check make sure its at least the minimum height
		If VB6.PixelsToTwipsY(Me.Height) < VB6.PixelsToTwipsY(TabStrip1.Height) * 1.5 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(TabStrip1.Height) * 1.5)
		End If
		
		'Check make sure its at least the minimum height
		If VB6.PixelsToTwipsY(Me.Height) < VB6.PixelsToTwipsY(lstCmdResult.Height) + VB6.PixelsToTwipsY(txtEntry.Height) * 6 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lstCmdResult.Height) + VB6.PixelsToTwipsY(txtEntry.Height) * 6)
		End If
		
		sb1.Items.Item("Hex").Text = "Sect:"
		
		'Re-arrange the controls
		ResizePanels()
		ArrangeControls()
		
		'if a prompt is up, reposition the prompt
		If PromptUp Then
			PromptForm.WindowState = System.Windows.Forms.FormWindowState.Normal
			'Put form in proper place
			PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
			PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		End If
		
		' Titleheight = Me.Height - Me.ScaleHeight   removed efj 8/2003
		
	End Sub
	
	Private Sub frmDrawMap_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim n As Short
		
		'save colors to the registry
		For n = 1 To NUMBER_OF_BASIC_COLORS
			SaveSetting(APPNAME, "Colors", "BasicColor" & CStr(n), CStr(BasicColors(n)))
			SaveSetting(APPNAME, "Colors", "BasicText" & CStr(n), CStr(BasicText(n)))
		Next n
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			SaveSetting(APPNAME, "Colors", "PlayerColors" & CStr(n), CStr(PlayerColors(n)))
			SaveSetting(APPNAME, "Colors", "PlayerText" & CStr(n), CStr(PlayerText(n)))
		Next n
		
		'save layout settings
		If MapMax Or UnitMax Then
			PercentageY = MMPercentageY
			PercentageX1 = MMPercentageX1
			PercentageX2 = MMPercentageX2
		End If
		
		SaveSetting(APPNAME, "Layout", "Magnification", CStr(Magnification))
		SaveSetting(APPNAME, "Layout", "Percentage Y", CStr(PercentageY))
		SaveSetting(APPNAME, "Layout", "Percentage X1", CStr(PercentageX1))
		SaveSetting(APPNAME, "Layout", "Percentage X2", CStr(PercentageX2))
		
		frmOptions.SaveOptions() '120203 rjk: Moved saving options to form
		
		'Set Nations = Nothing
		ShutDown = True
		frmEmpCmd.Close()
		frmEmpireLogin.Close()
		frmTelegram.Close()
		If frmEmpCmd.bEmpireHub Then
			frmChat.Close()
		End If
		frmChat.ControlFlashForm()
		
		If PromptUp Then
			PromptForm.Close()
		End If
		
		frmScript.Close()
		
		'Display forms that are not unloaded
		'Dim f As Object
		'For Each f In Forms
		'    MsgBox f.Name
		'Next
		
	End Sub
	
	Public Sub DisplayUnitBox()
		'rjk 05/13/03: Set the Sector Panel Tab Display to display the Unit Tab
		'
		'DisplaySectorPanel 3 rjk 05/13/03: switched over to tab in DisplaySectorPanel
		
		DisplaySectorPanel() '092803 rjk: Moved early to ensure unit is present for InRange option
		If TabStrip1.Tabs(TabStrip1.Tabs.Count).key = "Unit" Then
			TabStrip1.SelectedItem = TabStrip1.Tabs("Unit")
			SelectTabContent()
		End If
		
	End Sub
	
	Public Sub DisplayFirstSectorPanel()
		'Added rjk 05/13/03: Set the Sector Panel Tab Display to display the First Tab, normally the
		'Census Tab display
		'
		TabStrip1.SelectedItem = TabStrip1.Tabs(1)
		DisplaySectorPanel()
		
	End Sub
	
	Private Sub DisplaySectorPanel()
		' Changed 05/13/03: this function to determines which Tab to display and ensures the proper frame is displayed
		' Also removed the argument and made private, use the DisplayUnitBox or DisplayFirstSectorPanel functions instead
		'
		Dim CurrentTab As String
		
		'rjk 05/14/03: Determine what tab the Sector Panel was showing later decision on what to Tab show
		CurrentTab = TabStrip1.Tabs(TabStrip1.SelectedItem.Index).key
		
		TabStrip1.Tabs.Clear() 'Remove all Tabs
		'Added the appropriate Tabs for the sector
		If DisplaySelect <> enumUnitDisplay.udUNKNOWN Then
			TabStrip1.Tabs.Add(1, "Unit", "Units") 'Add first it will be at the end of the list
		End If
		If SelSectType = SEL_SECT_OWN Then
			TabStrip1.Tabs.Add(1, "Census", "Census")
			TabStrip1.Tabs.Add(2, "Dist", "Dist.")
		ElseIf SelSectType = SEL_SECT_ENEMY Then 
			TabStrip1.Tabs.Add(1, "Enemy", "Census")
		ElseIf SelSectType = SEL_SECT_SEA Then 
			TabStrip1.Tabs.Add(1, "Sea", "Census")
		Else
			If DisplaySelect = enumUnitDisplay.udUNKNOWN Then
				TabStrip1.Tabs.Add(1, "Unknown", vbNullString)
			End If
		End If
		
		'rjk 05/14/03: If the Sector Panel can not show the current tab then go back to the first tab.
		Select Case CurrentTab
			Case "Unit"
				If DisplaySelect <> enumUnitDisplay.udUNKNOWN Then
					TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.Tabs.Count)
				Else
					TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no unit tab
				End If
			Case "Dist"
				If TabStrip1.Tabs.Count >= 2 Then
					If TabStrip1.Tabs(2).key <> "Dist" Then
						TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no dist tab
					End If
				Else
					TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no dist tab
				End If
			Case Else
				If TabStrip1.SelectedItem.Index > TabStrip1.Tabs.Count Then
					TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if not enough tabs
				End If
		End Select
		
		SelectTabContent() 'Match the frame to the Selected Tab
	End Sub
	
	Private Sub SelectTabContent()
		'Added rjk 05/13/03: Display the appropriate frame for the tab selected.
		'drk 05/16/03: make everything hidden by default and only show what we need.  much fewer lines of code.
		
		Frame1(0).Visible = False
		Frame1(1).Visible = False
		Frame1(2).Visible = False
		Frame1(3).Visible = False
		Frame1(4).Visible = False
		
		Select Case TabStrip1.SelectedItem.key
			Case "Census"
				Frame1(1).Visible = True
			Case "Dist"
				Frame1(2).Visible = True
			Case "Unit"
				Frame1(3).Visible = True
			Case "Sea"
				Frame1(4).Visible = True
			Case "Enemy"
				Frame1(0).Visible = True
			Case "Unknown"
		End Select
	End Sub
	
	Public Function IsCensusTabActive() As Boolean
		If TabStrip1.SelectedItem.key = "Census" Then
			IsCensusTabActive = True
		Else
			IsCensusTabActive = False
		End If
	End Function
	
	Public Function IsAnyTabActive() As Boolean
		If TabStrip1.SelectedItem.key = "Unknown" Then
			IsAnyTabActive = False
		Else
			IsAnyTabActive = True
		End If
	End Function
	
	'UPGRADE_ISSUE: Frame event Frame1.DragDrop was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Frame1_DragDrop(ByRef Index As Short, ByRef Source As System.Windows.Forms.Control, ByRef X As Single, ByRef Y As Single)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub grid1_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grid1.DblClick
		Dim iRow As Short
		Dim iCol As Short
		Dim ix As Short
		Dim iy As Short
		
		iRow = grid1.MouseRow
		iCol = grid1.MouseCol
		If (iRow = 1 Or iCol = 0) And ((iRow > 0) And (iRow < grid1.Rows - 1)) Then
			Select Case DisplaySelect
				Case enumUnitDisplay.udSHIP
					ix = CShort(grid1.get_TextMatrix(iRow, 8))
					iy = CShort(grid1.get_TextMatrix(iRow, 9))
				Case enumUnitDisplay.udLAND
					ix = CShort(grid1.get_TextMatrix(iRow, 9))
					iy = CShort(grid1.get_TextMatrix(iRow, 10))
				Case enumUnitDisplay.udPLANE
					ix = CShort(grid1.get_TextMatrix(iRow, 7))
					iy = CShort(grid1.get_TextMatrix(iRow, 8))
				Case enumUnitDisplay.udNUKE, enumUnitDisplay.udENEMY
					ix = CShort(grid1.get_TextMatrix(iRow, 2))
					iy = CShort(grid1.get_TextMatrix(iRow, 3))
				Case Else
					Exit Sub
			End Select
			MoveMap(ix, iy)
		End If
	End Sub
	
	Private Sub grid1_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent) Handles grid1.MouseDownEvent
		Dim mx As Short
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim my_Renamed As Short
		Dim strTemp As String
		Dim StartX As Short
		Dim Finishx As Short
		Dim x1 As Short
		
		GridSelection = True
		'Unit Pop up window
		If eventArgs.Button = VB6.MouseButtonConstants.RightButton Then
			If Not PromptUp Then 'Don't prompt when already prompted
				GridSelection = True
				StartX = grid1.row
				Finishx = grid1.RowSel
				
				If Finishx < StartX Then
					x1 = StartX
					StartX = Finishx
					Finishx = x1
				End If
				
				'go thru selected rows, and add units to select String
				For x1 = StartX To Finishx
					'080304 rjk: Added row check, ensure the total line is not checked,
					'            also allows the selection of a ship actually selecting the line
					If x1 < grid1.Rows - 1 Then
						strTemp = strTemp & Trim(grid1.get_TextMatrix(x1, 1)) & "/"
					End If
				Next x1
				If Len(strTemp) > 0 Then
					strTemp = VB.Left(strTemp, Len(strTemp) - 1)
				Else
					mx = grid1.MouseRow
					my_Renamed = grid1.MouseCol
					If (my_Renamed = 1 Or my_Renamed = 0) And ((mx > 0) And (mx < grid1.Rows - 1)) Then
						strTemp = Trim(grid1.get_TextMatrix(mx, 1))
					End If
				End If
				PopUpMenuSource = DMAP_PUMS_GRID
				'Switched to DisplaySelect from BuildType, BuildType going private/local
				Select Case DisplaySelect
					Case enumUnitDisplay.udSHIP
						strShipid = strTemp
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuShip)
					Case enumUnitDisplay.udPLANE
						strPlaneid = strTemp
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuPlane)
					Case enumUnitDisplay.udLAND
						strLandid = strTemp
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuLand)
					Case enumUnitDisplay.udNUKE
						strNukeid = strTemp
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuNuke)
				End Select
			End If
		End If
	End Sub
	
	Private Sub grid1_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseMoveEvent) Handles grid1.MouseMoveEvent
		Dim mx As Short
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim my_Renamed As Short
		Dim id As Short
		Dim strTemp As String
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		
		mx = grid1.MouseRow
		my_Renamed = grid1.MouseCol
		strTemp = vbNullString
		
		If (my_Renamed = 1 Or my_Renamed = 0) And ((mx > 0) And (mx < grid1.Rows - 1)) Then
			id = CShort(Trim(grid1.get_TextMatrix(mx, 1)))
			'101403 rjk: Switched to DisplaySelect from BuildType, BuildType going local/private
			'strTemp = grid1.TextMatrix(mx, my) + "  " + BuildMissionString(BuildType, id)
			Select Case DisplaySelect
				Case enumUnitDisplay.udSHIP
					strTemp = grid1.get_TextMatrix(mx, 1) & "  " & BuildMissionString("s", id)
				Case enumUnitDisplay.udPLANE
					strTemp = grid1.get_TextMatrix(mx, 1) & "  " & BuildMissionString("p", id)
				Case enumUnitDisplay.udLAND
					strTemp = grid1.get_TextMatrix(mx, 1) & "  " & BuildMissionString("l", id)
				Case enumUnitDisplay.udNUKE
					strTemp = grid1.get_TextMatrix(mx, 1) & "  " & BuildMissionString("n", id)
			End Select
		End If
		sb1.Items.Item("Hex").Text = strTemp
		
	End Sub
	
	Private Sub lblSect_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblSect.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub lstCmdResult_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstCmdResult.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Public Sub mnuAttAttack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAttAttack.Click
		Dim Index As Short = mnuAttAttack.GetIndex(eventSender)
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		Select Case Index
			Case 0 'Probe
				frmEmpCmd.SubmitEmpireCommand("at1 /mil=0; /probe=y; /unt=N;", False)
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'frmEmpCmd.SubmitEmpireCommand "attack " + SectorString(MxPos, MyPos) + " 0 0 0 0", True
				frmEmpCmd.SubmitEmpireCommand("attack " & SectorString(SelX, SelY) & " n n n n", True)
				frmEmpCmd.SubmitEmpireCommand("at2", False)
				
			Case 1 'Attack
				'Setup Designation dialog and execute
				frmPromptAttack.Label2.Text = "Attack"
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'frmPromptAttack.txtOrigin = SectorString(MxPos, MyPos)
				frmPromptAttack.txtOrigin.Text = SectorString(SelX, SelY)
				'Put form in proper place
				frmPromptAttack.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptAttack.Width)) \ 2)
				frmPromptAttack.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptAttack.Height))
				PromptForm = frmPromptAttack
				PromptUp = True
				'Dialog box will take it from here
				frmPromptAttack.Show()
				
			Case 2 'Assault
				'Setup Designation dialog and execute
				frmPromptAttack.Label2.Text = "Assault"
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'frmPromptAttack.txtOrigin = SectorString(MxPos, MyPos)
				frmPromptAttack.txtOrigin.Text = SectorString(SelX, SelY)
				'Put form in proper place
				frmPromptAttack.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptAttack.Width)) \ 2)
				frmPromptAttack.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptAttack.Height))
				PromptForm = frmPromptAttack
				PromptUp = True
				
				'Dialog box will take it from here
				frmPromptAttack.Show()
				
			Case 3 'Bomb
				PromptForm = frmPromptBomb
				'Setup op sector hex
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
				'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtPath.Text = SectorString(SelX, SelY)
				'UPGRADE_ISSUE: Control ShowRange could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.ShowRange = True
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				PromptUp = True
				'Show Prompt
				PromptForm.Show()
				'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtUnit.SetFocus()
				
			Case 4 'Paradrop
				PromptForm = frmPromptRecon
				'Setup op sector hex
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Paradrop"
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
				'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtPath.Text = SectorString(SelX, SelY)
				'UPGRADE_ISSUE: Control ShowRange could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.ShowRange = True
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				PromptUp = True
				'Show Prompt
				PromptForm.Show()
				'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtUnit.SetFocus()
				
			Case 5 'Missile Strike
				PromptForm = frmPromptLaunch
				'Setup op sector hex
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "launch"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Launch"
				' 092303 rjk: Switched to SelX and SelY for top level menu access
				'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
				'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtPath.Text = SectorString(SelX, SelY)
				'UPGRADE_ISSUE: Control ShowRange could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.ShowRange = True
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				PromptUp = True
				'Show Prompt
				PromptForm.Show()
				'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtUnit.SetFocus()
				
			Case 6 'Center Map 112503 rjk: Moved mnuAttCenter to here to save a control
				MoveMap(SelX, SelY)
		End Select
	End Sub
	
	Public Sub mnuCopyReportWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCopyReportWindow.Click
		frmReport.Text = "Telegram/Server Output"
		frmReport.AddReportLine((rtbTelegram.Text))
		frmReport.Show()
	End Sub
	
	Public Sub mnuHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelp.Click
		Dim Index As Short = mnuHelp.GetIndex(eventSender)
		Dim strSubject As String
		
		Select Case Index
			Case 1 'Command List
				If frmOptions.bLocalHelp Then
					HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOPIC, "all.html")
				Else
					frmEmpCmd.SubmitEmpireCommand("list", True)
				End If
			Case 2 'Contents
				HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOC, 0)
			Case 3 'Index
				HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_INDEX, "")
			Case 4 'Info
				strSubject = InputBox("Enter Subject", "Empire Help")
				If frmOptions.bLocalHelp Then
					HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOPIC, strSubject & ".html")
				Else
					frmEmpCmd.SubmitEmpireCommand("info " & strSubject, True)
				End If
			Case 5 'Quick Start Guide
				HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOPIC, "Quick Start.mht")
			Case 6 'User Guide
				HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOPIC, "WinACE User's Guide.htm")
			Case 7 'About
				frmAbout.ShowDialog()
		End Select
	End Sub
	
	Public Sub mnuLandStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandStart.Click
		Dim Index As Short = mnuLandStart.GetIndex(eventSender)
		Select Case Index
			Case 0 'Start
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Start"
				frmPromptCensus.cmbType.SelectedIndex = 2
				ShowUnitPrompt("land")
			Case 1 'Stop
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Stop"
				frmPromptCensus.cmbType.SelectedIndex = 2
				ShowUnitPrompt("land")
		End Select
	End Sub
	
	Public Sub mnuNukeStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNukeStart.Click
		Dim Index As Short = mnuNukeStart.GetIndex(eventSender)
		Select Case Index
			Case 0 'Start
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Start"
				frmPromptCensus.cmbType.SelectedIndex = 4
				ShowUnitPrompt("nuke")
			Case 1 'Stop
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Stop"
				frmPromptCensus.cmbType.SelectedIndex = 4
				ShowUnitPrompt("nuke")
		End Select
	End Sub
	
	Public Sub mnuSectCmdEstimate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdEstimate.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		PromptForm = frmToolProduction
		
		'Setup Designation dialog and execute
		frmToolProduction.txtOrigin.Text = SectorString(SelX, SelY)
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		PromptUp = True
		
		'Dialog box will take it from here
		PromptForm.Show()
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.SetFocus()
	End Sub
	
	Public Sub mnuTelegramCenter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTelegramCenter.Click
		MoveMap(iTelegramSelX, iTelegramSelY)
	End Sub
	
	Public Sub mnuUnitAttAttack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitAttAttack.Click
		Dim Index As Short = mnuUnitAttAttack.GetIndex(eventSender)
		mnuAttAttack_Click(mnuAttAttack.Item(Index), New System.EventArgs())
	End Sub
	
	Public Sub mnuClearSeaIntelligence_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClearSeaIntelligence.Click
		mnuClearSectorIntelligence_Click(mnuClearSectorIntelligence, New System.EventArgs())
		
	End Sub
	
	Public Sub mnuClearSectorIntelligence_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClearSectorIntelligence.Click
		If rsEnemySect.BOF And rsEnemySect.EOF Then
			Exit Sub
		End If
		
		rsEnemySect.Seek("=", SelX, SelY)
		If Not rsEnemySect.NoMatch Then
			rsEnemySect.Delete()
			DrawHexes()
		End If
	End Sub
	
	Public Sub mnuDietyOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDietyOption.Click
		Dim Index As Short = mnuDietyOption.GetIndex(eventSender)
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		Select Case Index
			Case 1, 2, 3 'Give and set resource
				PromptForm = frmPromptLoad
				Select Case Index
					Case 1
						'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						PromptForm.strCmd = "give"
					Case 2
						'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						PromptForm.strCmd = "setsector"
					Case 3
						'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						PromptForm.strCmd = "setresource"
					Case 1
				End Select
				'    PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
				'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strSector = SectorString(SelX, SelY)
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 4
				If VersionCheck(4, 2, 21) = WinAceRoutines.enumVersion.VER_LESS Then
					frmPromptCensus.Label2.Text = "Hidden"
				Else
					frmPromptCensus.Label2.Text = "Peek"
				End If
				mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
		End Select
	End Sub
	
	Public Sub mnuDiploRelationsgrid_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploRelationsgrid.Click
		VB6.ShowForm(frmRelationsGrid, VB6.FormShowConstants.Modeless, Me)
		frmRelationsGrid.go()
	End Sub
	
	Public Sub mnuFileOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOption.Click
		Dim Index As Short = mnuFileOption.GetIndex(eventSender)
		Select Case Index
			Case 1 'Import Telegrams
				ImportTelegrams()
			Case 2 'Import Intelligence
				ImportIntelligenceOffset(vbNullString)
			Case 3 'Import sea information
				ImportSeaInformation()
			Case 5 'Export All Telegrams
				ExportTelegrams(True)
			Case 6 'Export Recent Telegrams
				ExportTelegrams(False)
			Case 7 'Export Intelligence
				'103003 rjk: Added its own form for selecting parameters.
				PromptForm = frmPromptExportIntelligence
				PromptUp = True
				PromptForm.Show()
			Case 8 'Export Nation
				ExportNation()
			Case 9 'Export Sea Information
				ExportSeaInformation()
			Case 11 'Copy Game Database
				frmPromptCopyGameDB.ShowDialog()
			Case 13 'Clear
				'112903 rjk: Created a separate form for clearing.
				frmClear.ShowDialog()
			Case 14 'Print
				frmPrintMap.Show()
			Case 15 'Exit
				Me.Close()
		End Select
		
	End Sub
	
	Public Sub mnuGenCmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGenCmd.Click
		Dim Index As Short = mnuGenCmd.GetIndex(eventSender)
		Select Case Index
			Case 0 'Break Sanctuary
				frmEmpCmd.SubmitEmpireCommand("break", True)
				
			Case 1 'Change
				'Put form in proper place
				frmPromptChange.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptChange.Width)) \ 2)
				frmPromptChange.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptChange.Height))
				PromptForm = frmPromptChange
				PromptUp = True
				'Dialog box will take it from here
				frmPromptChange.Show()
				
			Case 2 'Realm
				frmPromptConvert.Label2.Text = "Realm"
				frmPromptConvert.Label1.Text = "Sectors:"
				mnuSectCmdConvert_Click(mnuSectCmdConvert, New System.EventArgs())
				
			Case 3 'Territory
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptTerritory
				PromptUp = True
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				'Dialog box will take it from here
				PromptForm.Show()
				
			Case 4 'Set Owner
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptIsland
				PromptUp = True
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Set Owner"
				'UPGRADE_ISSUE: Control cbNations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.cbNations.Visible = True
				'Dialog box will take it from here
				PromptForm.Show()
		End Select
	End Sub
	
	Public Sub mnuGotoOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGotoOptions.Click
		Dim Index As Short = mnuGotoOptions.GetIndex(eventSender)
		On Error Resume Next
		Dim sx As Short
		Dim sy As Short
		
		If Not MapValid Then Exit Sub
		
		Select Case Index
			Case 1 ' Last event
				sx = EventX
				sy = EventY
			Case 2, 3, 4, 5, 6 ' Jump Point 1-5
				If Not ParseSectors(sx, sy, rsNation.Fields("JumpPoint" & CStr(Index - 1)).Value) Then
					Exit Sub
				End If
			Case Else
				Exit Sub
		End Select
		
		MoveMap(sx, sy)
	End Sub
	
	Public Sub mnuLandMisc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandMisc.Click
		Dim Index As Short = mnuLandMisc.GetIndex(eventSender)
		Dim sx, sy As Short
		Dim slist As Object
		Dim hldIndex As String
		Select Case Index
			Case 0 'Army
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptFleet
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2.Caption = "Army"
				If Len(strLandid) > 0 Then
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strLandid
				End If
				'UPGRADE_ISSUE: Control LoadCombo could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.LoadCombo()
				'Show Prompt
				ShowUnitPrompt("land")
			Case 1 'Center Map
				
				If strLandid <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object slist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					slist = ParseUnitString(strLandid)
					If IsArray(slist) Then
						hldIndex = rsLand.Index
						rsLand.Index = "PrimaryKey"
						rsLand.Seek("=", slist(1))
						If Not rsLand.NoMatch Then
							MoveMap(rsLand.Fields("x").Value, rsLand.Fields("y").Value)
						End If
						rsLand.Index = hldIndex
					End If
				End If
			Case 2 'Fortify
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				PromptUp = True
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "fortify"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Fortify"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "mobility"
				'Show Prompt
				ShowUnitPrompt("land")
			Case 3 'Mine
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptLook
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "lmine"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Land Mine"
				'Show Prompt
				ShowUnitPrompt("land")
			Case 4 'Range
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "lrange"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Range"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "land radius"
				'Show Prompt
				ShowUnitPrompt("land")
			Case 5 'Retreat
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "lretreat"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Retreat"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "path"
				'UPGRADE_ISSUE: Control Label3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label3.Visible = True
				'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Combo1.Visible = True
				'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				LoadRetreatBox(PromptForm.Combo1)
				'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Combo1.ListIndex = 0
				'Show Prompt
				ShowUnitPrompt("land")
		End Select
	End Sub
	
	Public Sub mnuLandWarload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandWarload.Click
		' added by thomas lullier
		' load max mil/shell/gun to land units
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptWarload
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "lwarload"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 09/14/03 rjk: strSector was removed 2.2.13, not used in function
		'PromptForm.SetSectorLabels - thomas lullier - what is it used for ?
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuMarketBuy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketBuy.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptCom
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Buy"
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuMarketMarket_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketMarket.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptCom
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Market"
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuMarketReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketReset.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptCom
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Reset"
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuMarketSell_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketSell.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptCom
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Sell"
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuMarketSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketSet.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptScrap
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Set"
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "set"
		'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option1 = True
		
		'Dialog box will take it from here
		ShowUnitPrompt("ship")
	End Sub
	
	Public Sub mnuNukeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNukeList.Click
		' 092303 rjk: Switched to SelX and SelY for top level menu access
		'frmEmpCmd.SubmitEmpireCommand "nuke " + SectorString(MxPos, MyPos), True
		frmEmpCmd.SubmitEmpireCommand("nuke " & SectorString(SelX, SelY), True)
	End Sub
	
	Public Sub mnuNukeTrans_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNukeTrans.Click
		PromptForm = frmPromptTrans
		PromptUp = True
		
		'Setup op sector hex
		' 092303 rjk: Switched to SelX and SelY for top level menu access
		'PromptForm.strSector = SectorString(MxPos, MyPos)
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		frmPromptTrans.Option2.Checked = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		PromptUp = True
		
		'Dialog box will take it from here
		ShowUnitPrompt("nuke")
		On Error Resume Next
		'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtUnit.SetFocus()
	End Sub
	
	Public Sub mnuPlaneMissile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneMissile.Click
		Dim Index As Short = mnuPlaneMissile.GetIndex(eventSender)
		Select Case Index
			Case 0 'Arm
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "arm"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Arm"
				If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
					'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.Label1 = "nuke id"
				Else
					'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.Label1 = "nuke type"
				End If
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 1 'Disarm
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptLook
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "disarm"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Disarm"
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 2 'Harden
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "harden"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Harden"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "amount"
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 3 'Launch
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptLaunch
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "launch"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Lanuch"
				'Show Prompt
				ShowUnitPrompt("plane")
		End Select
	End Sub
	
	Public Sub mnuPlaneOther_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneOther.Click
		Dim Index As Short = mnuPlaneOther.GetIndex(eventSender)
		Dim sx, sy As Short
		Dim slist As Object
		Dim hldIndex As String
		Select Case Index
			Case 0 'Center Map
				
				If strPlaneid <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object slist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					slist = ParseUnitString(strPlaneid)
					If IsArray(slist) Then
						hldIndex = rsPlane.Index
						rsPlane.Index = "PrimaryKey"
						rsPlane.Seek("=", slist(1))
						If Not rsPlane.NoMatch Then
							MoveMap(rsPlane.Fields("x").Value, rsPlane.Fields("y").Value)
						End If
						rsPlane.Index = hldIndex
					End If
				End If
			Case 1 'Sat
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "satellite"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Satellite"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "report"
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 2 'Scrap
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "scrap"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Scrap"
				'UPGRADE_ISSUE: Control Option3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option3 = True
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 3 'Scuttle
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "scuttle"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Scuttle"
				'UPGRADE_ISSUE: Control Option3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option3 = True
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 4 'Start
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Start"
				frmPromptCensus.cmbType.SelectedIndex = 3
				ShowUnitPrompt("plane")
			Case 5 'Stop
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Stop"
				frmPromptCensus.cmbType.SelectedIndex = 3
				ShowUnitPrompt("plane")
			Case 6 'Upgrade
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "upgrade "
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Upgrade"
				'UPGRADE_ISSUE: Control Option3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option3 = True
				'Show Prompt
				ShowUnitPrompt("plane")
			Case 7 'Wing
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptFleet
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2.Caption = "Wing"
				If Len(strPlaneid) > 0 Then
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strPlaneid
				End If
				'UPGRADE_ISSUE: Control LoadCombo could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.LoadCombo()
				'Show Prompt
				ShowUnitPrompt("plane")
		End Select
	End Sub
	
	Public Sub mnuRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRefresh.Click
		Dim Index As Short = mnuRefresh.GetIndex(eventSender)
		Select Case Index
			Case 1 'Update Sectors
				frmEmpCmd.ForceUpdates = True
				'Send Code to start database update
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				'Request an update of Sector information
				frmEmpCmd.SubmitEmpireCommand("cs1", False)
				GetSectors(True)
				'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
				GetCurrentStrength()
				GetOContent() '110605 rjk: Added ability to get OContent for Sea Sectors
				'send Code to end database update
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 2 'Update BMap
				frmEmpCmd.ForceUpdates = True
				If frmOptions.bSuppressBmapRefresh Then
					MsgBox("Bmap updates are suppressed. Change option and then request again")
					Exit Sub
				End If
				'Clear Bmap file
				DeleteAllRecords(rsBmap)
				'Send Code to start database update
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				'Request an update of map information
				frmEmpCmd.SubmitEmpireCommand("map *", False)
				frmEmpCmd.SubmitEmpireCommand("bmap *", False)
				'send Code to end database update
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 3 'Update Units
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				'Set parms to make sure all dumps are done
				frmEmpCmd.ForceUpdates = True
				frmEmpCmd.HaveLands = True
				frmEmpCmd.HaveShips = True
				frmEmpCmd.HavePlanes = True
				frmEmpCmd.HaveNukes = True '092503 rjk: Added to Units option
				
				On Error Resume Next 'MoveFirst will give error on empty dataset
				'Send Code to start database update
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				GetLost(True)
				frmEmpCmd.SubmitEmpireCommand("cu1", False)
				'Request an update of Plane information
				GetPlanes(True)
				'Request an update of land information
				GetLandUnits(True)
				'Request an update of ship information
				GetShips(True)
				'Request an update of nuke information 092503 rjk: Added to Units option
				GetNukes(True)
				'send Code to end database update
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 4 'Update Players List
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				'Request an update of player information
				If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
					GetNation()
					GetCountryList()
				Else
					GetCountryList()
					GetNation()
				End If
				'send Code to end database update
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 5 'Update All
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				frmEmpCmd.SubmitEmpireCommand("cs1", False)
				frmEmpCmd.SubmitEmpireCommand("cu1", False)
				GetUpdate(False)
				If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
					frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
				End If
				
				GetSectors(True)
				'101703 rjk: Added Strength fields to Sector database
				GetCurrentStrength() '112003 rjk: Added option to control
				GetOContent() '110605 rjk: Added ability to get OContent for Sea Sectors
				GetPlanes(True)
				GetLandUnits(True)
				GetShips(True)
				GetNukes(True)
				GetLost(True)
				If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
					GetNation()
					GetCountryList()
				Else
					GetCountryList()
					GetNation()
				End If
				frmToolShow.cmdRefreshAll_Click(Nothing, New System.EventArgs())
				frmEmpCmd.SubmitEmpireCommand("map *", False) '100103 rjk: Added to all options
				frmEmpCmd.SubmitEmpireCommand("bmap *", False) '100103 rjk: Added to all options
				If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
					GetOrders(False) 'inside if because GetShips has already been done
					DeleteAllRecords(rsMissions) '100103 rjk: Added to all options
					frmEmpCmd.SubmitEmpireCommand("mission land * q", False) '100103 rjk: Added to all options
					frmEmpCmd.SubmitEmpireCommand("mission ship * q", False) '100103 rjk: Added to all options
					frmEmpCmd.SubmitEmpireCommand("mission plane * q", False) '100103 rjk: Added to all options
				End If
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 6 'Update Changed
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				GetSectors()
				'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
				GetCurrentStrength(tsSectors)
				GetOContent() '110605 rjk: Added ability to get OContent for Sea Sectors
				GetPlanes()
				GetLandUnits()
				GetShips()
				GetNukes()
				GetLost()
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case 8 'Update Orders
				GetOrders(True)
			Case 9 'Update Missions
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
					DeleteAllRecords(rsMissions)
					frmEmpCmd.SubmitEmpireCommand("mission land * q", False)
					frmEmpCmd.SubmitEmpireCommand("mission ship * q", False)
					frmEmpCmd.SubmitEmpireCommand("mission plane * q", False)
					frmEmpCmd.SubmitEmpireCommand("db2", False)
				Else
					GetLandUnits(True)
					GetShips(True)
					GetPlanes(True)
				End If
		End Select
	End Sub
	
	Public Sub mnuReportRoute_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportRoute.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptCom
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Route"
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuReportShow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportShow.Click
		frmToolShow.Show()
	End Sub
	
	Public Sub mnuReportSpy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportSpy.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptRadar
		PromptUp = True
		
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Spy"
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = "*"
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuSeaCenter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSeaCenter.Click
		'092303 rjk: Switched to SelX and SelY for top level menu access
		'MoveMap MxPos, MyPos
		MoveMap(SelX, SelY)
	End Sub
	
	Public Sub mnuSectCmdAssault_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdAssault.Click
		mnuAttAttack_Click(mnuAttAttack.Item(2), New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdCenter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdCenter.Click
		'092303 rjk: Switched to SelX and SelY for top level menu access
		'MoveMap MxPos, MyPos
		MoveMap(SelX, SelY)
	End Sub
	
	Public Sub mnuSectCmdMoveAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdMoveAll.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "move"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		'Show Prompt
		'ShowUnitPrompt "ship"  'removed drk@unxsoft.com 2/26/02. replaced with the more sane code below...
		PromptUp = True
		PromptForm.Show()
		'UPGRADE_ISSUE: Control txtMultDest could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultDest.SetFocus() '111903 rjk: Multiple Sector Selection
	End Sub
	
	Public Sub mnuSectCmdTerr_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdTerr.Click
		If PromptUp Then
			Exit Sub
		End If
		PromptForm = frmPromptSimpleTerritory
		PromptUp = True
		
		frmPromptSimpleTerritory.txtMultOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuSetJumpPoint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSetJumpPoint.Click
		Dim Index As Short = mnuSetJumpPoint.GetIndex(eventSender)
		If rsNation.BOF And rsNation.EOF Then
			Exit Sub
		End If
		
		rsNation.MoveFirst()
		rsNation.Edit()
		rsNation.Fields("JumpPoint" & CStr(Index)).Value = SectorString(SelX, SelY)
		rsNation.Update()
	End Sub
	
	Public Sub mnuShipFollow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipFollow.Click '092303 rjk: Added, frmPromptFollow exists but no call to
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptFollow
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipMisc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipMisc.Click
		Dim Index As Short = mnuShipMisc.GetIndex(eventSender)
		Dim sx, sy As Short
		Dim slist As Object
		Dim hldIndex As String
		Select Case Index
			Case 0 'Center Map 112503 rjk Added the ability to center the map to the first ship in the list
				
				If strShipid <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object slist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					slist = ParseUnitString(strShipid)
					If IsArray(slist) Then
						hldIndex = rsShip.Index
						rsShip.Index = "PrimaryKey"
						rsShip.Seek("=", slist(1))
						If Not rsShip.NoMatch Then
							MoveMap(rsShip.Fields("x").Value, rsShip.Fields("y").Value)
						End If
						rsShip.Index = hldIndex
					End If
				End If
			Case 1 'Fleet
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptFleet
				PromptUp = True
				If Len(strShipid) > 0 Then
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strShipid
				End If
				'UPGRADE_ISSUE: Control LoadCombo could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.LoadCombo()
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 2 'Fuel
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptFuel
				'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option1.Value = True
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 3 'Name
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptSail
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "name"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Name"
				'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label1 = "to"
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 4 'Radar
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptLook
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "radar"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Radar"
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 5 'Scrap
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "scrap "
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Scrap"
				'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option1 = True
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 6 'Scuttle
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "scuttle"
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Scuttle"
				'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option1 = True
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 7 'Start
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Start"
				frmPromptCensus.cmbType.SelectedIndex = 1
				ShowUnitPrompt("ship")
			Case 8 'Stop
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptCensus
				frmPromptCensus.Label2.Text = "Stop"
				frmPromptCensus.cmbType.SelectedIndex = 1
				ShowUnitPrompt("ship")
			Case 9 'Tend
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptLoad
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "tend"
				'Show Prompt
				ShowUnitPrompt("ship")
			Case 10 'Upgrade
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				PromptForm = frmPromptScrap
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "upgrade "
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Upgrade"
				'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Option1 = True
				'Show Prompt
				ShowUnitPrompt("ship")
		End Select
	End Sub
	
	Public Sub mnuShipWarload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipWarload.Click
		' added by thomas lullier
		' load max mil/shell/gun to ship
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptWarload
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "swarload"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 09/14/03 rjk: strSector was removed 2.2.13, not used in function
		'PromptForm.SetSectorLabels - thomas lullier - what is it used for ?
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuToolsOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuToolsOption.Click
		Dim Index As Short = mnuToolsOption.GetIndex(eventSender)
		Dim strWarning As String
		
		strWarning = "Warning: This utility is a tool designed to quickly set up a test country for" & vbLf & "the Vampire Blitz.  It does not take into account GO_RENEW or BIG_CITIES." & vbLf & "and it does not do the most efficent job.  It should be used with care, " & vbLf & "and I recommend that it NEVER be used in a serious game. - Escher"
		
		Dim Reply As Short
		Select Case Index
			Case 1 'Script Files
				'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
				Load(frmScript)
				frmScript.Show()
				
			Case 2 'Build Projection
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				'Setup  dialog and execute
				'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
				Load(frmToolHarbor)
				frmToolHarbor.txtOrigin.Text = SectorString(SelX, SelY)
				PromptForm = frmToolHarbor
				
				'Put form in proper place
				frmToolHarbor.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmToolHarbor.Width))
				frmToolHarbor.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmToolHarbor.Height)) \ 2)
				PromptUp = True
				
				'Dialog box will take it from here
				frmToolHarbor.Show()
				
			Case 3
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				'Setup  dialog and execute
				'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
				Load(frmToolFeeder)
				frmToolFeeder.txtOrigin.Text = SectorString(SelX, SelY)
				PromptForm = frmToolFeeder
				
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width))
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height)) \ 2)
				PromptUp = True
				
				'Dialog box will take it from here
				frmToolFeeder.Show()
				
			Case 4
				frmToolNationLvl.ShowDialog()
			Case 5
				frmToolCheMarker.ShowDialog()
				DrawHexes()
			Case 6 'production summary
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				PromptForm = frmPromptProdSummary
				PromptUp = True
				
				frmPromptProdSummary.txtOrigin.Text = SectorString(SelX, SelY)
				
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				
				'Dialog box will take it from here
				PromptForm.Show()
			Case 7
				AirCombatReport()
			Case 8
				PromptForm = frmToolRange
				frmToolRange.txtOrigin.Text = SectorString(SelX, SelY)
				
				PromptUp = True
				'        PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
				'       PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
				
				'Dialog box will take it from here
				PromptForm.Show()
			Case 9
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				'Setup  dialog and execute
				'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
				Load(frmToolFortify)
				PromptForm = frmToolFeeder
				
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width))
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height)) \ 2)
				PromptUp = True
				
				'Dialog box will take it from here
				frmToolFortify.Show()
			Case 10
				frmSatellitePath.Show()
			Case 12
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				Reply = MsgBox("Run Blitz Setup ?" & vbLf & vbLf & strWarning, MsgBoxStyle.OKCancel + MsgBoxStyle.Question, "Tool Confirmation")
				If Reply = MsgBoxResult.Cancel Then
					Exit Sub
				End If
				
				BlitzSetup(True, 0, 0, 0, 0)
			Case 13
				Reply = MsgBox("Run Island Setup ?" & vbLf & vbLf & strWarning, MsgBoxStyle.OKCancel + MsgBoxStyle.Question, "Tool Confirmation")
				If Reply = MsgBoxResult.Cancel Then
					Exit Sub
				End If
				
				'Setup dialog and execute
				PromptForm = frmPromptIsland
				PromptUp = True
				
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				
				'Dialog box will take it from here
				PromptForm.Show()
			Case 14
				Reply = MsgBox("Run Island Develop Setup ?" & vbLf & vbLf & strWarning, MsgBoxStyle.OKCancel + MsgBoxStyle.Question, "Tool Confirmation")
				If Reply = MsgBoxResult.Cancel Then
					Exit Sub
				End If
				
				'check for current prompt
				If PromptUp Then
					Exit Sub
				End If
				
				PromptForm = frmPromptIsland
				PromptUp = True
				
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.Label2 = "Island Develop Tool"
				
				'Dialog box will take it from here
				PromptForm.Show()
				
				
			Case 16
				ImportBmap()
				
			Case 17 'WorldBuilder
				
				'If bDeity Then
				'Setup dialog and execute
				PromptForm = frmToolWorldBuild
				PromptUp = True
				
				'Put form in proper place
				PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
				PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
				
				'Dialog box will take it from here
				PromptForm.Show()
				'Else
				'MsgBox "World Builder is a Deity only Tool", vbOKOnly + vbExclamation, "Tool Confirmation"
				'End If
		End Select
	End Sub
	
	Public Sub mnuUnitShow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitShow.Click
		frmToolShow.Show()
	End Sub
	
	Public Sub mnuViewOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewOptions.Click
		Dim Index As Short = mnuViewOptions.GetIndex(eventSender)
		Static SaveUnitView As Boolean
		
		Dim sx As Short
		Dim sy As Short
		Dim cstring As String
		Select Case Index
			Case 1
				If ColorScheme = COLORSCHEME_NORMAL Then
					'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
					Load(frmColorScheme)
					frmColorScheme.optScheme(2).Checked = True
					frmColorScheme.ShowDialog()
				Else
					ColorScheme = COLORSCHEME_NORMAL
					If frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_CURRENT Or frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH_NEW Then
						HexGrid.dspDesignation = HexGrid.enumDisplayDesignation.DD_BOTH
						frmOptions.dspDesignationSave = HexGrid.enumDisplayDesignation.DD_BOTH
					End If
					frmUnitView.unitView.Load()
				End If
				DrawHexes()
			Case 2
				frmSelectColor.ShowDialog()
				DrawHexes()
			Case 3
				frmUnitView.Show()
			Case 4
				If ViewShipWake Or LastShip >= 0 Then
					mnuViewOptions(4).Text = "Show Ship Wakes"
					ViewShipWake = False
					LastShip = -1
				Else
					mnuViewOptions(4).Text = "Hide Ship Wakes"
					ViewShipWake = True
					LastShip = -1
				End If
				DrawHexes()
			Case 5
				EventMarkers.Clear()
				DrawHexes()
			Case 9
				On Error Resume Next
				
				If Not MapValid Then Exit Sub
				
				cstring = InputBox("Enter sector coordinates (x,y)", "Center Map", SectorString(SelX, SelY))
				
				'092403 rjk: Cancel should not return an error msgbox, this is also check blank entry
				If cstring = "" Then
					Exit Sub
				End If
				
				If ParseSectors(sx, sy, cstring) Then
					MoveMap(sx, sy)
				Else
					MsgBox("Sector Entry Invalid - Center Function Aborted", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly)
				End If
			Case 11 'Refresh screen
				DrawHexes()
			Case 13 'Options
				frmOptions.Show()
		End Select
	End Sub
	
	Public Sub mnuZoomOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuZoomOptions.Click
		Dim Index As Short = mnuZoomOptions.GetIndex(eventSender)
		Dim strTemp As String
		Select Case Index
			Case 1
				Magnification = Magnification * 0.25
			Case 2
				Magnification = Magnification * 0.5
			Case 3
				Magnification = Magnification * 0.8
			Case 4
				Magnification = Magnification * 1.25
			Case 5
				Magnification = Magnification * 1.75
			Case 6
				Magnification = Magnification * 2.5
			Case 8
				strTemp = InputBox("Current Magnification - " & VB6.Format(Magnification, "#0.00"), "Set Magnification")
				If Len(strTemp) > 0 Then
					If Val(strTemp) > 0 Then
						Magnification = Val(strTemp)
					End If
				End If
			Case 9
				Magnification = CSng(GetSetting(APPNAME, "Layout", "Magnification", CStr(1)))
				
		End Select
		
		'Make sure values are reasonable
		If Magnification > 6 Then
			Magnification = 6
		ElseIf Magnification < 0.2 Then 
			Magnification = 0.2
		End If
		
		'Erase the grid
		SetHexSideLength(picMap, (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * Magnification)
		DrawHexes()
	End Sub
	
	Public Sub mnuUnitsMarkNukes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitsMarkNukes.Click
		EventMarkNukes()
	End Sub
	
	'103003 rjk: Added Keys for moving about on the map
	'110603 rjk: Modified the Key pattern
	'UPGRADE_ISSUE: PictureBox event picMap.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picMap_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Dim bSectorFound As Boolean
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Left
				bSectorFound = FillSectorBox(SelX - 2, SelY)
			Case System.Windows.Forms.Keys.Right
				bSectorFound = FillSectorBox(SelX + 2, SelY)
			Case System.Windows.Forms.Keys.Up
				If (SelX Mod 2) = 1 Then
					bSectorFound = FillSectorBox(SelX - 1, SelY - 1)
				Else
					bSectorFound = FillSectorBox(SelX + 1, SelY - 1)
				End If
			Case System.Windows.Forms.Keys.Down
				If (SelX Mod 2) = 1 Then
					bSectorFound = FillSectorBox(SelX - 1, SelY + 1)
				Else
					bSectorFound = FillSectorBox(SelX + 1, SelY + 1)
				End If
			Case Else
				Exit Sub
		End Select
		
		DisplaySelect = enumUnitDisplay.udUNKNOWN
		rsEnemyUnit.Seek("=", SelX, SelY)
		If rsEnemyUnit.NoMatch Then
			rsNuke.Seek("=", SelX, SelY)
			If rsNuke.NoMatch Then
				rsPlane.Seek("=", SelX, SelY)
				If rsPlane.NoMatch Then
					rsLand.Seek("=", SelX, SelY)
					If rsLand.NoMatch Then
						rsShip.Seek("=", SelX, SelY)
						If Not rsShip.NoMatch Then
							DisplaySelect = enumUnitDisplay.udSHIP
							FillGrid()
						End If
					Else
						DisplaySelect = enumUnitDisplay.udLAND
						FillGrid()
					End If
				Else
					DisplaySelect = enumUnitDisplay.udPLANE
					FillGrid()
				End If
			Else
				DisplaySelect = enumUnitDisplay.udNUKE
				FillGrid()
			End If
		Else
			DisplaySelect = enumUnitDisplay.udENEMY
			FillGrid()
		End If
		DisplaySectorPanel()
		enableOrDisableMenus(bSectorFound) '112103 rjk: Added bSectorFound to better determine what menus are enabled
	End Sub
	
	Private Sub picMap_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMap.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim PosX As Single
		Dim PosY As Single
		Dim bSectorFound As Boolean
		' Dim bShift As Boolean   removed efj 8/2003
		Dim bCtl As Boolean
		Dim n As enumUnitDisplay
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		'Dim bAlt As Boolean
		
		' bShift = False   removed efj 8/2003
		If (Shift And VB6.ShiftConstants.ShiftMask) = 1 Then
			' bShift = True   removed efj 8/2003
		End If
		
		bCtl = False
		If (Shift And VB6.ShiftConstants.CtrlMask) = 2 Then
			bCtl = True
		End If
		
		'On Error GoTo picmap_MouseDown_End
		
		GridSelection = False
		
		If DrawingPath Then
			If Button = VB6.MouseButtonConstants.RightButton Then
				RemoveFromPath()
			Else
				AddtoPath()
			End If
			Exit Sub
		End If
		
		'fill the sector box
		bSectorFound = FillSectorBox(MxPos, MyPos)
		
		'if you are marking territories
		If MarkingTerritory And bSectorFound Then
			On Error GoTo 0
			ToggleTerritory(MxPos, MyPos)
			'UPGRADE_ISSUE: Control ToggleHex could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.ToggleHex(MxPos, MyPos, lblDes(1).Text)
			Exit Sub
		End If
		
		'if you are marking territories
		If WorldBuilding And bSectorFound Then
			On Error GoTo 0
			'UPGRADE_ISSUE: Control DesignateHex could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.DesignateHex(MxPos, MyPos, Button)
			If MapValid Then
				UpdateScreenInfo(MxPos, MyPos)
				'Get Center position and draw new hex
				Coord(MxPos, MyPos, PosX, PosY)
				DrawSector((Me.picMap), PosX, PosY, MxPos, MyPos)
			End If
			Exit Sub
		End If
		
		
		
		'Set Indexs
		rsLand.Index = "location"
		rsShip.Index = "location"
		rsPlane.Index = "location"
		rsEnemyUnit.Index = "location"
		rsNuke.Index = "location"
		bShips = False
		bLands = False
		bPlanes = False
		benemys = False
		bNukes = False
		
		'check for ship/land/plane/nuke
		'rjk 05/13/03: Remove the checks for PromptUp so the sector panel is always updated, also
		'              remove the checks for ViewUnits so the sector panel is always updated independent
		'              of show/hide unit option
		'              moved the tab control to DisplaySectorPanel
		
		n = enumUnitDisplay.udUNKNOWN
		rsNuke.Seek("=", MxPos, MyPos)
		If Not rsNuke.NoMatch Then
			bNukes = True
		End If
		rsEnemyUnit.Seek("=", MxPos, MyPos)
		If Not rsEnemyUnit.NoMatch Then
			n = enumUnitDisplay.udENEMY
			benemys = True
		End If
		
		rsNuke.Seek("=", MxPos, MyPos)
		If Not rsNuke.NoMatch Then
			n = enumUnitDisplay.udNUKE
			bNukes = True
		End If
		rsPlane.Seek("=", MxPos, MyPos)
		If Not rsPlane.NoMatch Then
			n = enumUnitDisplay.udPLANE
			bPlanes = True
		End If
		
		rsLand.Seek("=", MxPos, MyPos)
		If Not rsLand.NoMatch Then
			n = enumUnitDisplay.udLAND
			bLands = True
		End If
		
		rsShip.Seek("=", MxPos, MyPos)
		If Not rsShip.NoMatch Then
			n = enumUnitDisplay.udSHIP
			bShips = True
		End If
		
		If n <> enumUnitDisplay.udUNKNOWN Then
			'if you are on the list option, you don't need to switch unless it enemy only
			If DisplaySelect <> enumUnitDisplay.udList Or n = enumUnitDisplay.udENEMY Then
				'if whatever you are currently displaying is present in the sector, stay on that
				If Not ((DisplaySelect = enumUnitDisplay.udSHIP And bShips) Or (DisplaySelect = enumUnitDisplay.udLAND And bLands) Or (DisplaySelect = enumUnitDisplay.udPLANE And bPlanes) Or (DisplaySelect = enumUnitDisplay.udNUKE And bNukes) Or (DisplaySelect = enumUnitDisplay.udENEMY And benemys)) Then
					'if you are already not displaying the first choice, flip the selection
					If DisplaySelect <> n Then
						DisplaySelect = n
						SubType = enumUnitType.TYPE_ALL
					End If
				End If
			End If
			DisplaySubset = enumSubset.ssSECT
			' 092303 rjk: removed secx and secy and switched to SelX and SelY
			'secx = MxPos
			'secy = MyPos
			'FillSubSetFlag = True 092503 rjk: Eliminated
			FillGrid()
		Else
			DisplaySelect = enumUnitDisplay.udUNKNOWN
		End If
		
		DisplaySectorPanel() 'Updated the Tab display
		
		'This subroutine puts the clicked sector designation in the text box of
		'the prompt form, and activates the form
		
		Dim tb As System.Windows.Forms.TextBox
		If PromptUp Then
			If Not (PromptForm.ActiveControl Is Nothing) Then 'drk@unxsoft.com 7/2/02.  this seems to be required to prevent crashes in vb6?
				If Button = VB6.MouseButtonConstants.LeftButton And (TypeOf PromptForm.ActiveControl Is System.Windows.Forms.TextBox) Then 'drk@unxsoft.com: this seems to be required to prevent crashes in vb6?
					tb = PromptForm.ActiveControl
					If tb.Name = "txtOrigin" Or tb.Name = "txtDest" Or tb.Name = "txtPath" Or tb.Name = "txtSector" Then '111603 rjk: Added txtSector frmPromptSector
						tb.Text = SectorString(MxPos, MyPos)
						'101803 rjk: Added multiple sectors selection
					ElseIf tb.Name = "txtMultOrigin" Or tb.Name = "txtMultDest" Or tb.Name = "txtMultPath" Then 
						If bCtl And InStr(tb.Text, ";") = 0 Then 'Ensure no waypoints present
							If Len(tb.Text) > 0 Then
								If InStr(tb.Text, "\") = 0 Then 'one item
									If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (Len(tb.Text) = Len(SectorString(MxPos, MyPos))) Then
										tb.Text = ""
									Else
										tb.Text = tb.Text & "\" & SectorString(MxPos, MyPos)
									End If
								Else 'multiple items
									'check left position
									If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (InStr(tb.Text, "\") = Len(SectorString(MxPos, MyPos)) + 1) Then
										tb.Text = Mid(tb.Text, InStr(tb.Text, "\") + 1)
									Else
										If InStr(tb.Text, "\" & SectorString(MxPos, MyPos) & "\") > 0 Then 'check middle positions
											tb.Text = Replace(tb.Text, "\" & SectorString(MxPos, MyPos) & "\", "\")
										Else
											If (InStr(InStrRev(tb.Text, "\") + 1, tb.Text, SectorString(MxPos, MyPos)) > 0) And ((InStrRev(tb.Text, "\") + Len(SectorString(MxPos, MyPos))) = Len(tb.Text)) Then
												tb.Text = VB.Left(tb.Text, InStrRev(tb.Text, "\") - 1)
											Else
												tb.Text = tb.Text & "\" & SectorString(MxPos, MyPos)
											End If
										End If
									End If
								End If
							Else
								tb.Text = SectorString(MxPos, MyPos)
							End If
						Else
							tb.Text = SectorString(MxPos, MyPos)
						End If
					ElseIf tb.Name = "txtTarget" Then 
						If Len(strEnemyid) = 0 Then
							tb.Text = SectorString(MxPos, MyPos)
						ElseIf InStr(strEnemyid, "/") = 0 Then 
							tb.Text = strEnemyid
						Else
							DisplaySelect = enumUnitDisplay.udENEMY
							DisplaySubset = enumSubset.ssSECT
							' 092303 rjk: removed secx and secy and switched to SelX and SelY
							'secx = MxPos
							'secy = MyPos
							'FillSubSetFlag = True 092503 rjk: Eliminated
							SubType = enumUnitType.TYPE_ALL
							FillGrid()
							DisplayUnitBox() 'rjk 05/13/03: switched over DisplayUnitBox function instead of DisplaySectorPanel 3
						End If
						'092003 rjk: added check to make sure what SHIPs
						'            ElseIf tb.Name = "txtUnit" And Len(strShipid) > 0 And LastUnitDisplay = udSHIP Then
						'                tb.Text = strShipid
						'092003 rjk: added to get Plane
						'            ElseIf tb.Name = "txtUnit" And Len(strPlaneid) > 0 And LastUnitDisplay = udPLANE Then
						'                tb.Text = strPlaneid
						'092003 rjk: added to get Land Unit
						'            ElseIf tb.Name = "txtUnit" And Len(strLandid) > 0 And LastUnitDisplay = udLAND Then
						'                tb.Text = strLandid
						'101903 rjk: Removed frmPromptMove, it now uses txtMultPath
						'103103 rjk: Added back with Multiple Sector Selection capability
					ElseIf PromptForm.Name = "frmPromptMove" Then 
						'UPGRADE_ISSUE: Control txtMultPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						tb = PromptForm.txtMultPath
						If bCtl And InStr(tb.Text, ";") = 0 Then 'Ensure no waypoints present
							If Len(tb.Text) > 0 Then
								If InStr(tb.Text, "\") = 0 Then 'one item
									If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (Len(tb.Text) = Len(SectorString(MxPos, MyPos))) Then
										tb.Text = ""
									Else
										tb.Text = tb.Text & "\" & SectorString(MxPos, MyPos)
									End If
								Else 'multiple items
									'check left position
									If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (InStr(tb.Text, "\") = Len(SectorString(MxPos, MyPos)) + 1) Then
										tb.Text = Mid(tb.Text, InStr(tb.Text, "\") + 1)
									Else
										If InStr(tb.Text, "\" & SectorString(MxPos, MyPos) & "\") > 0 Then 'check middle positions
											tb.Text = Replace(tb.Text, "\" & SectorString(MxPos, MyPos) & "\", "\")
										Else
											If (InStr(InStrRev(tb.Text, "\") + 1, tb.Text, SectorString(MxPos, MyPos)) > 0) And ((InStrRev(tb.Text, "\") + Len(SectorString(MxPos, MyPos))) = Len(tb.Text)) Then
												tb.Text = VB.Left(tb.Text, InStrRev(tb.Text, "\") - 1)
											Else
												tb.Text = tb.Text & "\" & SectorString(MxPos, MyPos)
											End If
										End If
									End If
								End If
							Else
								tb.Text = SectorString(MxPos, MyPos)
							End If
						Else
							tb.Text = SectorString(MxPos, MyPos)
						End If
					ElseIf PromptForm.Name = "frmPromptLaunch" Or PromptForm.Name = "frmPromptTrans" Or PromptForm.Name = "frmPromptNav" Then 
						'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
					End If
					
					PromptForm.Activate()
				End If
			End If '(PromptForm.ActiveControl <> Nothing)
			'if control key is down, add to text line
		ElseIf bCtl Then 
			'    txtEntry.Text = txtEntry.Text + " " + SectorString(MxPos, MyPos)
			str_Renamed = " " & SectorString(MxPos, MyPos) & " "
			'check for a sector substitute.  If found, the put in sector and execute
			n = InStr(txtEntry.Text, "&s")
			If n = 0 Then n = InStr(txtEntry.Text, "&S")
			If n > 0 Then
				str_Renamed = VB.Left(txtEntry.Text, n - 1) & str_Renamed & Mid(txtEntry.Text, n + 2)
				frmEmpCmd.SubmitFromCommandLine(str_Renamed)
			ElseIf Not CommandCursorFocus Then 
				txtEntry.SelectedText = str_Renamed
				CommandCursorPos = CommandCursorPos + Len(str_Renamed)
				SetCommandFocus()
			Else
				txtEntry.SelectedText = str_Renamed
			End If
		End If
		'092503 rjk: Do when a new sector is selected, need for top level menu
		'112103 rjk: Added bSectorFound to better determine what menus are enabled
		enableOrDisableMenus(bSectorFound)
		Dim varShipNum As Object
		Dim NewShip As Short
		If Button = VB6.MouseButtonConstants.RightButton Then
			If Not PromptUp Then 'Don't prompt when already prompted
				PopUpMenuSource = DMAP_PUMS_MAP
				'enableOrDisableMenus 'drk 6/5/02 092503 rjk: switched to be when a select is selected
				If Not (bShips Or bLands Or bPlanes Or bNukes Or bDeity) Then
					If bSectorFound Then
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuSectCmd)
					Else
						rsBmap.Seek("=", MxPos, MyPos)
						If rsBmap.NoMatch Then
							'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							PopupMenu(mnuAttack)
						ElseIf rsBmap.Fields("des").Value <> "." And rsBmap.Fields("des").Value <> "X" Then  '112703 rjk: Added check for mined sea sector
							'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							PopupMenu(mnuAttack)
						Else
							mnuSeaCenter.Visible = True
							'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							PopupMenu(mnuSea) '091403 rjk: Added Center Map for Sea
							If Len(strEnemyid) > 0 Then
								'set enemp ship wake to next in sector
								'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object varShipNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varShipNum = ParseUnitString(strEnemyid)
								'UPGRADE_WARNING: Couldn't resolve default property of object varShipNum(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								NewShip = CShort(varShipNum(1))
								For n = 1 To UBound(varShipNum) - 1
									If LastShip = Int(CDbl(varShipNum(n))) Then
										NewShip = Int(CDbl(varShipNum(n)))
									End If
								Next n
								LastShip = NewShip
								DrawHexes()
							End If
						End If
					End If
				Else
					'092503 rjk: Moved to enableOrDisableMenus
					'mnuLand.Visible = blands
					'mnuShip.Visible = bships
					'mnuPlane.Visible = bplanes
					'mnuNuke.Visible = bnukes
					'mnuDeity.Visible = bDeity
					'112003 rjk: Suppress Sector menu when on sea sector, 112103 moved to enableOrDisableMenus
					'mnuSectCmd.Visible = bSectorFound
					If bSectorFound Then
						'mnuSectCmd.Visible = True '112003 rjk: Suppress Sector menu when on sea sector
						'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PopupMenu(mnuUnitCommand)
					Else '112703 rjk: Changed to show Attack and Unit menus together on enemy sectors
						rsBmap.Seek("=", MxPos, MyPos)
						If Not rsBmap.NoMatch Then
							If rsBmap.Fields("des").Value = "." Or rsBmap.Fields("des").Value = "X" Then
								If Not (bLands Or bPlanes Or bDeity) Then
									'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									PopupMenu(mnuShip)
								ElseIf Not (bLands Or bShips Or bDeity) Then 
									'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									PopupMenu(mnuPlane)
								ElseIf Not (bPlanes Or bShips Or bDeity) Then  '112003 rjk: Change one bships to bplanes
									'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									PopupMenu(mnuLand)
								Else
									'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									PopupMenu(mnuUnitCommand)
								End If
							Else
								'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								PopupMenu(mnuUnitCommand)
							End If
						Else
							'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							PopupMenu(mnuUnitCommand)
						End If
					End If
				End If
			Else
				'if there is an OK button, then run it.
				On Error Resume Next
				'UPGRADE_ISSUE: Control cmdOK_Click could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.cmdOK_Click()
				On Error GoTo picmap_MouseDown_End
			End If
		End If
picmap_MouseDown_End: 
		
	End Sub
	
	'112103 rjk: Added bSectorFound to better determine what menus should be active.
	Private Sub enableOrDisableMenus(ByRef bSectorFound As Boolean)
		'060502 drk: enable or disable menu items intelligently based on the sector designation
		'fixme, this section could be significantly expanded upon.
		'091303 rjk: added check for grind,Convert,Demobilize,Enlist,Shoot,Improve,Start,Stop,Spy,Explore
		'103003 rjk: Switched to SelX and SelY from MxPos and MyPos.
		'rsSectors.Seek "=", MxPos, MyPos
		rsSectors.Seek("=", SelX, SelY)
		If Not rsSectors.NoMatch Then
			With rsSectors.Fields("des")
				mnuSectCmdFire.Enabled = (.Value = "f")
				mnuSectCmdBuild.Enabled = ((.Value = "h") Or (.Value = "*") Or (.Value = "!") Or (.Value = "n") Or (rsSectors.Fields("coast").Value = 1))
				mnuSectCmdRadar.Enabled = (.Value = ")")
				mnuSectCmdCap.Enabled = (.Value = "c") Or (.Value = "^")
			End With
			Me.mnuSectCmdAssault.Enabled = rsSectors.Fields("coast").Value = 1
			mnuSectCmdGrind.Enabled = rsSectors.Fields("bar").Value > 0
			'100203 rjk: Changed to Anti to not have to be occuppied but must have mil. and mob.
			'mnuSectCmdAnti = (rsSectors.Fields("*").Value = "*") And (rsSectors.Fields("mob").Value > 0)
			mnuSectCmdAnti.Enabled = (rsSectors.Fields("mil").Value > 0) And (rsSectors.Fields("mob").Value > 0)
			mnuSectCmdConvert.Enabled = (rsSectors.Fields("*").Value = "*") And (rsSectors.Fields("civ").Value > 0) And (rsSectors.Fields("mob").Value > 0)
			mnuSectCmdDemob.Enabled = (rsSectors.Fields("mil").Value > 0) And (rsSectors.Fields("mob").Value > 0)
			If frmOptions.b2K4 Then
				mnuSectCmdEnlist.Enabled = (MilReserves > 0) And (rsSectors.Fields("mob").Value > 0)
			Else
				mnuSectCmdEnlist.Enabled = (rsSectors.Fields("civ").Value > 0) And (MilReserves > 0) And (rsSectors.Fields("mob").Value > 0)
			End If
			mnuSectCmdShoot.Enabled = ((rsSectors.Fields("civ").Value > 0) Or (rsSectors.Fields("uw").Value > 0)) And (rsSectors.Fields("mob").Value > 0)
			mnuSectCmdImprove.Enabled = (rsSectors.Fields("road").Value < 100) Or (rsSectors.Fields("rail").Value < 100) Or (rsSectors.Fields("defense").Value < 100) And (rsSectors.Fields("mob").Value > 0)
			mnuSectCmdStart.Enabled = rsSectors.Fields("off").Value = 1
			mnuSectCmdStop.Enabled = rsSectors.Fields("off").Value <> 1
			mnuSectCmdSpy.Enabled = rsSectors.Fields("mil").Value > 0
			mnuSectCmdExplore.Enabled = ((rsSectors.Fields("civ").Value + rsSectors.Fields("mil").Value) > 1) And (rsSectors.Fields("mob").Value > 0)
			If frmOptions.bSPAtlantis Then
				mnuSectCmdImprove.Enabled = True
				If rsSectors.Fields("gold").Value > 0 Then
					mnuSectCmdRadar.Enabled = True
				End If
				If rsSectors.Fields("fert").Value > 0 Then
					mnuSectCmdFire.Enabled = True
				End If
			End If
		End If
		If VersionCheck(4, 3, 6) = WinAceRoutines.enumVersion.VER_LESS Then
			mnuLandStart(0).Visible = False
			mnuLandStart(1).Visible = False
			mnuShipMisc(7).Visible = False
			mnuShipMisc(8).Visible = False
			mnuPlaneOther(4).Visible = False
			mnuPlaneOther(5).Visible = False
			mnuNukeStart(0).Visible = False
			mnuNukeStart(1).Visible = False
		End If
		
		'092503 rjk: moved from RightClick to allow modification for top level menu
		'112103 rjk: Set Visibility only if there is something to show
		If bSectorFound Or bLands Or bShips Or bPlanes Or bNukes Or bDeity Then
			mnuUnitCommand.Visible = True
			mnuDeity.Visible = True 'Need as least one menu visible
			mnuSectCmd.Visible = bSectorFound
			mnuLand.Visible = bLands
			mnuShip.Visible = bShips
			mnuPlane.Visible = bPlanes
			mnuNuke.Visible = bNukes
			mnuDeity.Visible = bDeity
			If Not bSectorFound Then '112703 rjk: Added to display attack and unit menu together
				rsBmap.Seek("=", MxPos, MyPos)
				If Not rsBmap.NoMatch Then
					If rsBmap.Fields("des").Value = "." Or rsBmap.Fields("des").Value = "X" Then
						mnuUnitAttack.Visible = False
					Else
						mnuUnitAttack.Visible = True
					End If
				Else
					mnuUnitAttack.Visible = True
				End If
			Else
				mnuUnitAttack.Visible = False
			End If
		Else
			mnuUnitCommand.Visible = False
		End If
	End Sub
	
	Public Function FillSectorBox(ByRef SxPos As Short, ByRef SyPos As Short) As Boolean
		Dim strOwner As String
		On Error GoTo Empire_Error
		
		'set the selected sector indicators
		SelX = SxPos
		SelY = SyPos
		
		'Unhighlight the old hex and highlight the current one
		Dim PosX As Single
		Dim PosY As Single
		Dim oldx As Short
		Dim oldy As Short
		oldx = CurrentSectorX
		oldy = CurrentSectorY
		CurrentSectorX = SxPos
		CurrentSectorY = SyPos
		
		If MapValid Then
			On Error Resume Next 'might not still be visible
			'Get Center position and redraw old current hex
			Coord(oldx, oldy, PosX, PosY)
			DrawSector((Me.picMap), PosX, PosY, oldx, oldy)
			DrawEvent(picMap, oldx, oldy)
			'Get Center position and draw new hex
			Coord(SxPos, SyPos, PosX, PosY)
			DrawSector((Me.picMap), PosX, PosY, SxPos, SyPos)
			DrawEvent(picMap, SxPos, SyPos)
			On Error GoTo Empire_Error
		End If
		
		'Get Sector Information
		FillSectorBox = False
		lblSect.Text = "Sector " & SectorString(SxPos, SyPos)
		
		
		rsSectors.Seek("=", SxPos, SyPos)
		Dim n As Short
		Dim v As productionDataType
		Dim strlab As String
		If Not rsSectors.NoMatch Then
			FillSectorBox = True
			SelSectType = SEL_SECT_OWN
			lblDes(1).Text = vbNullString
			'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(rsSectors.Fields(des).Value). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lblDes(1).Text = colDes.Item(rsSectors.Fields("des").Value)
			
			'drk 7/22/03
			If (CapSect = SectorString(SxPos, SyPos)) Then lblDes(1).Text = lblDes(1).Text & " : National capital"
			lblDes(1).Text = CStr(rsSectors.Fields("eff").Value) & "% " & lblDes(1).Text
			If Not bDeity Then
				If Len(Trim(rsSectors.Fields("sdes").Value)) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lblNDes.Text = "new = " + colDes.Item(Trim(rsSectors.Fields("sdes").Value))
				Else
					lblNDes.Text = vbNullString
				End If
			Else
				' for deities - store newdes in top label, put owner in the second
				If Len(Trim(rsSectors.Fields("sdes").Value)) > 0 Then
					lblDes(1).Text = lblDes(1).Text & " (newdes=" & Trim(rsSectors.Fields("sdes").Value) & ")" '121203 rjk: Added closing bracket
				End If
				lblNDes.Text = Nations.NationName((rsSectors.Fields("country").Value))
			End If
			
			lblDes(2).Text = lblNDes.Text 'rjk 5/12/03: added new designation info to distribution tab
			lblDes(0).Text = lblDes(1).Text
			'240104 rjk: Added distribution point title
			If SxPos = rsSectors.Fields("dist_x").Value And SyPos = rsSectors.Fields("dist_y").Value Then
				lblDes(4).Text = "No Dist. Point"
			Else
				lblDes(4).Text = SectorString((rsSectors.Fields("dist_x").Value), (rsSectors.Fields("dist_y").Value))
			End If
			'    Label1(3).Caption = CStr(rsSectors.Fields("civ").Value)
			'    Label1(4).Caption = CStr(rsSectors.Fields("mil").Value)
			'    Label1(5).Caption = CStr(rsSectors.Fields("uw").Value)
			Label1(3).Text = CStr(Items.FindByLetter("c").DatabaseValue(rsSectors))
			Label1(0).Text = Items.FindByLetter("c").SectorPanelLabel
			Label1(4).Text = CStr(Items.FindByLetter("m").DatabaseValue(rsSectors))
			Label1(1).Text = Items.FindByLetter("m").SectorPanelLabel
			Label1(5).Text = CStr(Items.FindByLetter("u").DatabaseValue(rsSectors))
			Label1(2).Text = Items.FindByLetter("u").SectorPanelLabel
			Label1(9).Text = CStr(rsSectors.Fields("mob").Value)
			Label1(11).Text = CStr(rsSectors.Fields("avail").Value)
			Label1(10).Text = CStr(rsSectors.Fields("work").Value)
			Label1(20).Text = CStr(rsSectors.Fields("min").Value)
			Label1(19).Text = CStr(rsSectors.Fields("gold").Value)
			Label1(18).Text = CStr(rsSectors.Fields("fert").Value)
			Label1(32).Text = CStr(rsSectors.Fields("ocontent").Value)
			Label1(31).Text = CStr(rsSectors.Fields("uran").Value)
			Label1(14).Text = CStr(rsSectors.Fields("road").Value)
			Label1(13).Text = CStr(rsSectors.Fields("rail").Value)
			Label1(12).Text = CStr(rsSectors.Fields("defense").Value)
			Label1(26).Text = CStr(rsSectors.Fields("fallout").Value)
			Label1(25).Text = CStr(rsSectors.Fields("terr").Value)
			If frmOptions.bSPAtlantis Then
				Label1(23).Text = "Run:"
				Label1(22).Text = "Rad:"
				Label1(21).Text = "Fort:"
				Label1(35).Text = "Nav:"
			End If
			
			'Commodities
			'    Label1(60).Caption = CStr(rsSectors.Fields("food").Value)
			'    Label1(61).Caption = CStr(rsSectors.Fields("shell").Value)
			'    Label1(62).Caption = CStr(rsSectors.Fields("gun").Value)
			'    Label1(63).Caption = CStr(rsSectors.Fields("pet").Value)
			'    Label1(64).Caption = CStr(rsSectors.Fields("iron").Value)
			'    Label1(65).Caption = CStr(rsSectors.Fields("dust").Value)
			'    Label1(66).Caption = CStr(rsSectors.Fields("bar").Value)
			'    Label1(67).Caption = CStr(rsSectors.Fields("oil").Value)
			'    Label1(68).Caption = CStr(rsSectors.Fields("lcm").Value)
			'    Label1(69).Caption = CStr(rsSectors.Fields("hcm").Value)
			'    Label1(70).Caption = CStr(rsSectors.Fields("rad").Value)
			Label1(60).Text = CStr(Items.FindByLetter("f").DatabaseValue(rsSectors))
			Label1(55).Text = Items.FindByLetter("f").SectorPanelLabel
			Label1(61).Text = CStr(Items.FindByLetter("s").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(54), "s", rsSectors)
			Label1(62).Text = CStr(Items.FindByLetter("g").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(53), "g", rsSectors)
			Label1(63).Text = CStr(Items.FindByLetter("p").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(24), "p", rsSectors)
			Label1(64).Text = CStr(Items.FindByLetter("i").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(27), "i", rsSectors)
			Label1(65).Text = CStr(Items.FindByLetter("d").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(30), "d", rsSectors)
			Label1(66).Text = CStr(Items.FindByLetter("b").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(49), "b", rsSectors)
			Label1(67).Text = CStr(Items.FindByLetter("o").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(48), "o", rsSectors)
			Label1(68).Text = CStr(Items.FindByLetter("l").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(47), "l", rsSectors)
			Label1(69).Text = CStr(Items.FindByLetter("h").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(38), "h", rsSectors)
			Label1(70).Text = CStr(Items.FindByLetter("r").DatabaseValue(rsSectors))
			DisplayCommodity(Label1(39), "r", rsSectors)
			
			'Distribution
			Label1(100).Text = CStr(rsSectors.Fields("c_dist").Value)
			Label1(99).Text = Items.FindByLetter("c").DistributionPanelLabel
			Label1(101).Text = CStr(rsSectors.Fields("m_dist").Value)
			Label1(80).Text = Items.FindByLetter("m").DistributionPanelLabel
			Label1(102).Text = CStr(rsSectors.Fields("u_dist").Value)
			Label1(94).Text = Items.FindByLetter("u").DistributionPanelLabel
			Label1(103).Text = CStr(rsSectors.Fields("f_dist").Value)
			Label1(76).Text = Items.FindByLetter("f").DistributionPanelLabel
			Label1(104).Text = CStr(rsSectors.Fields("s_dist").Value)
			Label1(77).Text = Items.FindByLetter("s").DistributionPanelLabel
			Label1(105).Text = CStr(rsSectors.Fields("g_dist").Value)
			Label1(78).Text = Items.FindByLetter("g").DistributionPanelLabel
			Label1(106).Text = CStr(rsSectors.Fields("p_dist").Value)
			Label1(97).Text = Items.FindByLetter("p").DistributionPanelLabel
			Label1(107).Text = CStr(rsSectors.Fields("i_dist").Value)
			Label1(96).Text = Items.FindByLetter("i").DistributionPanelLabel
			Label1(108).Text = CStr(rsSectors.Fields("d_dist").Value)
			Label1(95).Text = Items.FindByLetter("d").DistributionPanelLabel
			Label1(109).Text = CStr(rsSectors.Fields("b_dist").Value)
			Label1(82).Text = Items.FindByLetter("b").DistributionPanelLabel
			Label1(110).Text = CStr(rsSectors.Fields("o_dist").Value)
			Label1(83).Text = Items.FindByLetter("o").DistributionPanelLabel
			Label1(111).Text = CStr(rsSectors.Fields("l_dist").Value)
			Label1(84).Text = Items.FindByLetter("l").DistributionPanelLabel
			Label1(112).Text = CStr(rsSectors.Fields("h_dist").Value)
			Label1(91).Text = Items.FindByLetter("h").DistributionPanelLabel
			Label1(113).Text = CStr(rsSectors.Fields("r_dist").Value)
			Label1(90).Text = Items.FindByLetter("r").DistributionPanelLabel
			
			'Delivery
			Label1(120).Text = DeliveryLabel((rsSectors.Fields("c_cut").Value), (rsSectors.Fields("c_del").Value))
			Label1(33).Text = Items.FindByLetter("c").DistributionPanelLabel
			Label1(121).Text = DeliveryLabel((rsSectors.Fields("m_cut").Value), (rsSectors.Fields("m_del").Value))
			Label1(36).Text = Items.FindByLetter("m").DistributionPanelLabel
			Label1(122).Text = DeliveryLabel((rsSectors.Fields("u_cut").Value), (rsSectors.Fields("u_del").Value))
			Label1(42).Text = Items.FindByLetter("u").DistributionPanelLabel
			Label1(123).Text = DeliveryLabel((rsSectors.Fields("f_cut").Value), (rsSectors.Fields("f_del").Value))
			Label1(88).Text = Items.FindByLetter("f").DistributionPanelLabel
			Label1(124).Text = DeliveryLabel((rsSectors.Fields("s_cut").Value), (rsSectors.Fields("s_del").Value))
			Label1(87).Text = Items.FindByLetter("s").DistributionPanelLabel
			Label1(125).Text = DeliveryLabel((rsSectors.Fields("g_cut").Value), (rsSectors.Fields("g_del").Value))
			Label1(86).Text = Items.FindByLetter("g").DistributionPanelLabel
			Label1(126).Text = DeliveryLabel((rsSectors.Fields("p_cut").Value), (rsSectors.Fields("p_del").Value))
			Label1(43).Text = Items.FindByLetter("p").DistributionPanelLabel
			Label1(127).Text = DeliveryLabel((rsSectors.Fields("i_cut").Value), (rsSectors.Fields("i_del").Value))
			Label1(44).Text = Items.FindByLetter("i").DistributionPanelLabel
			Label1(128).Text = DeliveryLabel((rsSectors.Fields("d_cut").Value), (rsSectors.Fields("d_del").Value))
			Label1(45).Text = Items.FindByLetter("d").DistributionPanelLabel
			Label1(129).Text = DeliveryLabel((rsSectors.Fields("b_cut").Value), (rsSectors.Fields("b_del").Value))
			Label1(75).Text = Items.FindByLetter("b").DistributionPanelLabel
			Label1(130).Text = DeliveryLabel((rsSectors.Fields("o_cut").Value), (rsSectors.Fields("o_del").Value))
			Label1(74).Text = Items.FindByLetter("o").DistributionPanelLabel
			Label1(131).Text = DeliveryLabel((rsSectors.Fields("l_cut").Value), (rsSectors.Fields("l_del").Value))
			Label1(73).Text = Items.FindByLetter("l").DistributionPanelLabel
			Label1(132).Text = DeliveryLabel((rsSectors.Fields("h_cut").Value), (rsSectors.Fields("h_del").Value))
			Label1(52).Text = Items.FindByLetter("h").DistributionPanelLabel
			Label1(133).Text = DeliveryLabel((rsSectors.Fields("r_cut").Value), (rsSectors.Fields("r_del").Value))
			Label1(56).Text = Items.FindByLetter("r").DistributionPanelLabel
			'101703 rjk: Strength labels 147 to 158 - odd numbers the label even numbers the values
			'112003 rjk: Added a new option for controlling strength updates, blue if suppressed
			If frmOptions.bSuppressStrength = True Then
				Label1(148).ForeColor = System.Drawing.Color.Blue
				Label1(150).ForeColor = System.Drawing.Color.Blue
				Label1(152).ForeColor = System.Drawing.Color.Blue
				Label1(154).ForeColor = System.Drawing.Color.Blue
				Label1(156).ForeColor = System.Drawing.Color.Blue
				Label1(158).ForeColor = System.Drawing.Color.Blue
			Else
				Label1(148).ForeColor = System.Drawing.Color.Black
				Label1(150).ForeColor = System.Drawing.Color.Black
				Label1(152).ForeColor = System.Drawing.Color.Black
				Label1(154).ForeColor = System.Drawing.Color.Black
				Label1(156).ForeColor = System.Drawing.Color.Black
				Label1(158).ForeColor = System.Drawing.Color.Black
			End If
			'112503 rjk: If Null display unknown
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("units").Value) Then
				Label1(148).Text = CStr(rsSectors.Fields("units").Value)
			Else
				Label1(148).Text = "???"
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("runits").Value) Then
				Label1(150).Text = CStr(rsSectors.Fields("runits").Value)
			Else
				Label1(150).Text = "???"
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("lmines").Value) Then
				If rsSectors.Fields("lmines").Value = -1 Then '120303 rjk: display enemy mines in red
					Label1(152).ForeColor = System.Drawing.Color.Red
					Label1(152).Text = "?"
				Else
					Label1(152).Text = CStr(rsSectors.Fields("lmines").Value)
				End If
			Else
				Label1(152).Text = "???"
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("sec_mult").Value) Then
				Label1(154).Text = CStr(rsSectors.Fields("sec_mult").Value)
			Else
				Label1(154).Text = "???"
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("sec_def").Value) Then
				Label1(156).Text = CStr(rsSectors.Fields("sec_def").Value)
			Else
				Label1(156).Text = "???"
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsSectors.Fields("tot_def").Value) Then
				Label1(158).Text = CStr(rsSectors.Fields("tot_def").Value)
			Else
				Label1(158).Text = "???"
			End If
			
			'102303 rjk: Added Terr1/Terr2/Terr3 fields labels 159 to 164 - odd numbers the label even numbers the values
			Label1(160).Text = CStr(rsSectors.Fields("terr1").Value)
			Label1(162).Text = CStr(rsSectors.Fields("terr2").Value)
			Label1(164).Text = CStr(rsSectors.Fields("terr3").Value)
			
			'calculated fields
			n = FoodRequired(rsSectors)
			If n >= 0 Then
				Label1(58).Text = CStr(n)
			Else
				Label1(58).Text = "??"
			End If
			
			If rsSectors.Fields("food").Value - n < 0 Then
				Label1(58).ForeColor = System.Drawing.Color.Red
			Else
				Label1(58).ForeColor = System.Drawing.Color.Black
			End If
			
			
			Label1(41).ForeColor = System.Drawing.Color.Black
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			v = Production(rsSectors)
			If (v.prodValidFlag) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Label1(85).Text = CStr(v.ProdAmount)
				'UPGRADE_WARNING: Couldn't resolve default property of object v.MaxProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Label1(142).Text = CStr(v.MaxProdAmount)
				If v.prodMaxMaterialLimit Then
					Label1(142).ForeColor = System.Drawing.Color.Blue
				Else
					Label1(142).ForeColor = System.Drawing.Color.Black
				End If
				n = CShort(v.ExcessCiv)
				If n < 0 Then
					Label1(41).ForeColor = System.Drawing.Color.Red
					n = 0 - n
				End If
				Label1(41).Text = CStr(n)
				Label1(46).Text = CStr(v.NewEff)
			Else
				Label1(85).Text = vbNullString
				Label1(142).Text = vbNullString
				Label1(41).Text = vbNullString
				Label1(46).Text = vbNullString
			End If
			
			For n = 60 To 70
				If Label1(n).Text = "0" Then
					Label1(n).Text = "--"
				End If
			Next n
			For n = 100 To 113
				If Label1(n).Text = "0" Then
					Label1(n).Text = "--"
				End If
			Next n
			
			For n = 120 To 133
				If Label1(n).Text = "0." Then
					Label1(n).Text = "--"
				End If
			Next n
			
			lblWarn.Text = EventMarkers.Find(SxPos, SyPos)
			If Len(lblWarn.Text) = 0 Then
				If rsSectors.Fields("off").Value = 1 Then
					lblWarn.Text = "Stopped "
				End If
				If rsSectors.Fields("*").Value = "*" Then
					lblWarn.Text = lblWarn.Text & "Not Yours"
				End If
			End If
			
			lblPlayers.Text = ""
			If UBound(currentPlayersNumber) > 0 Then
				For n = 0 To UBound(currentPlayersNumber) - 1
					lblPlayers.Text = lblPlayers.Text & CStr(currentPlayersNumber(n)) & ":" & currentPlayersName(n) & " "
				Next n
				lblPlayers.Text = VB.Left(lblPlayers.Text, Len(lblPlayers.Text) - 1)
			End If
			
			'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
			'If mintCurFrame = 0 Then
			'    mintCurFrame = TabStrip1.SelectedItem.Index
			'End If
			
			'Frame1(0).Visible = False
			'Frame1(3).Visible = False
			'Frame1(mintCurFrame).Visible = True
		Else
			'Check for Enemy Sector
			rsEnemySect.Seek("=", SxPos, SyPos)
			If Not rsEnemySect.NoMatch Then
				'Frame1(mintCurFrame).Visible = False rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
				SelSectType = SEL_SECT_ENEMY
				lblDes(5).Text = vbNullString
				On Error Resume Next 'in case des is not in the collection
				'110304 rjk: Switched to use IsNull for EnemySector designation, display Unknown designation for null entry
				'            if you can not find the information in the rsBmap.
				'            Show ??? for unknown efficiency.
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rsEnemySect.Fields("des").Value) Then
					rsBmap.Seek("=", SxPos, SyPos)
					If rsBmap.NoMatch Then
						lblDes(5).Text = "Unknown designation"
					Else
						If rsBmap.Fields("des").Value <> "?" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(rsBmap.Fields(des)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							lblDes(5).Text = colDes.Item(rsBmap.Fields("des"))
							If rsEnemySect.Fields("eff").Value >= 0 Then
								lblDes(5).Text = CStr(rsEnemySect.Fields("eff").Value) & "% " & lblDes(5).Text
							Else
								lblDes(5).Text = "???% " & lblDes(5).Text
							End If
						Else
							lblDes(5).Text = "Unknown designation"
						End If
					End If
				ElseIf rsEnemySect.Fields("des").Value = "?" Then 
					rsBmap.Seek("=", SxPos, SyPos)
					If rsBmap.NoMatch Then
						lblDes(5).Text = "Unknown designation"
					Else
						If rsBmap.Fields("des").Value <> "?" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(rsBmap.Fields(des)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							lblDes(5).Text = colDes.Item(rsBmap.Fields("des"))
							If rsEnemySect.Fields("eff").Value >= 0 Then
								lblDes(5).Text = CStr(rsEnemySect.Fields("eff").Value) & "% " & lblDes(5).Text
							Else
								lblDes(5).Text = "???% " & lblDes(5).Text
							End If
						Else
							lblDes(5).Text = "Unknown designation"
						End If
					End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(rsEnemySect.Fields(des).Value). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lblDes(5).Text = colDes.Item(rsEnemySect.Fields("des").Value)
					If rsEnemySect.Fields("eff").Value >= 0 Then
						lblDes(5).Text = CStr(rsEnemySect.Fields("eff").Value) & "% " & lblDes(5).Text
					Else
						lblDes(5).Text = "???% " & lblDes(5).Text
					End If
				End If
				On Error GoTo Empire_Error
				
				strOwner = Nations.NationName((rsEnemySect.Fields("owner").Value))
				If strOwner = vbNullString Then
					strOwner = CStr(rsEnemySect.Fields("owner").Value)
				End If
				Label7.Text = "owner = " & strOwner
				n = rsEnemySect.Fields("oldowner").Value
				If n > 0 And n <> rsEnemySect.Fields("owner").Value Then
					Label7.Text = Label7.Text & "; oldown = " & CStr(rsEnemySect.Fields("oldowner").Value)
				End If
				
				'Set Enemy Sectors
				strlab = "098 114 115 138 137 117 118 116 081 079 136"
				For n = 0 To 10
					If rsEnemySect.Fields(n + 6).Value < 0 Then
						Label1(CShort(Mid(strlab, n * 4 + 1, 3))).Text = "???"
					Else
						Label1(CShort(Mid(strlab, n * 4 + 1, 3))).Text = rsEnemySect.Fields(n + 6).Value
					End If
				Next n
				Label8.Text = EventMarkers.Find(SxPos, SyPos)
				If Len(Label8.Text) = 0 Then
					Label8.Text = VB6.Format(ConvertToLocalTimeFromWinACEUTC((rsEnemySect.Fields("timestamp").Value)), "ddd mmm dd hh:mm:ss yyyy")
				End If
				
				'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
				'mintCurFrame = 0
				'Frame1(mintCurFrame).Visible = True
			Else 'Not rsEnemySect.NoMatch => not enemy either.  maybe it is a sea sector?  drk 5/13/03
				rsSea.Seek("=", SxPos, SyPos)
				If Not rsSea.NoMatch Then
					SelSectType = SEL_SECT_SEA
					'101503 rjk: Switch Oil and Fert fields so they display correctly on the screen
					If rsSea.Fields("ocontent").Value < 0 Then
						Label1(146).Text = "???"
					Else
						Label1(146).Text = CStr(rsSea.Fields("ocontent").Value)
					End If
					If rsSea.Fields("fert").Value < 0 Then
						Label1(143).Text = "???"
					Else
						Label1(143).Text = CStr(rsSea.Fields("fert").Value)
					End If
					Label4.Text = EventMarkers.Find(SxPos, SyPos)
				Else
					SelSectType = SEL_SECT_UNKNOWN
				End If
				'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
				'If mintCurFrame <> 3 Then
				'    Frame1(mintCurFrame).Visible = False
				'End If
			End If
			
		End If
		
Empire_Error: 
		
	End Function
	
	Private Function DeliveryLabel(ByRef vCutoff As Object, ByRef vDirection As Object) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object vDirection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If vDirection >= 0 And vDirection <= 7 And frmOptions.bStarWars Then
			'UPGRADE_WARNING: Couldn't resolve default property of object vDirection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vCutoff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DeliveryLabel = CStr(vCutoff) & ":" & CStr(vDirection)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object vDirection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vCutoff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DeliveryLabel = CStr(vCutoff) & CStr(vDirection)
		End If
	End Function
	
	Private Sub DisplayCommodity(ByRef lblComm As System.Windows.Forms.Label, ByRef strComm As String, ByRef rsSect As DAO.Recordset)
		Dim CommodityRatio As String
		
		If frmOptions.bCommodityRatio Then
			If rsSect.Fields("sdes").Value <> " " Then
				rsSectorType.Seek("=", rsSect.Fields("sdes"))
			Else
				rsSectorType.Seek("=", rsSect.Fields("des"))
			End If
			If Not rsSectorType.NoMatch Then
				If rsSectorType.Fields("use1s").Value = strComm Then
					CommodityRatio = " " & CStr(rsSectorType.Fields("use1n").Value)
				End If
				If rsSectorType.Fields("use2s").Value = strComm Then
					CommodityRatio = " " & CStr(rsSectorType.Fields("use2n").Value)
				End If
				If rsSectorType.Fields("use3s").Value = strComm Then
					CommodityRatio = " " & CStr(rsSectorType.Fields("use3n").Value)
				End If
			End If
		End If
		If Len(CommodityRatio) > 0 Then
			lblComm.ForeColor = System.Drawing.ColorTranslator.FromOle(vbMediumGreen)
			lblComm.Text = Items.FindByLetter(strComm).SectorPanelLabel & CommodityRatio
		Else
			lblComm.ForeColor = System.Drawing.Color.Black
			lblComm.Text = Items.FindByLetter(strComm).SectorPanelLabel
		End If
	End Sub
	
	Private Sub PicMap_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PicMap.DoubleClick
		
		Dim Sector As String
		' Dim n As Integer   removed efj 8/2003
		'Dont bring up box when drawing path
		If DrawingPath Then
			Exit Sub
		End If
		
		'Get Center position of click
		Dim PosX As Single
		Dim PosY As Single
		
		Coord(MxPos, MyPos, PosX, PosY)
		If bShips Or bLands Or bPlanes Or benemys Then
			If SubType = enumUnitType.TYPE_ALL Then
				DisplayUnitBox() 'rjk 05/13/03: Switch to DisplayUnitBox function from DisplaySectorPanel 3
			Else
				'        SetUnitType TYPE_ALL 091903 rjk: removed SetUnitType function, do same actions as TYPE_ALL from combo box
				SubType = enumUnitType.TYPE_ALL
				FillGrid()
			End If
			Exit Sub
		End If
		
		Sector = SectorString(MxPos, MyPos)
		rsSectors.Seek("=", MxPos, MyPos)
		If rsSectors.NoMatch Then
			Exit Sub
		End If
		
		rsSectorType.Seek("=", rsSectors.Fields("des"))
		If Not rsSectorType.NoMatch Then
			If rsSectorType.Fields("product").Value <> " " Then
				If (rsSectorType.Fields("des").Value <> 0) And (rsSectors.Fields("sdes").Value <> "") Then
					If (rsSectors.Fields("des").Value = "g" And rsSectors.Fields("gold").Value = 0) Or (rsSectors.Fields("des").Value = "o" And rsSectors.Fields("ocontent").Value = 0) Or (rsSectors.Fields("des").Value = "u" And rsSectors.Fields("uran").Value = 0) Then
						mnuSectCmdDes_Click(mnuSectCmdDes, New System.EventArgs())
						Exit Sub
					End If
				End If
				frmEmpCmd.SubmitEmpireCommand("production " & " " & Sector, True)
				Exit Sub
			End If
		End If
		
		
		Select Case rsSectors.Fields("des").Value
			
			Case ")"
				frmEmpCmd.SubmitEmpireCommand("radar " & " " & Sector, True)
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				frmEmpCmd.SubmitEmpireCommand("bmap *", False)
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			Case "h", "!", "*", "n"
				mnuUnitBuild_Click(mnuUnitBuild, New System.EventArgs())
			Case "-", "~"
				mnuSectCmdDes_Click(mnuSectCmdDes, New System.EventArgs())
			Case "f"
				mnuSectCmdFire_Click(mnuSectCmdFire, New System.EventArgs())
			Case "c"
				mnuReportNation_Click(mnuReportNation, New System.EventArgs())
			Case "b"
				mnuReportBudget_Click(mnuReportBudget, New System.EventArgs())
			Case "w"
				mnuSectCmdDist_Click(mnuSectCmdDist, New System.EventArgs())
				'    Case "j", "k", "l", "t", "d", "i", "m", "e", "%", "p", "a", "r"
				'        frmEmpCmd.SubmitEmpireCommand "production " + " " + Sector, True
				'    Case "g", "o", "u"
				'        If (rsSectors.Fields("des").Value = "g" And _
				''            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
				''            rsSectors.Fields("gold").Value = 0) Or _
				''            (rsSectors.Fields("des").Value = "o" And _
				''            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
				''            rsSectors.Fields("ocontent").Value = 0) Or _
				''            (rsSectors.Fields("des").Value = "u" And _
				''            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
				''            rsSectors.Fields("rad").Value = 0) Then
				'            mnuSectCmdDes_Click
				'        Else
				'            frmEmpCmd.SubmitEmpireCommand "production " + " " + Sector, True
				'        End If
		End Select
		
		
	End Sub
	
	'112903 rjk: Removed ViewUnits check, always display units on bottom status line
	Private Sub picMap_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMap.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim oldx As Short
		Dim oldy As Short
		Dim PosX As Single
		Dim PosY As Single
		Dim AdjX As Short
		Dim AdjY As Short
		Dim HalfLength As Single
		Dim HalfWidth As Single
		Dim hexWidth As Single
		Dim hexHeight As Single
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		If Not MapValid Then
			Exit Sub
		End If
		
		'Save the old position x and y to see if they have changed
		oldx = MxPos
		oldy = MyPos
		
		HalfLength = 0.5 * GetHexSideLength
		HalfWidth = 0.8667 * GetHexSideLength
		
		hexHeight = 2 * GetHexSideLength
		hexWidth = 0.8667 * 2 * GetHexSideLength
		
		MyPos = Int(Y / (hexHeight * 1.5))
		PosY = Y - ((hexHeight * 1.5) * MyPos)
		
		MyPos = MyPos * 2
		If PosY > hexHeight Then
			MyPos = MyPos + 1
			PosY = PosY - (0.75 * hexHeight)
		End If
		
		If (MyPos Mod 2) = 0 Then
			MxPos = Int(X / hexWidth)
			PosX = X - (hexWidth * MxPos)
			MxPos = MxPos * 2
		Else
			MxPos = Int((X - HalfWidth) / hexWidth)
			PosX = X - (hexWidth * MxPos) - HalfWidth
			MxPos = (MxPos * 2) + 1
		End If
		
		If PosY < HalfLength Then
			If PosX <= HalfWidth Then
				If PosY > (HalfLength - (HalfLength * PosX) / HalfWidth) Then
					AdjX = 0
					AdjY = 0
				Else
					AdjX = -1
					AdjY = -1
				End If
			Else
				If PosY > (HalfLength * (PosX - HalfWidth) / HalfWidth) Then
					AdjX = 0
					AdjY = 0
				Else
					AdjX = 1
					AdjY = -1
				End If
			End If
			
		ElseIf PosY > 3 * HalfLength Then 
			If PosX <= HalfWidth Then
				If PosY > (3 * HalfLength + (HalfLength * PosX) / HalfWidth) Then
					AdjX = -1
					AdjY = 1
				Else
					AdjX = 0
					AdjY = 0
				End If
			Else
				If PosY > (4 * HalfLength - (HalfLength * (PosX - HalfWidth) / HalfWidth)) Then
					AdjX = 1
					AdjY = 1
				Else
					AdjX = 0
					AdjY = 0
				End If
			End If
			
		Else
			AdjX = 0
			AdjY = 0
		End If
		
		'Calc new X and Y position of mouse
		MxPos = MxPos + AdjX + OriginX
		MyPos = MyPos + AdjY + OriginY
		
		If MyPos > (MapSizeY / 2) - 1 Then
			MyPos = MyPos - MapSizeY
		End If
		
		If MxPos > (MapSizeX / 2) - 1 Then
			MxPos = MxPos - MapSizeX
		End If
		
		'if we haven't changed hexes, then we're done
		If MxPos = oldx And MyPos = oldy Then
			Exit Sub
		End If
		
		'Output Label
		'sb1.Panels.Item("Hex").Text = "Sect: " + SectorString(MxPos, MyPos) + " "
		sb1.Items.Item("Hex").Text = SectorString(MxPos, MyPos) & " "
		
		'Top Third of Block
		'Get ship and plane info
		Dim ship As Short
		Dim land As Short
		Dim plane As Short
		Dim nuke As Short
		Dim enemy As Short
		Dim lname As String
		Dim strLand As String
		Dim sname As String
		Dim strShip As String
		Dim strEnemy As String
		' Dim strSect As String    removed efj 8/2003
		Dim strPlane As String
		Dim strNuke As String '100603 rjk: Added nuke information.
		Dim strOwner As String
		Dim strLastOwner As String
		Dim lastid As Short
		
		strEnemyid = vbNullString
		'Get Sector Information
		'rsSectors.Seek "=", MxPos, MyPos
		'If Not rsSectors.NoMatch Then
		'    sb1.Panels.Item("Hex").Text = sb1.Panels.Item("Hex").Text _
		''        + CStr(rsSectors.Fields("eff").Value) + "% "
		'    strSect = CStr(rsSectors.Fields("civ").Value) _
		''        + "c/" + CStr(rsSectors.Fields("mil").Value) _
		''        + "m/" + CStr(rsSectors.Fields("uw").Value) + "u"
		'End If
		
		ship = 0
		plane = 0
		land = 0
		nuke = 0
		enemy = 0
		strShipid = vbNullString
		strPlaneid = vbNullString
		strLandid = vbNullString
		strNuke = vbNullString '100603 rjk: Added nuke information.
		strNukeid = vbNullString '100603 rjk: Added nuke information.
		strEnemyid = vbNullString
		
		'Set Indexs
		rsLand.Index = "location"
		rsShip.Index = "location"
		rsPlane.Index = "location"
		rsNuke.Index = "location"
		rsEnemyUnit.Index = "location"
		
		'Load number of land units
		rsLand.Seek("=", MxPos, MyPos)
		If Not rsLand.NoMatch Then
			
land_loop: 
			If (Not rsLand.EOF) Then
				If rsLand.Fields("x").Value = MxPos And rsLand.Fields("y").Value = MyPos Then
					land = land + 1
					lname = rsLand.Fields("type").Value
					'rsBuild.Seek "=", "l", lname
					'If Not rsBuild.NoMatch Then
					'    If Len(rsBuild.Fields("desc")) > 0 Then
					'        lname = rsBuild.Fields("desc")
					'    End If
					'End If
					strLand = strLand & " " & lname & " #" & CStr(rsLand.Fields("id").Value) & ","
					If land = 1 Then
						strLandid = CStr(rsLand.Fields("id").Value)
					Else
						strLandid = strLandid & "/" & CStr(rsLand.Fields("id").Value)
					End If
					lastid = rsLand.Fields("id").Value
					rsLand.MoveNext()
					GoTo land_loop
				End If
			End If
			If land > 2 Then
				strLand = " " & CStr(land) & " units:" & strLand
			End If
			strLand = VB.Left(strLand, Len(strLand) - 1)
			If land = 1 Then
				strLand = strLand & " " & BuildMissionString("l", lastid)
			End If
		End If
		
		'Load number of ships
		rsShip.Seek("=", MxPos, MyPos)
		If Not rsShip.NoMatch Then
ship_loop: 
			If (Not rsShip.EOF) Then
				If rsShip.Fields("x").Value = MxPos And rsShip.Fields("y").Value = MyPos Then
					ship = ship + 1
					sname = rsShip.Fields("type").Value
					'rsBuild.Seek "=", "s", sname
					'If Not rsBuild.NoMatch Then
					'    If Len(rsBuild.Fields("desc")) > 0 Then
					'        sname = rsBuild.Fields("desc")
					'    End If
					'End If
					strShip = strShip & " " & sname & " #" & CStr(rsShip.Fields("id").Value) & ","
					If ship = 1 Then
						strShipid = CStr(rsShip.Fields("id").Value)
					Else
						strShipid = strShipid & "/" & CStr(rsShip.Fields("id").Value)
					End If
					lastid = rsShip.Fields("id").Value
					rsShip.MoveNext()
					GoTo ship_loop
				End If
			End If
			If ship > 2 Then
				strShip = " " & CStr(ship) & " ships:" & strShip
			End If
			strShip = VB.Left(strShip, Len(strShip) - 1)
			If ship = 1 Then
				strShip = strShip & " " & BuildMissionString("s", lastid)
			End If
		End If
		
		Dim cp1() As Short
		Dim cp2() As String
		Dim cpfound As Boolean
		ReDim cp1(1)
		ReDim cp2(1)
		
		'Load nubmer of planes
		rsPlane.Seek("=", MxPos, MyPos)
		If Not rsPlane.NoMatch Then
plane_loop: 
			If (Not rsPlane.EOF) Then
				If rsPlane.Fields("x").Value = MxPos And rsPlane.Fields("y").Value = MyPos Then
					plane = plane + 1
					lname = rsPlane.Fields("type").Value
					cpfound = False
					For X = LBound(cp1) To UBound(cp1)
						If cp2(X) = lname Then
							cp1(X) = cp1(X) + 1
							cpfound = True
							Exit For
						End If
					Next X
					If Not cpfound Then
						ReDim Preserve cp1(UBound(cp1) + 1)
						ReDim Preserve cp2(UBound(cp2) + 1)
						cp2(X) = lname
						cp1(X) = 1
					End If
					
					'rsBuild.Seek "=", "l", lname
					'If Not rsBuild.NoMatch Then
					'   If Len(rsBuild.Fields("desc")) > 0 Then
					'        lname = rsBuild.Fields("desc")
					'    End If
					'End If
					If plane = 1 Then
						strPlaneid = CStr(rsPlane.Fields("id").Value)
					Else
						strPlaneid = strPlaneid & "/" & CStr(rsPlane.Fields("id").Value)
					End If
					lastid = rsPlane.Fields("id").Value
					rsPlane.MoveNext()
					GoTo plane_loop
				End If
			End If
		End If
		' print out summation of planes
		For X = LBound(cp1) To UBound(cp1)
			If cp1(X) > 0 Then
				If cp1(X) > 1 Then
					strPlane = strPlane & " " & CStr(cp1(X))
				End If
				lname = cp2(X)
				rsBuild.Seek("=", "p", lname)
				If Not rsBuild.NoMatch Then
					If Len(rsBuild.Fields("desc").Value) > 0 Then
						lname = rsBuild.Fields("desc").Value
					End If
				End If
				strPlane = strPlane & " " & lname & ","
			End If
		Next X
		If Len(strPlane) > 1 Then
			strPlane = VB.Left(strPlane, Len(strPlane) - 1)
		End If
		If plane = 1 Then
			strPlane = strPlane & " " & BuildMissionString("p", lastid)
		End If
		
		'100603 rjk: Added nuke information.
		'Look for nukes
		rsNuke.Seek("=", MxPos, MyPos)
		If Not rsNuke.NoMatch Then
nuke_loop: 
			If (Not rsNuke.EOF) Then
				If rsNuke.Fields("x").Value = MxPos And rsNuke.Fields("y").Value = MyPos Then
					If Len(strNuke) = 0 Then
						strNuke = " Nukes:"
					End If
					strNuke = strNuke & " " + rsNuke.Fields("type").Value + "(" + CStr(rsNuke.Fields("num").Value) + "),"
					nuke = nuke + 1
					If nuke = 1 Then
						strNukeid = CStr(rsNuke.Fields("id").Value)
					Else
						strNukeid = strNukeid & "/" & CStr(rsNuke.Fields("id").Value)
					End If
					rsNuke.MoveNext()
					GoTo nuke_loop
				End If
			End If
			'remove last comma
			If Len(strNuke) > 0 Then
				strNuke = VB.Left(strNuke, Len(strNuke) - 1)
			End If
		End If
		
		'Look for enemy units
		strLastOwner = vbNullString
		rsEnemyUnit.Seek("=", MxPos, MyPos)
		If Not rsEnemyUnit.NoMatch Then
enemy_loop: 
			If (Not rsEnemyUnit.EOF) Then
				If rsEnemyUnit.Fields("x").Value = MxPos And rsEnemyUnit.Fields("y").Value = MyPos Then
					sname = rsEnemyUnit.Fields("type").Value
					rsBuild.Seek("=", rsEnemyUnit.Fields("class").Value, sname)
					If Not rsBuild.NoMatch Then
						If Len(rsBuild.Fields("desc").Value) > 0 Then
							sname = rsBuild.Fields("desc").Value
						End If
					End If
					
					'for enemy ships, build the enemy id string
					If rsEnemyUnit.Fields("class").Value = "s" Then
						enemy = enemy + 1
						If enemy = 1 Then
							strEnemyid = CStr(rsEnemyUnit.Fields("id").Value)
						Else
							strEnemyid = strEnemyid & "/" & CStr(rsEnemyUnit.Fields("id").Value)
						End If
					End If
					
					'get the owner
					strOwner = Nations.NationName((rsEnemyUnit.Fields("owner").Value))
					If strOwner = vbNullString Then
						strOwner = CStr(rsEnemyUnit.Fields("owner").Value)
					End If
					
					'use last owner string to prevent repeating the owner
					'name on multiple ships
					If strOwner = strLastOwner Then
						strOwner = vbNullString
					Else
						strLastOwner = strOwner
					End If
					
					strEnemy = strEnemy & " " & strOwner & " " & sname & " (#" & CStr(rsEnemyUnit.Fields("id").Value) & "),"
					rsEnemyUnit.MoveNext()
					GoTo enemy_loop
				End If
			End If
			'remove first space, last comma
			If Len(strEnemy) > 0 Then
				strEnemy = Trim(VB.Left(strEnemy, Len(strEnemy) - 1))
			End If
		End If
		'Combine strings
		'100603 rjk: Added nuke information.
		'If strShip = vbnullstring And strLand = vbnullstring And strEnemy = vbnullstring Then
		'    sb1.Panels.Item("Hex").Text = sb1.Panels.Item("Hex").Text + strSect
		'Else
		sb1.Items.Item("Hex").Text = sb1.Items.Item("Hex").Text & strEnemy & strShip & strLand & strNuke & strPlane
		'End If
	End Sub
	
	Private Sub lstCmdResult_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lstCmdResult.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'Trap Control C and Copy to Clipboard
		If KeyAscii = 3 Then
			CopyListBoxToClipboard(lstCmdResult)
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub ToggleAutoRead()
		If frmEmpCmd.bAutoRead Then
			frmEmpCmd.bAutoRead = False
		Else
			frmEmpCmd.bAutoRead = True
		End If
		UpdateAutoRead()
	End Sub
	
	Public Sub UpdateAutoRead()
		If frmEmpCmd.bAutoRead Then
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(4).ToolTipText = "AutoRead is On"
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(4).ImageIndex = 24
		Else
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(4).ToolTipText = "AutoRead is Off"
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(4).ImageIndex = 25
		End If
	End Sub
	
	Private Sub ToggleAutoUpdate()
		If frmEmpCmd.bAutoUpdate Then
			frmEmpCmd.bAutoUpdate = False
		Else
			frmEmpCmd.bAutoUpdate = True
			UpdateCommands()
			If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
				frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
			End If
		End If
		UpdateAutoUpdate()
	End Sub
	
	Public Sub UpdateAutoUpdate()
		If frmEmpCmd.bAutoUpdate Then
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(3).ToolTipText = "AutoUpdate is On"
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(3).ImageIndex = 2
		Else
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(3).ToolTipText = "AutoUpdate is Off"
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbMain.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbMain.Items.Item(3).ImageIndex = 23
		End If
	End Sub
	
	Public Sub UpdateBars()
		tbMain.Visible = frmOptions.bToolbar
		sb1.Visible = frmOptions.bStatusBar
		ResizePanels()
		ArrangeControls()
	End Sub
	
	'100903 rjk: Rename AcceptReport so as to not be confused with Accept from the Reject command
	Public Sub mnuDiploAcceptanceReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploAcceptanceReport.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptNation
		PromptUp = True
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control cbNations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbNations.Move(PromptForm.cbRelations.Left, PromptForm.cbRelations.top)
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbRelations.Visible = False
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1.Visible = False
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Accept"
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuDiploCollect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploCollect.Click
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptCollect
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuDiploConsider_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploConsider.Click
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptOffer
		PromptUp = True
		
		'100903 rjk: Moved the field logic to the form
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Consider"
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuDiploDeclare_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploDeclare.Click
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptNation
		PromptUp = True
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbRelations.Visible = True
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1.Visible = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Declare"
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		
	End Sub
	
	Public Sub mnuDiploOffer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploOffer.Click
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptOffer
		PromptUp = True
		
		'100903 rjk: Moved field logic to form
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	'100903 rjk: Split Reject into Reject and Accept separate menu items
	Public Sub mnuDiploReject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploReject.Click
		PromptForm = frmPromptOffer
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Reject"
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		'frmEmpCmd.SubmitEmpireCommand "reject", True
	End Sub
	
	Public Sub mnuDiploAccept_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploAccept.Click
		PromptForm = frmPromptOffer
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Accept"
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		'frmEmpCmd.SubmitEmpireCommand "reject", True
	End Sub
	
	Public Sub mnuDiploRelations_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploRelations.Click
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptNation
		PromptUp = True
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control cbNations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbNations.Move(PromptForm.cbRelations.Left, PromptForm.cbRelations.top)
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbRelations.Visible = False
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1.Visible = False
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Relations"
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuDiploRepay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploRepay.Click
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptCollect
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Repay"
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		
	End Sub
	
	Public Sub mnuDiploSharebmap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploSharebmap.Click
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptNation
		PromptUp = True
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.cbRelations.Visible = False
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1.Visible = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Sharebmap"
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Visible = True
		'UPGRADE_ISSUE: Control txtNum could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtNum.Visible = True
		'UPGRADE_ISSUE: Control Label3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label3.Visible = True
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.top = PromptForm.cbRelations.top
		'UPGRADE_ISSUE: Control txtNum could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtNum.top = PromptForm.cbRelations.top
		'UPGRADE_ISSUE: Control Label3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control cbRelations could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label3.top = PromptForm.cbRelations.top
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuDiploShark_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploShark.Click
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptCollect
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Shark"
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		
		
	End Sub
	
	Public Sub mnuDiploTreaty_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDiploTreaty.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("treaty *", True)
	End Sub
	
	Public Sub mnuLandCargo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandCargo.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Land Cargo"
		frmPromptLand.cmd = "lcargo"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
		
	End Sub
	
	Public Sub mnuLandFire_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandFire.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptFire
		
		'Show Prompt
		ShowUnitPrompt("land")
		
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2.Value = True 'need to set after loaded
		
	End Sub
	
	Public Sub mnuLandFuel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandFuel.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptFuel
		
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2.Value = True
		'PromptForm.strSector = SectorString(MxPos, MyPos) 091503 rjk: removed, strSector not used in frmPromptFuel
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandLoad_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandLoad.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		PromptUp = True
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "lload"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandLook_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandLook.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'If there is only one unit, don't bother prompting
		If InStr(strLandid, "/") = 0 And Len(strLandid) > 0 Then
			frmEmpCmd.SubmitEmpireCommand("llook " & strLandid, True)
			frmEmpCmd.SubmitEmpireCommand("db2", False) 'force map redraw
		Else
			PromptForm = frmPromptLook
			'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.strCmd = "llook"
			
			'Show Prompt
			ShowUnitPrompt("land")
		End If
		
	End Sub
	
	Public Sub mnulandMarch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnulandMarch.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptNav
		
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "March"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Show Prompt
		ShowUnitPrompt("land")
		
		'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If Len(PromptForm.txtUnit.Text) > 0 Then
			'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtPath.SetFocus()
		End If
	End Sub
	
	Public Sub mnuLandMission_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandMission.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptMission
		'Setup op sector hex
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin = "."
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2.Value = True
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandMorale_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandMorale.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptSail
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "morale"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Morale"
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1 = "retreat %"
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandRadar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandRadar.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "lradar"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Radar"
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandScrap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandScrap.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptScrap
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "scrap "
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Scrap"
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2 = True
		
		'Show Prompt
		ShowUnitPrompt("land")
	End Sub
	
	Public Sub mnuLandScuttle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandScuttle.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptScrap
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "scuttle"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Scuttle"
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2 = True
		
		'Show Prompt
		ShowUnitPrompt("land")
	End Sub
	
	Public Sub mnuLandStat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandStat.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Land Stats"
		frmPromptLand.cmd = "lstat"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
		
	End Sub
	
	Public Sub mnuLandSupply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandSupply.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "supply"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Supply"
		
		'Show Prompt
		ShowUnitPrompt("land")
		
	End Sub
	
	Public Sub mnuLandTend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandTend.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "ltend"
		'Show Prompt
		ShowUnitPrompt("ship") '100303 rjk: Switched to ship from land so initial field show ships.
	End Sub
	
	Public Sub mnuLandUnload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandUnload.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "lunload"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		'Show Prompt
		ShowUnitPrompt("land")
	End Sub
	
	Public Sub mnuLandUpgrade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLandUpgrade.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptScrap
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "upgrade "
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Upgrade"
		'UPGRADE_ISSUE: Control Option2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option2 = True
		
		'Show Prompt
		ShowUnitPrompt("land")
	End Sub
	
	Public Sub mnuMarketTrade_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMarketTrade.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("trade", True)
	End Sub
	
	Public Sub mnuPlaneBomb_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneBomb.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptBomb
		
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	'Private Sub mnuPlaneCargo_Click() 092303 rjk: empire command does not exist
	''check for current prompt
	'If PromptUp Then
	'    Exit Sub
	'End If
	'
	'frmPromptLand.Label2 = "Plane Cargo"
	'frmPromptLand.cmd = "pcargo"
	'mnuUnitLand_Click
	'End Sub
	
	Public Sub mnuPlaneDrop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneDrop.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptRecon
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Drop"
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneFly_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneFly.Click
		PromptForm = frmPromptRecon
		
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Fly"
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneMission_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneMission.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptMission
		
		'Setup op sector hex
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin = "."
		'UPGRADE_ISSUE: Control Option3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option3.Value = True
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlanePara_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlanePara.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptRecon
		
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Paradrop"
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneRange_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneRange.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptSail
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "range"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Range"
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1 = "radius"
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneRecon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneRecon.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptRecon
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Recon"
		
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneStats.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Plane Stats"
		frmPromptLand.cmd = "pstat"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
	End Sub
	
	Public Sub mnuPlaneSweep_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneSweep.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptRecon
		
		'Setup op sector hex
		'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'End If
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Sweep"
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuPlaneTrans_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlaneTrans.Click
		PromptForm = frmPromptTrans
		PromptUp = True
		
		'Setup op sector hex
		'PromptForm.strSector = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Show Prompt
		ShowUnitPrompt("plane")
		
	End Sub
	
	Public Sub mnuReportBudget_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportBudget.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("budget", True)
	End Sub
	
	Public Sub mnuReportCensus_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportCensus.Click
		'check for current prompt
		
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptCensus.txtOrigin = SectorString(MxPos, MyPos) 092103 rjk: Switch to SelX and SelY for top level menu
		frmPromptCensus.txtOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		frmPromptCensus.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptCensus.Width)) \ 2)
		frmPromptCensus.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptCensus.Height))
		PromptForm = frmPromptCensus
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptCensus.Show()
		
	End Sub
	
	Public Sub mnuReportCoastwatch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportCoastwatch.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("coastwatch *", True)
		frmEmpCmd.SubmitEmpireCommand("db2", False) 'redraw map
	End Sub
	
	'111903 rjk: Added Commodity Total Report
	Public Sub mnuReportCommodity_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportCommodity.Click
		Dim Index As Short = mnuReportCommodity.GetIndex(eventSender)
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		Select Case Index
			Case 1
				frmPromptCensus.Label2.Text = "Commodity"
				mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
			Case 2
				PromptForm = frmPromptLoad
				'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strCmd = "commoditytotal"
				'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.strSector = SectorString(SelX, SelY)
				PromptUp = True
				PromptForm.Show()
				'UPGRADE_ISSUE: Control txtMultOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				PromptForm.txtMultOrigin.SetFocus()
		End Select
	End Sub
	
	Public Sub mnuReportCountry_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportCountry.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("country *", True)
	End Sub
	
	Public Sub mnuReportCutoff_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportCutoff.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Cutoff"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
	End Sub
	
	Public Sub mnuReportFinancial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportFinancial.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("financial", True)
	End Sub
	
	Public Sub mnuReportLevel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportLevel.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Level"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
	End Sub
	
	Public Sub mnuReportMap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportMap.Click
		Dim Index As Short = mnuReportMap.GetIndex(eventSender)
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Map dialog and execute
		PromptForm = frmPromptMap
		PromptUp = True
		
		'prompt for parameters
		Dim strCmd As String
		strCmd = Trim(Me.mnuReportMap(Index).Text)
		'092303 rjk: Remove & from strCmd so submit works
		If InStr(strCmd, "&") <> 0 Then
			If InStr(strCmd, "&") = 1 Then
				strCmd = VB.Right(strCmd, Len(strCmd) - 1)
			Else
				strCmd = VB.Left(strCmd, InStr(strCmd, "&") - 1) & Mid(strCmd, InStr(strCmd, "&") + 1)
			End If
		End If
		frmPromptMap.Label2.Text = strCmd
		
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuReportMotd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportMotd.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("motd", True)
	End Sub
	
	Public Sub mnuReportNation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportNation.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("nation", True)
	End Sub
	
	Public Sub mnuReportNewsPaper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportNewsPaper.Click
		' Dim strCmd As String   removed efj 8/2003
		' Dim strConnect As String   removed efj 8/2003
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptNews
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'prompt for parameters
		frmPromptNews.Show()
		
	End Sub
	
	Public Sub mnuReportNuke_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportNuke.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Nukes"
		frmPromptLand.cmd = "nuke"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
	End Sub
	
	Public Sub mnuReportPlayers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportPlayers.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("players", True)
	End Sub
	
	Public Sub mnuReportPower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportPower.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("power", True)
	End Sub
	
	Public Sub mnuReportPowerNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportPowerNew.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("power new", True)
	End Sub
	
	Public Sub mnuReportProduction_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportProduction.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Production"
		frmPromptCensus.txtOrigin.Text = "*"
		'Put form in proper place
		frmPromptCensus.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptCensus.Width)) \ 2)
		frmPromptCensus.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptCensus.Height))
		PromptForm = frmPromptCensus
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptCensus.Show()
	End Sub
	
	'120203 rjk: Added to do multiple sector Start and Stops
	Public Sub mnuReportStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportStart.Click
		Dim Index As Short = mnuReportStart.GetIndex(eventSender)
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		If Index = 1 Then
			frmPromptCensus.Label2.Text = "Start"
		Else
			frmPromptCensus.Label2.Text = "Stop"
		End If
		frmPromptCensus.cmbType.SelectedIndex = 0
		
		frmPromptCensus.txtOrigin.Text = SectorString(SelX, SelY)
		'Put form in proper place
		frmPromptCensus.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptCensus.Width)) \ 2)
		frmPromptCensus.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptCensus.Height))
		PromptForm = frmPromptCensus
		PromptUp = True
		'Dialog box will take it from here
		frmPromptCensus.Show()
	End Sub
	
	Public Sub mnuReportRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportRead.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		SubmitTelegramRead(False, True)
	End Sub
	
	Public Sub mnuReportReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportReport.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("report *", True)
	End Sub
	
	Public Sub mnuReportResource_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportResource.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Resource"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
	End Sub
	
	Public Sub mnuReportSky_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportSky.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("skywatch *", True)
	End Sub
	
	Public Sub mnuReportsLedger_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportsLedger.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("ledger *", True)
	End Sub
	
	Public Sub mnuReportStarve_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportStarve.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmEmpCmd.SubmitEmpireCommand("bf1", False) 'run as batch file
		frmEmpCmd.SubmitEmpireCommand("starvation s *", True) 'ships
		frmEmpCmd.SubmitEmpireCommand("starvation l *", True) 'lands
		frmEmpCmd.SubmitEmpireCommand("starvation *", True) 'sectors
		frmEmpCmd.SubmitEmpireCommand("bf2", False) 'run as batch file
		frmEmpCmd.SubmitEmpireCommand("db2", False) 'force map redraw
	End Sub
	
	Public Sub mnuReportUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportUpdate.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand("show updates", True)
		Else
			frmEmpCmd.SubmitEmpireCommand("update", True)
		End If
	End Sub
	
	Public Sub mnuReportVersion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportVersion.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'If VersionCheck(4, 3, 10) <> VER_LESS Then
		'    frmEmpCmd.SubmitEmpireCommand "xdump ver", True
		'Else
		frmEmpCmd.SubmitEmpireCommand("version", True)
		'End If
	End Sub
	
	Public Sub mnuReportWire_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReportWire.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptWire
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuSectCmdAnti_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdAnti.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptAnti
		PromptUp = True
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switch to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuSectCmdBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdBoard.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup dialog
		'frmPromptBoard.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: switched to SelX and SelY for top level menu
		frmPromptBoard.txtOrigin.Text = SectorString(SelX, SelY)
		frmPromptBoard.txtOrigin.Visible = True
		frmPromptBoard.txtUnit.Visible = False
		frmPromptBoard.Frame3.Text = "Attacking Sector"
		frmPromptBoard.txtTarget.Text = strEnemyid
		PromptForm = frmPromptBoard
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		'UPGRADE_ISSUE: Control txtTarget could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtTarget.SetFocus() '092203 rjk: Added to set the focus to the enemy list automatically
	End Sub
	
	Public Sub mnuSectCmdBuild_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdBuild.Click
		mnuUnitBuild_Click(mnuUnitBuild, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdCap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdCap.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'CapSect = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		CapSect = SectorString(SelX, SelY)
		'frmEmpCmd.SubmitEmpireCommand "capital " + SectorString(MxPos, MyPos), True
		CapSect = SectorString(SelX, SelY)
		frmEmpCmd.SubmitEmpireCommand("capital " & SectorString(SelX, SelY), True)
		
	End Sub
	
	Public Sub mnuSectCmdConvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdConvert.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptConvert.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: switched to SelX and SelY for top level menu
		frmPromptConvert.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptConvert
		PromptUp = True
		
		'Put form in proper place
		frmPromptConvert.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptConvert.Width)) \ 2)
		frmPromptConvert.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptConvert.Height))
		PromptForm = frmPromptConvert
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptConvert.Show()
		
	End Sub
	
	Public Sub mnuSectCmdDel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdDel.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptMove
		PromptUp = True
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Deliver"
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
		'101903 rjk: Switched txtMultOrigin for multiple sector selection
		'UPGRADE_ISSUE: Control txtMultOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
		'100203 rjk: Moved the field logic to the form.
		
		'Dialog box will take it from here
		PromptForm.Show()
	End Sub
	
	Public Sub mnuSectCmdDemob_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdDemob.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptConvert.Label2.Text = "Demobilize"
		frmPromptConvert.Label1.Text = "military in"
		frmPromptConvert.chkReserve.Visible = True
		mnuSectCmdConvert_Click(mnuSectCmdConvert, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdDes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdDes.Click
		
		'Setup Designation dialog and execute
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'frmPromptDes.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		'101903 rjk: Switched form to Multiple Sector selection
		'frmPromptDes.txtOrigin.Text = SectorString(SelX, SelY)
		'101903 rjk: Switched txtMultOrigin for multiple sector selection
		frmPromptDes.txtMultOrigin.Text = SectorString(SelX, SelY)
		
		PromptForm = frmPromptDes
		PromptUp = True
		
		'Put form in proper place
		frmPromptDes.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptDes.Width)) \ 2)
		frmPromptDes.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptDes.Height))
		
		
		'Dialog box will take it from here
		frmPromptDes.Show()
		
	End Sub
	
	Public Sub mnuSectCmdDist_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdDist.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptDist.txtOrigin = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		frmPromptDist.txtOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		frmPromptDist.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptDist.Width)) \ 2)
		frmPromptDist.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptDist.Height))
		PromptForm = frmPromptDist
		PromptUp = True
		
		'Dialog box will take it from here
		PromptForm.Show()
		'UPGRADE_ISSUE: Control txtMultDest could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultDest.SetFocus() '101803 rjk: Switched txtMultDest for multiple sectors selection
		
	End Sub
	
	Public Sub mnuSectCmdEnlist_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdEnlist.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptConvert.Label2.Text = "Enlist"
		mnuSectCmdConvert_Click(mnuSectCmdConvert, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdExplore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdExplore.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptExplore.txtOrigin = SectorString(MxPos, MyPos) 0920303 rjk: Switched to SelX and SelY for top level menu
		frmPromptExplore.txtOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		frmPromptExplore.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptExplore.Width)) \ 2)
		frmPromptExplore.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptExplore.Height))
		PromptForm = frmPromptExplore
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptExplore.Show()
		
	End Sub
	
	Public Sub mnuSectCmdFire_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdFire.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptFire.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched for top level menu
		frmPromptFire.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptFire
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
		'UPGRADE_ISSUE: Control Option3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option3.Value = True
		
		On Error Resume Next
		'UPGRADE_ISSUE: Control txtTarget could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtTarget.SetFocus()
	End Sub
	
	Public Sub mnuSectCmdGrind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdGrind.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptConvert.Label2.Text = "Grind"
		frmPromptConvert.Label1.Text = "gold bars in"
		mnuSectCmdConvert_Click(mnuSectCmdConvert, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdImprove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdImprove.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptImprove.txtOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu access
		frmPromptImprove.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptImprove
		PromptUp = True
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuSectCmdMove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdMove.Click
		
		Dim n As Short
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptMove
		PromptUp = True
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
		'101903 rjk: Switched txtMultOrigin for multiple sector selection
		'UPGRADE_ISSUE: Control txtMultOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
		'100203 rjk: Moved the field logic to the form.
		
		'Dialog box will take it from here
		PromptForm.Show()
		'UPGRADE_ISSUE: Control txtNum could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtNum.SetFocus()
		
	End Sub
	
	Public Sub mnuSectCmdNeweff_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdNeweff.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Neweff"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdOrigin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdOrigin.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptOrigin.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Switched to SelX and SelY for top level menu
		frmPromptOrigin.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptOrigin
		PromptUp = True
		
		'Put form in proper place
		frmPromptOrigin.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptOrigin.Width)) \ 2)
		frmPromptOrigin.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptOrigin.Height))
		
		'Dialog box will take it from here
		frmPromptOrigin.Show()
		
	End Sub
	
	Public Sub mnuSectCmdProduction_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdProduction.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Production"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
		frmPromptCensus.cmdOK_Click(Nothing, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdRadar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdRadar.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptRadar
		PromptUp = True
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Switched to SelX and SelY for toplevel menu
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2.Caption = "Radar"
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		frmPromptRadar.Show()
		
	End Sub
	
	Public Sub mnuSectCmdShoot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdShoot.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptShoot.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Changed to SelX and SelY for top level menu
		frmPromptShoot.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptShoot
		PromptUp = True
		
		'Put form in proper place
		frmPromptShoot.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptShoot.Width)) \ 2)
		frmPromptShoot.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptShoot.Height))
		
		'default = number of civs in sector
		'092203 rjk: Changed to SelX and SelY for top level menu
		'rsSectors.Seek "=", MxPos, MyPos
		rsSectors.Seek("=", SelX, SelY)
		If Not rsSectors.NoMatch Then
			frmPromptShoot.txtNum.Text = CStr(rsSectors.Fields("civ").Value)
		End If
		'Dialog box will take it from here
		frmPromptShoot.Show()
		
	End Sub
	
	Public Sub mnuSectCmdSpy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdSpy.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		' 092303 rjk: Switched to SelX and SelY for top level menu
		'frmEmpCmd.SubmitEmpireCommand "spy " + SectorString(MxPos, MyPos), True
		frmEmpCmd.SubmitEmpireCommand("spy " & SectorString(SelX, SelY), True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'101603 rjk: Does not need a dump as none of our sectors change
		'GetSectors
		'101703 rjk: Added Strength fields to Sector database
		'frmEmpCmd.SubmitEmpireCommand "strength * ?timestamp>" + CStr(tsSectors), False
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
	End Sub
	
	Public Sub mnuSectCmdStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdStart.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Start"
		frmPromptCensus.cmbType.SelectedIndex = 0
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
		frmPromptCensus.cmdOK_Click(Nothing, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdStop.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Stop"
		frmPromptCensus.cmbType.SelectedIndex = 0
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
		frmPromptCensus.cmdOK_Click(Nothing, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdStrength_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdStrength.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptCensus.Label2.Text = "Strength"
		mnuReportCensus_Click(mnuReportCensus, New System.EventArgs())
	End Sub
	
	Public Sub mnuSectCmdThr_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdThr.Click
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptThresh
		PromptUp = True
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		'101903 rjk: Switched txtMultOrigin for multiple sector selection
		'UPGRADE_ISSUE: Control txtMultOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuSectCmdThrAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdThrAll.Click
		' Dim n As Integer 092103 rjk: functionality moved to form
		' Dim thresh As Integer   removed efj 8/2003
		
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptThreshAll
		PromptUp = True
		'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control txtMultOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
		
		PromptUp = True 'rjk 5/12/03: moved this line until we know it is valid sector and we will display the dialog box
		
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuSectCmdWipe_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSectCmdWipe.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'frmPromptWipe.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
		frmPromptWipe.txtOrigin.Text = SectorString(SelX, SelY)
		PromptForm = frmPromptWipe
		PromptUp = True
		
		'Put form in proper place
		frmPromptWipe.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptWipe.Width)) \ 2)
		frmPromptWipe.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptWipe.Height))
		
		'Dialog box will take it from here
		frmPromptWipe.Show()
	End Sub
	
	Public Sub mnuShipAssult_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipAssult.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup dialog
		frmPromptAttack.Label2.Text = "Assault"
		PromptForm = frmPromptAttack
		
		'Show Prompt
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		PromptUp = True
		
		'Fill the combo box with ships in the hex
		On Error GoTo PromptExit
		Dim slist As Object
		Dim n As Short
		If Len(strShipid) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object slist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			slist = ParseUnitString(strShipid)
			If IsArray(slist) Then
				For n = LBound(slist) To UBound(slist)
					If Len(slist(n)) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object slist(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						frmPromptAttack.cbShips.Items.Add(slist(n))
					End If
				Next n
				If frmPromptAttack.cbShips.Items.Count > 0 Then
					frmPromptAttack.cbShips.SelectedIndex = 0
				End If
			End If
		End If
		If (InStr(strShipid, "/") = 0 And Len(strShipid) > 0) Then
			'UPGRADE_ISSUE: Control cbShips could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.cbShips.Add(strShipid)
		End If
		
PromptExit: 
		'Dialog box will take it from here
		PromptForm.Show()
		
	End Sub
	
	Public Sub mnuShipBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipBoard.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup dialog
		frmPromptBoard.txtOrigin.Visible = False
		frmPromptBoard.txtUnit.Visible = True
		frmPromptBoard.Frame3.Text = "Attacking Ship"
		frmPromptBoard.txtTarget.Text = strEnemyid
		PromptForm = frmPromptBoard
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipCargo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipCargo.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Ship Cargo"
		frmPromptLand.cmd = "cargo"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
		
	End Sub
	
	
	Public Sub mnuShipFire_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipFire.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		PromptForm = frmPromptFire
		PromptUp = True
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
		'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option1.Value = True ' 091803 rjk ensure form is load, because there is SetFocus in Click event
		'PromptForm.txtUnit.Text = strShipid
		
		
		'if a firing ship has been filled in, switch to target
		'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If Len(PromptForm.txtUnit.Text) > 0 Then
			'UPGRADE_ISSUE: Control txtTarget could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtTarget.SetFocus()
		End If
	End Sub
	
	Public Sub mnuShipLoad_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipLoad.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "load"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		'Show Prompt
		ShowUnitPrompt("ship")
	End Sub
	
	Public Sub mnuShipLook_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipLook.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		If InStr(strShipid, "/") = 0 And Len(strShipid) > 0 Then
			frmEmpCmd.SubmitEmpireCommand("look " & strShipid, True)
			frmEmpCmd.SubmitEmpireCommand("db2", False) 'force map redraw
		Else
			PromptForm = frmPromptLook
			PromptUp = True
			'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.strCmd = "look"
			ShowUnitPrompt("ship")
		End If
		
	End Sub
	
	Public Sub mnuShipMine_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipMine.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "mine"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Mine"
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipMission_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipMission.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptMission
		'Setup op sector hex
		'UPGRADE_ISSUE: Control txtOrigin could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtOrigin = "."
		'UPGRADE_ISSUE: Control Option1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Option1.Value = True
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipMquota_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipMquota.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptSail
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "mquota"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Mquota"
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1 = "level"
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipNav_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipNav.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		Me.PromptForm = frmPromptNav
		Me.PromptUp = True
		'Setup op sector hex
		'092303 rjk: Switched to SelX and SelY for top level menu access
		'If PopUpMenuSource = DMAP_PUMS_MAP Then
		'    PromptForm.strSector = SectorString(MxPos, MyPos)
		'Else
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		'End If
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
		'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If Len(PromptForm.txtUnit.Text) > 0 Then
			'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtPath.SetFocus()
		End If
	End Sub
	
	Public Sub mnuShipNavMark_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipNavMark.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		PromptUp = True
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "marker"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Marker"
		
		'for single ships, just show the markers and put out the prompt
		If InStr(strShipid, "/") = 0 And Len(strShipid) > 0 Then
			'Put form in proper place
			PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
			PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
			'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtUnit = strShipid
			
			'Dialog box will take it from here
			PromptForm.Show()
			NavMarkerShip = strShipid
			DrawHexes()
		Else
			ShowUnitPrompt("ship")
		End If
		
	End Sub
	
	Public Sub mnuShipOrder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipOrder.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptOrder
		PromptUp = True
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipRetreat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipRetreat.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptSail
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "retreat"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Retreat"
		'UPGRADE_ISSUE: Control Label1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label1 = "path"
		'UPGRADE_ISSUE: Control Label3 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label3.Visible = True
		'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Combo1.Visible = True
		'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		LoadRetreatBox(PromptForm.Combo1)
		'UPGRADE_ISSUE: Control Combo1 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Combo1.ListIndex = 0
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipSail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipSail.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptSail
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "sail"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Sail"
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipSonar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipSonar.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "sonar"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Sonar"
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipStat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipStat.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Ship Stats"
		frmPromptLand.cmd = "sstat"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
	End Sub
	
	Public Sub mnuShipTorp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipTorp.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup dialog and execute
		PromptForm = frmPromptTorp
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuShipUnload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipUnload.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLoad
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "unload"
		'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
		'UPGRADE_ISSUE: Control strSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strSector = SectorString(SelX, SelY)
		
		'Show Prompt
		ShowUnitPrompt("ship")
	End Sub
	
	Public Sub mnuShipUnsail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuShipUnsail.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		PromptForm = frmPromptLook
		
		'UPGRADE_ISSUE: Control strCmd could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.strCmd = "unsail"
		'UPGRADE_ISSUE: Control Label2 could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.Label2 = "Unsail"
		
		'Show Prompt
		ShowUnitPrompt("ship")
		
	End Sub
	
	Public Sub mnuUnitBuild_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitBuild.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		'092203 rjk: Switched to SelX, SelY for top level menu
		'frmPromptBuild.txtOrigin = SectorString(MxPos, MyPos)
		'rsSectors.Seek "=", MxPos, MyPos
		frmPromptBuild.txtOrigin.Text = SectorString(SelX, SelY)
		rsSectors.Seek("=", SelX, SelY)
		
		'Put form in proper place
		frmPromptBuild.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptBuild.Width)) \ 2)
		frmPromptBuild.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptBuild.Height))
		PromptForm = frmPromptBuild
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptBuild.Show()
		
	End Sub
	
	Public Sub mnuUnitLand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitLand.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		'Setup Designation dialog and execute
		frmPromptLand.txtOrigin.Text = vbNullString
		
		'Put form in proper place
		frmPromptLand.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(frmPromptLand.Width)) \ 2)
		frmPromptLand.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(frmPromptLand.Height))
		PromptForm = frmPromptLand
		PromptUp = True
		
		'Dialog box will take it from here
		frmPromptLand.Show()
		
	End Sub
	
	Public Sub mnuUnitPlanes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitPlanes.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Plane"
		frmPromptLand.cmd = "plane"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
	End Sub
	
	Public Sub mnuUnitShip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnitShip.Click
		'check for current prompt
		If PromptUp Then
			Exit Sub
		End If
		
		frmPromptLand.Label2.Text = "Ship"
		frmPromptLand.cmd = "ship"
		mnuUnitLand_Click(mnuUnitLand, New System.EventArgs())
	End Sub
	
	
	Private Sub Msg_Timer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Msg_Timer.Tick
		
		Dim Msg As String
		Dim Done As Boolean
		Dim n As Short
		Dim i As Short
		Dim bwidth As Short
		Static lstlen As Short
		
		'If you have a pending message, then add it to the display box
		If MsgQ.Count() = 0 Then
			'UPGRADE_WARNING: Timer property Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
			Me.Msg_Timer.Interval = 0
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MsgQ(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Msg = MsgQ.Item(1)
		
		'add msg to the message box
		'break into multiple lines if too long
		n = 1
		Done = False
		bwidth = VB6.PixelsToTwipsX(rtbTelegram.Width) * 0.9
		
		While Not Done
			'UPGRADE_ISSUE: Form method frmDrawMap.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If Me.TextWidth(Msg) > bwidth Then
				i = 1
				Do 
					n = i
					i = InStr(n + 1, Msg, " ")
					'UPGRADE_ISSUE: Form method frmDrawMap.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Loop Until (i = 0 Or Me.TextWidth(VB.Left(Msg, i)) > bwidth)
			End If
			
			rtbTelegram.SelectionStart = 999999
			rtbTelegram.SelectionLength = 0
			If n = 1 Then
				If Not (lstlen = 0 And Len(Msg) = 0) Then
					rtbTelegram.SelectedText = vbCrLf & Msg
					lstlen = Len(Msg)
				End If
				Done = True
			Else
				rtbTelegram.SelectedText = vbCrLf & VB.Left(Msg, n)
				Msg = Mid(Msg, n + 1)
				lstlen = 99
				n = 1
			End If
		End While
		
		MsgQ.Remove(1)
		
	End Sub
	
	
	Private Sub Picture1_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture1.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	Private Sub Picture2_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture2.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	Private Sub Picture3_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture3.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	Private Sub Picture4_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture4.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub PlayersTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PlayersTimer.Tick
		Static playersMinutes As Short
		If playersMinutes <= 0 Then
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			frmEmpCmd.SubmitEmpireCommand("players", False)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
			playersMinutes = frmOptions.playersTimeInterval - 1
		Else
			playersMinutes = playersMinutes - 1
		End If
	End Sub
	
	Private Sub rtbTelegram_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtbTelegram.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lStartPos As Integer
		Dim lBegPos As Integer
		Dim lEndPos As Integer
		Dim strAddress As String
		Dim strUnit As String
		
		If Button = VB6.MouseButtonConstants.RightButton Then
			lStartPos = rtbTelegram.SelectionStart
			'UPGRADE_WARNING: UpTo was upgraded to SelectionStart and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			rtbTelegram.SelectionStart = rtbTelegram.Find(" (:@" & vbLf, System.Windows.Forms.RichTextBoxFinds.Reverse)
			lBegPos = rtbTelegram.SelectionStart + 1
			'UPGRADE_WARNING: UpTo was upgraded to SelectionStart and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			rtbTelegram.SelectionStart = rtbTelegram.Find(" ):!" & vbLf, System.Windows.Forms.RichTextBoxFinds.None)
			lEndPos = rtbTelegram.SelectionStart + 1
			rtbTelegram.SelectionStart = lStartPos
			strAddress = Mid(rtbTelegram.Text, lBegPos, lEndPos - lBegPos)
			If InStr(strAddress, ",") > 0 Then
				If ParseSectors(iTelegramSelX, iTelegramSelY, strAddress) Then
					mnuTelegramCenter.Visible = True
				Else
					mnuTelegramCenter.Visible = False
				End If
			Else
				mnuTelegramCenter.Visible = False
			End If
			'UPGRADE_ISSUE: Form method frmDrawMap.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			PopupMenu(mnuTelegram)
		End If
	End Sub
	
	Private Sub rtbTelegram_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtbTelegram.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub sb1_PanelClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _sb1_Panel1.Click
		Dim Panel As System.Windows.Forms.ToolStripStatusLabel = CType(eventSender, System.Windows.Forms.ToolStripStatusLabel)
		'UPGRADE_ISSUE: MSComctlLib.Panel property Panel.key was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Select Case Panel.key
			Case "Timer"
				GetUpdate(True)
				If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
					frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
				End If
				
			Case "Mail", "Mail Count"
				SubmitTelegramRead(True, False)
				sb1.Items.Item("Mail").Text = "..."
				sb1.Items.Item("Mail Count").Text = "..."
			Case "Anno"
				frmEmpCmd.SubmitEmpireCommand("wire y", True)
				'Case "Repeat" 102503 rjk: No Repeat Panel exists
				'    txtEntry.Text = frmEmpCmd.GetPreviousCommand
		End Select
	End Sub
	
	Private Sub TabStrip1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TabStrip1.ClickEvent
		'rjk 05/13/03: Moved the tab selection to DisplaySectorPanel
		'              Moved the frame code to SelectTabContent
		'    If TabStrip1.SelectedItem.Index = mintCurFrame _
		''        Then Exit Sub ' No need to change frame.
		'
		'   Select Case SelSectType
		'        Case SEL_SECT_OWN
		'            ' Otherwise, hide old frame, show new.
		'            Frame1(TabStrip1.SelectedItem.Index).Visible = True
		'            Frame1(mintCurFrame).Visible = False
		'            mintCurFrame = TabStrip1.SelectedItem.Index
		'        Case SEL_SECT_ENEMY
		'            If TabStrip1.SelectedItem.Index = 3 Then
		'                Frame1(3).Visible = True
		'                Frame1(mintCurFrame).Visible = False
		'                mintCurFrame = 3
		'            Else
		'                Frame1(mintCurFrame).Visible = False
		'                Frame1(0).Visible = True
		'                mintCurFrame = 0
		'            End If
		'        Case SEL_SECT_UNKNOWN
		'            If TabStrip1.SelectedItem.Index = 3 Then
		'                Frame1(3).Visible = True
		'                Frame1(mintCurFrame).Visible = False
		'                mintCurFrame = 3
		'            Else
		'                Frame1(mintCurFrame).Visible = False
		'                mintCurFrame = 0
		'            End If
		'    End Select
		SelectTabContent()
	End Sub
	
	Private Sub frmDrawMap_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'If F2 is pressed, repeat last command
		If KeyCode = System.Windows.Forms.Keys.F2 Then
			frmEmpCmd.SubmitLastCommand()
			Exit Sub
		End If
		
		'If F5 is pressed, Zoom or UnZoom
		If KeyCode = System.Windows.Forms.Keys.F5 Then
			If Shift > 0 Then
				'mnuMapUnzoom_Click
			Else
				'munMapZoom_Click
			End If
			frmDrawMap_Resize(Me, New System.EventArgs())
			Exit Sub
		End If
		
		'If F9 is pressed, move back thru command list
		'102503 rjk: Added shift F9 moves forward thru the command list
		If KeyCode = System.Windows.Forms.Keys.F9 Then
			If Not ((Shift And VB6.ShiftConstants.ShiftMask) = VB6.ShiftConstants.ShiftMask) Then
				txtEntry.Text = frmEmpCmd.GetPreviousCommand
			Else
				txtEntry.Text = frmEmpCmd.GetNextCommand
			End If
			txtEntry.SelectionStart = 999
			CommandCursorPos = 999
			SetCommandFocus()
			Exit Sub
		End If
		
		Dim n As Short
		If KeyCode = System.Windows.Forms.Keys.PageUp Then
			n = lstCmdResult.TopIndex - CShort(VB6.PixelsToTwipsY(lstCmdResult.Height) / VB6.PixelsToTwipsY(Me.txtEntry.Height))
			If n > 0 Then
				lstCmdResult.TopIndex = n
			Else
				lstCmdResult.TopIndex = 0
			End If
		End If
		If KeyCode = System.Windows.Forms.Keys.PageDown Then
			n = lstCmdResult.TopIndex + CShort(VB6.PixelsToTwipsY(lstCmdResult.Height) / VB6.PixelsToTwipsY(Me.txtEntry.Height))
			If n < lstCmdResult.Items.Count Then
				lstCmdResult.TopIndex = n
			ElseIf lstCmdResult.Items.Count > 0 Then 
				lstCmdResult.TopIndex = lstCmdResult.Items.Count - 1
			End If
		End If
		
		'check and see if the control key is down
		If Not ((Shift And VB6.ShiftConstants.CtrlMask) = 2) Then
			Exit Sub
		End If
		On Error Resume Next
		
		If KeyCode = System.Windows.Forms.Keys.D1 Then
			frmEmpCmd.WindowState = System.Windows.Forms.FormWindowState.Normal
			frmEmpCmd.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D2 Then
			frmTelegram.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D3 Then
			frmTelegram.WindowState = System.Windows.Forms.FormWindowState.Normal
			frmTelegram.Activate()
		End If
		
		'Shift Fonts
		If KeyCode = System.Windows.Forms.Keys.D4 Then
			On Error GoTo Font_exit
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmCommonDialog)
			' Set CancelError is True
			'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
			frmCommonDialog.cdb.CancelError = True
			' Set flags
			'UPGRADE_ISSUE: Constant cdlCFScreenFonts was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: MSComDlg.CommonDialog property cdb.Flags was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			frmCommonDialog.cdb.Flags = MSComDlg.FontsConstants.cdlCFScreenFonts
			frmCommonDialog.cdbFont.Font = VB6.FontChangeName(frmCommonDialog.cdbFont.Font, lstCmdResult.Font.Name)
			frmCommonDialog.cdbFont.Font = VB6.FontChangeSize(frmCommonDialog.cdbFont.Font, lstCmdResult.Font.SizeInPoints)
			
			' Display the font dialog box
			frmCommonDialog.cdbFont.ShowDialog()
			lstCmdResult.Font = VB6.FontChangeName(lstCmdResult.Font, frmCommonDialog.cdbFont.Font.Name)
			lstCmdResult.Font = VB6.FontChangeSize(lstCmdResult.Font, frmCommonDialog.cdbFont.Font.Size)
			frmDrawMap_Resize(Me, New System.EventArgs())
		End If
Font_exit: 
		
		
	End Sub
	
	
	Private Sub frmDrawMap_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		' Dim n As Integer   removed efj 8/2003
		' When Enter is pressed, process everything
		If KeyAscii = 13 Then
			If frmEmpCmd.ServerQuery Then
				frmEmpCmd.ExecuteEmpireCmd((txtEntry.Text))
				VB6.SetItemString(lstCmdResult, lstCmdResult.Items.Count - 1, VB6.GetItemString(lstCmdResult, lstCmdResult.Items.Count - 1) & txtEntry.Text)
			Else
				'now submit the command
				frmEmpCmd.SubmitFromCommandLine((txtEntry.Text))
			End If
			txtEntry.Text = vbNullString
			KeyAscii = 0
		End If
		
		'if an escape key is hit when a query is up
		If KeyAscii = 27 And frmEmpCmd.ServerQuery Then
			'frmEmpCmd.ExecuteEmpireCmd "ctld"
			frmEmpCmd.ExecuteEmpireCmd("aborted")
			VB6.SetItemString(lstCmdResult, lstCmdResult.Items.Count - 1, VB6.GetItemString(lstCmdResult, lstCmdResult.Items.Count - 1) & " aborted")
		End If
		
		'Any normal key - route to command box
		If KeyAscii >= 8 Or (KeyAscii >= 32 And KeyAscii <= 125) Then
			SetCommandFocus(KeyAscii)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub TabStrip1_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComctlLib.ITabStripEvents_MouseMoveEvent) Handles TabStrip1.MouseMoveEvent
		Me.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub tbMain_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _tbMain_Button1.Click, _tbMain_Button2.Click, _tbMain_Button3.Click, _tbMain_Button4.Click, _tbMain_Button5.Click, _tbMain_Button6.Click, _tbMain_Button7.Click, _tbMain_Button8.Click, _tbMain_Button9.Click, _tbMain_Button10.Click, _tbMain_Button11.Click, _tbMain_Button12.Click, _tbMain_Button13.Click, _tbMain_Button14.Click, _tbMain_Button15.Click, _tbMain_Button16.Click, _tbMain_Button17.Click, _tbMain_Button18.Click, _tbMain_Button19.Click, _tbMain_Button20.Click, _tbMain_Button21.Click, _tbMain_Button22.Click, _tbMain_Button23.Click, _tbMain_Button24.Click, _tbMain_Button25.Click, _tbMain_Button26.Click, _tbMain_Button27.Click, _tbMain_Button28.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		
		Select Case Button.Name
			Case "SetColors"
				mnuViewOptions_Click(mnuViewOptions.Item(1), New System.EventArgs()) 'Colors off the menu screen
			Case "AutoUpdate"
				ToggleAutoUpdate() '120203 rjk: Switch to frmOptions and boolean options.
			Case "AutoRead"
				ToggleAutoRead() '120203 rjk: Switch to frmOptions and boolean options.
			Case "UpdateSector"
				mnuRefresh_Click(mnuRefresh.Item(1), New System.EventArgs())
			Case "UpdateUnits"
				mnuRefresh_Click(mnuRefresh.Item(3), New System.EventArgs())
			Case "Budget"
				mnuReportBudget_Click(mnuReportBudget, New System.EventArgs())
			Case "Coastwatch"
				mnuReportCoastwatch_Click(mnuReportCoastwatch, New System.EventArgs())
			Case "Nation"
				mnuReportNation_Click(mnuReportNation, New System.EventArgs())
			Case "Newspaper"
				frmEmpCmd.SubmitEmpireCommand("newspaper 1")
			Case "PowerNew"
				mnuReportPowerNew_Click(mnuReportPowerNew, New System.EventArgs())
			Case "Report"
				mnuReportReport_Click(mnuReportReport, New System.EventArgs())
			Case "Starvation"
				mnuReportStarve_Click(mnuReportStarve, New System.EventArgs())
			Case "Version"
				mnuReportVersion_Click(mnuReportVersion, New System.EventArgs())
			Case "Wire"
				frmEmpCmd.SubmitEmpireCommand("wire 1")
			Case "Declare"
				mnuDiploDeclare_Click(mnuDiploDeclare, New System.EventArgs())
			Case "Relations"
				mnuDiploRelations_Click(mnuDiploRelations, New System.EventArgs())
			Case "Script"
				mnuToolsOption_Click(mnuToolsOption.Item(1), New System.EventArgs())
			Case "Build"
				mnuToolsOption_Click(mnuToolsOption.Item(2), New System.EventArgs())
			Case "Famine"
				mnuToolsOption_Click(mnuToolsOption.Item(3), New System.EventArgs())
			Case "NationLevels"
				mnuToolsOption_Click(mnuToolsOption.Item(4), New System.EventArgs())
			Case "Chi"
				mnuToolsOption_Click(mnuToolsOption.Item(5), New System.EventArgs())
			Case "ReportShow"
				mnuReportShow_Click(mnuReportShow, New System.EventArgs())
			Case "UpdateBmap"
				If frmOptions.bSPAtlantis Then
					frmEmpCmd.ForceUpdates = True
					If frmOptions.bSuppressBmapRefresh Then
						MsgBox("Bmap updates are suppressed. Change option and then request again")
						Exit Sub
					End If
					'Send Code to start database update
					frmEmpCmd.SubmitEmpireCommand("db1", False)
					'Request an update of map information
					frmEmpCmd.SubmitEmpireCommand("map *", False)
					frmEmpCmd.SubmitEmpireCommand("bmap *", False)
					'send Code to end database update
					frmEmpCmd.SubmitEmpireCommand("db2", False)
				Else
					mnuRefresh_Click(mnuRefresh.Item(2), New System.EventArgs())
				End If
			Case Else
				If InStr(Button.Name, "Script") = 1 Then
					ExecuteFile(CShort(Mid(Button.Name, 7)))
				End If
		End Select
	End Sub
	
	Private Sub txtEntry_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEntry.Enter
		On Error Resume Next
		CommandCursorFocus = True
		txtEntry.SelectionStart = CommandCursorPos
	End Sub
	
	Private Sub txtEntry_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEntry.Leave
		On Error Resume Next
		CommandCursorPos = txtEntry.SelectionStart
		CommandCursorFocus = False
	End Sub
	
	Private Sub txtEntry_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txtEntry.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Me.Cursor = System.Windows.Forms.Cursors.IBeam
	End Sub
	
	Private Sub UpdateTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UpdateTimer.Tick
		
		Dim totsec As Integer
		Dim totmin As Integer
		Dim sec As Integer
		Dim Min As Integer
		'UPGRADE_NOTE: hour was upgraded to hour_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim hour_Renamed As Integer
		Dim strMinus As String
		
		If TimerAtUpdate = 0 And SecondsToUpdate = 0 Then
			Me.sb1.Items.Item("Timer").Text = "No Update"
		Else
			totsec = VB.Timer() - TimerAtUpdate
			While totsec < 0
				totsec = totsec + 86400
			End While
			totsec = SecondsToUpdate - totsec
			If totsec <= -10 Then
				If totsec Mod 10 = 0 Then
					If frmEmpCmd.bAutoUpdate And frmEmpCmd.bUpdateEnabled Then
						If VersionCheck(4, 3, 10) = WinAceRoutines.enumVersion.VER_LESS Then
							UpdateCommands()
						Else
							frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
						End If
					Else
						GetUpdate(False)
						'Ensure it is still disabled
						If VersionCheck(4, 3, 10) <> WinAceRoutines.enumVersion.VER_LESS Then
							frmEmpCmd.SubmitEmpireCommand("xdump game *", False)
						End If
					End If
				End If
			End If
			sec = System.Math.Abs(totsec Mod 60)
			totmin = (totsec - sec) \ 60
			Min = System.Math.Abs(totmin Mod 60)
			hour_Renamed = System.Math.Abs(totmin \ 60)
			If hour_Renamed > 99 Then
				hour_Renamed = 99
			End If
			If totsec < 0 Then
				strMinus = "-"
			Else
				strMinus = ""
			End If
			Me.sb1.Items.Item("Timer").Text = strMinus & VB6.Format(hour_Renamed, "00") & ":" & VB6.Format(Min, "00") & ":" & VB6.Format(sec, "00")
		End If
		
		'if its time to run a pre-update script, run it now
		If BatchScript1Time > 0 And totsec <= BatchScript1Time Then
			frmScript.ExectuteBatchScript(BatchScript1File)
			BatchScript1Time = 0
			BatchScript1File = vbNullString
		End If
		
		'if you are idle for 10 minutes or more, submit a command to keep the line open
		TimeIdle = TimeIdle + 1
		If TimeIdle >= 600 Then
			GetUpdate(False)
			TimeIdle = 0
		End If
		
		'flash the lights if server query is up
		Static ToggleLights As Boolean
		If frmEmpCmd.ServerQuery Then
			'UPGRADE_WARNING: Lower bound of collection sb1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sb1.Items.Item(3).ToolTipText = "Empire Server Waiting for Response"
			If ToggleLights Then
				Me.sb1.Items.Item("Lights").Image = VB6.ImageToIPictureDisp(picBothLights)
				frmEmpCmd.imgLights.Image = picBothLights
			Else
				Me.sb1.Items.Item("Lights").Image = VB6.ImageToIPictureDisp(picGreenLight)
				frmEmpCmd.imgLights.Image = picGreenLight
			End If
			ToggleLights = Not ToggleLights
		End If
		
		'If a prompt help box needs positioned - do it
		If PositionHelp Then
			If PromptUp Then
				If frmReport.Visible Then
					If frmReport.WindowState = System.Windows.Forms.FormWindowState.Normal Then
						
						'Size Report
						frmReport.Top = Me.Top
						If VB6.PixelsToTwipsY(frmReport.Height) > (VB6.PixelsToTwipsY(PromptForm.Top) - VB6.PixelsToTwipsY(Me.Top)) Then
							frmReport.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(PromptForm.Top) - VB6.PixelsToTwipsY(Me.Top))
						End If
						
						'move report and shut off timer
						PositionHelp = False
					End If
				End If
			Else
				PositionHelp = False
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: vsMap.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vsMap.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vsMap_Change(ByVal newScrollValue As Integer)
		
		Static ChangeSkip As Boolean
		
		'Prevents second pass if we have to manually reset below due to
		'world wrap
		If ChangeSkip Then
			ChangeSkip = False
			Exit Sub
		End If
		
		'Reset Origin
		OriginY = Int(newScrollValue / 2) * 2
		If OriginY < MapSizeY / -2 Then
			OriginY = OriginY + MapSizeY
		End If
		If OriginY > (MapSizeY / 2) - 1 Then
			OriginY = OriginY - MapSizeY
		End If
		
		'Update index value if necessary
		If newScrollValue <> OriginY Then
			ChangeSkip = True
			newScrollValue = OriginY
		End If
		
		'Don't do a draw if the map's not valid
		If Not MapValid Then
			Exit Sub
		End If
		
		DrawHexes()
		
	End Sub
	'UPGRADE_NOTE: hsMap.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar event hsMap.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub hsMap_Change(ByVal newScrollValue As Integer)
		
		Static ChangeSkip As Boolean
		
		'Prevents second pass if we have to manually reset below due to
		'world wrap
		If ChangeSkip Then
			ChangeSkip = False
			Exit Sub
		End If
		
		'Reset Origin
		OriginX = Int(newScrollValue / 2) * 2
		If OriginX < MapSizeX / -2 Then
			OriginX = OriginX + MapSizeX
		End If
		If OriginX > (MapSizeX / 2) - 1 Then
			OriginX = OriginX - MapSizeX
		End If
		
		'Update index value if necessary
		If newScrollValue <> OriginX Then
			ChangeSkip = True
			newScrollValue = OriginX
		End If
		
		'Don't do a draw if the map's not valid
		If Not MapValid Then
			Exit Sub
		End If
		
		'Redraw hex Centers
		DrawHexes()
		
	End Sub
	
	
	Public Sub DrawHexes()
		Dim NewLargeChange As Short
		' Dim Y As Integer   removed efj 8/2003
		Dim X As Short
		' Dim nbrcol As Integer   removed efj 8/2003
		' Dim nbrrow As Integer   removed efj 8/2003
		Dim oldcolor As Object
		Dim oldfontsize As Object
		Dim oldWidth As Object
		Dim PosX As Single
		Dim PosY As Single
		Dim n As Short
		Dim strPath As String
		Dim startindex As Short
		
		'Don't do a draw if the map's not valid
		If Not MapValid Then
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: PictureBox property picMap.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.AutoRedraw = True
		'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.Cls()
		
		DrawGrid(picMap)
		
		If ViewShipWake Then
			DrawShipWake(-1)
		ElseIf LastShip >= 0 Then 
			DrawShipWake(LastShip)
		End If
		
		'handle event markers
		EventMarkers.RemoveExpired()
		If EventMarkers.Count > 0 Then
			DrawEvent(picMap)
		End If
		
		'If you were drawing path hexes - redraw them
		If DrawingPath Or MarkingTerritory Or DisplayingPath Then
			'Draw Hex
			'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oldWidth = picMap.DrawWidth
			'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			picMap.DrawWidth = 3 ' Set DrawWidth
			'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oldcolor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
			picMap.ForeColor = System.Drawing.Color.Red
			
			startindex = LBound(WorkingPath, 2)
			If MarkingTerritory Then
				startindex = startindex + 1
			End If
			
			For X = startindex To UBound(WorkingPath, 2)
				Coord(WorkingPath(1, X), WorkingPath(2, X), PosX, PosY)
				picMap.ForeColor = System.Drawing.Color.Red
				DrawHex(picMap, PosX, PosY)
				
				'draw center lines if appropriate
				If X > LBound(WorkingPath, 2) And DrawingPath Then
					picMap.ForeColor = System.Drawing.Color.White
					'Compute Direction
					strPath = EmpirePathDirection(WorkingPath(1, X - 1) - WorkingPath(1, X), WorkingPath(2, X - 1) - WorkingPath(2, X))
					n = InStr(HEX_DIR, strPath)
					DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, (InStr(HEX_DIR, strPath)))
					n = n - 3
					If n < 1 Then
						n = n + 6
					End If
					Coord(WorkingPath(1, X - 1), WorkingPath(2, X - 1), PosX, PosY)
					DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, n)
				End If
			Next X
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(oldcolor)
			'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			picMap.DrawWidth = oldWidth ' Set DrawWidth
		End If
		
		'if nav markers are up, redraw them
		If Len(NavMarkerShip) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oldfontsize = picMap.Font.SizeInPoints
			NavMarkers(NavMarkerShip, bNavMarkerShipIncludeUpdateMob, bNavMarkerShipIncludeUpdateEff, iNavMarkerShipUpdates, bNavMarkerShipLimitMobility)
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			picMap.Font = VB6.FontChangeSize(picMap.Font, oldfontsize)
		End If
		
		If Len(NukeDamageType) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oldfontsize = picMap.Font.SizeInPoints
			ShowNukeDamage(picMap)
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			picMap.Font = VB6.FontChangeSize(picMap.Font, oldfontsize)
		ElseIf Len(PlaneRangeType) > 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oldfontsize = picMap.Font.SizeInPoints
			ShowPlaneRange(picMap)
			'UPGRADE_WARNING: Couldn't resolve default property of object oldfontsize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			picMap.Font = VB6.FontChangeSize(picMap.Font, oldfontsize)
		End If
		
		'hsMap.Max = (MapSizeX / 2) - 1 - ((NbrHexWide - 1) * 2)
		'vsMap.Max = (MapSizeY / 2) - 1 - (NbrHexTall - 1)
		hsMap.Maximum = ((MapSizeX / 2) + hsMap.LargeChange - 1)
		vsMap.Maximum = ((MapSizeY / 2) + vsMap.LargeChange - 1)
		NewLargeChange = (NbrHexWide - 1)
		hsMap.Maximum = hsMap.Maximum + NewLargeChange - hsMap.LargeChange
		hsMap.LargeChange = NewLargeChange
		NewLargeChange = (NbrHexTall - 1) / 2
		vsMap.Maximum = vsMap.Maximum + NewLargeChange - vsMap.LargeChange
		vsMap.LargeChange = NewLargeChange
		
	End Sub
	
	Public Sub StartPath(ByRef strStartSect As String)
		Dim x1 As Short
		Dim y1 As Short
		If Not ParseSectors(x1, y1, strStartSect) Then
			Exit Sub
		End If
		
		ReDim WorkingPath(2, 0)
		WorkingPath(1, 0) = x1
		WorkingPath(2, 0) = y1
		DrawingPath = True
		
		'Draw Starting Hex
		
		Dim oldcolor As Object
		Dim oldWidth As Object
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldWidth = picMap.DrawWidth
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.DrawWidth = 3 ' Set DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		picMap.ForeColor = System.Drawing.Color.Red
		
		Dim PosX As Single
		Dim PosY As Single
		
		'Get Center position and draw hex
		Coord(x1, y1, PosX, PosY)
		DrawHex(picMap, PosX, PosY)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(oldcolor)
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = oldWidth ' Set DrawWidth
	End Sub
	
	'111803 rjk: Created one copy of the DisplayPath code
	Public Sub DisplayPath(ByRef strStart As String, ByRef strPath As String)
		Dim i As Short
		Dim ch As String
		
		StartPath(strStart)
		For i = 1 To Len(strPath)
			ch = Mid(strPath, i, 1)
			ExtendPath(ch)
		Next i
		DrawingPath = False
		DisplayingPath = True
	End Sub
	
	Public Sub AddtoPath()
		
		Dim lasthex As Short
		Dim DeltaX As Short
		Dim DeltaY As Short
		' Dim Index As Integer   removed efj 8/2003
		Dim strPath As String
		Dim n As Short
		
		On Error GoTo 0
		
		lasthex = UBound(WorkingPath, 2)
		
		'Make sure its an adjacent hex
		DeltaX = MxPos - WorkingPath(1, lasthex)
		DeltaY = MyPos - WorkingPath(2, lasthex)
		'111303 rjk: Added World wraparound checks
		If DeltaX > rsVersion.Fields("World X").Value / 2 Then
			DeltaX = DeltaX - rsVersion.Fields("World X").Value
		ElseIf DeltaX < -rsVersion.Fields("World X").Value / 2 Then 
			DeltaX = DeltaX + rsVersion.Fields("World X").Value
		End If
		If DeltaY > rsVersion.Fields("World Y").Value / 2 Then
			DeltaY = DeltaY - rsVersion.Fields("World Y").Value
		ElseIf DeltaY < -rsVersion.Fields("World Y").Value / 2 Then 
			DeltaY = DeltaY + rsVersion.Fields("World Y").Value
		End If
		If Not ((System.Math.Abs(DeltaX) = 2 And DeltaY = 0) Or (System.Math.Abs(DeltaX) = 1 And System.Math.Abs(DeltaY) = 1)) Then
			Exit Sub
		End If
		
		'Compute Direction
		strPath = EmpirePathDirection(DeltaX, DeltaY)
		
		'Add path to prompt text box
		If PromptForm.Name = "frmPromptPath" Then
			'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtPath.Text = PromptForm.txtPath.Text + strPath
			'UPGRADE_ISSUE: Control lblEndSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.lblEndSector.Caption = SectorString(MxPos, MyPos)
		End If
		
		ReDim Preserve WorkingPath(2, lasthex + 1)
		WorkingPath(1, lasthex + 1) = MxPos
		WorkingPath(2, lasthex + 1) = MyPos
		
		'Draw Hex
		Dim oldcolor As Object
		Dim oldWidth As Object
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldWidth = picMap.DrawWidth
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.DrawWidth = 3 ' Set DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		picMap.ForeColor = System.Drawing.Color.Red
		
		Dim PosX As Single
		Dim PosY As Single
		
		'Get Center position and draw hex
		Coord(MxPos, MyPos, PosX, PosY)
		DrawHex(picMap, PosX, PosY)
		
		picMap.ForeColor = System.Drawing.Color.White
		
		n = InStr(HEX_DIR, strPath)
		n = n - 3
		If n < 1 Then
			n = n + 6
		End If
		
		DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, n)
		
		Coord(WorkingPath(1, lasthex), WorkingPath(2, lasthex), PosX, PosY)
		DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, (InStr(HEX_DIR, strPath)))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(oldcolor)
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = oldWidth ' Set DrawWidth
		
	End Sub
	
	Public Function FinishPath() As String
		DrawingPath = False
		DrawHexes()
	End Function
	
	Public Sub ExtendPath(ByVal strDir As String)
		
		Dim lasthex As Short
		Dim x1 As Short
		Dim y1 As Short
		Dim X2 As Short
		Dim Y2 As Short
		
		'if we are drawing a path
		If Not DrawingPath Then
			Exit Sub
		End If
		
		'get the last hex
		lasthex = UBound(WorkingPath, 2)
		x1 = WorkingPath(1, lasthex)
		y1 = WorkingPath(2, lasthex)
		
		'Comput the new hex
		ComputeDestination(x1, y1, X2, Y2, strDir)
		
		'Set the new hex
		MxPos = X2
		MyPos = Y2
		
		'Now add to path
		AddtoPath()
		
	End Sub
	
	Public Sub RemoveFromPath()
		Dim lasthex As Short
		' Dim DeltaX As Integer   removed efj 8/2003
		' Dim DeltaY As Integer   removed efj 8/2003
		' Dim Index As Integer   removed efj 8/2003
		' Dim strPath As String   removed efj 8/2003
		On Error GoTo 0
		
		'Get the last hex
		lasthex = UBound(WorkingPath, 2)
		If lasthex <= LBound(WorkingPath, 2) Then
			Exit Sub
		End If
		
		Dim PosX As Single
		Dim PosY As Single
		
		'Get Center position and draw hex
		Coord(WorkingPath(1, lasthex), WorkingPath(2, lasthex), PosX, PosY)
		DrawSector(picMap, PosX, PosY, WorkingPath(1, lasthex), WorkingPath(2, lasthex))
		
		'Remove path from prompt text box
		'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If Len(PromptForm.txtPath.Text) > 0 Then
			'UPGRADE_ISSUE: Control txtPath could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			PromptForm.txtPath.Text = VB.Left(PromptForm.txtPath.Text, Len(PromptForm.txtPath.Text) - 1)
		End If
		'UPGRADE_ISSUE: Control lblEndSector could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.lblEndSector.Caption = CStr(WorkingPath(1, lasthex)) & "," & CStr(WorkingPath(2, lasthex))
		
		'remove hex from working path
		ReDim Preserve WorkingPath(2, lasthex - 1)
		
	End Sub
	
	Public Sub StartTerritory()
		MarkingTerritory = True
		ReDim WorkingPath(2, 0)
		DrawHexes() '100103 rjk: Clear Markers
	End Sub
	
	Public Sub ToggleTerritory(ByRef sx As Short, ByRef sy As Short)
		Dim lasthex As Short
		Dim found As Boolean
		Dim n As Short
		
		On Error Resume Next
		
		'if not drawing territory - 'quit
		If Not MarkingTerritory Then
			Exit Sub
		End If
		
		'Get the last hex
		lasthex = UBound(WorkingPath, 2)
		
		'search for the x,y coords
		found = False
		For n = LBound(WorkingPath, 2) + 1 To lasthex
			If WorkingPath(1, n) = sx And WorkingPath(2, n) = sy Then
				found = True
				WorkingPath(1, n) = WorkingPath(1, lasthex)
				WorkingPath(2, n) = WorkingPath(2, lasthex)
			End If
		Next n
		
		Dim PosX As Single
		Dim PosY As Single
		Dim oldcolor As Object
		Dim oldWidth As Object
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldWidth = picMap.DrawWidth
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.DrawWidth = 3 ' Set DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		
		If Not found Then
			'add hex
			ReDim Preserve WorkingPath(2, lasthex + 1)
			WorkingPath(1, lasthex + 1) = sx
			WorkingPath(2, lasthex + 1) = sy
			picMap.ForeColor = System.Drawing.Color.Red
		Else
			'remove hex from working path
			ReDim Preserve WorkingPath(2, lasthex - 1)
			picMap.ForeColor = System.Drawing.Color.Black
		End If
		
		'Get Center position and draw hex
		Coord(sx, sy, PosX, PosY)
		DrawHex(picMap, PosX, PosY)
		
		'Redraw all the hexs in case we messed up a border
		picMap.ForeColor = System.Drawing.Color.Red
		For n = LBound(WorkingPath, 2) + 1 To UBound(WorkingPath, 2)
			Coord(WorkingPath(1, n), WorkingPath(2, n), PosX, PosY)
			DrawHex(picMap, PosX, PosY)
		Next n
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(oldcolor)
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = oldWidth ' Set DrawWidth
		
	End Sub
	
	'Public Sub ShowCommandHistory()   removed efj 8/2003
	'
	''Shows the previous commands
	'Load frmRepeatCmd
	'frmRepeatCmd.Move Me.Left + Me.Width - frmRepeatCmd.Width, _
	''                        Me.top + Me.Height - frmRepeatCmd.Height
	'frmRepeatCmd.Show vbModal
	'
	'End Sub
	
	Public Sub ShowUnitPrompt(Optional ByRef UnitType As Object = Nothing)
		
		Dim FillBox As Boolean
		On Error Resume Next
		
		'Put form in proper place
		PromptForm.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + (VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(PromptForm.Width)) \ 2)
		PromptForm.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(PromptForm.Height))
		PromptUp = True
		
		'if there is only one unit, put it in the text box.  Otherwise,
		'show the unit dialog box
		' Dim oneunit As Boolean   removed efj 8/2003
		'UPGRADE_WARNING: Couldn't resolve default property of object UnitType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (UnitType = "ship" And InStr(strShipid, "/") = 0 And Len(strShipid) > 0) Or (UnitType = "plane" And InStr(strPlaneid, "/") = 0 And Len(strPlaneid) > 0) Or (UnitType = "land" And InStr(strLandid, "/") = 0 And Len(strLandid) > 0) Or (UnitType = "nuke" And InStr(strNukeid, "/") = 0 And Len(strNukeid) > 0) Or GridSelection = True Then
			Select Case UnitType
				Case "ship"
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strShipid
				Case "land"
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strLandid
				Case "plane"
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strPlaneid
				Case "nuke"
					'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					PromptForm.txtUnit.Text = strNukeid
			End Select
		Else
			'Bring up Unit dialog box
			DisplayUnitBox()
			
			'If specific unit requested then make sure that unit is showing in the unit box
			FillBox = False
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(UnitType) Then
				Select Case UnitType
					Case "ship"
						If DisplaySelect <> enumUnitDisplay.udSHIP Then
							DisplaySelect = enumUnitDisplay.udSHIP
							FillBox = True
						End If
						
					Case "land"
						If DisplaySelect <> enumUnitDisplay.udLAND Then
							DisplaySelect = enumUnitDisplay.udLAND
							FillBox = True
						End If
						
					Case "plane"
						If DisplaySelect <> enumUnitDisplay.udPLANE Then
							DisplaySelect = enumUnitDisplay.udPLANE
							FillBox = True
						End If
					Case "nuke"
						If DisplaySelect <> enumUnitDisplay.udNUKE Then
							DisplaySelect = enumUnitDisplay.udNUKE
							FillBox = True
						End If
				End Select
				
				'refill the box if necessary
				If FillBox Then
					DisplaySubset = enumSubset.ssSECT
					' 092303 rjk: Removed secx and secy and switched to SelX and SelY
					'secx = MxPos
					'secy = MyPos
					'FillSubSetFlag = True 092503 rjk: Eliminate
					FillGrid()
				End If
			End If
		End If
		
		
		'Dialog box will take it from here
		PromptForm.Show()
		On Error Resume Next
		'UPGRADE_ISSUE: Control txtUnit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		PromptForm.txtUnit.SetFocus()
	End Sub
	
	Private Sub NavMarkers(ByVal Shipnum As String, ByRef bIncludeUpdateMob As Boolean, ByRef bIncludeUpdateEff As Boolean, ByRef iUpdates As Short, ByRef bLimitMobility As Boolean)
		Const MaxRadius As Short = 20
		Dim mcost(MaxRadius * 2 + 1, MaxRadius * 2 + 1) As Single
		Dim path(MaxRadius * 2 + 1, MaxRadius * 2 + 1) As String
		Dim colCheckList As New Collection
		Dim StartX As Short
		Dim StartY As Short
		Dim iDistance As Short
		Dim iDirection As Short
		Dim iSpoke As Short
		Dim iAdjacent As Short
		Dim iAdjX As Short
		Dim iAdjY As Short
		Dim sx As Short
		Dim sy As Short
		Dim X As Short
		Dim Y As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim PosX As Single
		Dim PosY As Single
		Dim movecost As Single
		Dim hldcost As Single
		Dim hldIndex As String
		Dim shipid As Short
		Dim strSearchSector As String
		Dim yas, ya, xa, xas, pa As Object
		Dim iEfficiency As Short
		
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xa = New Object(){-2, -1, 1, 2, 1, -1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ya. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ya = New Object(){0, -1, -1, 0, 1, 1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xas = New Object(){1, 2, 1, -1, -2, -1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object yas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		yas = New Object(){-1, 0, 1, 1, 0, -1}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object pa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pa = New Object(){"g", "y", "u", "j", "n", "b"}
		
		'Save ship string to aid in redraw
		NavMarkerShip = Shipnum
		
		hldIndex = rsShip.Index
		rsShip.Index = "PrimaryKey"
		
		movecost = 0
		While Len(Shipnum) > 0
			X = InStr(Shipnum, "/")
			If X > 0 Then
				shipid = CShort(VB.Left(Shipnum, X - 1))
				Shipnum = Mid(Shipnum, X + 1)
			Else
				shipid = CShort(Shipnum)
				Shipnum = vbNullString
			End If
			
			rsShip.Seek("=", shipid)
			If rsShip.NoMatch Then
				rsShip.Index = hldIndex
				Exit Sub
			End If
			iEfficiency = rsShip.Fields("eff").Value
			If iEfficiency <> 100 And bIncludeUpdateEff Then
				iEfficiency = iEfficiency + iUpdates * rsVersion.Fields("Eff gain - Ship").Value
				If iEfficiency > 100 Then
					iEfficiency = 100
				End If
			End If
			hldcost = ShipMoveCost(rsShip.Fields("spd").Value, iEfficiency, rsShip.Fields("tech").Value, rsShip.Fields("type").Value)
			If hldcost > movecost Then
				movecost = hldcost
				mcost(MaxRadius, MaxRadius) = rsShip.Fields("mob").Value
				If bIncludeUpdateMob Then
					mcost(MaxRadius, MaxRadius) = mcost(MaxRadius, MaxRadius) + iUpdates * rsVersion.Fields("Mob gain - Ship").Value
					If mcost(MaxRadius, MaxRadius) > rsVersion.Fields("Max mob - Ship").Value And bLimitMobility Then
						mcost(MaxRadius, MaxRadius) = rsVersion.Fields("Max mob - Ship").Value
					End If
				End If
				StartX = rsShip.Fields("x").Value
				StartY = rsShip.Fields("y").Value
			End If
		End While
		
		colCheckList.Add(SectorString(0, 0))
		path(MaxRadius, MaxRadius) = ""
		
		While colCheckList.Count() > 0
			'UPGRADE_WARNING: Couldn't resolve default property of object colCheckList.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSearchSector = colCheckList.Item(1)
			ParseSectors(X, Y, strSearchSector)
			If Len(path(Int(X / 2) + MaxRadius, Y + MaxRadius)) < MaxRadius Then
				For iAdjacent = 0 To 5
					'UPGRADE_WARNING: Couldn't resolve default property of object xa(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iAdjX = X + xa(iAdjacent)
					'UPGRADE_WARNING: Couldn't resolve default property of object ya(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iAdjY = Y + ya(iAdjacent)
					If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pa(iAdjacent). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						path(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = path(Int(X / 2) + MaxRadius, Y + MaxRadius) + pa(iAdjacent)
						mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, EmpCommon.enumMobType.MT_SHIP) * movecost)
						If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) >= 0 Then
							colCheckList.Add(SectorString(iAdjX, iAdjY))
						End If
					Else
						If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) < mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, EmpCommon.enumMobType.MT_SHIP) * movecost) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object pa(iAdjacent). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							path(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = path(Int(X / 2) + MaxRadius, Y + MaxRadius) + pa(iAdjacent)
							mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, EmpCommon.enumMobType.MT_SHIP) * movecost)
							If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) >= 0 Then
								colCheckList.Add(SectorString(iAdjX, iAdjY))
							End If
						End If
					End If
				Next iAdjacent
			End If
			colCheckList.Remove(1)
		End While
		
		'Now put out the nav markers
		Dim HldColor As Integer
		Dim HldFontSize As Object
		HldColor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		picMap.ForeColor = System.Drawing.Color.White
		'UPGRADE_WARNING: Couldn't resolve default property of object HldFontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HldFontSize = picMap.Font.SizeInPoints
		picMap.Font = VB6.FontChangeSize(picMap.Font, picMap.Font.SizeInPoints * 0.7)
		
		For iDistance = 1 To MaxRadius - 1
			For iDirection = 0 To 5
				For iSpoke = 0 To MaxRadius - 2
					'UPGRADE_WARNING: Couldn't resolve default property of object xas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object xa(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
					'UPGRADE_WARNING: Couldn't resolve default property of object yas(iDirection). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ya(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
					Coord(StartX + X, StartY + Y, PosX, PosY)
					If mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) > 0 Then
						str_Renamed = CStr(CShort(mcost(Int(X / 2) + MaxRadius, Y + MaxRadius)))
						'UPGRADE_ISSUE: PictureBox method picMap.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						sx = PosX - (picMap.TextHeight(str_Renamed) / 2)
						'UPGRADE_ISSUE: PictureBox method picMap.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						sy = PosY - (picMap.TextWidth(str_Renamed) / 2)
						If sx >= 0 And sy >= 0 Then
							'UPGRADE_ISSUE: PictureBox property picMap.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							picMap.CurrentX = sx
							'UPGRADE_ISSUE: PictureBox property picMap.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							picMap.CurrentY = sy
							'UPGRADE_ISSUE: PictureBox method picMap.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							picMap.Print(str_Renamed)
						End If
					End If
				Next iSpoke
			Next iDirection
		Next iDistance
		
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(HldColor)
		'UPGRADE_WARNING: Couldn't resolve default property of object HldFontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.Font = VB6.FontChangeSize(picMap.Font, HldFontSize)
		rsShip.Index = hldIndex
	End Sub
	
	Private Sub DrawShipWake(ByRef Shipnum As Short)
		'ship number less than zero means draw all
		Dim found As Boolean
		Dim shippath As Object
		found = False
		With rsEnemyUnit
			If .BOF And .EOF Then
				Exit Sub
			End If
			.MoveFirst()
			While Not .EOF And Not found
				If .Fields("class").Value = "s" Then
					If .Fields("id").Value = Shipnum Then
						found = True
					End If
					
					If found Or Shipnum < 0 Then
						
						'can't draw a wake if the ship does not have one.
						If Len(.Fields("wake").Value) > 0 Then
							
							'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object shippath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							shippath = ParseWaypoints(.Fields("wake").Value)
							If Not IsArray(shippath) Then
								Exit Sub
							End If
							ReDim Preserve shippath(UBound(shippath) + 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object shippath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							shippath(UBound(shippath)) = SectorString(.Fields("x").Value, .Fields("y").Value)
							DrawWake(.Fields("id").Value, shippath)
						End If
					End If
				End If
				
				.MoveNext()
			End While
			
		End With
	End Sub
	
	Private Sub DrawWake(ByVal Shipnum As Short, ByRef shippath As Object)
		
		'find the enemy ship unit
		On Error GoTo DrawWake_Error
		
		
		Dim PosX As Single
		Dim PosY As Single
		Dim n As Short
		Dim X As Short
		Dim i As Short
		Dim strPath As String
		
		Dim sx As Short
		Dim sy As Short
		Dim sx2 As Short
		Dim sy2 As Short
		Dim dx As Short
		Dim dy As Short
		
		Dim strTemp As String
		
		'Draw Hex
		Dim oldcolor As Object
		Dim oldWidth As Object
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldWidth = picMap.DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = System.Drawing.ColorTranslator.ToOle(picMap.ForeColor)
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMap.DrawWidth = 3 'Set DrawWidth
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(BasicText(2))
		
		For X = LBound(shippath) To UBound(shippath)
			'UPGRADE_WARNING: Couldn't resolve default property of object shippath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTemp = CStr(shippath(X))
			If ParseSectors(sx, sy, strTemp) Then
				'draw center lines if appropriate
				If X > LBound(shippath) Then
					'Compute Direction
					'UPGRADE_WARNING: Couldn't resolve default property of object shippath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strTemp = CStr(shippath(X - 1))
					If ParseSectors(sx2, sy2, strTemp) Then
						dx = sx - sx2
						dy = sy - sy2
						'check for world wrap
						If dx > MapSizeX / 2 Then dx = dx - MapSizeX
						If dx < -1 * MapSizeX / 2 Then dx = dx + MapSizeX
						If dy > MapSizeY / 2 Then dy = dy - MapSizeY
						If dy < -1 * MapSizeY / 2 Then dy = dy + MapSizeY '112703 rjk: Fixed
						strPath = EmpirePathDirection(dx, dy)
						
						For i = 1 To Len(strPath)
							ComputeDestination(sx2, sy2, sx, sy, Mid(strPath, i, 1))
							n = InStr(HEX_DIR, Mid(strPath, i, 1))
							Coord(sx2, sy2, PosX, PosY)
							DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, n)
							n = n - 3
							If n < 1 Then
								n = n + 6
							End If
							If Not (X = UBound(shippath) And i = Len(strPath)) Then
								Coord(sx, sy, PosX, PosY)
								DrawHexCenterLine(picMap, PosX, PosY, GetHexSideLength, n)
							End If
							sx2 = sx : sy2 = sy
						Next i
					End If
				End If
			End If
		Next X
		
DrawWake_Error: 
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.ForeColor = System.Drawing.ColorTranslator.FromOle(oldcolor)
		'UPGRADE_ISSUE: PictureBox property picMap.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = oldWidth ' Set DrawWidth
		
	End Sub
	
	
	' Calculate index for use with TextArray property.
	Function gridIndex(ByRef row As Short, ByRef col As Short) As Integer
		gridIndex = row * grid1.Cols + col
	End Function
	
	Public Sub FillGrid()
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim fieldcnt As Short
		Dim X As Short
		Dim iIndex As Short
		Dim substype As String
		Dim substypekey As String
		Dim BuildType As String '101403 rjk: Moved to be private variable of FillGrid
		Dim hldIndex As String
		
		'UPGRADE_NOTE: PlaneRange was upgraded to PlaneRange_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim PlaneRange_Renamed As Short
		Static GridUpdateinProgress As Boolean
		
		'Airport vars
		Dim Airports() As String
		Dim i As Integer
		Dim strLast As String
		Dim strLoc As String
		Dim airport As Integer
		Dim found As Boolean
		Dim iCalculatedFields As Short
		
		ReDim Airports(0)
		'Prevent recursion
		If GridUpdateinProgress Then
			Exit Sub
		End If
		GridUpdateinProgress = True
		
		On Error Resume Next
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		PlaneRange_Renamed = 0
		
		FillTypeBox()
		
		'Determain which type of unit we are displaying
		Select Case DisplaySelect
			Case enumUnitDisplay.udSHIP
				rs = rsShipList
				rs.Requery()
				substype = "Fleet "
				substypekey = "flt"
				BuildType = "s"
				If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
					ComputeUnitCountsForShips()
				End If
			Case enumUnitDisplay.udLAND
				rs = rsLandList
				rs.Requery()
				substype = "Army "
				substypekey = "army"
				BuildType = "l"
				If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
					ComputeUnitCountsForLandUnits()
				End If
			Case enumUnitDisplay.udPLANE
				rs = rsPlaneList
				rs.Requery()
				substype = "Wing "
				substypekey = "wing"
				BuildType = "p"
			Case enumUnitDisplay.udNUKE
				rs = rsNuke
				substype = "Type "
				substypekey = "type"
				BuildType = "n"
			Case enumUnitDisplay.udENEMY
				rs = rsEnemyUnit
				'rs.Requery 100603 rjk: Removed, does not work this table.  Need to investigate more later.
				substype = "Nation "
				substypekey = "owner"
				
			Case enumUnitDisplay.udList
				FillUnitList()
				GridUpdateinProgress = False
				Exit Sub
			Case Else
				GridUpdateinProgress = False
				Exit Sub
		End Select
		If bDeity And DisplaySelect <> enumUnitDisplay.udENEMY Then
			substype = "Nation "
			substypekey = "country"
		End If
		
		'rs.Index = "Primary Key" 100603 rjk: Removed, does not work these tables.  Need to investigate more later.
		
		'092503 rjk: Rearrange the code to eliminate FillSubSetFlag
		ClearSubCombo()
		'if this is a "range" display, make that the only entry in sub box
		If DisplaySubset = enumSubset.ssPLANE_RANGE Or DisplaySubset = enumSubset.ssMISSILE_RANGE Or DisplaySubset = enumSubset.ssATTACK_RANGE Then
			cmbSub.Items.Add("In Range")
			cmbSub.SelectedIndex = 0
		Else
			cmbSub.Items.Add("All")
		End If
		
		fieldcnt = rs.Fields.Count
		rtbUnitList.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.Cols = fieldcnt
		grid1.FixedCols = 2
		
		'The following is to trap the size of the column width for column zero
		'so that if manually changed it is not always reset.
		Static ColZeroWidth1 As Single
		Static ColZeroWidth2 As Single
		Static ColZeroWidth3 As Single
		Static LastDisplayFlag As enumUnitDisplay
		Static bOldShortNameUnitGrid As Boolean
		
		If LastDisplayFlag <> enumUnitDisplay.udUNKNOWN Then
			If Not (frmOptions.bShortNameUnitGrid Xor bOldShortNameUnitGrid) Then
				If frmOptions.bShortNameUnitGrid Then
					ColZeroWidth3 = grid1.get_ColWidth(0)
				ElseIf LastDisplayFlag = enumUnitDisplay.udPLANE Or LastDisplayFlag = enumUnitDisplay.udENEMY Then 
					ColZeroWidth1 = grid1.get_ColWidth(0)
				Else
					ColZeroWidth2 = grid1.get_ColWidth(0)
				End If
			Else
				bOldShortNameUnitGrid = frmOptions.bShortNameUnitGrid
			End If
		Else
			grid1.set_ColWidth(0, -1) 'return to default
			grid1.set_ColWidth(-1, grid1.get_ColWidth(0) / 2.5)
			ColZeroWidth1 = grid1.get_ColWidth(0) * 3.9
			ColZeroWidth2 = grid1.get_ColWidth(0) * 2.8
			ColZeroWidth3 = grid1.get_ColWidth(0)
			If bDeity Then
				ColZeroWidth1 = ColZeroWidth1 + 400
				ColZeroWidth2 = ColZeroWidth2 + 400
				ColZeroWidth3 = ColZeroWidth3 + 400
			End If
			bOldShortNameUnitGrid = Not frmOptions.bShortNameUnitGrid
		End If
		LastDisplayFlag = DisplaySelect
		
		grid1.set_ColWidth(0, -1) 'return to default
		grid1.set_ColWidth(-1, grid1.get_ColWidth(0) / 2.5)
		
		If frmOptions.bShortNameUnitGrid Then
			grid1.set_ColWidth(0, ColZeroWidth3)
		ElseIf DisplaySelect = enumUnitDisplay.udPLANE Or DisplaySelect = enumUnitDisplay.udENEMY Then 
			grid1.set_ColWidth(0, ColZeroWidth1)
		Else
			grid1.set_ColWidth(0, ColZeroWidth2)
		End If
		
		If DisplaySelect = enumUnitDisplay.udENEMY Then
			grid1.set_ColWidth(8, grid1.get_ColWidth(0) * 1.25)
		End If
		
		grid1.set_ColData(0, 0)
		grid1.set_ColData(1, 1)
		grid1.row = 0
		'Fill in row headers
		grid1.col = 1
		grid1.Text = "id"
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		If DisplaySelect = enumUnitDisplay.udPLANE Then
			iCalculatedFields = iCalculatedFields + 1
			grid1.set_ColWidth(2, grid1.get_ColWidth(1))
			grid1.col = 2
			grid1.Text = "mis"
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
			grid1.set_ColData(2, 0) 'Text
		End If
		If DisplaySelect = enumUnitDisplay.udNUKE Then
			X = 1
			iCalculatedFields = 1
		Else
			X = 2
		End If
		Do While X < fieldcnt
			If DisplaySelect <> enumUnitDisplay.udNUKE Or X <> 4 Then
				grid1.col = X + iCalculatedFields
				grid1.Text = rs.Fields(X).Name
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
				If rs.Fields(X).Type = DAO.DataTypeEnum.dbText Then
					grid1.set_ColData(X + iCalculatedFields, 0) 'Text
				Else
					'rjk 8/11/03 Change the range column type to Airport Code so the grid can be sorted
					If (DisplaySubset = enumSubset.ssPLANE_RANGE Or DisplaySubset = enumSubset.ssMISSILE_RANGE) And grid1.Text = "range" Then
						grid1.set_ColData(X + iCalculatedFields, 2) 'Airport Code
					Else
						grid1.set_ColData(X + iCalculatedFields, 1) 'Number
					End If
				End If
				If DisplaySelect = enumUnitDisplay.udSHIP And rs.Fields(X).Name = "rng" Then
					For iIndex = 1 To 4
						iCalculatedFields = iCalculatedFields + 1
						grid1.Cols = fieldcnt + iCalculatedFields
						grid1.set_ColWidth(grid1.Cols - 1, grid1.get_ColWidth(grid1.Cols - 2))
						grid1.col = X + iCalculatedFields
						Select Case (iIndex)
							Case 1
								grid1.Text = "RFR"
							Case 2
								grid1.Text = "SMC"
							Case 3
								grid1.Text = "RPT"
							Case 4
								grid1.Text = "RMM"
						End Select
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
						grid1.set_ColData(X + iCalculatedFields, 1) 'Number
					Next iIndex
				End If
				If DisplaySelect = enumUnitDisplay.udLAND And rs.Fields(X).Name = "frg" Then
					iCalculatedFields = iCalculatedFields + 1
					grid1.Cols = fieldcnt + iCalculatedFields
					grid1.set_ColWidth(grid1.Cols - 1, grid1.get_ColWidth(grid1.Cols - 2))
					grid1.col = X + iCalculatedFields
					grid1.Text = "RFR"
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					grid1.set_ColData(X + iCalculatedFields, 1) 'Number
				End If
				grid1.CellFontBold = False
			Else
				iCalculatedFields = iCalculatedFields - 1
			End If
			X = X + 1
		Loop 
		
		'Fill in rows
		If Not (rs.BOF And rs.EOF) Then
			rs.MoveFirst()
		End If
		Dim movecost As Single
		While Not rs.EOF
			
			'Add to the sub combo box
			'092503 rjk: Eliminated FillSubSetFlag
			'If FillSubSetFlag Then
			If DisplaySubset <> enumSubset.ssPLANE_RANGE And DisplaySubset <> enumSubset.ssMISSILE_RANGE And DisplaySubset <> enumSubset.ssATTACK_RANGE Then
				AddSubBox(substype & CStr(rs.Fields(substypekey).Value), rs.Fields("x").Value, rs.Fields("y").Value) '100803 rjk: Added CStr for nation number
			End If
			
			If DisplaySelect = enumUnitDisplay.udPLANE Then
				PlaneRange_Renamed = rs.Fields("range").Value
			End If
			
			'Check to see if we are displaying this fleet
			' 092303 rjk: Removed secx and secy and switched to SelX and SelY
			'(DisplaySubset = DSP_SECT And secx = rs.Fields("x") And secy = rs.Fields("y")) Or _
			''(DisplaySubset = DSP_ATTACK_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y"))) = 1)) Or _
			''(DisplaySubset = DSP_PLANE_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y")) * 2) <= planerange)) Or _
			''(DisplaySubset = DSP_MISSILE_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y"))) <= planerange)) Then
			'092503 rjk: Switched Fleet to have whole string, allows separation ship fleets from plane wings (FillUnitList)
			'(DisplaySubset = ssFLEET And Fleet = Trim$(rs.Fields(substypekey))) Or
			If (DisplaySubset = enumSubset.ssALL) Or (DisplaySubset = enumSubset.ssFLEET And Fleet = substype & Trim(rs.Fields(substypekey).Value)) Or (DisplaySubset = enumSubset.ssSECT And SelX = rs.Fields("x").Value And SelY = rs.Fields("y").Value) Or (DisplaySubset = enumSubset.ssATTACK_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x").Value, rs.Fields("y").Value)) = 1)) Or (DisplaySubset = enumSubset.ssPLANE_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x").Value, rs.Fields("y").Value) * 2) <= PlaneRange_Renamed)) Or (DisplaySubset = enumSubset.ssMISSILE_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x").Value, rs.Fields("y").Value)) <= PlaneRange_Renamed)) Then
				
				'check to see if it is in the selected type subset
				If SubType = enumUnitType.TYPE_ALL Or InStr(strType(SubType), " " + rs.Fields(1).Value + " ") Or ((DisplaySelect = enumUnitDisplay.udLAND Or DisplaySelect = enumUnitDisplay.udPLANE) And SubType = enumUnitType.TYPE_SHIP_LOADED And rs.Fields("ship").Value <> -1) Or ((DisplaySelect = enumUnitDisplay.udLAND Or DisplaySelect = enumUnitDisplay.udPLANE) And SubType = enumUnitType.TYPE_SHIP_UNLOADED And rs.Fields("ship").Value = -1) Or ((DisplaySelect = enumUnitDisplay.udLAND Or DisplaySelect = enumUnitDisplay.udPLANE) And SubType = enumUnitType.TYPE_LAND_LOADED And rs.Fields("land").Value <> -1) Or ((DisplaySelect = enumUnitDisplay.udLAND Or DisplaySelect = enumUnitDisplay.udPLANE) And SubType = enumUnitType.TYPE_LAND_UNLOADED And rs.Fields("land").Value = -1) Then
					'check the mob and eff requirement
					If DisplaySelect = enumUnitDisplay.udENEMY Or (((Not NeedMob) Or rs.Fields("mob").Value > 0) And ((Not Needeff) Or rs.Fields("eff").Value > 59 Or (DisplaySelect = enumUnitDisplay.udPLANE And rs.Fields("eff").Value > 39))) Then
						
						'check if we are only displaying warships
						'092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
						'(ShowOnlyWarships And (rs.Fields("fir") > 0 Or rs.Fields("type") = "ls")) Then
						If DisplaySelect <> enumUnitDisplay.udSHIP Or Not frmOptions.bShowOnlyWarships Or (frmOptions.bShowOnlyWarships And (rs.Fields("fir").Value > 0 Or UnitCharacteristicCheck(enumUnitType.TYPE_S_LAND, rsShip.Fields("type").Value))) Then
							
							grid1.row = grid1.row + 1
							grid1.col = 0
							grid1.CellForeColor = System.Drawing.Color.Black
							
							If DisplaySelect = enumUnitDisplay.udENEMY Then
								BuildType = rs.Fields(9).Value 'Check the class
							Else
								'check for mission - use red if found
								rsMissions.Seek("=", BuildType, rs.Fields(0))
								If Not rsMissions.NoMatch Then
									grid1.CellForeColor = System.Drawing.Color.Red
								End If
								
								If DisplaySelect = enumUnitDisplay.udSHIP Then
									rsShipOrders.Seek("=", rs.Fields(0))
									If Not rsShipOrders.NoMatch Then
										If grid1.CellForeColor.equals(System.Drawing.Color.Black) Then
											grid1.CellForeColor = System.Drawing.Color.Blue
										Else
											grid1.CellForeColor = System.Drawing.Color.Magenta
										End If
									End If
								End If
							End If
							
							'Get the unit description
							If DisplaySelect = enumUnitDisplay.udNUKE Then
								grid1.Text = rs.Fields("type").Value
							Else
								grid1.Text = rs.Fields(1).Value ' Set in case lookup fails
								If Not frmOptions.bShortNameUnitGrid Then
									rsBuild.Seek("=", BuildType, rs.Fields(1))
									If Not rsBuild.NoMatch Then
										If Len(rsBuild.Fields("desc").Value) > 0 Then
											grid1.Text = rsBuild.Fields("desc").Value
										End If
									End If
								End If
							End If
							
							grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
							If bDeity Then
								grid1.Text = VB.Left(Nations.NationName(rs.Fields("country").Value), 4) & " " & grid1.Text
							End If
							grid1.col = 1
							grid1.CellForeColor = System.Drawing.Color.Black
							Select Case DisplaySelect
								Case enumUnitDisplay.udSHIP
									hldIndex = rsShip.Index
									rsShip.Index = "PrimaryKey"
									rsShip.Seek("=", rs.Fields("id"))
									If Not rsShip.NoMatch Then
										If rsShip.Fields("off").Value Then
											grid1.CellForeColor = System.Drawing.Color.Red
										End If
									End If
									rsShip.Index = hldIndex
								Case enumUnitDisplay.udLAND
									hldIndex = rsLand.Index
									rsLand.Index = "PrimaryKey"
									rsLand.Seek("=", rs.Fields("id"))
									If Not rsLand.NoMatch Then
										If rsLand.Fields("off").Value Then
											grid1.CellForeColor = System.Drawing.Color.Red
										End If
									End If
									rsLand.Index = hldIndex
								Case enumUnitDisplay.udPLANE
									hldIndex = rsPlane.Index
									rsPlane.Index = "PrimaryKey"
									rsPlane.Seek("=", rs.Fields("id"))
									If Not rsPlane.NoMatch Then
										If rsPlane.Fields("off").Value Then
											grid1.CellForeColor = System.Drawing.Color.Red
										End If
									End If
									rsPlane.Index = hldIndex
								Case enumUnitDisplay.udNUKE
									If rs.Fields("off").Value Then
										grid1.CellForeColor = System.Drawing.Color.Red
									End If
							End Select
							grid1.Text = CStr(rs.Fields(0).Value)
							grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
							
							iCalculatedFields = 0
							If DisplaySelect = enumUnitDisplay.udPLANE Then
								iCalculatedFields = iCalculatedFields + 1
								grid1.col = 2
								grid1.CellForeColor = System.Drawing.Color.Black
								grid1.Text = ""
								If Not rsMissions.NoMatch Then
									If Len(rsMissions.Fields("mission").Value) > 0 Then
										grid1.Text = StrConv(VB.Left(rsMissions.Fields("mission").Value, 1), VbStrConv.UpperCase) & "-" & CStr(rsMissions.Fields("radius").Value)
									End If
								End If
							End If
							If DisplaySelect = enumUnitDisplay.udNUKE Then
								X = 1
								iCalculatedFields = 1
							Else
								X = 2
							End If
							Do While X < fieldcnt
								If DisplaySelect <> enumUnitDisplay.udNUKE Or X <> 4 Then
									grid1.col = X + iCalculatedFields
									grid1.CellForeColor = System.Drawing.Color.Black
									
									'101803 rjk: Added check for x,y fields as -1 is valid for these fields.
									If (DisplaySelect = enumUnitDisplay.udENEMY) And (grid1.Text = "-1") And (rs.Fields(X).Name <> "x" And rs.Fields(X).Name <> "y") Then
										grid1.Text = "???"
									ElseIf rs.Fields(X).Name = "timestamp" Then 
										grid1.Text = VB6.Format(ConvertToLocalTimeFromWinACEUTC((rs.Fields("timestamp").Value)), "ddd mmm dd hh:mm:ss yyyy")
									ElseIf (DisplaySubset = enumSubset.ssPLANE_RANGE Or DisplaySubset = enumSubset.ssMISSILE_RANGE) And rs.Fields(X).Name = "range" Then 
										grid1.CellForeColor = System.Drawing.Color.Red
										On Error GoTo 0
										strLoc = SectorString(rs.Fields("x").Value, rs.Fields("y").Value)
										If strLoc <> strLast Then
											i = 0
											found = False
											strLast = strLoc
											While i < UBound(Airports) And Not found
												i = i + 1
												If Airports(i) = strLoc Then
													found = True
												End If
											End While
											'then add if necessary
											If Not found Then
												i = UBound(Airports) + 1
												ReDim Preserve Airports(i)
												Airports(i) = strLoc
											End If
											airport = 96 + i
										End If
										'Save last airport, as planes usually come out in order
										strLast = strLoc
										' 092303 rjk: Removed secx and secy and switched to SelX and SelY
										'grid1.Text = Chr$(airport) + "-" + CStr(SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y")))
										grid1.Text = Chr(airport) & "-" & CStr(SectorDistance(SelX, SelY, rs.Fields("x").Value, rs.Fields("y").Value))
										On Error Resume Next
									Else
										grid1.Text = rs.Fields(X).Value
									End If
									grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
									If (DisplaySelect = enumUnitDisplay.udSHIP And rs.Fields(X).Name = "rng") Then
										
										movecost = ShipMoveCost(rs.Fields("spd").Value, rs.Fields("eff").Value, CShort(rs.Fields("tech").Value), rs.Fields("type").Value)
										For iIndex = 1 To 4
											iCalculatedFields = iCalculatedFields + 1
											grid1.col = X + iCalculatedFields
											grid1.CellForeColor = System.Drawing.Color.Black
											Select Case iIndex
												Case 1
													grid1.Text = VB6.Format(UnitFireRange(rs.Fields("tech").Value, rs.Fields("rng").Value), "#0.00")
												Case 2
													grid1.Text = VB6.Format(movecost, "#0.0")
												Case 3
													grid1.Text = VB6.Format(rsVersion.Fields("Mob gain - Ship").Value / movecost, "#0.0")
												Case 4
													grid1.Text = VB6.Format(rsVersion.Fields("Max mob - Ship").Value / movecost, "#0.0")
											End Select
										Next iIndex
									End If
									If (DisplaySelect = enumUnitDisplay.udLAND And rs.Fields(X).Name = "frg") Then
										iCalculatedFields = iCalculatedFields + 1
										grid1.col = X + iCalculatedFields
										grid1.CellForeColor = System.Drawing.Color.Black
										grid1.Text = VB6.Format(UnitFireRange(rs.Fields("tech").Value, rs.Fields("frg").Value), "#0.00")
									End If
								Else
									iCalculatedFields = iCalculatedFields - 1
								End If
								X = X + 1
							Loop 
							grid1.Rows = grid1.Rows + 1
						End If
					End If
				End If
			End If
			rs.MoveNext()
		End While
		
		'Set the count
		If grid1.Rows > 1 Then
			grid1.row = grid1.row + 1
			grid1.col = 0
			grid1.Text = "Total: " & CStr(grid1.Rows - 2) & " units"
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
			grid1.CellFontBold = True
			grid1.CellFontName = "Arial"
		End If
		
		'Reshow the grid
		grid1.Visible = True
		
		'Set the option buttons to the correct entry
		'UPGRADE_WARNING: Lower bound of collection Toolbar1.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		CType(Toolbar1.Items.Item(DisplaySelect), ToolStripButton).Checked = True
		
		'Set the cmbSelect if box was refilled
		'092503 rjk: Moved to SetIndexSubCombo
		SetIndexSubCombo()
		
		'FillSubSetFlag = False  092503 rjk: Eliminated
		GridUpdateinProgress = False
	End Sub
	
	Private Sub ClearSubCombo()
		cmbSub.Items.Clear()
		strSubSect = vbNullString
		strSubFleet = vbNullString
	End Sub
	
	Private Sub SetIndexSubCombo()
		Dim tstx As Integer
		Dim tsty As Integer
		Dim X As Short
		
		For X = 0 To cmbSub.Items.Count
			Select Case DisplaySubset
				Case enumSubset.ssALL '112503 rjk: Moved All inside, as All is not always at the top
					If VB6.GetItemString(cmbSub, X) = "All" Then
						cmbSub.SelectedIndex = X
						Exit Sub
					End If
				Case enumSubset.ssFLEET
					If VB6.GetItemString(cmbSub, X) = Fleet Then
						cmbSub.SelectedIndex = X
						Exit Sub
					End If
				Case enumSubset.ssSECT
					If InStr(VB6.GetItemString(cmbSub, X), ",") > 0 Then
						tstx = VB6.GetItemData(cmbSub, X) / 10000
						tsty = VB6.GetItemData(cmbSub, X) - tstx * 10000
						If tstx >= 1000 Then
							tstx = 0 - tstx + 1000
						End If
						If tsty >= 1000 Then
							tsty = 0 - tsty + 1000
						End If
						If tstx = SelX And tsty = SelY Then
							cmbSub.SelectedIndex = X
							Exit Sub
						End If
					End If
			End Select
		Next 
	End Sub
	
	'UPGRADE_WARNING: Event cmbSub.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbSub_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSub.SelectedIndexChanged
		
		Dim strSelect As String
		Dim bUpdate As Boolean
		Dim n As Short
		' Dim n2 As Integer   removed efj 8/2003
		Dim NewSecX As Short
		Dim NewSecY As Short
		
		'Boolean keeps track if an update of grid is needed
		bUpdate = False
		
		'First check if all was selected
		strSelect = VB6.GetItemString(cmbSub, cmbSub.SelectedIndex)
		If strSelect = "All" Then
			If DisplaySubset <> enumSubset.ssALL Then
				bUpdate = True
				DisplaySubset = enumSubset.ssALL
			End If
			
			'Then See if a range is up - this never changes
		ElseIf VB.Left(strSelect, 3) = "In " Then 
			Exit Sub
			
			'Then See if a sector was selected
		ElseIf InStr(strSelect, ",") > 0 Then 
			n = InStr(strSelect, ",")
			NewSecX = CShort(Trim(VB.Left(strSelect, n - 1)))
			NewSecY = CShort(Trim(Mid(strSelect, n + 1)))
			
			'Make sure selection criteria has changed
			If DisplaySubset <> enumSubset.ssSECT Then
				bUpdate = True
				DisplaySubset = enumSubset.ssSECT
				' 092303 rjk: Removed secx and secy and switched to SelX and SelY
				FillSectorBox(NewSecX, NewSecY)
				'        secx = NewSecX
				'        secy = NewSecY
			Else
				' 092303 rjk: Removed secx and secy and switched to SelX and SelY
				'        If secx <> NewSecX Or secy <> NewSecY Then
				If SelX <> NewSecX Or SelY <> NewSecY Then
					bUpdate = True
					' 092303 rjk: Removed secx and secy and switched to SelX and SelY
					FillSectorBox(NewSecX, NewSecY)
					'            secx = NewSecX
					'           secy = NewSecY
				End If
			End If
		Else
			
			' Finally, must be a fleet/army/wing
			'092503 rjk: Switched Fleet to have whole name, used in FillUnitList to separate ship fleets from plane wings
			If DisplaySubset <> enumSubset.ssFLEET Then
				bUpdate = True
				DisplaySubset = enumSubset.ssFLEET
				Fleet = strSelect
			Else
				If Fleet <> strSelect Then
					Fleet = strSelect
					bUpdate = True
				End If
			End If
		End If
		
		If bUpdate Then
			FillGrid()
		End If
		
	End Sub
	
	'092503 rjk: Remove stype as not needed any more
	Private Sub AddSubBox(ByRef Fleet As String, ByRef X As Short, ByRef Y As Short)
		Dim strSec As String
		Dim n As Integer
		
		'If fleet has not been added
		'092503 rjk: Switched Fleet to have whole name, used for FillUnitList to separate ship fleets from plane wings
		If bDeity Or InStr(Fleet, "Nation") = 1 Then
			If InStr(strSubFleet, Mid(Fleet, 7)) = 0 Then
				strSubFleet = strSubFleet & Mid(Fleet, 7) & " "
				cmbSub.Items.Add(Fleet)
			End If
		ElseIf InStr(Fleet, "Type") = 1 Then 
			If InStr(strSubFleet, Fleet) = 0 Then
				strSubFleet = strSubFleet & Fleet & " "
				cmbSub.Items.Add(Fleet)
			End If
		ElseIf InStr(strSubFleet, VB.Right(Fleet, 1)) = 0 Then 
			strSubFleet = strSubFleet & VB.Right(Fleet, 1)
			cmbSub.Items.Add(Fleet)
		End If
		
		strSec = " " & SectorString(X, Y) & " "
		
		If InStr(strSubSect, strSec) = 0 Then
			strSubSect = strSubSect & strSec
			'Encode sector numbers in item data
			If X < 0 Then
				n = 0 - X + 1000
			Else
				n = X
			End If
			n = n * 10000
			If Y < 0 Then
				n = n - Y + 1000
			Else
				n = n + Y
			End If
			
			cmbSub.Items.Add(New VB6.ListBoxItem(Trim(strSec), n))
		End If
	End Sub
	
	Private Sub DeleteUnits()
		Dim StartX As Short
		Dim Finishx As Short
		Dim x1 As Short
		Dim gi As Short
		Dim id As Short
		Dim strOwner As String
		Dim strClass As String
		Dim hldIndex As String
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		StartX = grid1.row
		Finishx = grid1.RowSel
		
		If Finishx < StartX Then
			x1 = StartX
			StartX = Finishx
			Finishx = x1
		End If
		If Finishx > grid1.Rows - 2 Then
			Finishx = grid1.Rows - 2
		End If
		
		'set recordset
		Select Case DisplaySelect
			Case enumUnitDisplay.udSHIP
				rs = rsShip
			Case enumUnitDisplay.udLAND
				rs = rsLand
			Case enumUnitDisplay.udPLANE
				rs = rsPlane
			Case enumUnitDisplay.udNUKE
				rs = rsNuke
			Case enumUnitDisplay.udENEMY
				rs = rsEnemyUnit
			Case Else
				Exit Sub
		End Select
		
		'set index
		hldIndex = rs.Index
		rs.Index = "PrimaryKey"
		
		'go thru selected rows, delete units
		For x1 = StartX To Finishx
			gi = gridIndex(x1, 1)
			id = Val(grid1.get_TextArray(gi))
			If DisplaySelect = enumUnitDisplay.udENEMY Then
				gi = gridIndex(x1, 4)
				strOwner = grid1.get_TextArray(gi)
				gi = gridIndex(x1, 9)
				strClass = grid1.get_TextArray(gi)
				If strOwner = "???" Then '091603 rjk: CInt("???") gives run-time, it is trying to select '-1' row
					rs.Seek("=", -1, strClass, id)
				Else
					rs.Seek("=", CShort(strOwner), strClass, id)
				End If
			Else
				rs.Seek("=", id)
			End If
			If Not rs.NoMatch Then
				rs.Delete()
			End If
		Next x1
		rs.Index = hldIndex
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		
		'redraw the map
		Me.DrawHexes()
		FillGrid()
		
	End Sub
	
	'Private Sub SetUnitType(n As enumUnitType, Optional ReqMobility As Boolean, Optional ReqMinEff As Boolean)
	
	'Static SaveMobility As Variant
	'Static SaveEfficiency As Variant
	'Static SaveSubtype As Variant
	'Static SkipSave As Boolean
	
	'Save values if this is an option box request
	'If n = TYPE_RESTORE Then
	'    n = SaveSubtype
	'    Toolbar1.Buttons(6).Value = SaveMobility
	'    Toolbar1.Buttons(7).Value = SaveEfficiency
	'    NeedMob = (SaveMobility = tbrPressed)
	'    Needeff = (SaveEfficiency = tbrPressed)
	'    SkipSave = False
	'ElseIf (ReqMobility Or ReqMinEff) And (Not SkipSave) Then
	'    SaveMobility = Toolbar1.Buttons(6).Value
	'    SaveEfficiency = Toolbar1.Buttons(7).Value
	'    SaveSubtype = SubType
	'    SkipSave = True
	'End If
	
	'SubType = n
	'DisplayUnitBox 'rjk 05/13/03: Switched to DisplayUnitBox function from DisplaySectorPanel 3
	
	'NeedMob = ReqMobility
	'If ReqMobility Then
	'    Toolbar1.Buttons(6).Value = tbrPressed
	'Else
	'    Toolbar1.Buttons(6).Value = tbrUnpressed
	'End If
	
	'Needeff = ReqMinEff
	'If ReqMinEff = True Then
	'    Toolbar1.Buttons(7).Value = tbrPressed
	'Else
	'    Toolbar1.Buttons(7).Value = tbrUnpressed
	'End If
	
	'FillGrid
	'End Sub
	
	Private Sub grid1_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseUpEvent) Handles grid1.MouseUpEvent
		Dim tb As System.Windows.Forms.TextBox
		Dim StartX As Short
		Dim Finishx As Short
		Dim x1 As Short
		Dim n As Short
		Dim bCtl As Boolean
		Dim found As Boolean
		Dim strSel As String ' 092503 rjk: moved to here, only place used
		Dim CurrUnit As Object
		
		On Error Resume Next
		
		strTextBox = vbNullString
		
		'If clicking in top row, resort current box
		If grid1.MouseRow = 0 Then
			SortGrid((grid1.MouseCol))
			Exit Sub
		End If
		
		'this routine is used to fill prompt boxes
		If Not PromptUp Then
			Exit Sub
		End If
		
		'See if this is a box you add units to
		tb = PromptForm.ActiveControl
		If Not (tb.Name = "txtUnit" Or tb.Name = "txtUnit2" Or tb.Name = "txtEscort" Or tb.Name = "txtUnit3" Or (tb.Name = "txtTarget" And DisplaySelect = enumUnitDisplay.udENEMY)) Then
			Exit Sub
		End If
		
		' bShift = False   removed efj 8/2003
		If (eventArgs.Shift And VB6.ShiftConstants.ShiftMask) = 1 Then
			'    bShift = True   removed efj 8/2003
		End If
		
		bCtl = False
		If (eventArgs.Shift And VB6.ShiftConstants.CtrlMask) = 2 Then
			bCtl = True
		End If
		
		'use a temp string to avoid firing text boxs "change"
		'event muliple times
		strTextBox = tb.Text
		
		'If control key is down, clear existing units else
		'Get current units
		If Len(strTextBox) > 0 And Not bCtl Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CurrUnit = ParseUnitString(tb.Text)
		Else
			ReDim CurrUnit(0)
			'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CurrUnit(0) = vbNullString
			strTextBox = vbNullString
		End If
		
		StartX = grid1.row
		Finishx = grid1.RowSel
		
		If Finishx < StartX Then
			x1 = StartX
			StartX = Finishx
			Finishx = x1
		End If
		
		strSel = vbNullString
		
		'go thru selected rows, and add units to select String
		For x1 = StartX To Finishx
			strSel = Trim(grid1.get_TextMatrix(x1, 1))
			found = False
			For n = LBound(CurrUnit) To UBound(CurrUnit)
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CurrUnit(n) = strSel Then
					found = True
				End If
			Next n
			If Not found Then
				If Len(strTextBox) > 0 Then
					strTextBox = strTextBox & "/"
				End If
				strTextBox = strTextBox & strSel
			End If
		Next x1
		
		tb.Text = strTextBox
		Me.PromptForm.Activate()
		
	End Sub
	
	Private Sub SortGrid(ByRef col As Short)
		
		'This does a simple bubble sort of the grid in place.  Due to the small size
		'of the sample, a more efficent sort is usually not necessary.
		
		Dim n As Short
		Dim row As Short
		Dim var1 As Object
		Dim var2 As Object
		Dim sng1 As Single
		Dim sng2 As Single
		
		'If clicked multiple times, change the sort order
		If SortCol = col Then
			SortDecending = Not SortDecending
		End If
		SortCol = col
		
		grid1.Visible = False
		
		n = grid1.Rows - 3
		'rjk 8/11/03 Added sorting for Airport Code
		Select Case grid1.get_ColData(col)
			Case 0 'Text
				While n > 0
					For row = 1 To n
						'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						var1 = grid1.get_TextArray(gridIndex(row, col))
						'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						var2 = grid1.get_TextArray(gridIndex(row + 1, col))
						
						If SortDecending Then
							'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If var1 > var2 Then
								grid1.set_RowPosition(row + 1, row)
							End If
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If var1 < var2 Then
								grid1.set_RowPosition(row + 1, row)
							End If
						End If
					Next row
					n = n - 1
				End While
			Case 1 'Number
				While n > 0
					For row = 1 To n
						If Len(grid1.get_TextArray(gridIndex(row, col))) = 0 Then
							sng1 = 0
						Else
							sng1 = CSng(grid1.get_TextArray(gridIndex(row, col)))
						End If
						If Len(grid1.get_TextArray(gridIndex(row + 1, col))) = 0 Then
							sng2 = 0
						Else
							sng2 = CSng(grid1.get_TextArray(gridIndex(row + 1, col)))
						End If
						
						If SortDecending Then
							If sng1 > sng2 Then
								grid1.set_RowPosition(row + 1, row)
							End If
						Else
							If sng1 < sng2 Then
								grid1.set_RowPosition(row + 1, row)
							End If
						End If
					Next row
					n = n - 1
				End While
			Case 2 'Airport Code (letter 'dash' range (number))
				While n > 0
					'Sort by range first
					For row = 1 To n
						If Len(grid1.get_TextArray(gridIndex(row, col))) < 3 Then
							sng1 = 0
						Else
							sng1 = CSng(Mid(grid1.get_TextArray(gridIndex(row, col)), 3))
						End If
						If Len(grid1.get_TextArray(gridIndex(row + 1, col))) < 3 Then
							sng2 = 0
						Else
							sng2 = CSng(Mid(grid1.get_TextArray(gridIndex(row + 1, col)), 3))
						End If
						
						'If the range is equal then sort by the airport letter
						If sng1 = sng2 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							var1 = Mid(grid1.get_TextArray(gridIndex(row, col)), 1, 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							var2 = Mid(grid1.get_TextArray(gridIndex(row + 1, col)), 1, 1)
							
							If SortDecending Then
								'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If var1 > var2 Then
									grid1.set_RowPosition(row + 1, row)
								End If
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If var1 < var2 Then
									grid1.set_RowPosition(row + 1, row)
								End If
							End If
						Else
							If SortDecending Then
								If sng1 > sng2 Then
									grid1.set_RowPosition(row + 1, row)
								End If
							Else
								If sng1 < sng2 Then
									grid1.set_RowPosition(row + 1, row)
								End If
							End If
						End If
					Next row
					n = n - 1
				End While
		End Select
		
		grid1.Visible = True
		
	End Sub
	
	'091903 rjk: Removed ChangeUnitDisplay combined with Toolbar1_ButtonClick
	Private Sub Toolbar1_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _Toolbar1_Button1.Click, _Toolbar1_Button2.Click, _Toolbar1_Button3.Click, _Toolbar1_Button4.Click, _Toolbar1_Button5.Click, _Toolbar1_Button6.Click, _Toolbar1_Button7.Click, _Toolbar1_Button8.Click, _Toolbar1_Button9.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		'   ChangeUnitDisplay Button.key, (Button.Value = tbrPressed)
		'End Sub
		
		'Public Sub ChangeUnitDisplay(ByVal key As String, ByVal v As Boolean)
		' added by thomas lullier
		' revised by daniel to easily manage unit selection
		
		Dim n As enumUnitDisplay
		Select Case Button.Name
			Case "Ship"
				n = enumUnitDisplay.udSHIP
			Case "Land"
				n = enumUnitDisplay.udLAND
			Case "Plane"
				n = enumUnitDisplay.udPLANE
			Case "Nuke"
				n = enumUnitDisplay.udNUKE
			Case "Enemy"
				n = enumUnitDisplay.udENEMY
			Case "List"
				n = enumUnitDisplay.udList
				
			Case "Mob"
				NeedMob = (Button.Checked = True)
				FillGrid()
				Exit Sub
				
			Case "Delete"
				DeleteUnits()
				Exit Sub
				
			Case "Eff"
				Needeff = (Button.Checked = True)
				FillGrid()
				Exit Sub
		End Select
		
		'FillTypeBox - already done in FillGrid, and should not need to be redone, it just
		' reclicking button
		
		'If option Changed - Refill the grid
		If DisplaySelect <> n Then
			DisplaySelect = n
			SubType = enumUnitType.TYPE_ALL
			
			'if showing fleet/wing/army, shift to all
			If DisplaySubset = enumSubset.ssFLEET Then
				DisplaySubset = enumSubset.ssALL
			End If
			
			'Refill
			'FillSubSetFlag = True 092503 rjk: Eliminated
			FillGrid()
		End If
		
	End Sub
	
	Private Sub FillTypeBox()
		Static CurrentType As enumUnitDisplay
		'Dim hldSubtype As enumUnitType
		Dim Index As Short
		
		'loading the box can change the subtype, so we need
		'to store it and reload it
		'hldSubtype = SubType
		
		If CurrentType <> DisplaySelect Then
			CurrentType = DisplaySelect
			'optSelect(0).Tag = CStr(TYPE_ALL)
			'optSelect(0).Caption = "All"
			'optSelect(1).ToolTipText = "Show All"
			
			'show all
			'For n = 1 To 7
			'    optSelect(n).Visible = True
			'Next n
			
			cmbUnitFilter.Items.Clear()
			cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_ALL), enumUnitType.TYPE_ALL))
			
			Select Case DisplaySelect
				Case enumUnitDisplay.udSHIP
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_SC_FIRE), enumUnitType.TYPE_SC_FIRE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_FISH), enumUnitType.TYPE_S_FISH))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_TORP), enumUnitType.TYPE_S_TORP))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_DCHRG), enumUnitType.TYPE_S_DCHRG))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_PLANE), enumUnitType.TYPE_S_PLANE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_MISS), enumUnitType.TYPE_S_MISS))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_OIL), enumUnitType.TYPE_S_OIL))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_OILER), enumUnitType.TYPE_S_OILER))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SONAR), enumUnitType.TYPE_S_SONAR))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_MINE), enumUnitType.TYPE_S_MINE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SWEEP), enumUnitType.TYPE_S_SWEEP))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SUB), enumUnitType.TYPE_S_SUB))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SPY), enumUnitType.TYPE_S_SPY))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_LAND), enumUnitType.TYPE_S_LAND))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SEMI_LAND), enumUnitType.TYPE_S_SEMI_LAND))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_SUB_TORP), enumUnitType.TYPE_S_SUB_TORP))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_TRADE), enumUnitType.TYPE_S_TRADE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_S_ANTI_MISSILE), enumUnitType.TYPE_S_ANTI_MISSILE))
					cmbUnitFilter.Visible = True
				Case enumUnitDisplay.udLAND
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_LC_FIRE), enumUnitType.TYPE_LC_FIRE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_XLIGHT), enumUnitType.TYPE_L_XLIGHT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_ENGINEER), enumUnitType.TYPE_L_ENGINEER))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_SUPPLY), enumUnitType.TYPE_L_SUPPLY))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_SECURITY), enumUnitType.TYPE_L_SECURITY))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_LIGHT), enumUnitType.TYPE_L_LIGHT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_MARINE), enumUnitType.TYPE_L_MARINE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_RECON), enumUnitType.TYPE_L_RECON))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_RADAR), enumUnitType.TYPE_L_RADAR))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_L_ASSAULT), enumUnitType.TYPE_L_ASSAULT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_LAND_UNLOADED), enumUnitType.TYPE_LAND_UNLOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_LAND_LOADED), enumUnitType.TYPE_LAND_LOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_SHIP_UNLOADED), enumUnitType.TYPE_SHIP_UNLOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_SHIP_LOADED), enumUnitType.TYPE_SHIP_LOADED))
					cmbUnitFilter.Visible = True
				Case enumUnitDisplay.udPLANE
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_PC_BOMB), enumUnitType.TYPE_PC_BOMB))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_PG_ESCORTS), enumUnitType.TYPE_PG_ESCORTS))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_TACTICAL), enumUnitType.TYPE_P_TACTICAL))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_BOMBER), enumUnitType.TYPE_P_BOMBER))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_INTERCEPT), enumUnitType.TYPE_P_INTERCEPT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_ESCORT), enumUnitType.TYPE_P_ESCORT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_PC_DROP), enumUnitType.TYPE_PC_DROP))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_CARGO), enumUnitType.TYPE_P_CARGO))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_MINE), enumUnitType.TYPE_P_MINE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_SWEEP), enumUnitType.TYPE_P_SWEEP))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_PARA), enumUnitType.TYPE_P_PARA))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_SPY), enumUnitType.TYPE_P_SPY))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_PC_LAUNCH), enumUnitType.TYPE_PC_LAUNCH))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_MISSILE), enumUnitType.TYPE_P_MISSILE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_MARINE), enumUnitType.TYPE_P_MARINE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_SDI), enumUnitType.TYPE_P_SDI))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_SATELLITE), enumUnitType.TYPE_P_SATELLITE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_IMAGE), enumUnitType.TYPE_P_IMAGE))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_ANTISUB), enumUnitType.TYPE_P_ANTISUB))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_CHOPPER), enumUnitType.TYPE_P_CHOPPER))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_LIGHT), enumUnitType.TYPE_P_LIGHT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_X_LIGHT), enumUnitType.TYPE_P_X_LIGHT))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_STEALTH), enumUnitType.TYPE_P_STEALTH))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_HALF_STEALTH), enumUnitType.TYPE_P_HALF_STEALTH))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_P_VTOL), enumUnitType.TYPE_P_VTOL))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_LAND_UNLOADED), enumUnitType.TYPE_LAND_UNLOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_LAND_LOADED), enumUnitType.TYPE_LAND_LOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_SHIP_UNLOADED), enumUnitType.TYPE_SHIP_UNLOADED))
					cmbUnitFilter.Items.Add(New VB6.ListBoxItem(strTypeTitle(enumUnitType.TYPE_SHIP_LOADED), enumUnitType.TYPE_SHIP_LOADED))
					cmbUnitFilter.Visible = True
				Case enumUnitDisplay.udNUKE
					cmbUnitFilter.Visible = False
				Case enumUnitDisplay.udENEMY, enumUnitDisplay.udList
					cmbUnitFilter.Visible = False
			End Select
		End If
		
		For Index = 0 To cmbUnitFilter.Items.Count - 1
			If SubType = VB6.GetItemData(cmbUnitFilter, Index) Then
				cmbUnitFilter.SelectedIndex = Index
				Exit For
			End If
		Next Index
	End Sub
	
	Private Sub FillUnitList()
		Dim NeedHeader As Boolean
		Dim strText As String
		Dim ListType As enumUnitDisplay
		
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim X As Short
		Dim substype As String
		Dim substypekey As String
		Dim hldIndex As String
		
		rtbUnitList.Visible = False
		rtbUnitList.Text = vbNullString
		grid1.Visible = False
		
		ClearSubCombo()
		cmbSub.Items.Add("All")
		
		For ListType = enumUnitDisplay.udSHIP To enumUnitDisplay.udENEMY '100603 rjk: changed to udENEMY to include enemy list
			Select Case ListType
				Case enumUnitDisplay.udSHIP
					rs = rsShipList
					rs.Requery()
					substype = "Fleet "
					substypekey = "flt"
					'BuildType = "s" 100603 rjk: Not used in this procedure
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = PadString("Ship", 5, False)
					strText = strText & "   id mil gun  sh  fd"
					
				Case enumUnitDisplay.udLAND
					rs = rsLandList
					rs.Requery()
					substype = "Army "
					substypekey = "army"
					'BuildType = "l" 100603 rjk: Not used in this procedure
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = PadString("Land", 5, False)
					strText = strText & "   id mil gun  sh  fd"
					
				Case enumUnitDisplay.udPLANE
					rs = rsPlaneList
					rs.Requery()
					substype = "Wing "
					substypekey = "wing"
					'BuildType = "p" 100603 rjk: Not used in this procedure
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = PadString("Plane", 5, False)
					strText = strText & "   id range armed sts"
					
				Case enumUnitDisplay.udNUKE
					rs = rsNuke
					'rs.Requery 100606 rjk: Removed, does not work this table.  The others do queries.
					substype = "Type "
					substypekey = "type"
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = PadString("Nuke", 5, False)
					strText = strText & "   id range armed sts"
					
				Case enumUnitDisplay.udENEMY
					rs = rsEnemyUnit
					'rs.Requery 100603 rjk: Removed, does not work this table.  Need to investigate more later.
					substype = "Nation "
					substypekey = "owner"
					'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = PadString("Enemy", 5, False)
					strText = strText & "   id mil eff tech  type owner"
			End Select
			
			NeedHeader = True
			
			If Not (rs.BOF And rs.EOF) Then
				rs.MoveFirst()
			End If
			While Not rs.EOF
				
				'Add to the sub combo box
				'If FillSubSetFlag Then 092503 rjk: always include
				'100603 rjk: Added CStr for enemy nation to be converted to string.
				AddSubBox(substype & CStr(rs.Fields(substypekey).Value), rs.Fields("x").Value, rs.Fields("y").Value)
				'End If
				
				'Check to see if we are displaying this fleet
				' 092303 rjk: Removed secx and secy and switched to SelX and SelY
				'(DisplaySubset = DSP_SECT And secx = rs.Fields("x") And secy = rs.Fields("y")) Then
				' 092503 rjk: Switched Fleet to have the whole fleet name otherwise can separate ship from planes
				'(DisplaySubset = ssFLEET And Fleet = Trim$(rs.Fields(substypekey))) Or
				If (DisplaySubset = enumSubset.ssALL) Or (DisplaySubset = enumSubset.ssFLEET And Fleet = substype & Trim(rs.Fields(substypekey).Value)) Or (DisplaySubset = enumSubset.ssSECT And SelX = rs.Fields("x").Value And SelY = rs.Fields("y").Value) Then
					'check to see if it is in the selected type subset
					'If SubType = 0 Or InStr(strType(SubType), " " + rs.Fields(1) + " ") Then
					
					'check the mob and eff requirement
					If ListType = enumUnitDisplay.udENEMY Then '100603 rjk: Updated udENEMY so it produced output
						If NeedHeader Then
							NeedHeader = False
							rtbUnitList.Font = VB6.FontChangeBold(rtbUnitList.SelectionFont, True)
							rtbUnitList.SelectionStart = 99999
							rtbUnitList.SelectionLength = 0
							rtbUnitList.SelectedText = strText
							rtbUnitList.Font = VB6.FontChangeBold(rtbUnitList.SelectionFont, False)
						End If
						'                    If ListType = udENEMY Then 100603 rjk: not class in this procedure
						'                        BuildType = rs.Fields(9)    'Check the class
						'                    End If
						strText = rs.Fields(1).Value
						'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strText = PadString(strText, 5, False)
						'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strText = strText + PadString("#" & CStr(rs.Fields(0).Value), 5)
						If rs.Fields("mil").Value <> -1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strText = strText + PadString(rs.Fields("mil").Value, 4)
						Else
							strText = strText & " ???"
						End If
						If rs.Fields("eff").Value <> -1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strText = strText + PadString(rs.Fields("eff").Value, 4)
						Else
							strText = strText & " ???"
						End If
						If rs.Fields("tech").Value <> -1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strText = strText + PadString(rs.Fields("tech").Value, 5)
						Else
							strText = strText & "  ???"
						End If
						Select Case rs.Fields("class").Value
							Case "p"
								strText = strText & " Plane"
							Case "s"
								strText = strText & "  Ship"
							Case "l"
								strText = strText & "  Land"
						End Select
						strText = strText & " " & Nations.NationName(rs.Fields("owner").Value) & "(" & CStr(rs.Fields("owner").Value) & ")"
						rtbUnitList.SelectionStart = 99999
						rtbUnitList.SelectionLength = 0
						rtbUnitList.SelectedText = vbCrLf & strText
					Else
						If (((Not NeedMob) Or rs.Fields("mob").Value > 0) And ((Not Needeff) Or rs.Fields("eff").Value > 59 Or (ListType = enumUnitDisplay.udPLANE And rs.Fields("eff").Value > 39))) Then
							
							If NeedHeader Then
								NeedHeader = False
								rtbUnitList.Font = VB6.FontChangeBold(rtbUnitList.SelectionFont, True)
								rtbUnitList.SelectionStart = 99999
								rtbUnitList.SelectionLength = 0
								rtbUnitList.SelectedText = strText
								rtbUnitList.Font = VB6.FontChangeBold(rtbUnitList.SelectionFont, False)
							End If
							
							'Get the unit description
							strText = rs.Fields(1).Value
							'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strText = PadString(strText, 5, False)
							'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strText = strText + PadString("#" & CStr(rs.Fields(0).Value), 5)
							If ListType = enumUnitDisplay.udSHIP Or ListType = enumUnitDisplay.udLAND Then
								'UPGRADE_WARNING: Couldn't resolve default property of object PadString(rs.Fields(food), 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object PadString(rs.Fields(shell), 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object PadString(rs.Fields(gun), 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strText = strText + PadString(rs.Fields("mil").Value, 4) + PadString(rs.Fields("gun").Value, 4) + PadString(rs.Fields("shell").Value, 4) + PadString(rs.Fields("food").Value, 4)
							End If
							If ListType = enumUnitDisplay.udPLANE Then
								'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strText = strText + PadString(rs.Fields("range").Value, 6)
								If rs.Fields("nuke").Value <> "N/A" Then
									'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									strText = strText + PadString(rs.Fields("nuke").Value, 6)
								Else
									strText = strText & "      "
								End If
								If rs.Fields("eff").Value < 39 Or rs.Fields("mob").Value < 0 Then
									strText = strText & " out  "
								ElseIf rs.Fields("eff").Value > 89 Then 
									strText = strText & " OK   "
								Else
									strText = strText & " damgd"
								End If
								
							End If
							rtbUnitList.SelectionStart = 99999
							rtbUnitList.SelectionLength = 0
							rtbUnitList.SelectedText = vbCrLf & strText
							
							If ListType = enumUnitDisplay.udSHIP Then
								
								'list cargo
								For X = 12 To 24
									If InStr("pln civ uw hcm lcm oil petrol rad iron dust bar ", rs.Fields(X).Name) > 0 Then
										If rs.Fields(X).Value > 0 Then
											strText = "  cargo: " & CStr(rs.Fields(X).Value) & " " & rs.Fields(X).Name
											rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, True)
											rtbUnitList.SelectionStart = 99999
											rtbUnitList.SelectionLength = 0
											rtbUnitList.SelectedText = vbCrLf & strText
											rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, False)
										End If
									End If
								Next X
								
								'check for lands on ship
								If rs.Fields("land").Value > 0 Then
									hldIndex = rsLand.Index
									rsLand.Index = "location"
									rsLand.Seek("=", rs.Fields("x"), rs.Fields("y"))
									If Not rsLand.NoMatch Then
land_loop: 
										If (Not rsLand.EOF) Then
											If rsLand.Fields("x").Value = rs.Fields("x").Value And rsLand.Fields("y").Value = rs.Fields("y").Value Then
												If rsLand.Fields("ship").Value = rs.Fields(0).Value Then
													strText = "  cargo: " + rsLand.Fields("type").Value + " #" + CStr(rsLand.Fields(0).Value)
													rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, True)
													rtbUnitList.SelectionStart = 99999
													rtbUnitList.SelectionLength = 0
													rtbUnitList.SelectedText = vbCrLf & strText
													rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, False)
												End If
												rsLand.MoveNext()
												GoTo land_loop
											End If
										End If
									End If
									rsLand.Index = hldIndex
								End If
								
								'check for planes on ship
								If rs.Fields("pln").Value > 0 Then
									hldIndex = rsPlane.Index
									rsPlane.Index = "location"
									rsPlane.Seek("=", rs.Fields("x"), rs.Fields("y"))
									If Not rsPlane.NoMatch Then
plane_loop: 
										If (Not rsPlane.EOF) Then
											If rsPlane.Fields("x").Value = rs.Fields("x").Value And rsPlane.Fields("y").Value = rs.Fields("y").Value Then
												If rsPlane.Fields("ship").Value = rs.Fields(0).Value Then
													strText = "  cargo: " + rsPlane.Fields("type").Value + " #" + CStr(rsPlane.Fields(0).Value)
													If rsPlane.Fields("nuke").Value <> "N/A" Then
														strText = strText & " armed: " + rsPlane.Fields("nuke").Value
													End If
													rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, True)
													rtbUnitList.SelectionStart = 99999
													rtbUnitList.SelectionLength = 0
													rtbUnitList.SelectedText = vbCrLf & strText
													rtbUnitList.SelectionFont = VB6.FontChangeItalic(rtbUnitList.SelectionFont, False)
												End If
												rsPlane.MoveNext()
												GoTo plane_loop
											End If
										End If
									End If
									rsPlane.Index = hldIndex
								End If
							End If
						End If
					End If
				End If
				rs.MoveNext()
			End While
			If Not NeedHeader Then
				rtbUnitList.SelectionStart = 99999
				rtbUnitList.SelectionLength = 0
				rtbUnitList.SelectedText = vbCrLf & vbCrLf
			End If
			
		Next ListType
		
		SetIndexSubCombo() '092503 rjk: Added to ensure the combo is reflective of the grid contents.
		
		rtbUnitList.Visible = True
	End Sub
	
	Sub SetCommandFocus(Optional ByRef KeyAscii As Short = 0)
		'check for current prompt
		If CommandCursorFocus Or PromptUp Then
			Exit Sub
		End If
		
		On Error Resume Next
		Me.txtEntry.Focus()
		
		' if hitting a key caused the command box to get focus, then
		' insert the key hit in the proper place
		If KeyAscii > 0 Then
			If KeyAscii = 8 Then 'backspace
				CommandCursorPos = CommandCursorPos - 1
				If CommandCursorPos < 1 Then
					CommandCursorPos = 1
				Else
					txtEntry.SelectionStart = CommandCursorPos
					txtEntry.SelectionLength = 1
					txtEntry.SelectedText = vbNullString
				End If
			Else
				txtEntry.SelectionStart = CommandCursorPos
				txtEntry.SelectedText = Chr(KeyAscii)
				CommandCursorPos = CommandCursorPos + 1
				txtEntry.SelectionStart = CommandCursorPos
			End If
		End If
	End Sub
	
	Public Sub MoveMap(ByRef sx As Short, ByRef sy As Short)
		CenterMap(picMap, sx, sy)
		MapValid = False 'suppress screen redraws
		vsMap.Value = OriginY
		hsMap.Value = OriginX
		MapValid = True
		DrawHexes()
		FillSectorBox(sx, sy)
	End Sub
	
	Public Function BuildMissionString(ByRef BuildType As String, ByRef id As Short) As String
		
		Dim n As Short
		Dim strTemp As String
		Dim strTemp2 As String
		
		rsMissions.Seek("=", BuildType, id)
		If Not rsMissions.NoMatch Then
			strTemp = "Mission: " + rsMissions.Fields("mission").Value + ", "
			If Len(rsMissions.Fields("op sector").Value) > 0 Then
				strTemp = strTemp & "Center: " & CStr(rsMissions.Fields("op sector").Value) & ",  "
			End If
			strTemp = strTemp & " Radius: " & CStr(rsMissions.Fields("radius").Value)
		End If
		
		'check orders if ship
		If BuildType = "s" Then
			rsShipOrders.Seek("=", id)
			If Not rsShipOrders.NoMatch Then
				strTemp = strTemp & "  Orders - dest: " + rsShipOrders.Fields("start sector").Value
				If Len(rsShipOrders.Fields("end sector").Value) > 0 Then
					strTemp = strTemp & " => " + rsShipOrders.Fields("end sector").Value
				End If
				For n = 1 To 6
					If Len(rsShipOrders.Fields("start cargo " & CStr(n)).Value) > 0 Then
						strTemp2 = strTemp2 & " S " + rsShipOrders.Fields("start cargo " & CStr(n)).Value + " " + CStr(rsShipOrders.Fields("start level " & CStr(n)).Value)
					End If
					If Len(rsShipOrders.Fields("end cargo " & CStr(n)).Value) > 0 Then
						strTemp2 = strTemp2 & " E " + rsShipOrders.Fields("end cargo " & CStr(n)).Value + " " + CStr(rsShipOrders.Fields("end level " & CStr(n)).Value) + " "
					End If
				Next n
				If Len(strTemp2) > 0 Then
					strTemp = strTemp & "  Holds " & strTemp2
				End If
			End If
		End If
		
		BuildMissionString = strTemp
	End Function
	
	Public Sub DisplayPromptHelp(ByRef strSubject As String)
		If frmOptions.bLocalHelp Then
			HtmlHelp(Handle.ToInt32, My.Application.Info.DirectoryPath & "/help/winace.chm", HH_DISPLAY_TOPIC, strSubject & ".html")
		Else
			PositionHelp = True
			frmEmpCmd.SubmitEmpireCommand("info " & LCase(strSubject), True)
		End If
	End Sub
	
	Private Sub BuildUnitFilters()
		Dim n As Short
		Dim bAdd As Boolean
		
		For n = enumUnitType.TYPE_START To enumUnitType.TYPE_END
			strType(n) = " "
		Next n
		strTypeTitle(enumUnitType.TYPE_START) = "Start - Should never see"
		strTypeTitle(enumUnitType.TYPE_ALL) = "All"
		strTypeTitle(enumUnitType.TYPE_SHIP_UNLOADED) = "Not on Ship"
		strTypeTitle(enumUnitType.TYPE_LAND_UNLOADED) = "Not on Land"
		strTypeTitle(enumUnitType.TYPE_SHIP_LOADED) = "Loaded on Ship"
		strTypeTitle(enumUnitType.TYPE_LAND_LOADED) = "Loaded on Land"
		'ship capabilities/abilities
		strTypeTitle(enumUnitType.TYPE_S_FISH) = "Fish"
		strTypeTitle(enumUnitType.TYPE_S_TORP) = "Torpedo"
		strTypeTitle(enumUnitType.TYPE_S_DCHRG) = "Depth Charges"
		strTypeTitle(enumUnitType.TYPE_S_PLANE) = "Planes"
		strTypeTitle(enumUnitType.TYPE_S_MISS) = "Missiles"
		strTypeTitle(enumUnitType.TYPE_S_OIL) = "Drill Oil"
		strTypeTitle(enumUnitType.TYPE_S_OILER) = "Haul Oil"
		strTypeTitle(enumUnitType.TYPE_S_SONAR) = "Sonar"
		strTypeTitle(enumUnitType.TYPE_S_MINE) = "Mine"
		strTypeTitle(enumUnitType.TYPE_S_SWEEP) = "Sweep"
		strTypeTitle(enumUnitType.TYPE_S_SUB) = "Submarine"
		strTypeTitle(enumUnitType.TYPE_S_SPY) = "Spy"
		strTypeTitle(enumUnitType.TYPE_S_LAND) = "Land"
		strTypeTitle(enumUnitType.TYPE_S_SEMI_LAND) = "Semi-Land"
		strTypeTitle(enumUnitType.TYPE_S_SUB_TORP) = "Torpedo a Sub"
		strTypeTitle(enumUnitType.TYPE_S_TRADE) = "Trade"
		strTypeTitle(enumUnitType.TYPE_S_CANAL) = "Canal"
		strTypeTitle(enumUnitType.TYPE_S_ANTI_MISSILE) = "Anti Missile"
		'ship commands
		strTypeTitle(enumUnitType.TYPE_SC_FIRE) = "fire command"
		'plane capabilities/abilities
		strTypeTitle(enumUnitType.TYPE_P_TACTICAL) = "Pinpoint (Tactical) Bombers"
		strTypeTitle(enumUnitType.TYPE_P_BOMBER) = "Strategic Bombers"
		strTypeTitle(enumUnitType.TYPE_P_INTERCEPT) = "Interceptors"
		strTypeTitle(enumUnitType.TYPE_P_CARGO) = "Cargo"
		strTypeTitle(enumUnitType.TYPE_P_VTOL) = "VTOL"
		strTypeTitle(enumUnitType.TYPE_P_MISSILE) = "Missile"
		strTypeTitle(enumUnitType.TYPE_P_LIGHT) = "Light"
		strTypeTitle(enumUnitType.TYPE_P_SPY) = "Recon."
		strTypeTitle(enumUnitType.TYPE_P_IMAGE) = "Image"
		strTypeTitle(enumUnitType.TYPE_P_SATELLITE) = "Satellite"
		strTypeTitle(enumUnitType.TYPE_P_STEALTH) = "Stealth"
		strTypeTitle(enumUnitType.TYPE_P_SDI) = "SDI"
		strTypeTitle(enumUnitType.TYPE_P_HALF_STEALTH) = "Half-Stealth"
		strTypeTitle(enumUnitType.TYPE_P_X_LIGHT) = "x-light"
		strTypeTitle(enumUnitType.TYPE_P_CHOPPER) = "Chopper"
		strTypeTitle(enumUnitType.TYPE_P_ANTISUB) = "Anti-Submarine"
		strTypeTitle(enumUnitType.TYPE_P_PARA) = "Paratroop"
		strTypeTitle(enumUnitType.TYPE_P_ESCORT) = "Escort"
		strTypeTitle(enumUnitType.TYPE_P_MINE) = "Mine"
		strTypeTitle(enumUnitType.TYPE_P_SWEEP) = "Sweep"
		strTypeTitle(enumUnitType.TYPE_P_MARINE) = "Marine Missile"
		'plane groups
		strTypeTitle(enumUnitType.TYPE_PG_ESCORTS) = "escorts (intercept and escort)"
		'plane commands
		strTypeTitle(enumUnitType.TYPE_PC_BOMB) = "bomb command"
		strTypeTitle(enumUnitType.TYPE_PC_LAUNCH) = "Launch command"
		strTypeTitle(enumUnitType.TYPE_PC_DROP) = "Drop command"
		'land units capabilities/abilities
		strTypeTitle(enumUnitType.TYPE_L_XLIGHT) = "x-light"
		strTypeTitle(enumUnitType.TYPE_L_ENGINEER) = "Engineer"
		strTypeTitle(enumUnitType.TYPE_L_SUPPLY) = "Supply"
		strTypeTitle(enumUnitType.TYPE_L_SECURITY) = "Security"
		strTypeTitle(enumUnitType.TYPE_L_LIGHT) = "Light"
		strTypeTitle(enumUnitType.TYPE_L_MARINE) = "Marine"
		strTypeTitle(enumUnitType.TYPE_L_RECON) = "Recon."
		strTypeTitle(enumUnitType.TYPE_L_RADAR) = "Radar"
		strTypeTitle(enumUnitType.TYPE_L_ASSAULT) = "Assault"
		strTypeTitle(enumUnitType.TYPE_L_TRAIN) = "Train"
		'land unit commands
		strTypeTitle(enumUnitType.TYPE_LC_FIRE) = "fire command"
		strTypeTitle(enumUnitType.TYPE_END) = "End - Should never see"
		
		If rsBuild.RecordCount = 0 Then
			Exit Sub
		End If
		rsBuild.MoveFirst()
		While Not (rsBuild.EOF)
			For n = 32 To 51
				Select Case rsBuild.Fields("type").Value
					Case "s"
						Select Case rsBuild.Fields(n).Value
							'ship capabilities and abilities
							Case "fish"
								strType(enumUnitType.TYPE_S_FISH) = strType(enumUnitType.TYPE_S_FISH) + rsBuild.Fields("id").Value + " "
							Case "torp"
								strType(enumUnitType.TYPE_S_TORP) = strType(enumUnitType.TYPE_S_TORP) + rsBuild.Fields("id").Value + " "
							Case "dchrg"
								strType(enumUnitType.TYPE_S_DCHRG) = strType(enumUnitType.TYPE_S_DCHRG) + rsBuild.Fields("id").Value + " "
							Case "plane"
								strType(enumUnitType.TYPE_S_PLANE) = strType(enumUnitType.TYPE_S_PLANE) + rsBuild.Fields("id").Value + " "
							Case "miss"
								strType(enumUnitType.TYPE_S_MISS) = strType(enumUnitType.TYPE_S_MISS) + rsBuild.Fields("id").Value + " "
							Case "oil"
								strType(enumUnitType.TYPE_S_OIL) = strType(enumUnitType.TYPE_S_OIL) + rsBuild.Fields("id").Value + " "
							Case "oiler"
								strType(enumUnitType.TYPE_S_OILER) = strType(enumUnitType.TYPE_S_OILER) + rsBuild.Fields("id").Value + " "
							Case "sonar"
								strType(enumUnitType.TYPE_S_SONAR) = strType(enumUnitType.TYPE_S_SONAR) + rsBuild.Fields("id").Value + " "
							Case "mine"
								strType(enumUnitType.TYPE_S_MINE) = strType(enumUnitType.TYPE_S_MINE) + rsBuild.Fields("id").Value + " "
							Case "sweep"
								strType(enumUnitType.TYPE_S_SWEEP) = strType(enumUnitType.TYPE_S_SWEEP) + rsBuild.Fields("id").Value + " "
							Case "sub"
								strType(enumUnitType.TYPE_S_SUB) = strType(enumUnitType.TYPE_S_SUB) + rsBuild.Fields("id").Value + " "
							Case "spy"
								strType(enumUnitType.TYPE_S_SPY) = strType(enumUnitType.TYPE_S_SPY) + rsBuild.Fields("id").Value + " "
							Case "land"
								strType(enumUnitType.TYPE_S_LAND) = strType(enumUnitType.TYPE_S_LAND) + rsBuild.Fields("id").Value + " "
							Case "semi-land"
								strType(enumUnitType.TYPE_S_SEMI_LAND) = strType(enumUnitType.TYPE_S_SEMI_LAND) + rsBuild.Fields("id").Value + " "
							Case "sub-torp"
								strType(enumUnitType.TYPE_S_SUB_TORP) = strType(enumUnitType.TYPE_S_SUB_TORP) + rsBuild.Fields("id").Value + " "
							Case "trade"
								strType(enumUnitType.TYPE_S_TRADE) = strType(enumUnitType.TYPE_S_TRADE) + rsBuild.Fields("id").Value + " "
							Case "canal"
								strType(enumUnitType.TYPE_S_CANAL) = strType(enumUnitType.TYPE_S_CANAL) + rsBuild.Fields("id").Value + " "
							Case "anti-missile"
								strType(enumUnitType.TYPE_S_ANTI_MISSILE) = strType(enumUnitType.TYPE_S_ANTI_MISSILE) + rsBuild.Fields("id").Value + " "
						End Select
					Case "p" 'plane capabilities and abilities
						Select Case rsBuild.Fields(n).Value
							Case "tactical"
								strType(enumUnitType.TYPE_P_TACTICAL) = strType(enumUnitType.TYPE_P_TACTICAL) + rsBuild.Fields("id").Value + " "
							Case "bomber"
								strType(enumUnitType.TYPE_P_BOMBER) = strType(enumUnitType.TYPE_P_BOMBER) + rsBuild.Fields("id").Value + " "
							Case "intercept"
								strType(enumUnitType.TYPE_P_INTERCEPT) = strType(enumUnitType.TYPE_P_INTERCEPT) + rsBuild.Fields("id").Value + " "
							Case "cargo"
								strType(enumUnitType.TYPE_P_CARGO) = strType(enumUnitType.TYPE_P_CARGO) + rsBuild.Fields("id").Value + " "
							Case "VTOL"
								strType(enumUnitType.TYPE_P_VTOL) = strType(enumUnitType.TYPE_P_VTOL) + rsBuild.Fields("id").Value + " "
							Case "missile"
								strType(enumUnitType.TYPE_P_MISSILE) = strType(enumUnitType.TYPE_P_MISSILE) + rsBuild.Fields("id").Value + " "
							Case "light"
								strType(enumUnitType.TYPE_P_LIGHT) = strType(enumUnitType.TYPE_P_LIGHT) + rsBuild.Fields("id").Value + " "
							Case "spy"
								strType(enumUnitType.TYPE_P_SPY) = strType(enumUnitType.TYPE_P_SPY) + rsBuild.Fields("id").Value + " "
							Case "image"
								strType(enumUnitType.TYPE_P_IMAGE) = strType(enumUnitType.TYPE_P_IMAGE) + rsBuild.Fields("id").Value + " "
							Case "satellite"
								strType(enumUnitType.TYPE_P_SATELLITE) = strType(enumUnitType.TYPE_P_SATELLITE) + rsBuild.Fields("id").Value + " "
							Case "stealth"
								strType(enumUnitType.TYPE_P_STEALTH) = strType(enumUnitType.TYPE_P_STEALTH) + rsBuild.Fields("id").Value + " "
							Case "SDI"
								strType(enumUnitType.TYPE_P_SDI) = strType(enumUnitType.TYPE_P_SDI) + rsBuild.Fields("id").Value + " "
							Case "half-stealth"
								strType(enumUnitType.TYPE_P_HALF_STEALTH) = strType(enumUnitType.TYPE_P_HALF_STEALTH) + rsBuild.Fields("id").Value + " "
							Case "x-light"
								strType(enumUnitType.TYPE_P_X_LIGHT) = strType(enumUnitType.TYPE_P_X_LIGHT) + rsBuild.Fields("id").Value + " "
							Case "helo"
								strType(enumUnitType.TYPE_P_CHOPPER) = strType(enumUnitType.TYPE_P_CHOPPER) + rsBuild.Fields("id").Value + " "
							Case "ASW"
								strType(enumUnitType.TYPE_P_ANTISUB) = strType(enumUnitType.TYPE_P_ANTISUB) + rsBuild.Fields("id").Value + " "
							Case "para"
								strType(enumUnitType.TYPE_P_PARA) = strType(enumUnitType.TYPE_P_PARA) + rsBuild.Fields("id").Value + " "
							Case "escort"
								strType(enumUnitType.TYPE_P_ESCORT) = strType(enumUnitType.TYPE_P_ESCORT) + rsBuild.Fields("id").Value + " "
							Case "mine"
								strType(enumUnitType.TYPE_P_MINE) = strType(enumUnitType.TYPE_P_MINE) + rsBuild.Fields("id").Value + " "
							Case "sweep"
								strType(enumUnitType.TYPE_P_SWEEP) = strType(enumUnitType.TYPE_P_SWEEP) + rsBuild.Fields("id").Value + " "
							Case "marine"
								strType(enumUnitType.TYPE_P_MARINE) = strType(enumUnitType.TYPE_P_MARINE) + rsBuild.Fields("id").Value + " "
						End Select
					Case "l" 'land unit capabilities
						Select Case rsBuild.Fields(n).Value
							Case "xlight"
								strType(enumUnitType.TYPE_L_XLIGHT) = strType(enumUnitType.TYPE_L_XLIGHT) + rsBuild.Fields("id").Value + " "
							Case "engineer"
								strType(enumUnitType.TYPE_L_ENGINEER) = strType(enumUnitType.TYPE_L_ENGINEER) + rsBuild.Fields("id").Value + " "
							Case "supply"
								strType(enumUnitType.TYPE_L_SUPPLY) = strType(enumUnitType.TYPE_L_SUPPLY) + rsBuild.Fields("id").Value + " "
							Case "security"
								strType(enumUnitType.TYPE_L_SECURITY) = strType(enumUnitType.TYPE_L_SECURITY) + rsBuild.Fields("id").Value + " "
							Case "LIGHT"
								strType(enumUnitType.TYPE_L_LIGHT) = strType(enumUnitType.TYPE_L_LIGHT) + rsBuild.Fields("id").Value + " "
							Case "marine"
								strType(enumUnitType.TYPE_L_MARINE) = strType(enumUnitType.TYPE_L_MARINE) + rsBuild.Fields("id").Value + " "
							Case "recon"
								strType(enumUnitType.TYPE_L_RECON) = strType(enumUnitType.TYPE_L_RECON) + rsBuild.Fields("id").Value + " "
							Case "radar"
								strType(enumUnitType.TYPE_L_RADAR) = strType(enumUnitType.TYPE_L_RADAR) + rsBuild.Fields("id").Value + " "
							Case "assault"
								strType(enumUnitType.TYPE_L_ASSAULT) = strType(enumUnitType.TYPE_L_ASSAULT) + rsBuild.Fields("id").Value + " "
							Case "train"
								strType(enumUnitType.TYPE_L_TRAIN) = strType(enumUnitType.TYPE_L_TRAIN) + rsBuild.Fields("id").Value + " "
						End Select
				End Select
			Next n
			rsBuild.MoveNext()
		End While
		
		rsBuild.MoveFirst()
		Do While Not (rsBuild.EOF)
			Select Case rsBuild.Fields("type").Value
				Case "s" 'Ships
					If rsBuild.Fields(17).Value > 0 Then 'the number of guns can fire
						strType(enumUnitType.TYPE_SC_FIRE) = strType(enumUnitType.TYPE_SC_FIRE) + rsBuild.Fields("id").Value + " "
					End If
				Case "p" 'Planes
					bAdd = True
					If InStr(strType(enumUnitType.TYPE_P_INTERCEPT), rsBuild.Fields("id").Value) > 0 Or InStr(strType(enumUnitType.TYPE_P_ESCORT), rsBuild.Fields("id").Value) > 0 Then
						strType(enumUnitType.TYPE_PG_ESCORTS) = strType(enumUnitType.TYPE_PG_ESCORTS) + rsBuild.Fields("id").Value + " "
					End If
					If InStr(strType(enumUnitType.TYPE_P_MISSILE), rsBuild.Fields("id").Value) > 0 Or InStr(strType(enumUnitType.TYPE_P_SATELLITE), rsBuild.Fields("id").Value) > 0 Then
						strType(enumUnitType.TYPE_PC_LAUNCH) = strType(enumUnitType.TYPE_PC_LAUNCH) + rsBuild.Fields("id").Value + " "
					End If
					If InStr(strType(enumUnitType.TYPE_P_CARGO), rsBuild.Fields("id").Value) > 0 Or InStr(strType(enumUnitType.TYPE_P_MINE), rsBuild.Fields("id").Value) > 0 Then
						strType(enumUnitType.TYPE_PC_DROP) = strType(enumUnitType.TYPE_PC_DROP) + rsBuild.Fields("id").Value + " "
					End If
					If InStr(strType(enumUnitType.TYPE_P_CARGO), rsBuild.Fields("id").Value) = 0 Or InStr(strType(enumUnitType.TYPE_P_TACTICAL), rsBuild.Fields("id").Value) > 0 Then
						strType(enumUnitType.TYPE_PC_BOMB) = strType(enumUnitType.TYPE_PC_BOMB) + rsBuild.Fields("id").Value + " "
					End If
				Case "l" 'Land Units
					If rsBuild.Fields(21).Value > 0 Then 'the number of guns can fire
						strType(enumUnitType.TYPE_LC_FIRE) = strType(enumUnitType.TYPE_LC_FIRE) + rsBuild.Fields("id").Value + " "
					End If
			End Select
			rsBuild.MoveNext()
		Loop 
	End Sub
	
	Public Sub SetUnitDisplay(ByRef UnitDisplay As enumUnitDisplay, ByRef UnitFilter As enumUnitType, ByRef bMob As Boolean, ByRef bEff As Boolean, ByRef bRange As Boolean, Optional ByRef strTarget As String = "")
		Static LastUnitFilter As enumUnitType
		Static bLastRange As Boolean
		Static bLastMob As Boolean
		Static bLastEff As Boolean
		Static strLastTarget As String
		Dim tstSectX As Short
		Dim tstSectY As Short
		
		If TabStrip1.Tabs(TabStrip1.SelectedItem.Index).key <> "Unit" Or LastUnitDisplay <> UnitDisplay Or bLastRange <> bRange Or LastUnitFilter <> UnitFilter Or bLastMob <> bMob Or bLastEff <> bEff Then
			'Mob Button
			NeedMob = bMob
			If bMob Then
				'UPGRADE_WARNING: Lower bound of collection Toolbar1.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				CType(Toolbar1.Items.Item(7), ToolStripButton).Checked = True
			Else
				'UPGRADE_WARNING: Lower bound of collection Toolbar1.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				CType(Toolbar1.Items.Item(7), ToolStripButton).Checked = False
			End If
			'Eff Button
			Needeff = bEff
			If bEff Then
				'UPGRADE_WARNING: Lower bound of collection Toolbar1.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				CType(Toolbar1.Items.Item(8), ToolStripButton).Checked = True
			Else
				'UPGRADE_WARNING: Lower bound of collection Toolbar1.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				CType(Toolbar1.Items.Item(8), ToolStripButton).Checked = False
			End If
			
			SubType = UnitFilter
			
			DisplaySubset = enumSubset.ssSECT
			
			Select Case UnitDisplay
				Case enumUnitDisplay.udSHIP
					DisplaySelect = enumUnitDisplay.udSHIP
				Case enumUnitDisplay.udLAND
					If bRange Then
						If strLastTarget <> strTarget Then
							If ParseSectors(tstSectX, tstSectY, strTarget) Then
								DisplaySelect = enumUnitDisplay.udLAND
								DisplaySubset = enumSubset.ssATTACK_RANGE
							End If
							strTarget = strLastTarget
						End If
					Else
						DisplaySelect = enumUnitDisplay.udLAND
					End If
				Case enumUnitDisplay.udPLANE
					If bRange Then
						If strLastTarget <> strTarget Then
							If ParseSectors(tstSectX, tstSectY, strTarget) Then
								DisplaySelect = enumUnitDisplay.udPLANE
								If UnitFilter = enumUnitType.TYPE_P_MISSILE Then
									DisplaySubset = enumSubset.ssMISSILE_RANGE
								Else
									DisplaySubset = enumSubset.ssPLANE_RANGE
								End If
							End If
							strTarget = strLastTarget
						End If
					Else
						DisplaySelect = enumUnitDisplay.udPLANE
					End If
				Case enumUnitDisplay.udNUKE
					DisplaySelect = enumUnitDisplay.udNUKE
				Case enumUnitDisplay.udENEMY
					DisplaySelect = enumUnitDisplay.udENEMY
				Case enumUnitDisplay.udList
					DisplaySelect = enumUnitDisplay.udList
			End Select
			
			FillGrid()
			DisplayUnitBox()
			If bRange Then
				SortGrid(3)
			Else
				SortGrid(1)
			End If
			'Save current system system (not user settings)
			LastUnitDisplay = UnitDisplay
			bLastRange = bRange
			LastUnitFilter = UnitFilter
			bLastMob = bMob
			bLastEff = bEff
		End If
	End Sub
	
	Public Function UnitCharacteristicCheck(ByRef UnitType As enumUnitType, ByRef strUnit As String) As Boolean
		If InStr(strType(UnitType), " " & strUnit & " ") > 0 Then
			UnitCharacteristicCheck = True
		Else
			UnitCharacteristicCheck = False
		End If
	End Function
	
	Public Sub SetPlayerTimeInterval()
		If frmOptions.playersTimeInterval = 0 Then
			PlayersTimer.Enabled = False
		Else
			PlayersTimer.Interval = 60000
			PlayersTimer.Enabled = True
		End If
	End Sub
	
	Private Sub DrawEvent(ByRef picMap As Object, Optional ByRef secx As Short = 0, Optional ByRef secy As Short = -1)
		Dim oldWidth As Object
		Dim oldcolor As Object
		Dim natColor As Short
		Dim iIndex As Short
		Dim bSectorCheck As Boolean
		Dim PosX As Single
		Dim PosY As Single
		
		
		If secx = 0 And secy = -1 Then
			bSectorCheck = False
		Else
			bSectorCheck = True
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object picMap.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldWidth = picMap.DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object picMap.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = 3 ' Set DrawWidth
		'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oldcolor = picMap.ForeColor
		For iIndex = 1 To EventMarkers.Count
			If (EventMarkers.X(iIndex) = secx And EventMarkers.Y(iIndex) = secy) Or Not bSectorCheck Then
				natColor = EventMarkers.Nation(iIndex)
				If natColor >= LBound(PlayerColors) And natColor <= UBound(PlayerColors) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.ForeColor = PlayerColors(natColor)
				ElseIf natColor = -2 Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.ForeColor = BasicText(2) 'water text color
				ElseIf natColor = -3 Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.ForeColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.ForeColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
				End If
				Coord(EventMarkers.X(iIndex), EventMarkers.Y(iIndex), PosX, PosY)
				DrawHex(picMap, PosX, PosY)
				If bSectorCheck Then
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.ForeColor = oldcolor
					'UPGRADE_WARNING: Couldn't resolve default property of object picMap.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					picMap.DrawWidth = oldWidth ' Set DrawWidth
					Exit Sub
				End If
			End If
		Next iIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object picMap.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldcolor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.ForeColor = oldcolor
		'UPGRADE_WARNING: Couldn't resolve default property of object picMap.DrawWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object oldWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		picMap.DrawWidth = oldWidth ' Set DrawWidth
	End Sub
	
	Private Sub ExecuteFile(ByRef iIndex As Short)
		Dim sx As Short
		Dim sy As Short
		Dim strCommand As String
		Dim fs As Object
		Dim fileObject As Object
		If tScriptButtons(iIndex).bJumpPoint Then
			
			rsNation.MoveFirst()
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsNation.Fields("JumpPoint" & CStr(tScriptButtons(iIndex).iJumpPoint + 1)).Value) And tScriptButtons(iIndex).iJumpPoint >= 0 Then
				If ParseSectors(sx, sy, rsNation.Fields("JumpPoint" & CStr(tScriptButtons(iIndex).iJumpPoint + 1)).Value) Then
					MoveMap(sx, sy)
				End If
			End If
		Else
			
			If Len(tScriptButtons(iIndex).strFileName) = 0 Then
				Exit Sub
			End If
			If Not tScriptButtons(iIndex).bSendOutputReportWindow Then
				frmEmpCmd.SubmitEmpireCommand("bf1", False)
			Else
				frmEmpCmd.SubmitEmpireCommand("br1", False)
			End If
			On Error GoTo Empire_Error
			fs = CreateObject("Scripting.FileSystemObject")
			'UPGRADE_WARNING: Couldn't resolve default property of object fs.OpenTextFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fileObject = fs.OpenTextFile(tScriptButtons(iIndex).strFileName)
			'UPGRADE_WARNING: Couldn't resolve default property of object fileObject.AtEndOfStream. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			While fileObject.AtEndOfStream <> True
				'UPGRADE_WARNING: Couldn't resolve default property of object fileObject.readline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCommand = fileObject.readline
				frmEmpCmd.SubmitEmpireCommand(CStr(strCommand), False)
			End While
			'UPGRADE_WARNING: Couldn't resolve default property of object fileObject.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fileObject.Close()
			
			'filenumber = FreeFile
			'On Error GoTo Empire_Error
			'Open tScriptButtons(iIndex).strFileName For Input As #filenumber
			'While Not EOF(filenumber)
			'    Input #filenumber, vCommand
			'    frmEmpCmd.SubmitEmpireCommand CStr(vCommand), False
			'Wend
			'Close #filenumber
			
			If Not tScriptButtons(iIndex).bSendOutputReportWindow Then
				frmEmpCmd.SubmitEmpireCommand("bf2", False)
			Else
				frmEmpCmd.SubmitEmpireCommand("br2", False)
			End If
		End If
		Exit Sub
		
Empire_Error: 
		EmpireError("Failed to run script file for custom script button", CStr(iIndex), tScriptButtons(iIndex).strFileName)
	End Sub
	Private Sub vsMap_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vsMap.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vsMap_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub hsMap_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles hsMap.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				hsMap_Change(eventArgs.newValue)
		End Select
	End Sub
End Class