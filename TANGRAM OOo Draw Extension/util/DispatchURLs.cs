using System;

namespace tud.mci.tangram.util
{
    /*
     * The following classes list all available commands that are accessible through 
     * the GUI. The command names are used with the ".uno:" protocol scheme to 
     * define framework commands. These command URLs can be used to dispatch/execute, 
     * like ".uno:About" shows the about box. 
     * Abbreviations:
     *      A = accelerator
     *      M = menu
     *      T = toolbox
     *      S = status bar
     * */

    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands"/>
    /// </summary>
    public static class DispatchURLs
    {
        ///<summary>Media Player</summary>
        public const String SID_AVMEDIA_PLAYER = ".uno:AVMediaPlayer"; //AMT - 6694
        ///<summary>About OpenOffice.org</summary>
        public const String SID_ABOUT = ".uno:About"; //AMT - 5301
        ///<summary>Absolute Record</summary>
        public const String SID_FM_RECORD_ABSOLUTE = ".uno:AbsoluteRecord"; //AMT - 10622
        ///<summary>Extended Tips</summary>
        public const String SID_HELPBALLOONS = ".uno:ActiveHelp"; //AMT - 5403
        ///<summary>Date Field</summary>
        public const String SID_INSERT_DATEFIELD = ".uno:AddDateField"; //AMT - 10936
        ///<summary>New</summary>
        public const String SID_NEWDOCDIRECT = ".uno:AddDirect"; //AMT - 5537
        ///<summary>Add Field</summary>
        public const String SID_FM_ADD_FIELD = ".uno:AddField"; //AMT - 10623
        ///<summary>Add Table</summary>
        public const String SID_FM_ADDTABLE = ".uno:AddTable"; //AMT - 10722
        ///<summary>Enable Watch</summary>
        public const String SID_BASICIDE_ADDWATCH = ".uno:AddWatch"; //AMT - 30769
        ///<summary>Address Book Source</summary>
        public const String SID_TEMPLATE_ADDRESSBOKSOURCE = ".uno:AddressBookSource"; //AMT - 6655
        ///<summary>Centered</summary>
        public const String SID_OBJECT_ALIGN_CENTER = ".uno:AlignCenter"; //AMT - 10132
        ///<summary>Bottom</summary>
        public const String SID_OBJECT_ALIGN_DOWN = ".uno:AlignDown"; //AMT - 10136
        ///<summary>Center</summary>
        public const String SID_OBJECT_ALIGN_MIDDLE = ".uno:AlignMiddle"; //AMT - 10135
        ///<summary>Top</summary>
        public const String SID_OBJECT_ALIGN_UP = ".uno:AlignUp"; //AMT - 10134
        ///<summary>Arc</summary>
        public const String SID_DRAW_ARC = ".uno:Arc"; //AMT - 10114
        ///<summary>Block Arrows</summary>
        public const String SID_DRAWTBX_CS_ARROW = ".uno:ArrowShapes"; //AMT - 11049
        ///<summary>Automatic Control Focus</summary>
        public const String SID_FM_AUTOCONTROLFOCUS = ".uno:AutoControlFocus"; //AMT - 10763
        ///<summary>AutoCorrect</summary>
        public const String SID_AUTO_CORRECT_DLG = ".uno:AutoCorrectDlg"; //AMT - 10424
        ///<summary>AutoFilter</summary>
        public const String SID_FM_AUTOFILTER = ".uno:AutoFilter"; //AMT - 10716
        ///<summary>AutoFormat</summary>
        public const String SID_AUTOFORMAT = ".uno:AutoFormat"; //AMT - 10242
        ///<summary>AutoPilot: Address Data Source</summary>
        public const String SID_ADDRESS_DATA_SOURCE = ".uno:AutoPilotAddressDataSource"; //AMT - 10934
        ///<summary>Wizards</summary>
        public const String SID_AUTOPILOTMENU = ".uno:AutoPilotMenu"; //AMT - 6381
        ///<summary>AutoPilot: Presentation</summary>
        public const String SID_SD_AUTOPILOT = ".uno:AutoPilotPresentations"; //AMT - 10425
        ///<summary>Toolbars</summary>
        public const String SID_AVAILABLE_TOOLBARS = ".uno:AvailableToolbars"; //AMT - 6698
        ///<summary>Background Color</summary>
        public const String SID_BACKGROUND_COLOR = ".uno:BackgroundColor"; //AMT - 10185
        ///<summary>Background Pattern</summary>
        public const String SID_BACKGROUND_PATTERN = ".uno:BackgroundPatternController"; //AMT - 10186
        ///<summary>Interrupt Macro</summary>
        public const String SID_BASICBREAK = ".uno:BasicBreak"; //AMT - 6521
        ///<summary>Edit Macros</summary>
        public const String SID_BASICIDE_APPEAR = ".uno:BasicIDEAppear"; //AMT - 30783
        ///<summary>Basic Shapes</summary>
        public const String SID_DRAWTBX_CS_BASIC = ".uno:BasicShapes"; //AMT - 11047
        ///<summary>Step Into</summary>
        public const String SID_BASICSTEPINTO = ".uno:BasicStepInto"; //AMT - 5956
        ///<summary>Step Out</summary>
        public const String SID_BASICSTEPOUT = ".uno:BasicStepOut"; //AMT - 5963
        ///<summary>Step Over</summary>
        public const String SID_BASICSTEPOVER = ".uno:BasicStepOver"; //AMT - 5957
        ///<summary>Stop Macro</summary>
        public const String SID_BASICSTOP = ".uno:BasicStop"; //AMT - 5958
        ///<summary>Close Bezier</summary>
        public const String SID_BEZIER_CLOSE = ".uno:BezierClose"; //AMT - 10122
        ///<summary>Convert to Curve</summary>
        public const String SID_BEZIER_CONVERT = ".uno:BezierConvert"; //AMT - 27065
        ///<summary>Split Curve</summary>
        public const String SID_BEZIER_CUTLINE = ".uno:BezierCutLine"; //AMT - 10127
        ///<summary>Delete Points</summary>
        public const String SID_BEZIER_DELETE = ".uno:BezierDelete"; //AMT - 10120
        ///<summary>Corner Point</summary>
        public const String SID_BEZIER_EDGE = ".uno:BezierEdge"; //AMT - 27066
        ///<summary>Eliminate Points</summary>
        public const String SID_BEZIER_ELIMINATE_POINTS = ".uno:BezierEliminatePoints"; //AMT - 27030
        ///<summary>Curve, Filled</summary>
        public const String SID_DRAW_BEZIER_FILL = ".uno:BezierFill"; //AMT - 10118
        ///<summary>Insert Points</summary>
        public const String SID_BEZIER_INSERT = ".uno:BezierInsert"; //AMT - 10119
        ///<summary>Move Points</summary>
        public const String SID_BEZIER_MOVE = ".uno:BezierMove"; //AMT - 10121
        ///<summary>Smooth Transition</summary>
        public const String SID_BEZIER_SMOOTH = ".uno:BezierSmooth"; //AMT - 10123
        ///<summary>Symmetric Transition</summary>
        public const String SID_BEZIER_SYMMTR = ".uno:BezierSymmetric"; //AMT - 27067
        ///<summary>Curve</summary>
        public const String SID_DRAW_BEZIER_NOFILL = ".uno:Bezier_Unfilled"; //AMT - 10397
        ///<summary>Bibliography Database</summary>
        public const String SID_COMP_BIBLIOGRAPHY = ".uno:BibliographyComponent"; //AMT - 10880
        ///<summary>Eyedropper</summary>
        public const String SID_BMPMASK = ".uno:BmpMask"; //AMT - 10350
        ///<summary>Bold</summary>
        public const String SID_ATTR_CHAR_WEIGHT = ".uno:Bold"; //AMT - 10009
        ///<summary>Bring to Front</summary>
        public const String SID_FRAME_TO_TOP = ".uno:BringToFront"; //AMT - 10286
        ///<summary>Web Layout</summary>
        public const String SID_BROWSER_MODE = ".uno:BrowseView"; //AMT - 6313
        ///<summary>Callouts</summary>
        public const String SID_DRAWTBX_CS_CALLOUT = ".uno:CalloutShapes"; //AMT - 11051
        ///<summary>Centered</summary>
        public const String SID_ATTR_PARA_ADJUST_CENTER = ".uno:CenterPara"; //AMT - 10030
        ///<summary>Full-width</summary>
        public const String SID_TRANSLITERATE_FULLWIDTH = ".uno:ChangeCaseToFullWidth"; //AMT - 10915
        ///<summary>Half-width</summary>
        public const String SID_TRANSLITERATE_HALFWIDTH = ".uno:ChangeCaseToHalfWidth"; //AMT - 10914
        ///<summary>Hiragana</summary>
        public const String SID_TRANSLITERATE_HIRAGANA = ".uno:ChangeCaseToHiragana"; //AMT - 10916
        ///<summary>Katakana</summary>
        public const String SID_TRANSLITERATE_KATAGANA = ".uno:ChangeCaseToKatakana"; //AMT - 10917
        ///<summary>Lowercase</summary>
        public const String SID_TRANSLITERATE_LOWER = ".uno:ChangeCaseToLower"; //AMT - 10913
        ///<summary>Uppercase</summary>
        public const String SID_TRANSLITERATE_UPPER = ".uno:ChangeCaseToUpper"; //AMT - 10912
        ///<summary>Font Name</summary>
        public const String SID_ATTR_CHAR_FONT = ".uno:CharFontName"; //AMT - 10007
        ///<summary>Check Box</summary>
        public const String SID_FM_CHECKBOX = ".uno:CheckBox"; //AMT - 10596
        ///<summary>Check Box</summary>
        public const String SID_INSERT_CHECKBOX = ".uno:Checkbox"; //AMT - 10148
        ///<summary>Chinese translation</summary>
        public const String SID_CHINESE_CONVERSION = ".uno:ChineseConversion"; //AMT - 11016
        ///<summary>Insert Controls</summary>
        public const String SID_CHOOSE_CONTROLS = ".uno:ChooseControls"; //AMT - 10144
        ///<summary>Select Macro</summary>
        public const String SID_BASICIDE_CHOOSEMACRO = ".uno:ChooseMacro"; //AMT - 30770
        ///<summary>Symbol Selection</summary>
        public const String SID_CHOOSE_POLYGON = ".uno:ChoosePolygon"; //AMT - 10162
        ///<summary>Circle</summary>
        public const String SID_DRAW_CIRCLE = ".uno:Circle"; //AMT - 10385
        ///<summary>Circle Arc</summary>
        public const String SID_DRAW_CIRCLEARC = ".uno:CircleArc"; //AMT - 10390
        ///<summary>Circle Segment</summary>
        public const String SID_DRAW_CIRCLECUT = ".uno:CircleCut"; //AMT - 10115
        ///<summary>Circle Segment, Unfilled</summary>
        public const String SID_DRAW_CIRCLECUT_NOFILL = ".uno:CircleCut_Unfilled"; //AMT - 10391
        ///<summary>Circle Pie</summary>
        public const String SID_DRAW_CIRCLEPIE = ".uno:CirclePie"; //AMT - 10388
        ///<summary>Circle Pie, Unfilled</summary>
        public const String SID_DRAW_CIRCLEPIE_NOFILL = ".uno:CirclePie_Unfilled"; //AMT - 10389
        ///<summary>Circle, Unfilled</summary>
        public const String SID_DRAW_CIRCLE_NOFILL = ".uno:Circle_Unfilled"; //AMT - 10386
        ///<summary>Delete History</summary>
        public const String SID_CLEARHISTORY = ".uno:ClearHistory"; //AMT - 5703
        ///<summary>Remove</summary>
        public const String SID_OUTLINE_DELETEALL = ".uno:ClearOutline"; //AMT - 10234
        ///<summary>Close</summary>
        public const String SID_CLOSEDOC = ".uno:CloseDoc"; //AMT - 5503
        ///<summary>Close Window</summary>
        public const String SID_CLOSEWIN = ".uno:CloseWin"; //AMT - 5621
        ///<summary>Font Color</summary>
        public const String SID_ATTR_CHAR_COLOR = ".uno:Color"; //AMT - 10017
        ///<summary>Color Bar</summary>
        public const String SID_COLOR_CONTROL = ".uno:ColorControl"; //AMT - 10417
        ///<summary>Color</summary>
        public const String SID_COLOR_SETTINGS = ".uno:ColorSettings"; //AMT - 11044
        ///<summary>Combo Box</summary>
        public const String SID_FM_COMBOBOX = ".uno:ComboBox"; //AMT - 10601
        ///<summary>Combo Box</summary>
        public const String SID_INSERT_COMBOBOX = ".uno:Combobox"; //AMT - 10192
        ///<summary>Bottom</summary>
        public const String SID_ALIGN_ANY_BOTTOM = ".uno:CommonAlignBottom"; //AMT - 11008
        ///<summary>Centered</summary>
        public const String SID_ALIGN_ANY_HCENTER = ".uno:CommonAlignHorizontalCenter"; //AMT - 11003
        ///<summary>Default</summary>
        public const String SID_ALIGN_ANY_HDEFAULT = ".uno:CommonAlignHorizontalDefault"; //AMT - 11009
        ///<summary>Justified</summary>
        public const String SID_ALIGN_ANY_JUSTIFIED = ".uno:CommonAlignJustified"; //AMT - 11005
        ///<summary>Left</summary>
        public const String SID_ALIGN_ANY_LEFT = ".uno:CommonAlignLeft"; //AMT - 11002
        ///<summary>Right</summary>
        public const String SID_ALIGN_ANY_RIGHT = ".uno:CommonAlignRight"; //AMT - 11004
        ///<summary>Top</summary>
        public const String SID_ALIGN_ANY_TOP = ".uno:CommonAlignTop"; //AMT - 11006
        ///<summary>Center</summary>
        public const String SID_ALIGN_ANY_VCENTER = ".uno:CommonAlignVerticalCenter"; //AMT - 11007
        ///<summary>Default</summary>
        public const String SID_ALIGN_ANY_VDEFAULT = ".uno:CommonAlignVerticalDefault"; //AMT - 11010
        ///<summary>Compare Document</summary>
        public const String SID_DOCUMENT_COMPARE = ".uno:CompareDocuments"; //AMT - 6586
        ///<summary>Compile</summary>
        public const String SID_BASICCOMPILE = ".uno:CompileBasic"; //AMT - 5954
        ///<summary>Controls</summary>
        public const String SID_FM_CONFIG = ".uno:Config"; //AMT - 10593
        ///<summary>Customize</summary>
        public const String SID_CONFIG = ".uno:ConfigureDialog"; //AMT - 5904
        ///<summary>Current Context</summary>
        public const String SID_CONTEXT = ".uno:Context"; //S - 5310
        ///<summary>Edit Contour</summary>
        public const String SID_CONTOUR_DLG = ".uno:ContourDialog"; //AMT - 10334
        ///<summary>Control</summary>
        public const String SID_FM_CTL_PROPERTIES = ".uno:ControlProperties"; //AMT - 10613
        ///<summary>Replace with Button</summary>
        public const String SID_FM_CONVERTTO_BUTTON = ".uno:ConvertToButton"; //AMT - 10735
        ///<summary>Replace with Check Box</summary>
        public const String SID_FM_CONVERTTO_CHECKBOX = ".uno:ConvertToCheckBox"; //AMT - 10738
        ///<summary>Replace with Combo Box</summary>
        public const String SID_FM_CONVERTTO_COMBOBOX = ".uno:ConvertToCombo"; //AMT - 10741
        ///<summary>Replace with Currency Field</summary>
        public const String SID_FM_CONVERTTO_CURRENCY = ".uno:ConvertToCurrency"; //AMT - 10748
        ///<summary>Replace with Date Field</summary>
        public const String SID_FM_CONVERTTO_DATE = ".uno:ConvertToDate"; //AMT - 10745
        ///<summary>Replace with Text Box</summary>
        public const String SID_FM_CONVERTTO_EDIT = ".uno:ConvertToEdit"; //AMT - 10734
        ///<summary>Replace with File Selection</summary>
        public const String SID_FM_CONVERTTO_FILECONTROL = ".uno:ConvertToFileControl"; //AMT - 10744
        ///<summary>Replace with Label Field</summary>
        public const String SID_FM_CONVERTTO_FIXEDTEXT = ".uno:ConvertToFixed"; //AMT - 10736
        ///<summary>Replace with Formatted Field</summary>
        public const String SID_FM_CONVERTTO_FORMATTED = ".uno:ConvertToFormatted"; //AMT - 10751
        ///<summary>Replace with Group Box</summary>
        public const String SID_FM_CONVERTTO_GROUPBOX = ".uno:ConvertToGroup"; //AMT - 10740
        ///<summary>Replace with Image Button</summary>
        public const String SID_FM_CONVERTTO_IMAGEBUTTON = ".uno:ConvertToImageBtn"; //AMT - 10743
        ///<summary>Replace with Image Control</summary>
        public const String SID_FM_CONVERTTO_IMAGECONTROL = ".uno:ConvertToImageControl"; //AMT - 10750
        ///<summary>Replace with List Box</summary>
        public const String SID_FM_CONVERTTO_LISTBOX = ".uno:ConvertToList"; //AMT - 10737
        ///<summary>Replace with Navigation Bar</summary>
        public const String SID_FM_CONVERTTO_NAVIGATIONBAR = ".uno:ConvertToNavigationBar"; //AMT - 10772
        ///<summary>Replace with Numerical Field</summary>
        public const String SID_FM_CONVERTTO_NUMERIC = ".uno:ConvertToNumeric"; //AMT - 10747
        ///<summary>Replace with Pattern Field</summary>
        public const String SID_FM_CONVERTTO_PATTERN = ".uno:ConvertToPattern"; //AMT - 10749
        ///<summary>Replace with Radio Button</summary>
        public const String SID_FM_CONVERTTO_RADIOBUTTON = ".uno:ConvertToRadio"; //AMT - 10739
        ///<summary>Replace with Scrollbar</summary>
        public const String SID_FM_CONVERTTO_SCROLLBAR = ".uno:ConvertToScrollBar"; //AMT - 10770
        ///<summary>Replace with Spin Button</summary>
        public const String SID_FM_CONVERTTO_SPINBUTTON = ".uno:ConvertToSpinButton"; //AMT - 10771
        ///<summary>Replace with Time Field</summary>
        public const String SID_FM_CONVERTTO_TIME = ".uno:ConvertToTime"; //AMT - 10746
        ///<summary>Copy</summary>
        public const String SID_COPY = ".uno:Copy"; //AMT - 5711
        ///<summary>Records</summary>
        public const String SID_FM_COUNTALL = ".uno:CountAll"; //AMT - 10717
        ///<summary>Currency Field</summary>
        public const String SID_FM_CURRENCYFIELD = ".uno:CurrencyField"; //AMT - 10707
        ///<summary>Current Date</summary>
        public const String SID_CURRENTDATE = ".uno:CurrentDate"; //S - 5312
        ///<summary>Current Time</summary>
        public const String SID_CURRENTTIME = ".uno:CurrentTime"; //S - 5311
        ///<summary>Cut</summary>
        public const String SID_CUT = ".uno:Cut"; //AMT - 5710
        ///<summary>Explorer On/Off</summary>
        public const String SID_DSBROWSER_EXPLORER = ".uno:DSBrowserExplorer"; //AMT - 10764
        ///<summary>Date Field</summary>
        public const String SID_FM_DATEFIELD = ".uno:DateField"; //AMT - 10704
        ///<summary>Decrease Indent</summary>
        public const String SID_DEC_INDENT = ".uno:DecrementIndent"; //AMT - 10461
        ///<summary>Bullets On/Off</summary>
        public const String FN_NUM_BULLET_ON = ".uno:DefaultBullet"; //AMT - 20138
        ///<summary>Numbering On/Off</summary>
        public const String FN_NUM_NUMBERING_ON = ".uno:DefaultNumbering"; //AMT - 20144
        ///<summary>Delete Contents</summary>
        public const String SID_DELETE = ".uno:Delete"; //AMT - 5713
        ///<summary>Delete Frame</summary>
        public const String SID_DELETE_FRAME = ".uno:DeleteFrame"; //AMT - 5652
        ///<summary>Delete Record</summary>
        public const String SID_FM_RECORD_DELETE = ".uno:DeleteRecord"; //AMT - 10621
        ///<summary>Styles and Formatting</summary>
        public const String SID_STYLE_DESIGNER = ".uno:DesignerDialog"; //AMT - 5539
        ///<summary>Distribution</summary>
        public const String SID_DISTRIBUTE_DLG = ".uno:DistributeSelection"; //AMT - 5683
        ///<summary>Callouts</summary>
        public const String SID_DRAW_CAPTION = ".uno:DrawCaption"; //AMT - 10254
        ///<summary>Text</summary>
        public const String SID_DRAW_TEXT = ".uno:DrawText"; //AMT - 10253
        ///<summary>Text Box</summary>
        public const String SID_FM_EDIT = ".uno:Edit"; //AMT - 10599
        ///<summary>Edit File</summary>
        public const String SID_EDITDOC = ".uno:EditDoc"; //AMT - 6312
        ///<summary>Edit FrameSet</summary>
        public const String SID_EDIT_FRAMESET = ".uno:EditFrameSet"; //AMT - 5646
        ///<summary>Ellipse</summary>
        public const String SID_DRAW_ELLIPSE = ".uno:Ellipse"; //AMT - 10110
        ///<summary>Ellipse Segment</summary>
        public const String SID_DRAW_ELLIPSECUT = ".uno:EllipseCut"; //AMT - 10392
        ///<summary>Ellipse Segment, unfilled</summary>
        public const String SID_DRAW_ELLIPSECUT_NOFILL = ".uno:EllipseCut_Unfilled"; //AMT - 10393
        ///<summary>Ellipse, Unfilled</summary>
        public const String SID_DRAW_ELLIPSE_NOFILL = ".uno:Ellipse_Unfilled"; //AMT - 10384
        ///<summary>Enter Group</summary>
        public const String SID_ENTER_GROUP = ".uno:EnterGroup"; //AMT - 27096
        ///<summary>Export Directly as PDF</summary>
        public const String SID_DIRECTEXPORTDOCASPDF = ".uno:ExportDirectToPDF"; //AMT - 6674
        ///<summary>Export</summary>
        public const String SID_EXPORTDOC = ".uno:ExportTo"; //AMT - 5829
        ///<summary>Export as PDF</summary>
        public const String SID_EXPORTDOCASPDF = ".uno:ExportToPDF"; //AMT - 6673
        ///<summary>What's This?</summary>
        public const String SID_EXTENDEDHELP = ".uno:ExtendedHelp"; //AMT - 5402
        ///<summary>3D Color</summary>
        public const String SID_EXTRUSION_3D_COLOR = ".uno:Extrusion3DColor"; //AMT - 10969
        ///<summary>Extrusion</summary>
        public const String SID_EXTRUSION_DEPTH = ".uno:ExtrusionDepth"; //AMT - 10970
        ///<summary>Extrusion Depth</summary>
        public const String SID_EXTRUSION_DEPTH_DIALOG = ".uno:ExtrusionDepthDialog"; //AMT - 10976
        ///<summary>Depth</summary>
        public const String SID_EXTRUSION_DEPTH_FLOATER = ".uno:ExtrusionDepthFloater"; //AMT - 10965
        ///<summary>Direction</summary>
        public const String SID_EXTRUSION_DIRECTION_FLOATER = ".uno:ExtrusionDirectionFloater"; //AMT - 10966
        ///<summary>Lighting</summary>
        public const String SID_EXTRUSION_LIGHTING_FLOATER = ".uno:ExtrusionLightingFloater"; //AMT - 10967
        ///<summary>Surface</summary>
        public const String SID_EXTRUSION_SURFACE_FLOATER = ".uno:ExtrusionSurfaceFloater"; //AMT - 10968
        ///<summary>Tilt Down</summary>
        public const String SID_EXTRUSION_TILT_DOWN = ".uno:ExtrusionTiltDown"; //AMT - 10961
        ///<summary>Tilt Left</summary>
        public const String SID_EXTRUSION_TILT_LEFT = ".uno:ExtrusionTiltLeft"; //AMT - 10963
        ///<summary>Tilt Right</summary>
        public const String SID_EXTRUSION_TILT_RIGHT = ".uno:ExtrusionTiltRight"; //AMT - 10964
        ///<summary>Tilt Up</summary>
        public const String SID_EXTRUSION_TILT_UP = ".uno:ExtrusionTiltUp"; //AMT - 10962
        ///<summary>Extrusion On/Off</summary>
        public const String SID_EXTRUSION_TOOGLE = ".uno:ExtrusionToggle"; //AMT - 10960
        ///<summary>File Selection</summary>
        public const String SID_FM_FILECONTROL = ".uno:FileControl"; //AMT - 10605
        ///<summary>File Document</summary>
        public const String SID_SAVEDOCTOBOOKMARK = ".uno:FileDocument"; //AMT - 6315
        ///<summary>Fill Color</summary>
        public const String SID_ATTR_FILL_COLOR = ".uno:FillColor"; //AMT - 10165
        ///<summary>Shadow</summary>
        public const String SID_ATTR_FILL_SHADOW = ".uno:FillShadow"; //AMT - 10299
        ///<summary>Area Style / Filling</summary>
        public const String SID_ATTR_FILL_STYLE = ".uno:FillStyle"; //AMT - 10164
        ///<summary>Standard Filter</summary>
        public const String SID_FM_FILTERCRIT = ".uno:FilterCrit"; //AMT - 10715
        ///<summary>First Record</summary>
        public const String SID_FM_RECORD_FIRST = ".uno:FirstRecord"; //AMT - 10616
        ///<summary>Flash</summary>
        public const String SID_ATTR_FLASH = ".uno:Flash"; //AMT - 10406
        ///<summary>Flowcharts</summary>
        public const String SID_DRAWTBX_CS_FLOWCHART = ".uno:FlowChartShapes"; //AMT - 11050
        ///<summary>Character</summary>
        public const String SID_CHAR_DLG = ".uno:FontDialog"; //AMT - 10296
        ///<summary>Font Size</summary>
        public const String SID_ATTR_CHAR_FONTHEIGHT = ".uno:FontHeight"; //AMT - 10015
        ///<summary>Fontwork</summary>
        public const String SID_FONTWORK = ".uno:FontWork"; //AMT - 10256
        ///<summary>Fontwork Alignment</summary>
        public const String SID_FONTWORK_ALIGNMENT_FLOATER = ".uno:FontworkAlignmentFloater"; //AMT - 10981
        ///<summary>Fontwork Character Spacing</summary>
        public const String SID_FONTWORK_CHARACTER_SPACING_FLOATER = ".uno:FontworkCharacterSpacingFloater"; //AMT - 10982
        ///<summary>Fontwork Gallery</summary>
        public const String SID_FONTWORK_GALLERY_FLOATER = ".uno:FontworkGalleryFloater"; //AMT - 10977
        ///<summary>Fontwork Same Letter Heights</summary>
        public const String SID_FONTWORK_SAME_LETTER_HEIGHTS = ".uno:FontworkSameLetterHeights"; //AMT - 10980
        ///<summary>Fontwork Shape</summary>
        public const String SID_FONTWORK_SHAPE_TYPES = ".uno:FontworkShapeTypes"; //AMT - 10978
        ///<summary>Form Design</summary>
        public const String SID_FM_FORM_DESIGN_TOOLS = ".uno:FormDesignTools"; //AMT - 11046
        ///<summary>Form-Based Filters</summary>
        public const String SID_FM_FILTER_START = ".uno:FormFilter"; //AMT - 10729
        ///<summary>Apply Form-Based Filter</summary>
        public const String SID_FM_FILTER_EXECUTE = ".uno:FormFilterExecute"; //AMT - 10731
        ///<summary>Close</summary>
        public const String SID_FM_FILTER_EXIT = ".uno:FormFilterExit"; //AMT - 10730
        ///<summary>Filter Navigation</summary>
        public const String SID_FM_FILTER_NAVIGATOR = ".uno:FormFilterNavigator"; //AMT - 10732
        ///<summary>Apply Filter</summary>
        public const String SID_FM_FORM_FILTERED = ".uno:FormFiltered"; //AMT - 10723
        ///<summary>Form</summary>
        public const String SID_FM_PROPERTIES = ".uno:FormProperties"; //AMT - 10614
        ///<summary>Area</summary>
        public const String SID_ATTRIBUTES_AREA = ".uno:FormatArea"; //AMT - 10142
        ///<summary>Group</summary>
        public const String SID_GROUP = ".uno:FormatGroup"; //AMT - 10454
        ///<summary>Line</summary>
        public const String SID_ATTRIBUTES_LINE = ".uno:FormatLine"; //AMT - 10143
        ///<summary>Format</summary>
        public const String SID_FORMATMENU = ".uno:FormatMenu"; //AMT - 5780
        ///<summary>Format Paintbrush</summary>
        public const String SID_FORMATPAINTBRUSH = ".uno:FormatPaintbrush"; //AMT - 5715
        ///<summary>Ungroup</summary>
        public const String SID_UNGROUP = ".uno:FormatUngroup"; //AMT - 10455
        ///<summary>Formatted Field</summary>
        public const String SID_FM_FORMATTEDFIELD = ".uno:FormattedField"; //AMT - 10728
        ///<summary>Contents</summary>
        public const String SID_FRAME_CONTENT = ".uno:FrameContent"; //AMT - 5826
        ///<summary>Line Color (of the border)</summary>
        public const String SID_FRAME_LINECOLOR = ".uno:FrameLineColor"; //AMT - 10201
        ///<summary>Name</summary>
        public const String SID_FRAME_NAME = ".uno:FrameName"; //AMT - 5825
        ///<summary>FrameSet Spacing</summary>
        public const String SID_FRAMESPACING = ".uno:FrameSpacing"; //AMT - 6507
        ///<summary>Freeform Line, Filled</summary>
        public const String SID_DRAW_FREELINE = ".uno:Freeline"; //AMT - 10463
        ///<summary>Freeform Line</summary>
        public const String SID_DRAW_FREELINE_NOFILL = ".uno:Freeline_Unfilled"; //AMT - 10464
        ///<summary>Full Screen</summary>
        public const String SID_WIN_FULLSCREEN = ".uno:FullScreen"; //AMT - 5627
        ///<summary>Gallery</summary>
        public const String SID_GALLERY = ".uno:Gallery"; //AMT - 5960
        ///<summary>Color Palette</summary>
        public const String SID_GET_COLORTABLE = ".uno:GetColorTable"; //AMT - 10441
        ///<summary>Move Down</summary>
        public const String SID_CURSORDOWN = ".uno:GoDown"; //AMT - 5731
        ///<summary>Page Down</summary>
        public const String SID_CURSORPAGEDOWN = ".uno:GoDownBlock"; //AMT - 5735
        ///<summary>Select Page Down</summary>
        public const String SID_CURSORPAGEDOWN_SEL = ".uno:GoDownBlockSel"; //AMT - 26525
        ///<summary>Select Down</summary>
        public const String SID_CURSORDOWN_SEL = ".uno:GoDownSel"; //AMT - 26521
        ///<summary>Move Left</summary>
        public const String SID_CURSORLEFT = ".uno:GoLeft"; //AMT - 5733
        ///<summary>Page Left</summary>
        public const String SID_CURSORPAGELEFT = ".uno:GoLeftBlock"; //AMT - 5738
        ///<summary>Select Page Left</summary>
        public const String SID_CURSORPAGELEFT_SEL = ".uno:GoLeftBlockSel"; //AMT - 26528
        ///<summary>Select Left</summary>
        public const String SID_CURSORLEFT_SEL = ".uno:GoLeftSel"; //AMT - 26523
        ///<summary>Move Right</summary>
        public const String SID_CURSORRIGHT = ".uno:GoRight"; //AMT - 5734
        ///<summary>Select Right</summary>
        public const String SID_CURSORRIGHT_SEL = ".uno:GoRightSel"; //AMT - 26524
        ///<summary>To File End</summary>
        public const String SID_CURSORENDOFFILE = ".uno:GoToEndOfData"; //AMT - 5741
        ///<summary>Select to File End</summary>
        public const String SID_CURSORENDOFFILE_SEL = ".uno:GoToEndOfDataSel"; //AMT - 26532
        ///<summary>To Document End</summary>
        public const String SID_CURSOREND = ".uno:GoToEndOfRow"; //AMT - 5746
        ///<summary>Select to Document End</summary>
        public const String SID_CURSOREND_SEL = ".uno:GoToEndOfRowSel"; //AMT - 26534
        ///<summary>To File Begin</summary>
        public const String SID_CURSORTOPOFFILE = ".uno:GoToStart"; //AMT - 5742
        ///<summary>To Document Begin</summary>
        public const String SID_CURSORHOME = ".uno:GoToStartOfRow"; //AMT - 5745
        ///<summary>Select to Document Begin</summary>
        public const String SID_CURSORHOME_SEL = ".uno:GoToStartOfRowSel"; //AMT - 26533
        ///<summary>Select to File Begin</summary>
        public const String SID_CURSORTOPOFFILE_SEL = ".uno:GoToStartSel"; //AMT - 26531
        ///<summary>Move Up</summary>
        public const String SID_CURSORUP = ".uno:GoUp"; //AMT - 5732
        ///<summary>Page Up</summary>
        public const String SID_CURSORPAGEUP = ".uno:GoUpBlock"; //AMT - 5736
        ///<summary>Select Page Up</summary>
        public const String SID_CURSORPAGEUP_SEL = ".uno:GoUpBlockSel"; //AMT - 26526
        ///<summary>Select Up</summary>
        public const String SID_CURSORUP_SEL = ".uno:GoUpSel"; //AMT - 26522
        ///<summary>Control Focus</summary>
        public const String SID_FM_GRABCONTROLFOCUS = ".uno:GrabControlFocus"; //AMT - 10767
        ///<summary>Crop</summary>
        public const String SID_ATTR_GRAF_CROP = ".uno:GrafAttrCrop"; //AMT - 10883
        ///<summary>Blue</summary>
        public const String SID_ATTR_GRAF_BLUE = ".uno:GrafBlue"; //AMT - 10867
        ///<summary>Contrast</summary>
        public const String SID_ATTR_GRAF_CONTRAST = ".uno:GrafContrast"; //AMT - 10864
        ///<summary>Gamma</summary>
        public const String SID_ATTR_GRAF_GAMMA = ".uno:GrafGamma"; //AMT - 10868
        ///<summary>Green</summary>
        public const String SID_ATTR_GRAF_GREEN = ".uno:GrafGreen"; //AMT - 10866
        ///<summary>Invert</summary>
        public const String SID_ATTR_GRAF_INVERT = ".uno:GrafInvert"; //AMT - 10870
        ///<summary>Brightness</summary>
        public const String SID_ATTR_GRAF_LUMINANCE = ".uno:GrafLuminance"; //AMT - 10863
        ///<summary>Graphics mode</summary>
        public const String SID_ATTR_GRAF_MODE = ".uno:GrafMode"; //AMT - 10871
        ///<summary>Red</summary>
        public const String SID_ATTR_GRAF_RED = ".uno:GrafRed"; //AMT - 10865
        ///<summary>Transparency</summary>
        public const String SID_ATTR_GRAF_TRANSPARENCE = ".uno:GrafTransparence"; //AMT - 10869
        ///<summary>Invert</summary>
        public const String SID_GRFFILTER_INVERT = ".uno:GraphicFilterInvert"; //AMT - 10470
        ///<summary>Mosaic</summary>
        public const String SID_GRFFILTER_MOSAIC = ".uno:GraphicFilterMosaic"; //AMT - 10475
        ///<summary>Pop Art</summary>
        public const String SID_GRFFILTER_POPART = ".uno:GraphicFilterPopart"; //AMT - 10478
        ///<summary>Posterize</summary>
        public const String SID_GRFFILTER_POSTER = ".uno:GraphicFilterPoster"; //AMT - 10477
        ///<summary>Relief</summary>
        public const String SID_GRFFILTER_EMBOSS = ".uno:GraphicFilterRelief"; //AMT - 10476
        ///<summary>Remove Noise</summary>
        public const String SID_GRFFILTER_REMOVENOISE = ".uno:GraphicFilterRemoveNoise"; //AMT - 10473
        ///<summary>Aging</summary>
        public const String SID_GRFFILTER_SEPIA = ".uno:GraphicFilterSepia"; //AMT - 10479
        ///<summary>Sharpen</summary>
        public const String SID_GRFFILTER_SHARPEN = ".uno:GraphicFilterSharpen"; //AMT - 10472
        ///<summary>Smooth</summary>
        public const String SID_GRFFILTER_SMOOTH = ".uno:GraphicFilterSmooth"; //AMT - 10471
        ///<summary>Charcoal Sketch</summary>
        public const String SID_GRFFILTER_SOBEL = ".uno:GraphicFilterSobel"; //AMT - 10474
        ///<summary>Solarization</summary>
        public const String SID_GRFFILTER_SOLARIZE = ".uno:GraphicFilterSolarize"; //AMT - 10480
        ///<summary>Filter</summary>
        public const String SID_GRFFILTER = ".uno:GraphicFilterToolbox"; //AMT - 10469
        ///<summary>Table Control</summary>
        public const String SID_FM_DBGRID = ".uno:Grid"; //AMT - 10603
        ///<summary>Snap to Grid</summary>
        public const String SID_GRID_USE = ".uno:GridUse"; //AMT - 27154
        ///<summary>Display Grid</summary>
        public const String SID_GRID_VISIBLE = ".uno:GridVisible"; //AMT - 27322
        ///<summary>Group</summary>
        public const String SID_OUTLINE_MAKE = ".uno:Group"; //AMT - 26331
        ///<summary>Group Box</summary>
        public const String SID_FM_GROUPBOX = ".uno:GroupBox"; //AMT - 10598
        ///<summary>Group Box</summary>
        public const String SID_INSERT_GROUPBOX = ".uno:Groupbox"; //AMT - 10189
        ///<summary>Horizontal Line</summary>
        public const String SID_INSERT_HFIXEDLINE = ".uno:HFixedLine"; //AMT - 10928
        ///<summary>Horizontal Scroll Bar</summary>
        public const String SID_INSERT_HSCROLLBAR = ".uno:HScrollbar"; //AMT - 10194
        ///<summary>Hangul/Hanja Conversion</summary>
        public const String SID_HANGUL_HANJA_CONVERSION = ".uno:HangulHanjaConversion"; //AMT - 10959
        ///<summary>Choose Help File</summary>
        public const String SID_HELP_HELPFILEBOX = ".uno:HelpChooseFile"; //AMT - 5419
        ///<summary>OpenOffice.org Help</summary>
        public const String SID_HELPINDEX = ".uno:HelpIndex"; //AMT - 5401
        ///<summary>Help</summary>
        public const String SID_HELPMENU = ".uno:HelpMenu"; //AMT - 5410
        ///<summary>Help on Help</summary>
        public const String SID_HELPONHELP = ".uno:HelpOnHelp"; //AMT - 5400
        ///<summary>Support</summary>
        public const String SID_HELP_SUPPORTPAGE = ".uno:HelpSupport"; //AMT - 6683
        ///<summary>Help Agent</summary>
        public const String SID_HELP_PI = ".uno:HelperDialog"; //AMT - 5962
        ///<summary>Guides When Moving</summary>
        public const String SID_HELPLINES_MOVE = ".uno:HelplinesMove"; //AMT - 27153
        ///<summary>Hide Details</summary>
        public const String SID_OUTLINE_HIDE = ".uno:HideDetail"; //AMT - 26329
        ///<summary>Do Not Mark Errors</summary>
        public const String SID_AUTOSPELL_MARKOFF = ".uno:HideSpellMark"; //AMT - 12022
        ///<summary>Hyperlink</summary>
        public const String SID_HYPERLINK_DIALOG = ".uno:HyperlinkDialog"; //AMT - 5678
        ///<summary>Image Control</summary>
        public const String SID_FM_IMAGECONTROL = ".uno:ImageControl"; //AMT - 10710
        ///<summary>ImageMap</summary>
        public const String SID_IMAP = ".uno:ImageMapDialog"; //AMT - 10371
        ///<summary>Image Button</summary>
        public const String SID_FM_IMAGEBUTTON = ".uno:Imagebutton"; //AMT - 10604
        ///<summary>Increase Indent</summary>
        public const String SID_INC_INDENT = ".uno:IncrementIndent"; //AMT - 10462
        ///<summary>Movie and Sound</summary>
        public const String SID_INSERT_AVMEDIA = ".uno:InsertAVMedia"; //AMT - 6696
        ///<summary>Note</summary>
        public const String SID_INSERT_POSTIT = ".uno:InsertAnnotation"; //AMT - 26276
        ///<summary>Applet</summary>
        public const String SID_INSERT_APPLET = ".uno:InsertApplet"; //AMT - 5673
        ///<summary>Insert business cards</summary>
        public const String FN_BUSINESS_CARD = ".uno:InsertBusinessCard"; //AMT - 21052
        ///<summary>Currency Field</summary>
        public const String SID_INSERT_CURRENCYFIELD = ".uno:InsertCurrencyField"; //AMT - 10939
        ///<summary>File</summary>
        public const String SID_INSERTDOC = ".uno:InsertDoc"; //AMT - 5532
        ///<summary>Show Draw Functions</summary>
        public const String SID_INSERT_DRAW = ".uno:InsertDraw"; //AMT - 10244
        ///<summary>Text Box</summary>
        public const String SID_INSERT_EDIT = ".uno:InsertEdit"; //AMT - 10190
        ///<summary>File Selection</summary>
        public const String SID_INSERT_FILECONTROL = ".uno:InsertFileControl"; //AMT - 10942
        ///<summary>Label field</summary>
        public const String SID_INSERT_FIXEDTEXT = ".uno:InsertFixedText"; //AMT - 10188
        ///<summary>Formatted Field</summary>
        public const String SID_INSERT_FORMATTEDFIELD = ".uno:InsertFormattedField"; //AMT - 10940
        ///<summary>From File</summary>
        public const String SID_INSERT_GRAPHIC = ".uno:InsertGraphic"; //AMT - 10241
        ///<summary>Hyperlink Bar</summary>
        public const String SID_HYPERLINK_INSERT = ".uno:InsertHyperlink"; //AMT - 10360
        ///<summary>Insert from Image Editor</summary>
        public const String SID_INSERT_IMAGE = ".uno:InsertImage"; //AMT - 27105
        ///<summary>Image Control</summary>
        public const String SID_INSERT_IMAGECONTROL = ".uno:InsertImageControl"; //AMT - 10926
        ///<summary>Insert Labels</summary>
        public const String FN_LABEL = ".uno:InsertLabels"; //AMT - 21051
        ///<summary>List Box</summary>
        public const String SID_INSERT_LISTBOX = ".uno:InsertListbox"; //AMT - 10191
        ///<summary>Formula</summary>
        public const String SID_INSERT_MATH = ".uno:InsertMath"; //AMT - 27106
        ///<summary>Insert Mode</summary>
        public const String SID_ATTR_INSERT = ".uno:InsertMode"; //AMT - 10221
        ///<summary>Numeric Field</summary>
        public const String SID_INSERT_NUMERICFIELD = ".uno:InsertNumericField"; //AMT - 10938
        ///<summary>OLE Object</summary>
        public const String SID_INSERT_OBJECT = ".uno:InsertObject"; //AMT - 5561
        ///<summary>Chart</summary>
        public const String SID_INSERT_DIAGRAM = ".uno:InsertObjectChart"; //AMT - 10140
        ///<summary>Floating Frame</summary>
        public const String SID_INSERT_FLOATINGFRAME = ".uno:InsertObjectFloatingFrame"; //AMT - 5563
        ///<summary>Pattern Field</summary>
        public const String SID_INSERT_PATTERNFIELD = ".uno:InsertPatternField"; //AMT - 10941
        ///<summary>Plug-in</summary>
        public const String SID_INSERT_PLUGIN = ".uno:InsertPlugin"; //AMT - 5672
        ///<summary>Button</summary>
        public const String SID_INSERT_PUSHBUTTON = ".uno:InsertPushbutton"; //AMT - 10146
        ///<summary>Sound</summary>
        public const String SID_INSERT_SOUND = ".uno:InsertSound"; //AMT - 5676
        ///<summary>Spreadsheet</summary>
        public const String SID_ATTR_TABLE = ".uno:InsertSpreadsheet"; //AMT - 10217
        ///<summary>Special Character</summary>
        public const String SID_CHARMAP = ".uno:InsertSymbol"; //AMT - 10503
        ///<summary>Insert Text Frame</summary>
        public const String SID_INSERT_FRAME = ".uno:InsertTextFrame"; //AMT - 10240
        ///<summary>Time Field</summary>
        public const String SID_INSERT_TIMEFIELD = ".uno:InsertTimeField"; //AMT - 10937
        ///<summary>Video</summary>
        public const String SID_INSERT_VIDEO = ".uno:InsertVideo"; //AMT - 5677
        ///<summary>Internet Options</summary>
        public const String SID_INET_DLG = ".uno:InternetDialog"; //AMT - 10416
        ///<summary>Intersect</summary>
        public const String SID_POLY_INTERSECT = ".uno:Intersect"; //AMT - 5681
        ///<summary>Load Document</summary>
        public const String SID_DOC_LOADING = ".uno:IsLoading"; //AMT - 5585
        ///<summary>Italic</summary>
        public const String SID_ATTR_CHAR_POSTURE = ".uno:Italic"; //AMT - 10008
        ///<summary>Justified</summary>
        public const String SID_ATTR_PARA_ADJUST_BLOCK = ".uno:JustifyPara"; //AMT - 10031
        ///<summary>Label Field</summary>
        public const String SID_FM_FIXEDTEXT = ".uno:Label"; //AMT - 10597
        ///<summary>Last Record</summary>
        public const String SID_FM_RECORD_LAST = ".uno:LastRecord"; //AMT - 10619
        ///<summary>Start Image Editor</summary>
        public const String SID_SIM_START = ".uno:LaunchStarImage"; //AMT - 30000
        ///<summary>Exit group</summary>
        public const String SID_LEAVE_GROUP = ".uno:LeaveGroup"; //AMT - 27097
        ///<summary>Align Left</summary>
        public const String SID_ATTR_PARA_ADJUST_LEFT = ".uno:LeftPara"; //AMT - 10028
        ///<summary>Current Library</summary>
        public const String SID_BASICIDE_LIBSELECTOR = ".uno:LibSelector"; //AMT - 30787
        ///<summary>Line</summary>
        public const String SID_DRAW_LINE = ".uno:Line"; //AMT - 10102
        ///<summary>Line Dash/Dot</summary>
        public const String SID_ATTR_LINE_DASH = ".uno:LineDash"; //AMT - 10170
        ///<summary>Arrow Style</summary>
        public const String SID_ATTR_LINEEND_STYLE = ".uno:LineEndStyle"; //AMT - 10301
        ///<summary>Line Style</summary>
        public const String SID_FRAME_LINESTYLE = ".uno:LineStyle"; //AMT - 10200
        ///<summary>Line Width</summary>
        public const String SID_ATTR_LINE_WIDTH = ".uno:LineWidth"; //AMT - 10171
        ///<summary>Line (45%)</summary>
        public const String SID_DRAW_XLINE = ".uno:Line_Diagonal"; //AMT - 10103
        ///<summary>List Box</summary>
        public const String SID_FM_LISTBOX = ".uno:ListBox"; //AMT - 10600
        ///<summary>Insert BASIC Source</summary>
        public const String SID_BASICLOAD = ".uno:LoadBasic"; //AMT - 5951
        ///<summary>Load Configuration</summary>
        public const String SID_LOADCONFIG = ".uno:LoadConfiguration"; //AMT - 5933
        ///<summary>OpenOffice.org Basic</summary>
        public const String SID_BASICCHOOSER = ".uno:MacroDialog"; //AMT - 5959
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_MACROORGANIZER = ".uno:MacroOrganizer"; //AMT - 6691
        ///<summary>Record Macro</summary>
        public const String SID_RECORDMACRO = ".uno:MacroRecorder"; //AMT - 6669
        ///<summary>Digital Signature</summary>
        public const String SID_MACRO_SIGNATURE = ".uno:MacroSignature"; //AMT - 6704
        ///<summary>Manage Breakpoints</summary>
        public const String SID_BASICIDE_MANAGEBRKPNTS = ".uno:ManageBreakPoints"; //AMT - 30810
        ///<summary>Find Parenthesis</summary>
        public const String SID_BASICIDE_MATCHGROUP = ".uno:MatchGroup"; //AMT - 30782
        ///<summary>Menu On/Off</summary>
        public const String SID_TOGGLE_MENUBAR = ".uno:MenuBarVisible"; //AMT - 6661
        ///<summary>Merge</summary>
        public const String SID_POLY_MERGE = ".uno:Merge"; //AMT - 5679
        ///<summary>Merge Document</summary>
        public const String SID_DOCUMENT_MERGE = ".uno:MergeDocuments"; //AMT - 6587
        ///<summary>Document Modified</summary>
        public const String SID_DOC_MODIFIED = ".uno:ModifiedStatus"; //S - 5584
        ///<summary>Frame Properties</summary>
        public const String SID_MODIFY_FRAME = ".uno:ModifyFrame"; //AMT - 5651
        ///<summary>Select Module</summary>
        public const String SID_BASICIDE_MODULEDLG = ".uno:ModuleDialog"; //AMT - 30773
        ///<summary>More Controls</summary>
        public const String SID_FM_MORE_CONTROLS = ".uno:MoreControls"; //AMT - 11045
        ///<summary>Navigation Bar</summary>
        public const String SID_FM_NAVIGATIONBAR = ".uno:NavigationBar"; //AMT - 10607
        ///<summary>Navigator</summary>
        public const String SID_NAVIGATOR = ".uno:Navigator"; //AMT - 10366
        ///<summary>New Document From Template</summary>
        public const String SID_NEWDOC = ".uno:NewDoc"; //AMT - 5500
        ///<summary>New FrameSet</summary>
        public const String SID_NEWFRAMESET = ".uno:NewFrameSet"; //AMT - 6400
        ///<summary>New Presentation</summary>
        public const String SID_NEWSD = ".uno:NewPresentation"; //AMT - 6686
        ///<summary>New Record</summary>
        public const String SID_FM_RECORD_NEW = ".uno:NewRecord"; //AMT - 10620
        ///<summary>New Window</summary>
        public const String SID_NEWWINDOW = ".uno:NewWindow"; //AMT - 5620
        ///<summary>Next Record</summary>
        public const String SID_FM_RECORD_NEXT = ".uno:NextRecord"; //AMT - 10617
        ///<summary>Numerical Field</summary>
        public const String SID_FM_NUMERICFIELD = ".uno:NumericField"; //AMT - 10706
        ///<summary>Alignment</summary>
        public const String SID_OBJECT_ALIGN = ".uno:ObjectAlign"; //AMT - 10130
        ///<summary>Left</summary>
        public const String SID_OBJECT_ALIGN_LEFT = ".uno:ObjectAlignLeft"; //AMT - 10131
        ///<summary>Right</summary>
        public const String SID_OBJECT_ALIGN_RIGHT = ".uno:ObjectAlignRight"; //AMT - 10133
        ///<summary>Back One</summary>
        public const String SID_FRAME_DOWN = ".uno:ObjectBackOne"; //AMT - 26408
        ///<summary>Object Catalog</summary>
        public const String SID_BASICIDE_OBJCAT = ".uno:ObjectCatalog"; //AMT - 30774
        ///<summary>Forward One</summary>
        public const String SID_FRAME_UP = ".uno:ObjectForwardOne"; //AMT - 26407
        ///<summary>Object</summary>
        public const String SID_OBJECT = ".uno:ObjectMenue"; //AMT - 5575
        ///<summary>Registration</summary>
        public const String SID_ONLINE_REGISTRATION = ".uno:OnlineRegistrationDlg"; //AMT - 6537
        ///<summary>Open</summary>
        public const String SID_OPENDOC = ".uno:Open"; //AMT - 5501
        ///<summary>Open Hyperlink</summary>
        public const String SID_OPEN_HYPERLINK = ".uno:OpenHyperlinkOnCursor"; //AMT - 10955
        ///<summary>Open in Design Mode</summary>
        public const String SID_FM_OPEN_READONLY = ".uno:OpenReadOnly"; //AMT - 10709
        ///<summary>Edit</summary>
        public const String SID_OPENTEMPLATE = ".uno:OpenTemplate"; //AMT - 5594
        ///<summary>Load URL</summary>
        public const String SID_OPENURL = ".uno:OpenUrl"; //AMT - 5596
        ///<summary>XML Filter Settings</summary>
        public const String SID_OPEN_XML_FILTERSETTINGS = ".uno:OpenXMLFilterSettings"; //AMT - 10958
        ///<summary>Options</summary>
        public const String SID_OPTIONS_TREEDIALOG = ".uno:OptionsTreeDialog"; //AMT - 31630
        ///<summary>Sort</summary>
        public const String SID_FM_ORDERCRIT = ".uno:OrderCrit"; //AMT - 10714
        ///<summary>Organize</summary>
        public const String SID_ORGANIZER = ".uno:Organizer"; //AMT - 5540
        ///<summary>Bullets and Numbering</summary>
        public const String SID_OUTLINE_BULLET = ".uno:OutlineBullet"; //AMT - 10156
        ///<summary>Hide Subpoints</summary>
        public const String SID_OUTLINE_COLLAPSE = ".uno:OutlineCollapse"; //AMT - 10231
        ///<summary>First Level</summary>
        public const String SID_OUTLINE_COLLAPSE_ALL = ".uno:OutlineCollapseAll"; //AMT - 10155
        ///<summary>Move Down</summary>
        public const String SID_OUTLINE_DOWN = ".uno:OutlineDown"; //AMT - 10151
        ///<summary>Show Subpoints</summary>
        public const String SID_OUTLINE_EXPAND = ".uno:OutlineExpand"; //AMT - 10233
        ///<summary>All Levels</summary>
        public const String SID_OUTLINE_EXPAND_ALL = ".uno:OutlineExpandAll"; //AMT - 10232
        ///<summary>Outline</summary>
        public const String SID_ATTR_CHAR_CONTOUR = ".uno:OutlineFont"; //AMT - 10012
        ///<summary>Formatting On/Off</summary>
        public const String SID_OUTLINE_FORMAT = ".uno:OutlineFormat"; //AMT - 10154
        ///<summary>Promote</summary>
        public const String SID_OUTLINE_LEFT = ".uno:OutlineLeft"; //AMT - 10152
        ///<summary>Demote</summary>
        public const String SID_OUTLINE_RIGHT = ".uno:OutlineRight"; //AMT - 10153
        ///<summary>Move Up</summary>
        public const String SID_OUTLINE_UP = ".uno:OutlineUp"; //AMT - 10150
        ///<summary>Left-To-Right</summary>
        public const String SID_ATTR_PARA_LEFT_TO_RIGHT = ".uno:ParaLeftToRight"; //AMT - 10950
        ///<summary>Right-To-Left</summary>
        public const String SID_ATTR_PARA_RIGHT_TO_LEFT = ".uno:ParaRightToLeft"; //AMT - 10951
        ///<summary>Paragraph</summary>
        public const String SID_PARA_DLG = ".uno:ParagraphDialog"; //AMT - 10297
        ///<summary>Paste</summary>
        public const String SID_PASTE = ".uno:Paste"; //AMT - 5712
        ///<summary>Pattern Field</summary>
        public const String SID_FM_PATTERNFIELD = ".uno:PatternField"; //AMT - 10708
        ///<summary>File</summary>
        public const String SID_PICKLIST = ".uno:PickList"; //AMT - 5510
        ///<summary>Ellipse Pie</summary>
        public const String SID_DRAW_PIE = ".uno:Pie"; //AMT - 10112
        ///<summary>Ellipse Pie, Unfilled</summary>
        public const String SID_DRAW_PIE_NOFILL = ".uno:Pie_Unfilled"; //AMT - 10387
        ///<summary>Plug-in</summary>
        public const String SID_PLUGINS_ACTIVE = ".uno:PlugInsActive"; //AMT - 6314
        ///<summary>Shapes</summary>
        public const String SID_POLY_FORMEN = ".uno:PolyFormen"; //AMT - 5682
        ///<summary>Polygon (45%), Filled</summary>
        public const String SID_DRAW_XPOLYGON = ".uno:Polygon_Diagonal"; //AMT - 10394
        ///<summary>Polygon (45%)</summary>
        public const String SID_DRAW_XPOLYGON_NOFILL = ".uno:Polygon_Diagonal_Unfilled"; //AMT - 10396
        ///<summary>Polygon</summary>
        public const String SID_DRAW_POLYGON_NOFILL = ".uno:Polygon_Unfilled"; //AMT - 10395
        ///<summary>Previous Record</summary>
        public const String SID_FM_RECORD_PREV = ".uno:PrevRecord"; //AMT - 10618
        ///<summary>Preview</summary>
        public const String SID_INSERT_PREVIEW = ".uno:Preview"; //AMT - 10196
        ///<summary>Print</summary>
        public const String SID_PRINTDOC = ".uno:Print"; //AMT - 5504
        ///<summary>Print File Directly</summary>
        public const String SID_PRINTDOCDIRECT = ".uno:PrintDefault"; //AMT - 5509
        ///<summary>Page Preview</summary>
        public const String SID_PRINTPREVIEW = ".uno:PrintPreview"; //AMT - 5325
        ///<summary>Printer Settings</summary>
        public const String SID_SETUPPRINTER = ".uno:PrinterSetup"; //AMT - 5302
        ///<summary>Progress Bar</summary>
        public const String SID_INSERT_PROGRESSBAR = ".uno:ProgressBar"; //AMT - 10927
        ///<summary>Push Button</summary>
        public const String SID_FM_PUSHBUTTON = ".uno:Pushbutton"; //AMT - 10594
        ///<summary>Exit</summary>
        public const String SID_QUITAPP = ".uno:Quit"; //AMT - 5300
        ///<summary>Option Button</summary>
        public const String SID_FM_RADIOBUTTON = ".uno:RadioButton"; //AMT - 10595
        ///<summary>Option Button</summary>
        public const String SID_INSERT_RADIOBUTTON = ".uno:Radiobutton"; //AMT - 10147
        ///<summary>Text -> Record</summary>
        public const String SID_FM_RECORD_FROM_TEXT = ".uno:RecFromText"; //AMT - 10625
        ///<summary>Save Record</summary>
        public const String SID_FM_RECORD_SAVE = ".uno:RecSave"; //AMT - 10627
        ///<summary>Find Record</summary>
        public const String SID_FM_SEARCH = ".uno:RecSearch"; //AMT - 10725
        ///<summary>Record</summary>
        public const String SID_FM_RECORD_TEXT = ".uno:RecText"; //AMT - 10624
        ///<summary>Total No. of Records</summary>
        public const String SID_FM_RECORD_TOTAL = ".uno:RecTotal"; //AMT - 10626
        ///<summary>Undo: Data entry</summary>
        public const String SID_FM_RECORD_UNDO = ".uno:RecUndo"; //AMT - 10630
        ///<summary>Recent Documents</summary>
        public const String SID_RECENTFILELIST = ".uno:RecentFileList"; //AMT - 6697
        ///<summary>Rectangle</summary>
        public const String SID_DRAW_RECT = ".uno:Rect"; //AMT - 10104
        ///<summary>Rectangle, Rounded</summary>
        public const String SID_DRAW_RECT_ROUND = ".uno:Rect_Rounded"; //AMT - 10105
        ///<summary>Rounded Rectangle, Unfilled</summary>
        public const String SID_DRAW_RECT_ROUND_NOFILL = ".uno:Rect_Rounded_Unfilled"; //AMT - 10379
        ///<summary>Rectangle, Unfilled</summary>
        public const String SID_DRAW_RECT_NOFILL = ".uno:Rect_Unfilled"; //AMT - 10378
        ///<summary>Can't Restore</summary>
        public const String SID_REDO = ".uno:Redo"; //AMT - 5700
        ///<summary>Refresh</summary>
        public const String SID_FM_REFRESH = ".uno:Refresh"; //AMT - 10724
        ///<summary>Reload</summary>
        public const String SID_RELOAD = ".uno:Reload"; //AMT - 5508
        ///<summary>Remove Filter</summary>
        public const String SID_FM_FILTER_REMOVE = ".uno:RemoveFilter"; //AMT - 10762
        ///<summary>Remove Filter/Sort</summary>
        public const String SID_FM_REMOVE_FILTER_SORT = ".uno:RemoveFilterSort"; //AMT - 10711
        ///<summary>Redraw</summary>
        public const String SID_REPAINT = ".uno:Repaint"; //AMT - 26012
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_REPEAT = ".uno:RepeatAction"; //AMT - 5702
        ///<summary>Align Right</summary>
        public const String SID_ATTR_PARA_ADJUST_RIGHT = ".uno:RightPara"; //AMT - 10029
        ///<summary>Asian phonetic guide</summary>
        public const String SID_RUBY_DIALOG = ".uno:RubyDialog"; //AMT - 6656
        ///<summary>Run BASIC</summary>
        public const String SID_BASICRUN = ".uno:RunBasic"; //AMT - 5955
        ///<summary>Run Macro</summary>
        public const String SID_RUNMACRO = ".uno:RunMacro"; //AMT - 6692
        ///<summary>Save</summary>
        public const String SID_SAVEDOC = ".uno:Save"; //AMT - 5505
        ///<summary>Save All</summary>
        public const String SID_SAVEDOCS = ".uno:SaveAll"; //AMT - 5309
        ///<summary>Save As</summary>
        public const String SID_SAVEASDOC = ".uno:SaveAs"; //AMT - 5502
        ///<summary>Save</summary>
        public const String SID_DOCTEMPLATE = ".uno:SaveAsTemplate"; //AMT - 5538
        ///<summary>Save BASIC</summary>
        public const String SID_BASICSAVEAS = ".uno:SaveBasicAs"; //AMT - 5953
        ///<summary>Save configuration</summary>
        public const String SID_SAVECONFIG = ".uno:SaveConfiguration"; //AMT - 5930
        ///<summary>Run Query</summary>
        public const String SID_FM_EXECUTE = ".uno:SbaExecuteSql"; //AMT - 10721
        ///<summary>Run SQL command directly</summary>
        public const String SID_FM_NATIVESQL = ".uno:SbaNativeSql"; //AMT - 10720
        ///<summary>Spreadsheet Options</summary>
        public const String SID_SC_EDITOPTIONS = ".uno:ScEditOptions"; //AMT - 10435
        ///<summary>Scan</summary>
        public const String SID_SCAN = ".uno:Scan"; //AMT - 10330
        ///<summary>Chart Options</summary>
        public const String SID_SCH_EDITOPTIONS = ".uno:SchEditOptions"; //AMT - 10437
        ///<summary>Organize Macros</summary>
        public const String SID_SCRIPTORGANIZER = ".uno:ScriptOrganizer"; //AMT - 6690
        ///<summary>Scrollbar</summary>
        public const String SID_FM_SCROLLBAR = ".uno:ScrollBar"; //AMT - 10768
        ///<summary>Presentation Options</summary>
        public const String SID_SD_EDITOPTIONS = ".uno:SdEditOptions"; //AMT - 10434
        ///<summary>Presentation Graphic Options</summary>
        public const String SID_SD_GRAPHIC_OPTIONS = ".uno:SdGraphicOptions"; //AMT - 10447
        ///<summary>Find & Replace</summary>
        public const String SID_SEARCH_DLG = ".uno:SearchDialog"; //AMT - 5961
        ///<summary>Select All</summary>
        public const String SID_SELECT = ".uno:Select"; //AMT - 5720
        ///<summary>Select All</summary>
        public const String SID_SELECTALL = ".uno:SelectAll"; //AMT - 5723
        ///<summary>Select</summary>
        public const String SID_INSERT_SELECT = ".uno:SelectMode"; //AMT - 10198
        ///<summary>Select</summary>
        public const String SID_OBJECT_SELECT = ".uno:SelectObject"; //AMT - 10128
        ///<summary>Send Default Fax</summary>
        public const String FN_FAX = ".uno:SendFax"; //AMT - 20028
        ///<summary>Document as E-mail</summary>
        public const String SID_MAIL_SENDDOC = ".uno:SendMail"; //AMT - 5331
        ///<summary>Document as PDF Attachment</summary>
        public const String SID_MAIL_SENDDOCASPDF = ".uno:SendMailDocAsPDF"; //AMT - 6672
        ///<summary>Send to Back</summary>
        public const String SID_FRAME_TO_BOTTOM = ".uno:SendToBack"; //AMT - 10287
        ///<summary>Borders 10187</summary>
        public const String SID_ATTR_BORDER = ".uno:SetBorderStyle"; //AMT - etDefault
        ///<summary>Properties</summary>
        public const String SID_DOCINFO = ".uno:SetDocumentProperties"; //AMT - 5535
        ///<summary>To Background</summary>
        public const String SID_OBJECT_HELL = ".uno:SetObjectToBackground"; //AMT - 10282
        ///<summary>To Foreground</summary>
        public const String SID_OBJECT_HEAVEN = ".uno:SetObjectToForeground"; //AMT - 10283
        ///<summary>Shadow</summary>
        public const String SID_ATTR_CHAR_SHADOWED = ".uno:Shadowed"; //AMT - 10010
        ///<summary>Display Properties</summary>
        public const String SID_SHOW_BROWSER = ".uno:ShowBrowser"; //AMT - 10163
        ///<summary>Data Navigator</summary>
        public const String SID_FM_SHOW_DATANAVIGATOR = ".uno:ShowDataNavigator"; //AMT - 10773
        ///<summary>Show Details</summary>
        public const String SID_OUTLINE_SHOW = ".uno:ShowDetail"; //AMT - 26330
        ///<summary>Form Navigator</summary>
        public const String SID_FM_SHOW_FMEXPLORER = ".uno:ShowFmExplorer"; //AMT - 10633
        ///<summary>Input Method Status</summary>
        public const String SID_SHOW_IME_STATUS_WINDOW = ".uno:ShowImeStatusWindow"; //AMT - 6680
        ///<summary>Item Browser On/Off</summary>
        public const String SID_SHOW_ITEMBROWSER = ".uno:ShowItemBrowser"; //AMT - 11001
        ///<summary>Properties</summary>
        public const String SID_SHOW_PROPERTYBROWSER = ".uno:ShowPropBrowser"; //AMT - 10943
        ///<summary>Digital Signatures</summary>
        public const String SID_SIGNATURE = ".uno:Signature"; //AMT - 6643
        ///<summary>Image Options</summary>
        public const String SID_SIM_EDITOPTIONS = ".uno:SimEditOptions"; //AMT - 10438
        ///<summary>Size</summary>
        public const String SID_ATTR_SIZE = ".uno:Size"; //S - 10224
        ///<summary>Formula Options</summary>
        public const String SID_SM_EDITOPTIONS = ".uno:SmEditOptions"; //AMT - 10436
        ///<summary>Sort Descending</summary>
        public const String SID_FM_SORTDOWN = ".uno:SortDown"; //AMT - 10713
        ///<summary>Sort Ascending</summary>
        public const String SID_FM_SORTUP = ".uno:Sortup"; //AMT - 10712
        ///<summary>HTML Source</summary>
        public const String SID_SOURCEVIEW = ".uno:SourceView"; //AMT - 5675
        ///<summary>Line Spacing: 1</summary>
        public const String SID_ATTR_PARA_LINESPACE_10 = ".uno:SpacePara1"; //AMT - 10034
        ///<summary>Line Spacing : 1.5</summary>
        public const String SID_ATTR_PARA_LINESPACE_15 = ".uno:SpacePara15"; //AMT - 10035
        ///<summary>Line Spacing : 2</summary>
        public const String SID_ATTR_PARA_LINESPACE_20 = ".uno:SpacePara2"; //AMT - 10036
        ///<summary>Spellcheck</summary>
        public const String SID_SPELL_DIALOG = ".uno:SpellDialog"; //AMT - 10243
        ///<summary>AutoSpellcheck</summary>
        public const String SID_AUTOSPELL_CHECK = ".uno:SpellOnline"; //AMT - 12021
        ///<summary>Spin Button</summary>
        public const String SID_FM_SPINBUTTON = ".uno:SpinButton"; //AMT - 10769
        ///<summary>Spin Button</summary>
        public const String SID_INSERT_SPINBUTTON = ".uno:Spinbutton"; //AMT - 10193
        ///<summary>Split Frame Horizontally</summary>
        public const String SID_SPLIT_HORIZONTAL = ".uno:SplitHorizontal"; //AMT - 5647
        ///<summary>Split FrameSet Horizontally</summary>
        public const String SID_SPLIT_PARENT_HORIZONTAL = ".uno:SplitParentHorizontal"; //AMT - 5649
        ///<summary>Split FrameSet Vertically</summary>
        public const String SID_SPLIT_PARENT_VERTICAL = ".uno:SplitParentVertical"; //AMT - 5650
        ///<summary>Split Frame Vertically</summary>
        public const String SID_SPLIT_VERTICAL = ".uno:SplitVertical"; //AMT - 5648
        ///<summary>Square</summary>
        public const String SID_DRAW_SQUARE = ".uno:Square"; //AMT - 10380
        ///<summary>Rounded Square</summary>
        public const String SID_DRAW_SQUARE_ROUND = ".uno:Square_Rounded"; //AMT - 10381
        ///<summary>Rounded Square, Unfilled</summary>
        public const String SID_DRAW_SQUARE_ROUND_NOFILL = ".uno:Square_Rounded_Unfilled"; //AMT - 10383
        ///<summary>Square, Unfilled</summary>
        public const String SID_DRAW_SQUARE_NOFILL = ".uno:Square_Unfilled"; //AMT - 10382
        ///<summary>Stars</summary>
        public const String SID_DRAWTBX_CS_STAR = ".uno:StarShapes"; //AMT - 11052
        ///<summary>Cell</summary>
        public const String SID_TABLE_CELL = ".uno:StateTableCell"; //S - 10225
        ///<summary>Status Bar</summary>
        public const String SID_TOGGLESTATUSBAR = ".uno:StatusBarVisible"; //AMT - 5920
        ///<summary>Position</summary>
        public const String SID_BASICIDE_STAT_POS = ".uno:StatusGetPosition"; //S - 30806
        ///<summary>Current Basic Module</summary>
        public const String SID_BASICIDE_STAT_TITLE = ".uno:StatusGetTitle"; //S - 30808
        ///<summary>Stop Loading</summary>
        public const String SID_BROWSE_STOP = ".uno:Stop"; //AMT - 6302
        ///<summary>Stop Recording</summary>
        public const String SID_STOP_RECORDING = ".uno:StopRecording"; //AMT - 6671
        ///<summary>Strikethrough</summary>
        public const String SID_ATTR_CHAR_STRIKEOUT = ".uno:Strikeout"; //AMT - 10013
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_STYLE_APPLY = ".uno:StyleApplyState"; //AMT - 5552
        ///<summary>Catalog</summary>
        public const String SID_STYLE_CATALOG = ".uno:StyleCatalog"; //AMT - 5573
        ///<summary>New Style from Selection</summary>
        public const String SID_STYLE_NEW_BY_EXAMPLE = ".uno:StyleNewByExample"; //AMT - 5555
        ///<summary>Update Style</summary>
        public const String SID_STYLE_UPDATE_BY_EXAMPLE = ".uno:StyleUpdateByExample"; //AMT - 5556
        ///<summary>Subscript</summary>
        public const String SID_SET_SUB_SCRIPT = ".uno:SubScript"; //AMT - 10295
        ///<summary>Subtract</summary>
        public const String SID_POLY_SUBSTRACT = ".uno:Substract"; //AMT - 5680
        ///<summary>Superscript</summary>
        public const String SID_SET_SUPER_SCRIPT = ".uno:SuperScript"; //AMT - 10294
        ///<summary>Text Document Options</summary>
        public const String SID_SW_EDITOPTIONS = ".uno:SwEditOptions"; //AMT - 10433
        ///<summary>Design Mode On/Off</summary>
        public const String SID_FM_DESIGN_MODE = ".uno:SwitchControlDesignMode"; //AMT - 10629
        ///<summary>Symbol Shapes</summary>
        public const String SID_DRAWTBX_CS_SYMBOL = ".uno:SymbolShapes"; //AMT - 11048
        ///<summary>Activation Order</summary>
        public const String SID_FM_TAB_DIALOG = ".uno:TabDialog"; //AMT - 10615
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_TASKBAR = ".uno:TaskBarVisible"; //AMT - 5931
        ///<summary>Test Mode On/Off</summary>
        public const String SID_DIALOG_TESTMODE = ".uno:TestMode"; //AMT - 10199
        ///<summary>Text</summary>
        public const String SID_ATTR_CHAR = ".uno:Text"; //AMT - 10006
        ///<summary>Fit to Frame</summary>
        public const String SID_ATTR_TEXT_FITTOSIZE = ".uno:TextFitToSize"; //AMT - 10367
        ///<summary>Text Animation</summary>
        public const String SID_DRAW_TEXT_MARQUEE = ".uno:Text_Marquee"; //AMT - 10465
        ///<summary>Text direction from left to right</summary>
        public const String SID_TEXTDIRECTION_LEFT_TO_RIGHT = ".uno:TextdirectionLeftToRight"; //AMT - 10907
        ///<summary>Text direction from top to bottom</summary>
        public const String SID_TEXTDIRECTION_TOP_TO_BOTTOM = ".uno:TextdirectionTopToBottom"; //AMT - 10908
        ///<summary>Thesaurus</summary>
        public const String SID_THESAURUS = ".uno:Thesaurus"; //AMT - 10245
        ///<summary>Time Field</summary>
        public const String SID_FM_TIMEFIELD = ".uno:TimeField"; //AMT - 10705
        ///<summary>Breakpoint On/Off</summary>
        public const String SID_BASICIDE_TOGGLEBRKPNT = ".uno:ToggleBreakPoint"; //AMT - 30768
        ///<summary>Breakpoint Enabled/Disabled</summary>
        public const String SID_BASICIDE_TOGGLEBRKPNTENABLED = ".uno:ToggleBreakPointEnabled"; //AMT - 30811
        ///<summary>Points</summary>
        public const String SID_BEZIER_EDIT = ".uno:ToggleObjectBezierMode"; //AMT - 10126
        ///<summary>Rotate</summary>
        public const String SID_OBJECT_ROTATE = ".uno:ToggleObjectRotateMode"; //AMT - 10129
        ///<summary>Edit Macros</summary>
        public const String SID_EDITMACRO = ".uno:ToolsMacroEdit"; //AMT - 5802
        ///<summary>Position and Size</summary>
        public const String SID_ATTR_TRANSFORM = ".uno:TransformDialog"; //AMT - 10087
        ///<summary>Select Source</summary>
        public const String SID_TWAIN_SELECT = ".uno:TwainSelect"; //AMT - 10331
        ///<summary>Request</summary>
        public const String SID_TWAIN_TRANSFER = ".uno:TwainTransfer"; //AMT - 10332
        ///<summary>URL Button</summary>
        public const String SID_INSERT_URLBUTTON = ".uno:URLButton"; //AMT - 10197
        ///<summary>Underline</summary>
        public const String SID_ATTR_CHAR_UNDERLINE = ".uno:Underline"; //AMT - 10014
        ///<summary>Can't Undo</summary>
        public const String SID_UNDO = ".uno:Undo"; //AMT - 5701
        ///<summary>Ungroup</summary>
        public const String SID_OUTLINE_REMOVE = ".uno:Ungroup"; //AMT - 26332
        ///<summary>Wizards On/Off</summary>
        public const String SID_FM_USE_WIZARDS = ".uno:UseWizards"; //AMT - 10727
        ///<summary>Vertical Line</summary>
        public const String SID_INSERT_VFIXEDLINE = ".uno:VFixedLine"; //AMT - 10929
        ///<summary>Vertical Scroll Bar</summary>
        public const String SID_INSERT_VSCROLLBAR = ".uno:VScrollbar"; //AMT - 10195
        ///<summary>Versions</summary>
        public const String SID_VERSION = ".uno:VersionDialog"; //AMT - 6583
        ///<summary>Version Visible</summary>
        public const String SID_VERSION_VISIBLE = ".uno:VersionVisible"; //AMT - 5313
        ///<summary>Vertical Callouts</summary>
        public const String SID_DRAW_CAPTION_VERTICAL = ".uno:VerticalCaption"; //AMT - 10906
        ///<summary>Vertical Text</summary>
        public const String SID_DRAW_TEXT_VERTICAL = ".uno:VerticalText"; //AMT - 10905
        ///<summary>Data Sources</summary>
        public const String SID_VIEW_DATA_SOURCE_BROWSER = ".uno:ViewDataSourceBrowser"; //AMT - 6660
        ///<summary>Data source as Table</summary>
        public const String SID_FM_VIEW_AS_GRID = ".uno:ViewFormAsGrid"; //AMT - 10761
        ///<summary>3D Effects</summary>
        public const String SID_3D_WIN = ".uno:Window3D"; //AMT - 10644
        ///<summary>Line Color</summary>
        public const String SID_ATTR_LINE_COLOR = ".uno:XLineColor"; //AMT - 10172
        ///<summary>Line Style</summary>
        public const String SID_ATTR_LINE_STYLE = ".uno:XLineStyle"; //AMT - 10169
        ///<summary>Zoom</summary>
        public const String SID_ATTR_ZOOM = ".uno:Zoom"; //AMT - 10000
        ///<summary>Zoom 100</summary>
        public const String SID_SIZE_REAL = ".uno:Zoom100Percent"; //AMT - 
        ///<summary>Zoom Out</summary>
        public const String SID_ZOOM_IN = ".uno:ZoomMinus"; //AMT - 10098
        ///<summary>Zoom Next</summary>
        public const String SID_ZOOM_NEXT = ".uno:ZoomNext"; //AMT - 10402
        ///<summary>Object Zoom</summary>
        public const String SID_SIZE_OPTIMAL = ".uno:ZoomObjects"; //AMT - 27099
        ///<summary>Optimal</summary>
        public const String SID_SIZE_ALL = ".uno:ZoomOptimal"; //AMT - 10101
        ///<summary>Zoom Page</summary>
        public const String SID_SIZE_PAGE = ".uno:ZoomPage"; //AMT - 10100
        ///<summary>Zoom Page Width</summary>
        public const String SID_SIZE_PAGE_WIDTH = ".uno:ZoomPageWidth"; //AMT - 27098
        ///<summary>Zoom In</summary>
        public const String SID_ZOOM_OUT = ".uno:ZoomPlus"; //AMT - 10097
        ///<summary>Zoom Previous</summary>
        public const String SID_ZOOM_PREV = ".uno:ZoomPrevious"; //AMT - 10403
        ///<summary>Zoom</summary>
        public const String SID_ZOOM_TOOLBOX = ".uno:ZoomToolBox"; //AMT - 10096

    }

    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#"/>
    /// </summary>
    public static class DispatchURLs_WriterCommands
    {
        ///<summary>Accept or Reject</summary>
        public const String FN_REDLINE_ACCEPT = ".uno:AcceptTrackedChanges"; //AMT - 21829
        ///<summary>Add Unknown Words</summary>
        public const String FN_ADD_UNKNOWN = ".uno:AddAllUnknownWords"; //AMT - 20606
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_ALIGN_VERT_BOTTOM = ".uno:AlignBottom"; //AMT - 20479
        ///<summary>Align to Bottom of Character</summary>
        public const String FN_FRAME_ALIGN_VERT_CHAR_BOTTOM = ".uno:AlignCharBottom"; //AMT - 20569
        ///<summary>Align to Top of Character</summary>
        public const String FN_FRAME_ALIGN_VERT_CHAR_TOP = ".uno:AlignCharTop"; //AMT - 20568
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_ALIGN_HORZ_CENTER = ".uno:AlignHorizontalCenter"; //AMT - 20477
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_ALIGN_HORZ_LEFT = ".uno:AlignLeft"; //AMT - 20475
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_ALIGN_HORZ_RIGHT = ".uno:AlignRight"; //AMT - 20476
        ///<summary>Align to Bottom of Line</summary>
        public const String FN_FRAME_ALIGN_VERT_ROW_BOTTOM = ".uno:AlignRowBottom"; //AMT - 20566
        ///<summary>Align to Top of Line</summary>
        public const String FN_FRAME_ALIGN_VERT_ROW_TOP = ".uno:AlignRowTop"; //AMT - 20565
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_ALIGN_VERT_TOP = ".uno:AlignTop"; //AMT - 20478
        ///<summary>Align Vertical Center</summary>
        public const String FN_FRAME_ALIGN_VERT_CENTER = ".uno:AlignVerticalCenter"; //AMT - 20480
        ///<summary>Align to Vertical Center of Character</summary>
        public const String FN_FRAME_ALIGN_VERT_CHAR_CENTER = ".uno:AlignVerticalCharCenter"; //AMT - 20570
        ///<summary>Align to Vertical Center of Line</summary>
        public const String FN_FRAME_ALIGN_VERT_ROW_CENTER = ".uno:AlignVerticalRowCenter"; //AMT - 20567
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_APPLY_DEFAULT = ".uno:ApplyStyleDefault"; //AMT - 21757
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_APPLY_HEAD1 = ".uno:ApplyStyleHead1"; //AMT - 21754
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_APPLY_HEAD2 = ".uno:ApplyStyleHead2"; //AMT - 21755
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_APPLY_HEAD3 = ".uno:ApplyStyleHead3"; //AMT - 21756
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_APPLY_TEXTBODY = ".uno:ApplyStyleTextbody"; //AMT - 21758
        ///<summary>Bibliography Entry</summary>
        public const String FN_EDIT_AUTH_ENTRY_DLG = ".uno:AuthoritiesEntryDialog"; //AMT - 21833
        ///<summary>Apply</summary>
        public const String FN_AUTOFORMAT_APPLY = ".uno:AutoFormatApply"; //AMT - 20401
        ///<summary>Apply and Edit Changes</summary>
        public const String FN_AUTOFORMAT_REDLINE_APPLY = ".uno:AutoFormatRedlineApply"; //AMT - 20406
        ///<summary>Sum</summary>
        public const String FN_TABLE_AUTOSUM = ".uno:AutoSum"; //AMT - 20595
        ///<summary>Highlighting</summary>
        public const String SID_ATTR_CHAR_COLOR_BACKGROUND = ".uno:BackColor"; //AMT - 10489
        ///<summary>Background</summary>
        public const String FN_FORMAT_BACKGROUND_DLG = ".uno:BackgroundDialog"; //AMT - 20450
        ///<summary>Borders</summary>
        public const String FN_FORMAT_BORDER_DLG = ".uno:BorderDialog"; //AMT - 20448
        ///<summary>Bullets and Numbering</summary>
        public const String FN_NUMBER_BULLETS = ".uno:BulletsAndNumberingDialog"; //AMT - 20121
        ///<summary>Calculate Table</summary>
        public const String FN_CALC_TABLE = ".uno:Calc"; //AMT - 20129
        ///<summary>Calculate</summary>
        public const String FN_CALCULATE = ".uno:CalculateSel"; //AMT - 20615
        ///<summary>Bottom</summary>
        public const String FN_TABLE_VERT_BOTTOM = ".uno:CellVertBottom"; //AMT - 20587
        ///<summary>Center ( vertical )</summary>
        public const String FN_TABLE_VERT_CENTER = ".uno:CellVertCenter"; //AMT - 20586
        ///<summary>Top</summary>
        public const String FN_TABLE_VERT_NONE = ".uno:CellVertTop"; //AMT - 20585
        ///<summary>Link Frames</summary>
        public const String FN_FRAME_CHAIN = ".uno:ChainFrames"; //AMT - 21736
        ///<summary>Exchange Database</summary>
        public const String FN_CHANGE_DBFIELD = ".uno:ChangeDatabaseField"; //AMT - 20309
        ///<summary>Outline Numbering</summary>
        public const String FN_NUMBERING_OUTLINE_DLG = ".uno:ChapterNumberingDialog"; //AMT - 20612
        ///<summary>Highlight Fill</summary>
        public const String SID_ATTR_CHAR_COLOR_BACKGROUND_EXT = ".uno:CharBackgroundExt"; //AMT - 10490
        ///<summary>Font Color Fill</summary>
        public const String SID_ATTR_CHAR_COLOR_EXT = ".uno:CharColorExt"; //AMT - 10488
        ///<summary>Select Character Left</summary>
        public const String FN_CHAR_LEFT_SEL = ".uno:CharLeftSel"; //AMT - 20801
        ///<summary>Select Character Right</summary>
        public const String FN_CHAR_RIGHT_SEL = ".uno:CharRightSel"; //AMT - 20802
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_CLOSE_PAGEPREVIEW = ".uno:ClosePreview"; //AMT - 21254
        ///<summary>Comment</summary>
        public const String FN_REDLINE_COMMENT = ".uno:CommentChangeTracking"; //AMT - 21827
        ///<summary>Nonprinting Characters</summary>
        public const String FN_VIEW_META_CHARS = ".uno:ControlCodes"; //AMT - 20224
        ///<summary>Text <-> Table</summary>
        public const String FN_CONVERT_TEXT_TABLE = ".uno:ConvertTableText"; //AMT - 20500
        ///<summary>Table to Text</summary>
        public const String FN_CONVERT_TABLE_TO_TEXT = ".uno:ConvertTableToText"; //AMT - 20532
        ///<summary>Text to Table</summary>
        public const String FN_CONVERT_TEXT_TO_TABLE = ".uno:ConvertTextToTable"; //AMT - 20531
        ///<summary>Create AutoAbstract</summary>
        public const String FN_ABSTRACT_NEWDOC = ".uno:CreateAbstract"; //AMT - 21612
        ///<summary>Decrement Indent Value</summary>
        public const String FN_DEC_INDENT_OFFSET = ".uno:DecrementIndentValue"; //AMT - 21751
        ///<summary>Down One Level</summary>
        public const String FN_NUM_BULLET_DOWN = ".uno:DecrementLevel"; //AMT - 20130
        ///<summary>Move Down with Subpoints</summary>
        public const String FN_NUM_BULLET_OUTLINE_DOWN = ".uno:DecrementSubLevels"; //AMT - 20139
        ///<summary>Delete Row</summary>
        public const String FN_DELETE_WHOLE_LINE = ".uno:DelLine"; //AMT - 20935
        ///<summary>Delete to End of Line</summary>
        public const String FN_DELETE_LINE = ".uno:DelToEndOfLine"; //AMT - 20931
        ///<summary>Delete to End of Paragraph</summary>
        public const String FN_DELETE_PARA = ".uno:DelToEndOfPara"; //AMT - 20933
        ///<summary>Delete to End of Sentence</summary>
        public const String FN_DELETE_SENT = ".uno:DelToEndOfSentence"; //AMT - 20927
        ///<summary>Delete to End of Word</summary>
        public const String FN_DELETE_WORD = ".uno:DelToEndOfWord"; //AMT - 20929
        ///<summary>Delete to Start of Line</summary>
        public const String FN_DELETE_BACK_LINE = ".uno:DelToStartOfLine"; //AMT - 20932
        ///<summary>Delete to Start of Paragraph</summary>
        public const String FN_DELETE_BACK_PARA = ".uno:DelToStartOfPara"; //AMT - 20934
        ///<summary>Delete to Start of Sentence</summary>
        public const String FN_DELETE_BACK_SENT = ".uno:DelToStartOfSentence"; //AMT - 20928
        ///<summary>Delete to Start of Word</summary>
        public const String FN_DELETE_BACK_WORD = ".uno:DelToStartOfWord"; //AMT - 20930
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_DELETE_COL = ".uno:DeleteColumns"; //AMT - 20504
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_DELETE_ROW = ".uno:DeleteRows"; //AMT - 20503
        ///<summary>Table</summary>
        public const String FN_TABLE_DELETE_TABLE = ".uno:DeleteTable"; //AMT - 20529
        ///<summary>Distribute Columns Evenly</summary>
        public const String FN_TABLE_BALANCE_CELLS = ".uno:DistributeColumns"; //AMT - 20582
        ///<summary>Distribute Rows Equally</summary>
        public const String FN_TABLE_BALANCE_ROWS = ".uno:DistributeRows"; //AMT - 20583
        ///<summary>Edit index</summary>
        public const String FN_EDIT_CURRENT_TOX = ".uno:EditCurIndex"; //AMT - 21832
        ///<summary>Footnote</summary>
        public const String FN_EDIT_FOOTNOTE = ".uno:EditFootnote"; //AMT - 20162
        ///<summary>AutoText</summary>
        public const String FN_GLOSSARY_DLG = ".uno:EditGlossary"; //AMT - 20620
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_EDIT_HYPERLINK = ".uno:EditHyperlink"; //AMT - 21835
        ///<summary>Sections</summary>
        public const String FN_EDIT_REGION = ".uno:EditRegion"; //AMT - 20165
        ///<summary>Select to Document End</summary>
        public const String FN_END_OF_DOCUMENT_SEL = ".uno:EndOfDocumentSel"; //AMT - 20808
        ///<summary>Select to End of Line</summary>
        public const String FN_END_OF_LINE_SEL = ".uno:EndOfLineSel"; //AMT - 20806
        ///<summary>Select to Paragraph End</summary>
        public const String FN_END_OF_PARA_SEL = ".uno:EndOfParaSel"; //AMT - 20820
        ///<summary>Cells</summary>
        public const String FN_TABLE_SELECT_CELL = ".uno:EntireCell"; //AMT - 20530
        ///<summary>Select Column</summary>
        public const String FN_TABLE_SELECT_COL = ".uno:EntireColumn"; //AMT - 20514
        ///<summary>Select Rows</summary>
        public const String FN_TABLE_SELECT_ROW = ".uno:EntireRow"; //AMT - 20513
        ///<summary>Cancel</summary>
        public const String FN_ESCAPE = ".uno:Escape"; //AMT - 20941
        ///<summary>Hyperlinks Active</summary>
        public const String FN_STAT_HYPERLINKS = ".uno:ExecHyperlinks"; //S - 21186
        ///<summary>Run Macro Field</summary>
        public const String FN_EXECUTE_MACROFIELD = ".uno:ExecuteMacroField"; //AMT - 20127
        ///<summary>Run AutoText Entry</summary>
        public const String FN_EXPAND_GLOSSARY = ".uno:ExpandGlossary"; //AMT - 20628
        ///<summary>Fields</summary>
        public const String FN_EDIT_FIELD = ".uno:FieldDialog"; //AMT - 20104
        ///<summary>Field Names</summary>
        public const String FN_VIEW_FIELDNAME = ".uno:Fieldnames"; //AMT - 20226
        ///<summary>Fields</summary>
        public const String FN_VIEW_FIELDS = ".uno:Fields"; //AMT - 20215
        ///<summary>Flip Vertically</summary>
        public const String FN_FLIP_HORZ_GRAFIC = ".uno:FlipHorizontal"; //AMT - 20425
        ///<summary>Flip Horizontally</summary>
        public const String FN_FLIP_VERT_GRAFIC = ".uno:FlipVertical"; //AMT - 20426
        ///<summary>Font Color</summary>
        public const String SID_ATTR_CHAR_COLOR2 = ".uno:FontColor"; //AMT - 10537
        ///<summary>Footnotes</summary>
        public const String FN_FORMAT_FOOTNOTE_DLG = ".uno:FootnoteDialog"; //AMT - 20468
        ///<summary>Columns</summary>
        public const String FN_FORMAT_COLUMN = ".uno:FormatColumns"; //AMT - 20453
        ///<summary>Drop Caps</summary>
        public const String FN_FORMAT_DROPCAPS = ".uno:FormatDropcap"; //AMT - 20454
        ///<summary>Frame Properties</summary>
        public const String FN_FORMAT_FRAME_DLG = ".uno:FrameDialog"; //AMT - 20456
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_LINE_DOWN = ".uno:GoDown"; //AMT - 20904
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_CHAR_LEFT = ".uno:GoLeft"; //AMT - 20901
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_CHAR_RIGHT = ".uno:GoRight"; //AMT - 20902
        ///<summary>Set Cursor To Anchor</summary>
        public const String FN_FRAME_TO_ANCHOR = ".uno:GoToAnchor"; //AMT - 20959
        ///<summary>To Table End</summary>
        public const String FN_END_TABLE = ".uno:GoToEnd"; //AMT - 20948
        ///<summary>To Column End</summary>
        public const String FN_END_OF_COLUMN = ".uno:GoToEndOfColumn"; //AMT - 20918
        ///<summary>To Document End</summary>
        public const String FN_END_OF_DOCUMENT = ".uno:GoToEndOfDoc"; //AMT - 20908
        ///<summary>To End of Line</summary>
        public const String FN_END_OF_LINE = ".uno:GoToEndOfLine"; //AMT - 20906
        ///<summary>To End of Next Column</summary>
        public const String FN_END_OF_NEXT_COLUMN = ".uno:GoToEndOfNextColumn"; //AMT - 20952
        ///<summary>To End of Next Page</summary>
        public const String FN_END_OF_NEXT_PAGE = ".uno:GoToEndOfNextPage"; //AMT - 20910
        ///<summary>Select to End of Next Page</summary>
        public const String FN_END_OF_NEXT_PAGE_SEL = ".uno:GoToEndOfNextPageSel"; //AMT - 20810
        ///<summary>To Page End</summary>
        public const String FN_END_OF_PAGE = ".uno:GoToEndOfPage"; //AMT - 20914
        ///<summary>Select to Page End</summary>
        public const String FN_END_OF_PAGE_SEL = ".uno:GoToEndOfPageSel"; //AMT - 20814
        ///<summary>To Paragraph End</summary>
        public const String FN_END_OF_PARA = ".uno:GoToEndOfPara"; //AMT - 20920
        ///<summary>To Previous Column</summary>
        public const String FN_END_OF_PREV_COLUMN = ".uno:GoToEndOfPrevColumn"; //AMT - 20954
        ///<summary>To End of Previous Page</summary>
        public const String FN_END_OF_PREV_PAGE = ".uno:GoToEndOfPrevPage"; //AMT - 20912
        ///<summary>Select to End of Previous Page</summary>
        public const String FN_END_OF_PREV_PAGE_SEL = ".uno:GoToEndOfPrevPageSel"; //AMT - 20812
        ///<summary>To Next Paragraph</summary>
        public const String FN_NEXT_PARA = ".uno:GoToNextPara"; //AMT - 20975
        ///<summary>To Next Sentence</summary>
        public const String FN_NEXT_SENT = ".uno:GoToNextSentence"; //AMT - 20923
        ///<summary>To Word Right</summary>
        public const String FN_NEXT_WORD = ".uno:GoToNextWord"; //AMT - 20921
        ///<summary>To Previous Paragraph</summary>
        public const String FN_PREV_PARA = ".uno:GoToPrevPara"; //AMT - 20974
        ///<summary>To Previous Sentence</summary>
        public const String FN_PREV_SENT = ".uno:GoToPrevSentence"; //AMT - 20924
        ///<summary>To Word Left</summary>
        public const String FN_PREV_WORD = ".uno:GoToPrevWord"; //AMT - 20922
        ///<summary>To Column Begin</summary>
        public const String FN_START_OF_COLUMN = ".uno:GoToStartOfColumn"; //AMT - 20917
        ///<summary>To Document Begin</summary>
        public const String FN_START_OF_DOCUMENT = ".uno:GoToStartOfDoc"; //AMT - 20907
        ///<summary>To Line Begin</summary>
        public const String FN_START_OF_LINE = ".uno:GoToStartOfLine"; //AMT - 20905
        ///<summary>To Begin of Next Column</summary>
        public const String FN_START_OF_NEXT_COLUMN = ".uno:GoToStartOfNextColumn"; //AMT - 20951
        ///<summary>To Begin of Next Page</summary>
        public const String FN_START_OF_NEXT_PAGE = ".uno:GoToStartOfNextPage"; //AMT - 20909
        ///<summary>Select to Begin of Next Page</summary>
        public const String FN_START_OF_NEXT_PAGE_SEL = ".uno:GoToStartOfNextPageSel"; //AMT - 20809
        ///<summary>To Page Begin</summary>
        public const String FN_START_OF_PAGE = ".uno:GoToStartOfPage"; //AMT - 20913
        ///<summary>Select to Page Begin</summary>
        public const String FN_START_OF_PAGE_SEL = ".uno:GoToStartOfPageSel"; //AMT - 20813
        ///<summary>To Paragraph Begin</summary>
        public const String FN_START_OF_PARA = ".uno:GoToStartOfPara"; //AMT - 20919
        ///<summary>To Begin of Previous Column</summary>
        public const String FN_START_OF_PREV_COLUMN = ".uno:GoToStartOfPrevColumn"; //AMT - 20953
        ///<summary>To Begin of Previous Page</summary>
        public const String FN_START_OF_PREV_PAGE = ".uno:GoToStartOfPrevPage"; //AMT - 20911
        ///<summary>Select to Begin of Previous Page</summary>
        public const String FN_START_OF_PREV_PAGE_SEL = ".uno:GoToStartOfPrevPageSel"; //AMT - 20811
        ///<summary>To Table Begin</summary>
        public const String FN_START_TABLE = ".uno:GoToStartOfTable"; //AMT - 20947
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_LINE_UP = ".uno:GoUp"; //AMT - 20903
        ///<summary>Go to Next Index Mark</summary>
        public const String FN_NEXT_TOXMARK = ".uno:GotoNextIndexMark"; //AMT - 20983
        ///<summary>To Next Input Field</summary>
        public const String FN_GOTO_NEXT_INPUTFLD = ".uno:GotoNextInputField"; //AMT - 20147
        ///<summary>To Next Object</summary>
        public const String FN_GOTO_NEXT_OBJ = ".uno:GotoNextObject"; //AMT - 20944
        ///<summary>To Next Placeholder</summary>
        public const String FN_GOTO_NEXT_MARK = ".uno:GotoNextPlacemarker"; //AMT - 20976
        ///<summary>Select to Next Sentence</summary>
        public const String FN_NEXT_SENT_SEL = ".uno:GotoNextSentenceSel"; //AMT - 20823
        ///<summary>Go to next table formula</summary>
        public const String FN_NEXT_TBLFML = ".uno:GotoNextTableFormula"; //AMT - 20985
        ///<summary>Go to next faulty table formula</summary>
        public const String FN_NEXT_TBLFML_ERR = ".uno:GotoNextWrongTableFormula"; //AMT - 20987
        ///<summary>To Page</summary>
        public const String FN_NAVIGATION_PI_GOTO_PAGE = ".uno:GotoPage"; //AMT - 20659
        ///<summary>Go to Previous Index Mark</summary>
        public const String FN_PREV_TOXMARK = ".uno:GotoPrevIndexMark"; //AMT - 20984
        ///<summary>To Previous Input Field</summary>
        public const String FN_GOTO_PREV_INPUTFLD = ".uno:GotoPrevInputField"; //AMT - 20148
        ///<summary>To Previous Object</summary>
        public const String FN_GOTO_PREV_OBJ = ".uno:GotoPrevObject"; //AMT - 20945
        ///<summary>To Previous Placeholder</summary>
        public const String FN_GOTO_PREV_MARK = ".uno:GotoPrevPlacemarker"; //AMT - 20977
        ///<summary>Select to Previous Sentence</summary>
        public const String FN_PREV_SENT_SEL = ".uno:GotoPrevSentenceSel"; //AMT - 20824
        ///<summary>Go to previous table formula</summary>
        public const String FN_PREV_TBLFML = ".uno:GotoPrevTableFormula"; //AMT - 20986
        ///<summary>Go to previous faulty table formula</summary>
        public const String FN_PREV_TBLFML_ERR = ".uno:GotoPrevWrongTableFormula"; //AMT - 20988
        ///<summary>Graphics On/Off</summary>
        public const String FN_VIEW_GRAPHIC = ".uno:Graphic"; //AMT - 20213
        ///<summary>Picture</summary>
        public const String FN_FORMAT_GRAFIC_DLG = ".uno:GraphicDialog"; //AMT - 20458
        ///<summary>Increase Font</summary>
        public const String FN_GROW_FONT_SIZE = ".uno:Grow"; //AMT - 20403
        ///<summary>Scroll Horizontal</summary>
        public const String FN_HSCROLLBAR = ".uno:HScroll"; //AMT - 20218
        ///<summary>Heading rows repeat</summary>
        public const String FN_TABLE_HEADLINE_REPEAT = ".uno:HeadingRowsRepeat"; //AMT - 20520
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_HYPHENATE_OPT_DLG = ".uno:Hyphenate"; //AMT - 20605
        ///<summary>Increment Indent Value</summary>
        public const String FN_INC_INDENT_OFFSET = ".uno:IncrementIndentValue"; //AMT - 21750
        ///<summary>Up One Level</summary>
        public const String FN_NUM_BULLET_UP = ".uno:IncrementLevel"; //AMT - 20131
        ///<summary>Move Up with Subpoints</summary>
        public const String FN_NUM_BULLET_OUTLINE_UP = ".uno:IncrementSubLevels"; //AMT - 20140
        ///<summary>Index Entry</summary>
        public const String FN_EDIT_IDX_ENTRY_DLG = ".uno:IndexEntryDialog"; //AMT - 20123
        ///<summary>Index Mark to Index</summary>
        public const String FN_IDX_MARK_TO_IDX = ".uno:IndexMarkToIndex"; //AMT - 20962
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_POSTIT = ".uno:InsertAnnotation"; //AMT - 20329
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_FLD_AUTHOR = ".uno:InsertAuthorField"; //AMT - 20398
        ///<summary>Bibliography Entry</summary>
        public const String FN_INSERT_AUTH_ENTRY_DLG = ".uno:InsertAuthoritiesEntry"; //AMT - 21421
        ///<summary>Bookmark</summary>
        public const String FN_INSERT_BOOKMARK = ".uno:InsertBookmark"; //AMT - 20302
        ///<summary>Manual Break</summary>
        public const String FN_INSERT_BREAK_DLG = ".uno:InsertBreak"; //AMT - 20304
        ///<summary>Caption</summary>
        public const String FN_INSERT_CAPTION = ".uno:InsertCaptionDialog"; //AMT - 20310
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_COLUMN_BREAK = ".uno:InsertColumnBreak"; //AMT - 20305
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_INSERT_COL = ".uno:InsertColumns"; //AMT - 20502
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_CTRL = ".uno:InsertCtrl"; //AMT - 20389
        ///<summary>Date</summary>
        public const String FN_INSERT_FLD_DATE = ".uno:InsertDateField"; //AMT - 20392
        ///<summary>Insert Endnote Directly</summary>
        public const String FN_INSERT_ENDNOTE = ".uno:InsertEndnote"; //AMT - 21418
        ///<summary>Envelope</summary>
        public const String FN_ENVELOP = ".uno:InsertEnvelope"; //AMT - 21050
        ///<summary>Other</summary>
        public const String FN_INSERT_FIELD = ".uno:InsertField"; //AMT - 20308
        ///<summary>Insert Fields</summary>
        public const String FN_INSERT_FIELD_CTRL = ".uno:InsertFieldCtrl"; //AMT - 20391
        ///<summary>Insert Footnote Directly</summary>
        public const String FN_INSERT_FOOTNOTE = ".uno:InsertFootnote"; //AMT - 20399
        ///<summary>Footnote</summary>
        public const String FN_INSERT_FOOTNOTE_DLG = ".uno:InsertFootnoteDialog"; //AMT - 20312
        ///<summary>Formula</summary>
        public const String FN_EDIT_FORMULA = ".uno:InsertFormula"; //AMT - 20128
        ///<summary>Frame</summary>
        public const String FN_INSERT_FRAME = ".uno:InsertFrame"; //AMT - 20334
        ///<summary>Insert Frame Manually</summary>
        public const String FN_INSERT_FRAME_INTERACT = ".uno:InsertFrameInteract"; //AMT - 20333
        ///<summary>Insert single-column frame manually</summary>
        public const String FN_INSERT_FRAME_INTERACT_NOCOL = ".uno:InsertFrameInteractNoColumns"; //AMT - 20336
        ///<summary>Horizontal Ruler</summary>
        public const String FN_INSERT_HRULER = ".uno:InsertGraphicRuler"; //AMT - 21411
        ///<summary>Insert Non-breaking Hyphen</summary>
        public const String FN_INSERT_HARDHYPHEN = ".uno:InsertHardHyphen"; //AMT - 20385
        ///<summary>Insert Hyperlink</summary>
        public const String FN_INSERT_HYPERLINK = ".uno:InsertHyperlinkDlg"; //AMT - 20314
        ///<summary>Entry</summary>
        public const String FN_INSERT_IDX_ENTRY_DLG = ".uno:InsertIndexesEntry"; //AMT - 20335
        ///<summary>Insert Manual Row Break</summary>
        public const String FN_INSERT_LINEBREAK = ".uno:InsertLinebreak"; //AMT - 20318
        ///<summary>Indexes and Tables</summary>
        public const String FN_INSERT_MULTI_TOX = ".uno:InsertMultiIndex"; //AMT - 21420
        ///<summary>Insert Unnumbered Entry</summary>
        public const String FN_NUM_BULLET_NONUM = ".uno:InsertNeutralParagraph"; //AMT - 20136
        ///<summary>Insert Non-breaking Space</summary>
        public const String FN_INSERT_HARD_SPACE = ".uno:InsertNonBreakingSpace"; //AMT - 20344
        ///<summary>Insert Object</summary>
        public const String FN_INSERT_OBJ_CTRL = ".uno:InsertObjCtrl"; //AMT - 20390
        ///<summary>Insert Other Objects</summary>
        public const String FN_INSERT_OBJECT_DLG = ".uno:InsertObjectDialog"; //AMT - 20322
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_SMA = ".uno:InsertObjectStarMath"; //AMT - 20369
        ///<summary>Page Count</summary>
        public const String FN_INSERT_FLD_PGCOUNT = ".uno:InsertPageCountField"; //AMT - 20395
        ///<summary>Footer</summary>
        public const String FN_INSERT_PAGEFOOTER = ".uno:InsertPageFooter"; //AMT - 21414
        ///<summary>Header</summary>
        public const String FN_INSERT_PAGEHEADER = ".uno:InsertPageHeader"; //AMT - 21413
        ///<summary>Page Number</summary>
        public const String FN_INSERT_FLD_PGNUMBER = ".uno:InsertPageNumberField"; //AMT - 20394
        ///<summary>Insert Manual Page Break</summary>
        public const String FN_INSERT_PAGEBREAK = ".uno:InsertPagebreak"; //AMT - 20323
        ///<summary>Insert Paragraph</summary>
        public const String FN_INSERT_BREAK = ".uno:InsertPara"; //AMT - 20303
        ///<summary>Cross-reference</summary>
        public const String FN_INSERT_REF_FIELD = ".uno:InsertReferenceField"; //AMT - 20313
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_INSERT_ROW = ".uno:InsertRows"; //AMT - 20501
        ///<summary>Script</summary>
        public const String FN_JAVAEDIT = ".uno:InsertScript"; //AMT - 21410
        ///<summary>Section</summary>
        public const String FN_INSERT_REGION = ".uno:InsertSection"; //AMT - 21419
        ///<summary>Insert Optional Hyphen</summary>
        public const String FN_INSERT_SOFT_HYPHEN = ".uno:InsertSoftHyphen"; //AMT - 20343
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_SYMBOL = ".uno:InsertSymbol"; //AMT - 20328
        ///<summary>Table</summary>
        public const String FN_INSERT_TABLE = ".uno:InsertTable"; //AMT - 20330
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_INSERT_FLD_TIME = ".uno:InsertTimeField"; //AMT - 20393
        ///<summary>Title</summary>
        public const String FN_INSERT_FLD_TITLE = ".uno:InsertTitleField"; //AMT - 20397
        ///<summary>Subject</summary>
        public const String FN_INSERT_FLD_TOPIC = ".uno:InsertTopicField"; //AMT - 20396
        ///<summary>To Next Paragraph in Level</summary>
        public const String FN_NUM_BULLET_NEXT = ".uno:JumpDownThisLevel"; //AMT - 20133
        ///<summary>Directly to Document End</summary>
        public const String FN_END_DOC_DIRECT = ".uno:JumpToEndOfDoc"; //AMT - 20979
        ///<summary>To Footer</summary>
        public const String FN_TO_FOOTER = ".uno:JumpToFooter"; //AMT - 20961
        ///<summary>Edit Footnote/Endnote</summary>
        public const String FN_TO_FOOTNOTE_AREA = ".uno:JumpToFootnoteArea"; //AMT - 20963
        ///<summary>To Footnote Anchor</summary>
        public const String FN_FOOTNOTE_TO_ANCHOR = ".uno:JumpToFootnoteOrAnchor"; //AMT - 20955
        ///<summary>To Header</summary>
        public const String FN_TO_HEADER = ".uno:JumpToHeader"; //AMT - 20960
        ///<summary>To Next Bookmark</summary>
        public const String FN_NEXT_BOOKMARK = ".uno:JumpToNextBookmark"; //AMT - 20168
        ///<summary>To Next Footnote</summary>
        public const String FN_NEXT_FOOTNOTE = ".uno:JumpToNextFootnote"; //AMT - 20956
        ///<summary>To Next Frame</summary>
        public const String FN_CNTNT_TO_NEXT_FRAME = ".uno:JumpToNextFrame"; //AMT - 20958
        ///<summary>To Next Section</summary>
        public const String FN_GOTO_NEXT_REGION = ".uno:JumpToNextRegion"; //AMT - 21609
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NEXT_TABLE = ".uno:JumpToNextTable"; //AMT - 20949
        ///<summary>To Previous Bookmark</summary>
        public const String FN_PREV_BOOKMARK = ".uno:JumpToPrevBookmark"; //AMT - 20169
        ///<summary>To Previous Footnote</summary>
        public const String FN_PREV_FOOTNOTE = ".uno:JumpToPrevFootnote"; //AMT - 20957
        ///<summary>To Previous Section</summary>
        public const String FN_GOTO_PREV_REGION = ".uno:JumpToPrevRegion"; //AMT - 21610
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_PREV_TABLE = ".uno:JumpToPrevTable"; //AMT - 20950
        ///<summary>To Reference</summary>
        public const String FN_GOTO_REFERENCE = ".uno:JumpToReference"; //AMT - 20166
        ///<summary>Directly to Document Begin</summary>
        public const String FN_START_DOC_DIRECT = ".uno:JumpToStartOfDoc"; //AMT - 20978
        ///<summary>To Previous Paragraph in Level</summary>
        public const String FN_NUM_BULLET_PREV = ".uno:JumpUpThisLevel"; //AMT - 20132
        ///<summary>Select Down</summary>
        public const String FN_LINE_DOWN_SEL = ".uno:LineDownSel"; //AMT - 20804
        ///<summary>Line Numbering</summary>
        public const String FN_LINE_NUMBERING_DLG = ".uno:LineNumberingDialog"; //AMT - 20602
        ///<summary>Select to Top Line</summary>
        public const String FN_LINE_UP_SEL = ".uno:LineUpSel"; //AMT - 20803
        ///<summary>Links</summary>
        public const String FN_EDIT_LINK_DLG = ".uno:LinkDialog"; //AMT - 20109
        ///<summary>Load Styles</summary>
        public const String SID_TEMPLATE_LOAD = ".uno:LoadStyles"; //AMT - 5663
        ///<summary>Mail Merge Wizard</summary>
        public const String FN_MAILMERGE_WIZARD = ".uno:MailMergeWizard"; //AMT - 20364
        ///<summary>Field Shadings</summary>
        public const String FN_VIEW_MARKS = ".uno:Marks"; //AMT - 20225
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_MERGE_CELLS = ".uno:MergeCells"; //AMT - 20506
        ///<summary>Mail Merge</summary>
        public const String FN_QRY_MERGE = ".uno:MergeDialog"; //AMT - 20367
        ///<summary>Merge Table</summary>
        public const String FN_TABLE_MERGE_TABLE = ".uno:MergeTable"; //AMT - 21752
        ///<summary>Flip Graphics on Even Pages</summary>
        public const String FN_GRAPHIC_MIRROR_ON_EVEN_PAGES = ".uno:MirrorGraphicOnEvenPages"; //AMT - 21741
        ///<summary>Mirror Object on Even Pages</summary>
        public const String FN_FRAME_MIRROR_ON_EVEN_PAGES = ".uno:MirrorOnEvenPages"; //AMT - 21740
        ///<summary>Move Down</summary>
        public const String FN_NUM_BULLET_MOVEDOWN = ".uno:MoveDown"; //AMT - 20135
        ///<summary>Move Down with Subpoints</summary>
        public const String FN_NUM_BULLET_OUTLINE_MOVEDOWN = ".uno:MoveDownSubItems"; //AMT - 20142
        ///<summary>Move Up</summary>
        public const String FN_NUM_BULLET_MOVEUP = ".uno:MoveUp"; //AMT - 20134
        ///<summary>Move Up with Subpoints</summary>
        public const String FN_NUM_BULLET_OUTLINE_MOVEUP = ".uno:MoveUpSubItems"; //AMT - 20141
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NAME_GROUP = ".uno:NameGroup"; //AMT - 21614
        ///<summary>Create Master Document</summary>
        public const String FN_NEW_GLOBAL_DOC = ".uno:NewGlobalDoc"; //AMT - 20004
        ///<summary>Create HTML Document</summary>
        public const String FN_NEW_HTML_DOC = ".uno:NewHtmlDoc"; //AMT - 20040
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_CURRENCY = ".uno:NumberFormatCurrency"; //AMT - 21727
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_DATE = ".uno:NumberFormatDate"; //AMT - 21725
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_TWODEC = ".uno:NumberFormatDecimal"; //AMT - 21723
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_PERCENT = ".uno:NumberFormatPercent"; //AMT - 21728
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_SCIENTIFIC = ".uno:NumberFormatScientific"; //AMT - 21724
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_STANDARD = ".uno:NumberFormatStandard"; //AMT - 21721
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_NUMBER_TIME = ".uno:NumberFormatTime"; //AMT - 21726
        ///<summary>Numbering On/Off</summary>
        public const String FN_NUM_OR_NONUM = ".uno:NumberOrNoNumber"; //AMT - 20146
        ///<summary>Restart Numbering</summary>
        public const String FN_NUMBER_NEWSTART = ".uno:NumberingStart"; //AMT - 21738
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_UP = ".uno:ObjectBackOne"; //AMT - 20522
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FRAME_DOWN = ".uno:ObjectForwardOne"; //AMT - 20523
        ///<summary>While Typing</summary>
        public const String FN_AUTOFORMAT_AUTO = ".uno:OnlineAutoFormat"; //AMT - 20402
        ///<summary>Optimize</summary>
        public const String FN_OPTIMIZE_TABLE = ".uno:OptimizeTable"; //AMT - 20510
        ///<summary>Page Columns</summary>
        public const String FN_FORMAT_PAGE_COLUMN_DLG = ".uno:PageColumnDialog"; //AMT - 20449
        ///<summary>Page Settings</summary>
        public const String FN_FORMAT_PAGE_DLG = ".uno:PageDialog"; //AMT - 20452
        ///<summary>Next Page</summary>
        public const String FN_PAGEDOWN = ".uno:PageDown"; //AMT - 20938
        ///<summary>Select to Next Page</summary>
        public const String FN_PAGEDOWN_SEL = ".uno:PageDownSel"; //AMT - 20830
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_CHANGE_PAGENUM = ".uno:PageOffset"; //AMT - 20634
        ///<summary>Apply Page Style</summary>
        public const String FN_SET_PAGE_STYLE = ".uno:PageStyleApply"; //AMT - 20493
        ///<summary>Page Style</summary>
        public const String FN_STAT_TEMPLATE = ".uno:PageStyleName"; //S - 21182
        ///<summary>Previous Page</summary>
        public const String FN_PAGEUP = ".uno:PageUp"; //AMT - 20937
        ///<summary>Select to Previous Page</summary>
        public const String FN_PAGEUP_SEL = ".uno:PageUpSel"; //AMT - 20829
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_PASTESPECIAL = ".uno:PasteSpecial"; //AMT - 20114
        ///<summary>Print options page view</summary>
        public const String FN_PREVIEW_PRINT_OPTIONS = ".uno:PreviewPrintOptions"; //AMT - 20250
        ///<summary>Preview Zoom</summary>
        public const String FN_PREVIEW_ZOOM = ".uno:PreviewZoom"; //AMT - 20251
        ///<summary>Print Layout</summary>
        public const String FN_PRINT_LAYOUT = ".uno:PrintLayout"; //AMT - 20237
        ///<summary>Print page view</summary>
        public const String FN_PRINT_PAGEPREVIEW = ".uno:PrintPagePreView"; //AMT - 21253
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_SET_READ_ONLY_CELLS = ".uno:Protect"; //AMT - 20517
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_REDLINE_PROTECT = ".uno:ProtectTraceChangeMode"; //AMT - 21823
        ///<summary>Restore View</summary>
        public const String FN_REFRESH_VIEW = ".uno:RefreshView"; //AMT - 20201
        ///<summary>Numbering Off</summary>
        public const String FN_NUM_BULLET_OFF = ".uno:RemoveBullets"; //AMT - 20137
        ///<summary>Remove Direct Character Formats</summary>
        public const String FN_REMOVE_DIRECT_CHAR_FORMATS = ".uno:RemoveDirectCharFormats"; //AMT - 21759
        ///<summary>Delete index</summary>
        public const String FN_REMOVE_CUR_TOX = ".uno:RemoveTableOf"; //AMT - 20655
        ///<summary>Page Formatting</summary>
        public const String FN_REPAGINATE = ".uno:Repaginate"; //AMT - 20161
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_REPEAT_SEARCH = ".uno:RepeatSearch"; //AMT - 20150
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_FORMAT_RESET = ".uno:ResetAttributes"; //AMT - 20469
        ///<summary>Unprotect sheet</summary>
        public const String FN_TABLE_UNSET_READ_ONLY = ".uno:ResetTableProtection"; //AMT - 20559
        ///<summary>Allow Row to Break Across Pages and Columns</summary>
        public const String FN_TABLE_ROW_SPLIT = ".uno:RowSplit"; //AMT - 21753
        ///<summary>Ruler</summary>
        public const String FN_RULER = ".uno:Ruler"; //AMT - 20211
        ///<summary>Select Table</summary>
        public const String FN_TABLE_SELECT_ALL = ".uno:SelectTable"; //AMT - 20515
        ///<summary>Select Paragraph</summary>
        public const String FN_SELECT_PARA = ".uno:SelectText"; //AMT - 20197
        ///<summary>Select Text</summary>
        public const String FN_READONLY_SELECTION_MODE = ".uno:SelectTextMode"; //AMT - 20989
        ///<summary>Select Word</summary>
        public const String FN_SELECT_WORD = ".uno:SelectWord"; //AMT - 20943
        ///<summary>Selection Mode</summary>
        public const String FN_STAT_SELMODE = ".uno:SelectionMode"; //S - 21185
        ///<summary>AutoAbstract to Presentation</summary>
        public const String FN_ABSTRACT_STARIMPRESS = ".uno:SendAbstractToStarImpress"; //AMT - 21613
        ///<summary>Outline to Clipboard</summary>
        public const String FN_OUTLINE_TO_CLIPBOARD = ".uno:SendOutlineToClipboard"; //AMT - 20037
        ///<summary>Outline to Presentation</summary>
        public const String FN_OUTLINE_TO_IMPRESS = ".uno:SendOutlineToStarImpress"; //AMT - 20036
        ///<summary>Anchor to Character</summary>
        public const String FN_TOOL_ANKER_AT_CHAR = ".uno:SetAnchorAtChar"; //AMT - 21412
        ///<summary>Anchor as Character</summary>
        public const String FN_TOOL_ANKER_CHAR = ".uno:SetAnchorToChar"; //AMT - 20384
        ///<summary>Anchor To Frame</summary>
        public const String FN_TOOL_ANKER_FRAME = ".uno:SetAnchorToFrame"; //AMT - 20366
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TOOL_ANKER_PAGE = ".uno:SetAnchorToPage"; //AMT - 20350
        ///<summary>Anchor To Paragraph</summary>
        public const String FN_TOOL_ANKER_PARAGRAPH = ".uno:SetAnchorToPara"; //AMT - 20351
        ///<summary>Extended Selection On</summary>
        public const String FN_SET_EXT_MODE = ".uno:SetExtSelection"; //AMT - 20940
        ///<summary>MultiSelection On</summary>
        public const String FN_SET_ADD_MODE = ".uno:SetMultiSelection"; //AMT - 20939
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_ADJUST_CELLS = ".uno:SetOptimalColumnWidth"; //AMT - 20521
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_OPTIMAL_HEIGHT = ".uno:SetOptimalRowHeight"; //AMT - 20528
        ///<summary>Row Height</summary>
        public const String FN_TABLE_SET_ROW_HEIGHT = ".uno:SetRowHeight"; //AMT - 20507
        ///<summary>Direct Cursor On/Off</summary>
        public const String FN_SHADOWCURSOR = ".uno:ShadowCursor"; //AMT - 22204
        ///<summary>Backspace</summary>
        public const String FN_SHIFT_BACKSPACE = ".uno:ShiftBackspace"; //AMT - 20942
        ///<summary>Book Preview</summary>
        public const String FN_SHOW_BOOKVIEW = ".uno:ShowBookview"; //AMT - 21255
        ///<summary>Hidden Paragraphs</summary>
        public const String FN_VIEW_HIDDEN_PARA = ".uno:ShowHiddenParagraphs"; //AMT - 20242
        ///<summary>Page Preview: Multiple Pages</summary>
        public const String FN_SHOW_MULTIPLE_PAGES = ".uno:ShowMultiplePages"; //AMT - 21252
        ///<summary>Show</summary>
        public const String FN_REDLINE_SHOW = ".uno:ShowTrackedChanges"; //AMT - 21826
        ///<summary>Page Preview: Two Pages</summary>
        public const String FN_SHOW_TWO_PAGES = ".uno:ShowTwoPages"; //AMT - 21251
        ///<summary>Reduce Font</summary>
        public const String FN_SHRINK_FONT_SIZE = ".uno:Shrink"; //AMT - 20404
        ///<summary>Sort</summary>
        public const String FN_SORTING_DLG = ".uno:SortDialog"; //AMT - 20614
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TABLE_SPLIT_CELLS = ".uno:SplitCell"; //AMT - 20505
        ///<summary>Split Table</summary>
        public const String FN_TABLE_SPLIT_TABLE = ".uno:SplitTable"; //AMT - 21742
        ///<summary>AutoCorrect</summary>
        public const String FN_AUTO_CORRECT = ".uno:StartAutoCorrect"; //AMT - 20649
        ///<summary>Select to Document Begin</summary>
        public const String FN_START_OF_DOCUMENT_SEL = ".uno:StartOfDocumentSel"; //AMT - 20807
        ///<summary>Select to Begin of Line</summary>
        public const String FN_START_OF_LINE_SEL = ".uno:StartOfLineSel"; //AMT - 20805
        ///<summary>Select to Paragraph Begin</summary>
        public const String FN_START_OF_PARA_SEL = ".uno:StartOfParaSel"; //AMT - 20819
        ///<summary>Page Number</summary>
        public const String FN_STAT_PAGE = ".uno:StatePageNumber"; //S - 21181
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_SET_SUB_SCRIPT = ".uno:SubScript"; //AMT - 20412
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_SET_SUPER_SCRIPT = ".uno:SuperScript"; //AMT - 20411
        ///<summary>Backspace</summary>
        public const String FN_BACKSPACE = ".uno:SwBackspace"; //AMT - 20926
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_XFORMS_DESIGN_MODE = ".uno:SwitchXFormsDesignMode"; //AMT - 22300
        ///<summary>Table Boundaries</summary>
        public const String FN_VIEW_TABLEGRID = ".uno:TableBoundaries"; //AMT - 20227
        ///<summary>Table Properties</summary>
        public const String FN_FORMAT_TABLE_DLG = ".uno:TableDialog"; //AMT - 20460
        ///<summary>Table: Fixed</summary>
        public const String FN_TABLE_MODE_FIX = ".uno:TableModeFix"; //AMT - 20589
        ///<summary>Table: Fixed, Proportional</summary>
        public const String FN_TABLE_MODE_FIX_PROP = ".uno:TableModeFixProp"; //AMT - 20590
        ///<summary>Table: Variable</summary>
        public const String FN_TABLE_MODE_VARIABLE = ".uno:TableModeVariable"; //AMT - 20591
        ///<summary>Number Format</summary>
        public const String FN_NUM_FORMAT_TABLE_DLG = ".uno:TableNumberFormatDialog"; //AMT - 20445
        ///<summary>Number Recognition</summary>
        public const String FN_SET_MODOPT_TBLNUMFMT = ".uno:TableNumberRecognition"; //AMT - 20252
        ///<summary>Sort</summary>
        public const String FN_TABLE_SORT_DIALOG = ".uno:TableSort"; //AMT - 20533
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_DRAWTEXT_ATTR_DLG = ".uno:TextAttributes"; //AMT - 20376
        ///<summary>Text Wrap</summary>
        public const String FN_DRAW_WRAP_DLG = ".uno:TextWrap"; //AMT - 20203
        ///<summary>Thesaurus</summary>
        public const String FN_THESAURUS_DLG = ".uno:ThesaurusDialog"; //AMT - 20603
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_TOOL_ANKER = ".uno:ToggleAnchorType"; //AMT - 20349
        ///<summary>Change Position</summary>
        public const String FN_TOOL_HIERARCHIE = ".uno:ToggleObjectLayer"; //AMT - 20352
        ///<summary>Record</summary>
        public const String FN_REDLINE_ON = ".uno:TrackChanges"; //AMT - 21825
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FN_UNDERLINE_DOUBLE = ".uno:UnderlineDouble"; //AMT - 20405
        ///<summary>Unlink Frames</summary>
        public const String FN_FRAME_UNCHAIN = ".uno:UnhainFrames"; //AMT - 21737
        ///<summary>Unprotect cells</summary>
        public const String FN_TABLE_UNSET_READ_ONLY_CELLS = ".uno:UnsetCellsReadOnly"; //AMT - 20519
        ///<summary>Update All</summary>
        public const String FN_UPDATE_ALL = ".uno:UpdateAll"; //AMT - 21828
        ///<summary>All Indexes and Tables</summary>
        public const String FN_UPDATE_TOX = ".uno:UpdateAllIndexes"; //AMT - 20653
        ///<summary>Links</summary>
        public const String FN_UPDATE_ALL_LINKS = ".uno:UpdateAllLinks"; //AMT - 21824
        ///<summary>All Charts</summary>
        public const String FN_UPDATE_CHARTS = ".uno:UpdateCharts"; //AMT - 21834
        ///<summary>Current Index</summary>
        public const String FN_UPDATE_CUR_TOX = ".uno:UpdateCurIndex"; //AMT - 20654
        ///<summary>Fields</summary>
        public const String FN_UPDATE_FIELDS = ".uno:UpdateFields"; //AMT - 20126
        ///<summary>Update Input Fields</summary>
        public const String FN_UPDATE_INPUTFIELDS = ".uno:UpdateInputFields"; //AMT - 20143
        ///<summary>Vertical Ruler</summary>
        public const String FN_VLINEAL = ".uno:VRuler"; //AMT - 20216
        ///<summary>Vertical Scroll Bar</summary>
        public const String FN_VSCROLLBAR = ".uno:VScroll"; //AMT - 20217
        ///<summary>Text Boundaries</summary>
        public const String FN_VIEW_BOUNDS = ".uno:ViewBounds"; //AMT - 20214
        ///<summary>Word Count</summary>
        public const String FN_WORDCOUNT_DIALOG = ".uno:WordCountDialog"; //AMT - 22298
        ///<summary>Select to Begin of Word</summary>
        public const String FN_PREV_WORD_SEL = ".uno:WordLeftSel"; //AMT - 20822
        ///<summary>Select to Word Right</summary>
        public const String FN_NEXT_WORD_SEL = ".uno:WordRightSel"; //AMT - 20821
        ///<summary>Wrap First Paragraph</summary>
        public const String FN_WRAP_ANCHOR_ONLY = ".uno:WrapAnchorOnly"; //AMT - 20581
        ///<summary>Wrap Contour On</summary>
        public const String FN_FRAME_WRAP_CONTOUR = ".uno:WrapContour"; //AMT - 20584
        ///<summary>Optimal Page Wrap</summary>
        public const String FN_FRAME_WRAP_IDEAL = ".uno:WrapIdeal"; //AMT - 20563
        ///<summary>Wrap Left</summary>
        public const String FN_FRAME_WRAP_LEFT = ".uno:WrapLeft"; //AMT - 20572
        ///<summary>Wrap Off</summary>
        public const String FN_FRAME_NOWRAP = ".uno:WrapOff"; //AMT - 20472
        ///<summary>Page Wrap</summary>
        public const String FN_FRAME_WRAP = ".uno:WrapOn"; //AMT - 20473
        ///<summary>Wrap Right</summary>
        public const String FN_FRAME_WRAP_RIGHT = ".uno:WrapRight"; //AMT - 20573
        ///<summary>Wrap Through</summary>
        public const String FN_FRAME_WRAPTHRU = ".uno:WrapThrough"; //AMT - 20474
        ///<summary>In Background</summary>
        public const String FN_FRAME_WRAPTHRU_TRANSP = ".uno:WrapThroughTransparent"; //AMT - 20564

    }

    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#"/>
    /// </summary>
    public static class DispatchURLs_CalcCommands
    {
        ///<summary>Accept or Reject</summary>
        public const String FID_CHG_ACCEPT = ".uno:AcceptChanges"; //AMT - 26258
        ///<summary>Append Sheet</summary>
        public const String FID_TAB_APPEND = ".uno:Add"; //AMT - 26350
        ///<summary>Add</summary>
        public const String SID_ADD_PRINTAREA = ".uno:AddPrintArea"; //AMT - 26651
        ///<summary>Adjust Scale</summary>
        public const String FID_ADJUST_PRINTZOOM = ".uno:AdjustPrintZoom"; //AMT - 26652
        ///<summary>Justified</summary>
        public const String SID_ALIGNBLOCK = ".uno:AlignBlock"; //AMT - 26374
        ///<summary>Align Bottom</summary>
        public const String SID_ALIGNBOTTOM = ".uno:AlignBottom"; //AMT - 26376
        ///<summary>Align Center Horizontally</summary>
        public const String SID_ALIGNCENTERHOR = ".uno:AlignHorizontalCenter"; //AMT - 26373
        ///<summary>Align Left</summary>
        public const String SID_ALIGNLEFT = ".uno:AlignLeft"; //AMT - 26371
        ///<summary>Align Right</summary>
        public const String SID_ALIGNRIGHT = ".uno:AlignRight"; //AMT - 26372
        ///<summary>Align Top</summary>
        public const String SID_ALIGNTOP = ".uno:AlignTop"; //AMT - 26375
        ///<summary>Align Center Vertically</summary>
        public const String SID_ALIGNCENTERVER = ".uno:AlignVCenter"; //AMT - 26377
        ///<summary>Assign Names</summary>
        public const String FID_APPLY_NAME = ".uno:ApplyNames"; //AMT - 26274
        ///<summary>Fill Mode</summary>
        public const String SID_DETECTIVE_FILLMODE = ".uno:AuditingFillMode"; //AMT - 26462
        ///<summary>AutoInput</summary>
        public const String FID_AUTOCOMPLETE = ".uno:AutoComplete"; //AMT - 26319
        ///<summary>AutoFill Data Series: automatic</summary>
        public const String FID_FILL_AUTO = ".uno:AutoFill"; //AMT - 26556
        ///<summary>AutoOutline</summary>
        public const String SID_AUTO_OUTLINE = ".uno:AutoOutline"; //AMT - 26333
        ///<summary>AutoRefresh</summary>
        public const String SID_DETECTIVE_AUTO = ".uno:AutoRefreshArrows"; //AMT - 26471
        ///<summary>AutoCalculate</summary>
        public const String FID_AUTO_CALC = ".uno:AutomaticCalculation"; //AMT - 26303
        ///<summary>Recalculate</summary>
        public const String FID_RECALC = ".uno:Calculate"; //AMT - 26304
        ///<summary>Recalculate Hard</summary>
        public const String FID_HARD_RECALC = ".uno:CalculateHard"; //AMT - 26318
        ///<summary>Cancel</summary>
        public const String SID_CANCEL = ".uno:Cancel"; //AMT - 26557
        ///<summary>Choose Themes</summary>
        public const String SID_CHOOSE_DESIGN = ".uno:ChooseDesign"; //AMT - 26082
        ///<summary>Remove Dependents</summary>
        public const String SID_DETECTIVE_DEL_SUCC = ".uno:ClearArrowDependents"; //AMT - 26459
        ///<summary>Remove Precedents</summary>
        public const String SID_DETECTIVE_DEL_PRED = ".uno:ClearArrowPrecedents"; //AMT - 26457
        ///<summary>Remove All Traces</summary>
        public const String SID_DETECTIVE_DEL_ALL = ".uno:ClearArrows"; //AMT - 26461
        ///<summary>Delete Contents</summary>
        public const String SID_DELETE_CONTENTS = ".uno:ClearContents"; //AMT - 26553
        ///<summary>Close Preview</summary>
        public const String SID_PREVIEW_CLOSE = ".uno:ClosePreview"; //AMT - 26503
        ///<summary>Width</summary>
        public const String FID_COL_WIDTH = ".uno:ColumnWidth"; //AMT - 26285
        ///<summary>Comments</summary>
        public const String FID_CHG_COMMENT = ".uno:CommentChange"; //AMT - 26259
        ///<summary>Conditional Formatting</summary>
        public const String SID_OPENDLG_CONDFRMT = ".uno:ConditionalFormatDialog"; //AMT - 26159
        ///<summary>Create</summary>
        public const String FID_USE_NAME = ".uno:CreateNames"; //AMT - 26273
        ///<summary>Refresh Range</summary>
        public const String SID_REFRESH_DBAREA = ".uno:DataAreaRefresh"; //AMT - 26643
        ///<summary>Consolidate</summary>
        public const String SID_OPENDLG_CONSOLIDATE = ".uno:DataConsolidate"; //AMT - 26150
        ///<summary>Start</summary>
        public const String SID_OPENDLG_PIVOTTABLE = ".uno:DataDataPilotRun"; //AMT - 26151
        ///<summary>AutoFilter</summary>
        public const String SID_AUTO_FILTER = ".uno:DataFilterAutoFilter"; //AMT - 26325
        ///<summary>Hide AutoFilter</summary>
        public const String SID_AUTOFILTER_HIDE = ".uno:DataFilterHideAutoFilter"; //AMT - 26341
        ///<summary>Remove Filter</summary>
        public const String SID_UNFILTER = ".uno:DataFilterRemoveFilter"; //AMT - 26326
        ///<summary>Advanced Filter</summary>
        public const String SID_SPECIAL_FILTER = ".uno:DataFilterSpecialFilter"; //AMT - 26324
        ///<summary>Standard Filter</summary>
        public const String SID_FILTER = ".uno:DataFilterStandardFilter"; //AMT - 26323
        ///<summary>Import Data</summary>
        public const String SID_IMPORT_DATA = ".uno:DataImport"; //AMT - 26335
        ///<summary>DataPilot Filter</summary>
        public const String SID_DP_FILTER = ".uno:DataPilotFilter"; //AMT - 26091
        ///<summary>Refresh Data Import</summary>
        public const String SID_REIMPORT_DATA = ".uno:DataReImport"; //AMT - 26336
        ///<summary>Selection List</summary>
        public const String SID_DATA_SELECT = ".uno:DataSelect"; //AMT - 26610
        ///<summary>Sort</summary>
        public const String SID_SORT = ".uno:DataSort"; //AMT - 26322
        ///<summary>Subtotals</summary>
        public const String SID_SUBTOTALS = ".uno:DataSubTotals"; //AMT - 26328
        ///<summary>Define Range</summary>
        public const String SID_DEFINE_DBNAME = ".uno:DefineDBName"; //AMT - 26320
        ///<summary>Labels</summary>
        public const String SID_DEFINE_COLROWNAMERANGES = ".uno:DefineLabelRange"; //AMT - 26629
        ///<summary>Define</summary>
        public const String FID_DEFINE_NAME = ".uno:DefineName"; //AMT - 26271
        ///<summary>Define</summary>
        public const String SID_DEFINE_PRINTAREA = ".uno:DefinePrintArea"; //AMT - 26602
        ///<summary>Delete Page Breaks</summary>
        public const String FID_DEL_MANUALBREAKS = ".uno:DeleteAllBreaks"; //AMT - 26650
        ///<summary>Delete Cells</summary>
        public const String FID_DELETE_CELL = ".uno:DeleteCell"; //AMT - 26222
        ///<summary>Column Break</summary>
        public const String FID_DEL_COLBRK = ".uno:DeleteColumnbreak"; //AMT - 26264
        ///<summary>Delete Columns</summary>
        public const String SID_DEL_COLS = ".uno:DeleteColumns"; //AMT - 26237
        ///<summary>Delete</summary>
        public const String SID_PIVOT_KILL = ".uno:DeletePivotTable"; //AMT - 26315
        ///<summary>Remove</summary>
        public const String SID_DELETE_PRINTAREA = ".uno:DeletePrintArea"; //AMT - 26603
        ///<summary>Row Break</summary>
        public const String FID_DEL_ROWBRK = ".uno:DeleteRowbreak"; //AMT - 26263
        ///<summary>Delete Rows</summary>
        public const String SID_DEL_ROWS = ".uno:DeleteRows"; //AMT - 26236
        ///<summary>Undo Selection</summary>
        public const String SID_SELECT_NONE = ".uno:Deselect"; //AMT - 26549
        ///<summary>Insert Chart</summary>
        public const String SID_DRAW_CHART = ".uno:DrawChart"; //AMT - 26071
        ///<summary>Headers & Footers</summary>
        public const String SID_HFEDIT = ".uno:EditHeaderAndFooter"; //AMT - 26235
        ///<summary>Links</summary>
        public const String SID_LINKS = ".uno:EditLinks"; //AMT - 26060
        ///<summary>Edit</summary>
        public const String SID_OPENDLG_EDIT_PRINTAREA = ".uno:EditPrintArea"; //AMT - 26605
        ///<summary>Euro Converter</summary>
        public const String SID_EURO_CONVERTER = ".uno:EuroConverter"; //AMT - 26083
        ///<summary>Down</summary>
        public const String FID_FILL_TO_BOTTOM = ".uno:FillDown"; //AMT - 26224
        ///<summary>Left</summary>
        public const String FID_FILL_TO_LEFT = ".uno:FillLeft"; //AMT - 26227
        ///<summary>Right</summary>
        public const String FID_FILL_TO_RIGHT = ".uno:FillRight"; //AMT - 26225
        ///<summary>Series</summary>
        public const String FID_FILL_SERIES = ".uno:FillSeries"; //AMT - 26229
        ///<summary>Sheet</summary>
        public const String FID_FILL_TAB = ".uno:FillTable"; //AMT - 26228
        ///<summary>Up</summary>
        public const String FID_FILL_TO_TOP = ".uno:FillUp"; //AMT - 26226
        ///<summary>First Page</summary>
        public const String SID_PREVIEW_FIRST = ".uno:FirstPage"; //AMT - 26498
        ///<summary>Sheet Area Input Field</summary>
        public const String FID_FOCUS_POSWND = ".uno:FocusCellAddress"; //AMT - 26645
        ///<summary>Input Line</summary>
        public const String SID_FOCUS_INPUTLINE = ".uno:FocusInputLine"; //AMT - 26089
        ///<summary>Cells</summary>
        public const String FID_CELL_FORMAT = ".uno:FormatCellDialog"; //AMT - 26280
        ///<summary>Freeze</summary>
        public const String SID_WINDOW_FIX = ".uno:FreezePanes"; //AMT - 26070
        ///<summary>Function List</summary>
        public const String FID_FUNCTION_BOX = ".uno:FunctionBox"; //AMT - 26248
        ///<summary>Function</summary>
        public const String SID_OPENDLG_FUNCTION = ".uno:FunctionDialog"; //AMT - 26152
        ///<summary>To Lower Block Margin</summary>
        public const String SID_CURSORBLKDOWN = ".uno:GoDownToEndOfData"; //AMT - 26536
        ///<summary>Select to Lower Block Margin</summary>
        public const String SID_CURSORBLKDOWN_SEL = ".uno:GoDownToEndOfDataSel"; //AMT - 26540
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_CURSORPAGELEFT_ = ".uno:GoLeftBlock"; //AMT - 26527
        ///<summary>To Left Block Margin</summary>
        public const String SID_CURSORBLKLEFT = ".uno:GoLeftToStartOfData"; //AMT - 26537
        ///<summary>Select to Left Block Margin</summary>
        public const String SID_CURSORBLKLEFT_SEL = ".uno:GoLeftToStartOfDataSel"; //AMT - 26541
        ///<summary>Page Right</summary>
        public const String SID_CURSORPAGERIGHT_ = ".uno:GoRightBlock"; //AMT - 26529
        ///<summary>Select to Page Right</summary>
        public const String SID_CURSORPAGERIGHT_SEL = ".uno:GoRightBlockSel"; //AMT - 26530
        ///<summary>To Right Block Margin</summary>
        public const String SID_CURSORBLKRIGHT = ".uno:GoRightToEndOfData"; //AMT - 26538
        ///<summary>Select to Right Block Margin</summary>
        public const String SID_CURSORBLKRIGHT_SEL = ".uno:GoRightToEndOfDataSel"; //AMT - 26542
        ///<summary>To Current Cell</summary>
        public const String SID_ALIGNCURSOR = ".uno:GoToCurrentCell"; //AMT - 26550
        ///<summary>To Upper Block Margin</summary>
        public const String SID_CURSORBLKUP = ".uno:GoUpToStartOfData"; //AMT - 26535
        ///<summary>Select to Upper Block Margin</summary>
        public const String SID_CURSORBLKUP_SEL = ".uno:GoUpToStartOfDataSel"; //AMT - 26539
        ///<summary>Goal Seek</summary>
        public const String SID_OPENDLG_SOLVE = ".uno:GoalSeekDialog"; //AMT - 26153
        ///<summary>Hide</summary>
        public const String FID_TABLE_HIDE = ".uno:Hide"; //AMT - 26289
        ///<summary>Hide</summary>
        public const String FID_COL_HIDE = ".uno:HideColumn"; //AMT - 26287
        ///<summary>Hide</summary>
        public const String FID_ROW_HIDE = ".uno:HideRow"; //AMT - 26283
        ///<summary>Hyphenation</summary>
        public const String SID_ENABLE_HYPHENATION = ".uno:Hyphenate"; //AMT - 26087
        ///<summary>Formula Bar</summary>
        public const String FID_TOGGLEINPUTLINE = ".uno:InputLineVisible"; //AMT - 26241
        ///<summary>Insert Cells</summary>
        public const String SID_TBXCTL_INSCELLS = ".uno:InsCellsCtrl"; //AMT - 26627
        ///<summary>Insert Object</summary>
        public const String SID_TBXCTL_INSOBJ = ".uno:InsObjCtrl"; //AMT - 26628
        ///<summary>Sheet</summary>
        public const String FID_INS_TABLE = ".uno:Insert"; //AMT - 26269
        ///<summary>Cells</summary>
        public const String FID_INS_CELL = ".uno:InsertCell"; //AMT - 26266
        ///<summary>Insert Cells Down</summary>
        public const String FID_INS_CELLSDOWN = ".uno:InsertCellsDown"; //AMT - 26278
        ///<summary>Insert Cells Right</summary>
        public const String FID_INS_CELLSRIGHT = ".uno:InsertCellsRight"; //AMT - 26279
        ///<summary>Column Break</summary>
        public const String FID_INS_COLBRK = ".uno:InsertColumnBreak"; //AMT - 26262
        ///<summary>Insert Columns</summary>
        public const String FID_INS_COLUMN = ".uno:InsertColumns"; //AMT - 26268
        ///<summary>Paste Special</summary>
        public const String FID_INS_CELL_CONTENTS = ".uno:InsertContents"; //AMT - 26265
        ///<summary>Insert</summary>
        public const String SID_TBXCTL_INSERT = ".uno:InsertCtrl"; //AMT - 26626
        ///<summary>Link to External Data</summary>
        public const String SID_EXTERNAL_SOURCE = ".uno:InsertExternalDataSource"; //AMT - 26085
        ///<summary>Insert</summary>
        public const String FID_INSERT_NAME = ".uno:InsertName"; //AMT - 26272
        ///<summary>Insert From Image Editor</summary>
        public const String SID_INSERT_SIMAGE = ".uno:InsertObjectStarImage"; //AMT - 26061
        ///<summary>Formula</summary>
        public const String SID_INSERT_SMATH = ".uno:InsertObjectStarMath"; //AMT - 26063
        ///<summary>Row Break</summary>
        public const String FID_INS_ROWBRK = ".uno:InsertRowBreak"; //AMT - 26261
        ///<summary>Insert Rows</summary>
        public const String FID_INS_ROW = ".uno:InsertRows"; //AMT - 26267
        ///<summary>Sheet From File</summary>
        public const String FID_INS_TABLE_EXT = ".uno:InsertSheetFromFile"; //AMT - 26275
        ///<summary>To Next Sheet</summary>
        public const String SID_NEXT_TABLE = ".uno:JumpToNextTable"; //AMT - 26543
        ///<summary>Select to Next Sheet</summary>
        public const String SID_NEXT_TABLE_SEL = ".uno:JumpToNextTableSel"; //AMT - 26561
        ///<summary>To Next Unprotected Cell</summary>
        public const String SID_NEXT_UNPROTECT = ".uno:JumpToNextUnprotected"; //AMT - 26545
        ///<summary>To Previous Sheet</summary>
        public const String SID_PREV_TABLE = ".uno:JumpToPrevTable"; //AMT - 26544
        ///<summary>Select to Previous Sheet</summary>
        public const String SID_PREV_TABLE_SEL = ".uno:JumpToPrevTableSel"; //AMT - 26562
        ///<summary>To Previous Unprotected Cell</summary>
        public const String SID_PREV_UNPROTECT = ".uno:JumpToPreviousUnprotected"; //AMT - 26546
        ///<summary>Last Page</summary>
        public const String SID_PREVIEW_LAST = ".uno:LastPage"; //AMT - 26499
        ///<summary>Merge Cells</summary>
        public const String FID_MERGE_ON = ".uno:MergeCells"; //AMT - 26293
        ///<summary>Move/Copy</summary>
        public const String FID_TAB_MOVE = ".uno:Move"; //AMT - 26348
        ///<summary>Rename Sheet</summary>
        public const String FID_TAB_RENAME = ".uno:Name"; //AMT - 26347
        ///<summary>Next Page</summary>
        public const String SID_PREVIEW_NEXT = ".uno:NextPage"; //AMT - 26496
        ///<summary>Normal</summary>
        public const String FID_NORMALVIEWMODE = ".uno:NormalViewMode"; //AMT - 26249
        ///<summary>Show Note</summary>
        public const String FID_NOTE_VISIBLE = ".uno:NoteVisible"; //AMT - 26630
        ///<summary>Number Format: Currency</summary>
        public const String SID_NUMBER_CURRENCY = ".uno:NumberFormatCurrency"; //AMT - 26045
        ///<summary>Number Format : Date</summary>
        public const String SID_NUMBER_DATE = ".uno:NumberFormatDate"; //AMT - 26053
        ///<summary>Number Format: Delete Decimal Place</summary>
        public const String SID_NUMBER_DECDEC = ".uno:NumberFormatDecDecimals"; //AMT - 26058
        ///<summary>Number Format: Decimal</summary>
        public const String SID_NUMBER_TWODEC = ".uno:NumberFormatDecimal"; //AMT - 26054
        ///<summary>Number Format: Add Decimal Place</summary>
        public const String SID_NUMBER_INCDEC = ".uno:NumberFormatIncDecimals"; //AMT - 26057
        ///<summary>Number Format: Percent</summary>
        public const String SID_NUMBER_PERCENT = ".uno:NumberFormatPercent"; //AMT - 26046
        ///<summary>Number Format: Exponential</summary>
        public const String SID_NUMBER_SCIENTIFIC = ".uno:NumberFormatScientific"; //AMT - 26055
        ///<summary>Number Format: Standard</summary>
        public const String SID_NUMBER_STANDARD = ".uno:NumberFormatStandard"; //AMT - 26052
        ///<summary>Number Format: Time</summary>
        public const String SID_NUMBER_TIME = ".uno:NumberFormatTime"; //AMT - 26056
        ///<summary>Flip Object Horizontally</summary>
        public const String SID_MIRROR_HORIZONTAL = ".uno:ObjectMirrorHorizontal"; //AMT - 26066
        ///<summary>Flip Vertically</summary>
        public const String SID_MIRROR_VERTICAL = ".uno:ObjectMirrorVertical"; //AMT - 26065
        ///<summary>Page</summary>
        public const String SID_FORMATPAGE = ".uno:PageFormatDialog"; //AMT - 26295
        ///<summary>Page Break Preview</summary>
        public const String FID_PAGEBREAKMODE = ".uno:PagebreakMode"; //AMT - 26247
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String FID_PASTE_CONTENTS = ".uno:PasteSpecial"; //AMT - 26220
        ///<summary>Previous Page</summary>
        public const String SID_PREVIEW_PREVIOUS = ".uno:PreviousPage"; //AMT - 26497
        ///<summary>Sheet</summary>
        public const String FID_PROTECT_TABLE = ".uno:Protect"; //AMT - 26306
        ///<summary>Protect Records</summary>
        public const String SID_CHG_PROTECT = ".uno:ProtectTraceChangeMode"; //AMT - 26084
        ///<summary>Refresh</summary>
        public const String SID_PIVOT_RECALC = ".uno:RecalcPivotTable"; //AMT - 26314
        ///<summary>Refresh Traces</summary>
        public const String SID_DETECTIVE_REFRESH = ".uno:RefreshArrows"; //AMT - 26470
        ///<summary>Delete</summary>
        public const String FID_DELETE_TABLE = ".uno:Remove"; //AMT - 26223
        ///<summary>Name Object</summary>
        public const String SID_RENAME_OBJECT = ".uno:RenameObject"; //AMT - 26088
        ///<summary>Rename</summary>
        public const String FID_TAB_MENU_RENAME = ".uno:RenameTable"; //AMT - 26346
        ///<summary>Repeat Search</summary>
        public const String FID_REPEAT_SEARCH = ".uno:RepeatSearch"; //AMT - 26612
        ///<summary>Default Formatting</summary>
        public const String SID_CELL_FORMAT_RESET = ".uno:ResetAttributes"; //AMT - 26067
        ///<summary>Reset Scale</summary>
        public const String FID_RESET_PRINTZOOM = ".uno:ResetPrintZoom"; //AMT - 26653
        ///<summary>Height</summary>
        public const String FID_ROW_HEIGHT = ".uno:RowHeight"; //AMT - 26281
        ///<summary>Scale Screen Display</summary>
        public const String FID_SCALE = ".uno:Scale"; //AMT - 26244
        ///<summary>Scenarios</summary>
        public const String SID_SCENARIOS = ".uno:ScenarioManager"; //AMT - 26312
        ///<summary>Select Array Formula</summary>
        public const String SID_MARKARRAYFORMULA = ".uno:SelectArrayFormula"; //AMT - 26560
        ///<summary>Select Column</summary>
        public const String SID_SELECT_COL = ".uno:SelectColumn"; //AMT - 26547
        ///<summary>Select Range</summary>
        public const String SID_SELECT_DB = ".uno:SelectDB"; //AMT - 26321
        ///<summary>Select Data Area</summary>
        public const String SID_MARKDATAAREA = ".uno:SelectData"; //AMT - 26551
        ///<summary>Select Row</summary>
        public const String SID_SELECT_ROW = ".uno:SelectRow"; //AMT - 26548
        ///<summary>Select Scenario</summary>
        public const String SID_SELECT_SCENARIO = ".uno:SelectScenario"; //AMT - 26378
        ///<summary>Select</summary>
        public const String SID_SELECT_TABLES = ".uno:SelectTables"; //AMT - 26090
        ///<summary>Set Input Mode</summary>
        public const String SID_SETINPUTMODE = ".uno:SetInputMode"; //AMT - 26552
        ///<summary>Optimal Width</summary>
        public const String FID_COL_OPT_WIDTH = ".uno:SetOptimalColumnWidth"; //AMT - 26286
        ///<summary>Optimal Column Width, direct</summary>
        public const String FID_COL_OPT_DIRECT = ".uno:SetOptimalColumnWidthDirect"; //AMT - 26299
        ///<summary>Optimal Height</summary>
        public const String FID_ROW_OPT_HEIGHT = ".uno:SetOptimalRowHeight"; //AMT - 26282
        ///<summary>Right-To-Left</summary>
        public const String FID_TAB_RTL = ".uno:SheetRightToLeft"; //AMT - 26352
        ///<summary>Show</summary>
        public const String FID_TABLE_SHOW = ".uno:Show"; //AMT - 26290
        ///<summary>Show</summary>
        public const String FID_CHG_SHOW = ".uno:ShowChanges"; //AMT - 26239
        ///<summary>Show</summary>
        public const String FID_COL_SHOW = ".uno:ShowColumn"; //AMT - 26288
        ///<summary>Trace Dependents</summary>
        public const String SID_DETECTIVE_ADD_SUCC = ".uno:ShowDependents"; //AMT - 26458
        ///<summary>Trace Error</summary>
        public const String SID_DETECTIVE_ADD_ERR = ".uno:ShowErrors"; //AMT - 26460
        ///<summary>Mark Invalid Data</summary>
        public const String SID_DETECTIVE_INVALID = ".uno:ShowInvalid"; //AMT - 26469
        ///<summary>Trace Precedents</summary>
        public const String SID_DETECTIVE_ADD_PRED = ".uno:ShowPrecedents"; //AMT - 26456
        ///<summary>Show</summary>
        public const String FID_ROW_SHOW = ".uno:ShowRow"; //AMT - 26284
        ///<summary>Enter References</summary>
        public const String WID_SIMPLE_REF = ".uno:SimpleReferenz"; //AMT - 25728
        ///<summary>Sort Ascending</summary>
        public const String SID_SORT_ASCENDING = ".uno:SortAscending"; //AMT - 26344
        ///<summary>Sort Descending</summary>
        public const String SID_SORT_DESCENDING = ".uno:SortDescending"; //AMT - 26343
        ///<summary>Split Cells</summary>
        public const String FID_MERGE_OFF = ".uno:SplitCell"; //AMT - 26294
        ///<summary>Split</summary>
        public const String SID_WINDOW_SPLIT = ".uno:SplitWindow"; //AMT - 26069
        ///<summary>Standard Text Attributes</summary>
        public const String SID_TEXT_STANDARD = ".uno:StandardTextAttributes"; //AMT - 26296
        ///<summary>Modify Chart Data Area</summary>
        public const String SID_OPENDLG_MODCHART = ".uno:StarChartDataDialog"; //AMT - 26158
        ///<summary>Chart</summary>
        public const String SID_OPENDLG_CHART = ".uno:StarChartDialog"; //AMT - 26155
        ///<summary>Position in Document</summary>
        public const String SID_STATUS_DOCPOS = ".uno:StatusDocPos"; //S - 26114
        ///<summary>Page Format</summary>
        public const String SID_STATUS_PAGESTYLE = ".uno:StatusPageStyle"; //S - 26115
        ///<summary>Selection Mode</summary>
        public const String SID_STATUS_SELMODE = ".uno:StatusSelectionMode"; //S - 26116
        ///<summary>Status Expanded Selection</summary>
        public const String SID_STATUS_SELMODE_ERG = ".uno:StatusSelectionModeExp"; //AMT - 26122
        ///<summary>Status Extended Selection</summary>
        public const String SID_STATUS_SELMODE_ERW = ".uno:StatusSelectionModeExt"; //AMT - 26123
        ///<summary>Multiple Operations</summary>
        public const String SID_OPENDLG_TABOP = ".uno:TableOperationDialog"; //AMT - 26154
        ///<summary>Select All Sheets</summary>
        public const String FID_TAB_SELECTALL = ".uno:TableSelectAll"; //AMT - 26349
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_DRAWTEXT_ATTR_DLG = ".uno:TextAttributes"; //AMT - 26297
        ///<summary>Change Anchor</summary>
        public const String SID_ANCHOR_TOGGLE = ".uno:ToggleAnchorType"; //AMT - 26412
        ///<summary>Merge Cells</summary>
        public const String FID_MERGE_TOGGLE = ".uno:ToggleMergeCells"; //AMT - 26581
        ///<summary>Relative/Absolute References</summary>
        public const String SID_TOGGLE_REL = ".uno:ToggleRelative"; //AMT - 26609
        ///<summary>Document</summary>
        public const String FID_PROTECT_DOC = ".uno:ToolProtectionDocument"; //AMT - 26307
        ///<summary>Spreadsheet Options</summary>
        public const String SID_SCOPTIONS = ".uno:ToolsOptions"; //AMT - 26309
        ///<summary>Record</summary>
        public const String FID_CHG_RECORD = ".uno:TraceChangeMode"; //AMT - 26238
        ///<summary>Underline: Dotted</summary>
        public const String SID_ULINE_VAL_DOTTED = ".uno:UnderlineDotted"; //AMT - 26649
        ///<summary>Underline: Double</summary>
        public const String SID_ULINE_VAL_DOUBLE = ".uno:UnderlineDouble"; //AMT - 26648
        ///<summary>Underline: Off</summary>
        public const String SID_ULINE_VAL_NONE = ".uno:UnderlineNone"; //AMT - 26646
        ///<summary>Underline: Single</summary>
        public const String SID_ULINE_VAL_SINGLE = ".uno:UnderlineSingle"; //AMT - 26647
        ///<summary>Redraw Chart</summary>
        public const String SID_UPDATECHART = ".uno:UpdateChart"; //AMT - 26013
        ///<summary>Validity</summary>
        public const String FID_VALIDATION = ".uno:Validation"; //AMT - 26625
        ///<summary>Column & Row Headers</summary>
        public const String FID_TOGGLEHEADERS = ".uno:ViewRowColumnHeaders"; //AMT - 26242
        ///<summary>Value Highlighting</summary>
        public const String FID_TOGGLESYNTAX = ".uno:ViewValueHighlighting"; //AMT - 26245
        ///<summary>Automatic Row Break</summary>
        public const String SID_ATTR_ALIGN_LINEBREAK = ".uno:WrapText"; //AMT - 10230
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_PREVIEW_ZOOMIN = ".uno:ZoomIn"; //AMT - 26501
        ///<summary>Zoom Out</summary>
        public const String SID_PREVIEW_ZOOMOUT = ".uno:ZoomOut"; //AMT - 26502

    }
    
    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#Draw.2FImpress_commands"/>
    /// </summary>
    public static class DispatchURLs_DrawImpressCommands
    {

        ///<summary>Allow Interaction</summary>
        public const String SID_ACTIONMODE = ".uno:ActionMode"; //AMT - 27060
        ///<summary>Effects</summary>
        public const String SID_OBJECT_CHOOSE_MODE = ".uno:AdvancedMode"; //AMT - 27095
        ///<summary>Interaction</summary>
        public const String SID_ANIMATION_EFFECTS = ".uno:AnimationEffects"; //AMT - 27063
        ///<summary>Animated Image</summary>
        public const String SID_ANIMATION_OBJECTS = ".uno:AnimationObjects"; //AMT - 27062
        ///<summary>Lines and Arrows</summary>
        public const String SID_DRAWTBX_ARROWS = ".uno:ArrowsToolbox"; //AMT - 27171
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_ASSIGN_LAYOUT = ".uno:AssignLayout"; //AMT - 27435
        ///<summary>Send Backward</summary>
        public const String SID_MOREBACK = ".uno:Backward"; //AMT - 27032
        ///<summary>In Front of Object</summary>
        public const String SID_BEFORE_OBJ = ".uno:BeforeObject"; //AMT - 27326
        ///<summary>Behind Object</summary>
        public const String SID_BEHIND_OBJ = ".uno:BehindObject"; //AMT - 27116
        ///<summary>Large Handles</summary>
        public const String SID_BIG_HANDLES = ".uno:BigHandles"; //AMT - 27168
        ///<summary>Break</summary>
        public const String SID_BREAK = ".uno:Break"; //AMT - 27094
        ///<summary>Special Character</summary>
        public const String SID_BULLET = ".uno:Bullet"; //AMT - 27019
        ///<summary>Insert Snap Point/Line</summary>
        public const String SID_CAPTUREPOINT = ".uno:CapturePoint"; //AMT - 27038
        ///<summary>To Curve</summary>
        public const String SID_CHANGEBEZIER = ".uno:ChangeBezier"; //AMT - 27036
        ///<summary>To Polygon</summary>
        public const String SID_CHANGEPOLYGON = ".uno:ChangePolygon"; //AMT - 27037
        ///<summary>Rotation Mode after Clicking Object</summary>
        public const String SID_CLICK_CHANGE_ROTATION = ".uno:ClickChangeRotation"; //AMT - 27170
        ///<summary>Close Master View</summary>
        public const String SID_CLOSE_MASTER_VIEW = ".uno:CloseMasterView"; //AMT - 27434
        ///<summary>Black & White View</summary>
        public const String SID_COLORVIEW = ".uno:ColorView"; //AMT - 27257
        ///<summary>Combine</summary>
        public const String SID_COMBINE = ".uno:Combine"; //AMT - 27026
        ///<summary>Cone</summary>
        public const String SID_3D_CONE = ".uno:Cone"; //AMT - 27299
        ///<summary>Connect</summary>
        public const String SID_CONNECT = ".uno:Connect"; //AMT - 27093
        ///<summary>Connector</summary>
        public const String SID_TOOL_CONNECTOR = ".uno:Connector"; //AMT - 27058
        ///<summary>Connector Ends with Arrow</summary>
        public const String SID_CONNECTOR_ARROW_END = ".uno:ConnectorArrowEnd"; //AMT - 27120
        ///<summary>Connector Starts with Arrow</summary>
        public const String SID_CONNECTOR_ARROW_START = ".uno:ConnectorArrowStart"; //AMT - 27119
        ///<summary>Connector with Arrows</summary>
        public const String SID_CONNECTOR_ARROWS = ".uno:ConnectorArrows"; //AMT - 27121
        ///<summary>Connector</summary>
        public const String SID_CONNECTION_DLG = ".uno:ConnectorAttributes"; //AMT - 27338
        ///<summary>Connector Ends with Circle</summary>
        public const String SID_CONNECTOR_CIRCLE_END = ".uno:ConnectorCircleEnd"; //AMT - 27123
        ///<summary>Connector Starts with Circle</summary>
        public const String SID_CONNECTOR_CIRCLE_START = ".uno:ConnectorCircleStart"; //AMT - 27122
        ///<summary>Connector with Circles</summary>
        public const String SID_CONNECTOR_CIRCLES = ".uno:ConnectorCircles"; //AMT - 27124
        ///<summary>Curved Connector</summary>
        public const String SID_CONNECTOR_CURVE = ".uno:ConnectorCurve"; //AMT - 27132
        ///<summary>Curved Connector Ends with Arrow</summary>
        public const String SID_CONNECTOR_CURVE_ARROW_END = ".uno:ConnectorCurveArrowEnd"; //AMT - 27134
        ///<summary>Curved Connector Starts with Arrow</summary>
        public const String SID_CONNECTOR_CURVE_ARROW_START = ".uno:ConnectorCurveArrowStart"; //AMT - 27133
        ///<summary>Curved Connector with Arrows</summary>
        public const String SID_CONNECTOR_CURVE_ARROWS = ".uno:ConnectorCurveArrows"; //AMT - 27135
        ///<summary>Curved Connector Ends with Circle</summary>
        public const String SID_CONNECTOR_CURVE_CIRCLE_END = ".uno:ConnectorCurveCircleEnd"; //AMT - 27137
        ///<summary>Curved Connector Starts with Circle</summary>
        public const String SID_CONNECTOR_CURVE_CIRCLE_START = ".uno:ConnectorCurveCircleStart"; //AMT - 27136
        ///<summary>Curved Connector with Circles</summary>
        public const String SID_CONNECTOR_CURVE_CIRCLES = ".uno:ConnectorCurveCircles"; //AMT - 27138
        ///<summary>Straight Connector</summary>
        public const String SID_CONNECTOR_LINE = ".uno:ConnectorLine"; //AMT - 27125
        ///<summary>Straight Connector ends with Arrow</summary>
        public const String SID_CONNECTOR_LINE_ARROW_END = ".uno:ConnectorLineArrowEnd"; //AMT - 27127
        ///<summary>Straight Connector starts with Arrow</summary>
        public const String SID_CONNECTOR_LINE_ARROW_START = ".uno:ConnectorLineArrowStart"; //AMT - 27126
        ///<summary>Straight Connector with Arrows</summary>
        public const String SID_CONNECTOR_LINE_ARROWS = ".uno:ConnectorLineArrows"; //AMT - 27128
        ///<summary>Straight Connector ends with Circle</summary>
        public const String SID_CONNECTOR_LINE_CIRCLE_END = ".uno:ConnectorLineCircleEnd"; //AMT - 27130
        ///<summary>Straight Connector starts with Circle</summary>
        public const String SID_CONNECTOR_LINE_CIRCLE_START = ".uno:ConnectorLineCircleStart"; //AMT - 27129
        ///<summary>Straight Connector with Circles</summary>
        public const String SID_CONNECTOR_LINE_CIRCLES = ".uno:ConnectorLineCircles"; //AMT - 27131
        ///<summary>Line Connector</summary>
        public const String SID_CONNECTOR_LINES = ".uno:ConnectorLines"; //AMT - 27139
        ///<summary>Line Connector Ends with Arrow</summary>
        public const String SID_CONNECTOR_LINES_ARROW_END = ".uno:ConnectorLinesArrowEnd"; //AMT - 27141
        ///<summary>Line Connector Starts with Arrow</summary>
        public const String SID_CONNECTOR_LINES_ARROW_START = ".uno:ConnectorLinesArrowStart"; //AMT - 27140
        ///<summary>Line Connector with Arrows</summary>
        public const String SID_CONNECTOR_LINES_ARROWS = ".uno:ConnectorLinesArrows"; //AMT - 27142
        ///<summary>Line Connector Ends with Circle</summary>
        public const String SID_CONNECTOR_LINES_CIRCLE_END = ".uno:ConnectorLinesCircleEnd"; //AMT - 27144
        ///<summary>Line Connector Starts with Circle</summary>
        public const String SID_CONNECTOR_LINES_CIRCLE_START = ".uno:ConnectorLinesCircleStart"; //AMT - 27143
        ///<summary>Line Connector with Circles</summary>
        public const String SID_CONNECTOR_LINES_CIRCLES = ".uno:ConnectorLinesCircles"; //AMT - 27145
        ///<summary>Connector</summary>
        public const String SID_DRAWTBX_CONNECTORS = ".uno:ConnectorToolbox"; //AMT - 27028
        ///<summary>To 3D</summary>
        public const String SID_CONVERT_TO_3D = ".uno:ConvertInto3D"; //AMT - 10648
        ///<summary>In 3D Rotation Object</summary>
        public const String SID_CONVERT_TO_3D_LATHE = ".uno:ConvertInto3DLathe"; //AMT - 27008
        ///<summary>To 3D Rotation Object</summary>
        public const String SID_CONVERT_TO_3D_LATHE_FAST = ".uno:ConvertInto3DLatheFast"; //AMT - 10649
        ///<summary>To Bitmap</summary>
        public const String SID_CONVERT_TO_BITMAP = ".uno:ConvertIntoBitmap"; //AMT - 27378
        ///<summary>To Metafile</summary>
        public const String SID_CONVERT_TO_METAFILE = ".uno:ConvertIntoMetaFile"; //AMT - 27379
        ///<summary>1 Bit Dithered</summary>
        public const String SID_CONVERT_TO_1BIT_MATRIX = ".uno:ConvertTo1BitMatrix"; //AMT - 27162
        ///<summary>1 Bit Threshold</summary>
        public const String SID_CONVERT_TO_1BIT_THRESHOLD = ".uno:ConvertTo1BitThreshold"; //AMT - 27161
        ///<summary>4 Bit color palette</summary>
        public const String SID_CONVERT_TO_4BIT_COLORS = ".uno:ConvertTo4BitColors"; //AMT - 27164
        ///<summary>4 Bit grayscales</summary>
        public const String SID_CONVERT_TO_4BIT_GRAYS = ".uno:ConvertTo4BitGrays"; //AMT - 27163
        ///<summary>8 Bit color palette</summary>
        public const String SID_CONVERT_TO_8BIT_COLORS = ".uno:ConvertTo8BitColors"; //AMT - 27166
        ///<summary>8 Bit Grayscales</summary>
        public const String SID_CONVERT_TO_8BIT_GRAYS = ".uno:ConvertTo8BitGrays"; //AMT - 27165
        ///<summary>24 Bit True Color</summary>
        public const String SID_CONVERT_TO_24BIT = ".uno:ConvertToTrueColor"; //AMT - 27167
        ///<summary>Duplicate</summary>
        public const String SID_COPYOBJECTS = ".uno:CopyObjects"; //AMT - 27004
        ///<summary>Set in Circle (perspective)</summary>
        public const String SID_OBJECT_CROOK_ROTATE = ".uno:CrookRotate"; //AMT - 27090
        ///<summary>Set to circle (slant)</summary>
        public const String SID_OBJECT_CROOK_SLANT = ".uno:CrookSlant"; //AMT - 27091
        ///<summary>Set in Circle (distort)</summary>
        public const String SID_OBJECT_CROOK_STRETCH = ".uno:CrookStretch"; //AMT - 27092
        ///<summary>Cube</summary>
        public const String SID_3D_CUBE = ".uno:Cube"; //AMT - 27296
        ///<summary>Custom Animation</summary>
        public const String SID_CUSTOM_ANIMATION_PANEL = ".uno:CustomAnimation"; //AMT - 27328
        ///<summary>Animation Schemes</summary>
        public const String SID_CUSTOM_ANIMATION_SCHEMES_PANEL = ".uno:CustomAnimationSchemes"; //AMT - 27333
        ///<summary>Custom Slide Show</summary>
        public const String SID_CUSTOMSHOW_DLG = ".uno:CustomShowDialog"; //AMT - 27365
        ///<summary>Cylinder</summary>
        public const String SID_3D_CYLINDER = ".uno:Cylinder"; //AMT - 27298
        ///<summary>Pyramid</summary>
        public const String SID_3D_PYRAMID = ".uno:Cyramid"; //AMT - 27300
        ///<summary>Delete</summary>
        public const String SID_DELETE_LAYER = ".uno:DeleteLayer"; //AMT - 27081
        ///<summary>Delete Master</summary>
        public const String SID_DELETE_MASTER_PAGE = ".uno:DeleteMasterPage"; //AMT - 27432
        ///<summary>Delete Slide</summary>
        public const String SID_DELETE_PAGE = ".uno:DeletePage"; //AMT - 27080
        ///<summary>Slide Sorter</summary>
        public const String SID_DIAMODE = ".uno:DiaMode"; //AMT - 27011
        ///<summary>Split</summary>
        public const String SID_DISMANTLE = ".uno:Dismantle"; //AMT - 27082
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_DISPLAY_MASTER_BACKGROUND = ".uno:DisplayMasterBackground"; //AMT - 27436
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_DISPLAY_MASTER_OBJECTS = ".uno:DisplayMasterObjects"; //AMT - 27437
        ///<summary>Double-click to edit Text</summary>
        public const String SID_DOUBLECLICK_TEXTEDIT = ".uno:DoubleClickTextEdit"; //AMT - 27169
        ///<summary>Drawing View</summary>
        public const String SID_DRAWINGMODE = ".uno:DrawingMode"; //AMT - 27009
        ///<summary>Duplicate Slide</summary>
        public const String SID_DUPLICATE_PAGE = ".uno:DuplicatePage"; //AMT - 27342
        ///<summary>Hyperlink</summary>
        public const String SID_EDIT_HYPERLINK = ".uno:EditHyperlink"; //AMT - 27382
        ///<summary>Ellipse</summary>
        public const String SID_DRAWTBX_ELLIPSES = ".uno:EllipseToolbox"; //AMT - 10400
        ///<summary>Expand Slide</summary>
        public const String SID_EXPAND_PAGE = ".uno:ExpandPage"; //AMT - 27343
        ///<summary>Contour Mode</summary>
        public const String SID_FILL_DRAFT = ".uno:FillDraft"; //AMT - 27147
        ///<summary>Bring Forward</summary>
        public const String SID_MOREFRONT = ".uno:Forward"; //AMT - 27031
        ///<summary>Glue Points</summary>
        public const String SID_GLUE_EDITMODE = ".uno:GlueEditMode"; //AMT - 27301
        ///<summary>Exit Direction</summary>
        public const String SID_GLUE_ESCDIR = ".uno:GlueEscapeDirection"; //AMT - 27304
        ///<summary>Exit Direction Bottom</summary>
        public const String SID_GLUE_ESCDIR_BOTTOM = ".uno:GlueEscapeDirectionBottom"; //AMT - 27317
        ///<summary>Exit Direction Left</summary>
        public const String SID_GLUE_ESCDIR_LEFT = ".uno:GlueEscapeDirectionLeft"; //AMT - 27314
        ///<summary>Exit Direction Right</summary>
        public const String SID_GLUE_ESCDIR_RIGHT = ".uno:GlueEscapeDirectionRight"; //AMT - 27315
        ///<summary>Exit Direction Top</summary>
        public const String SID_GLUE_ESCDIR_TOP = ".uno:GlueEscapeDirectionTop"; //AMT - 27316
        ///<summary>Glue Point Horizontal Center</summary>
        public const String SID_GLUE_HORZALIGN_CENTER = ".uno:GlueHorzAlignCenter"; //AMT - 27305
        ///<summary>Glue Point Horizontal Left</summary>
        public const String SID_GLUE_HORZALIGN_LEFT = ".uno:GlueHorzAlignLeft"; //AMT - 27306
        ///<summary>Glue Point Horizontal Right</summary>
        public const String SID_GLUE_HORZALIGN_RIGHT = ".uno:GlueHorzAlignRight"; //AMT - 27307
        ///<summary>Insert Glue Point</summary>
        public const String SID_GLUE_INSERT_POINT = ".uno:GlueInsertPoint"; //AMT - 27302
        ///<summary>Glue Point Relative</summary>
        public const String SID_GLUE_PERCENT = ".uno:GluePercent"; //AMT - 27303
        ///<summary>Glue Point Vertical Bottom</summary>
        public const String SID_GLUE_VERTALIGN_BOTTOM = ".uno:GlueVertAlignBottom"; //AMT - 27310
        ///<summary>Glue Point Vertical Center</summary>
        public const String SID_GLUE_VERTALIGN_CENTER = ".uno:GlueVertAlignCenter"; //AMT - 27308
        ///<summary>Glue Point Vertical Top</summary>
        public const String SID_GLUE_VERTALIGN_TOP = ".uno:GlueVertAlignTop"; //AMT - 27309
        ///<summary>Picture Placeholders</summary>
        public const String SID_GRAPHIC_DRAFT = ".uno:GraphicDraft"; //AMT - 27146
        ///<summary>Grid to Front</summary>
        public const String SID_GRID_FRONT = ".uno:GridFront"; //AMT - 27323
        ///<summary>Half-Sphere</summary>
        public const String SID_3D_HALF_SPHERE = ".uno:HalfSphere"; //AMT - 27313
        ///<summary>Simple Handles</summary>
        public const String SID_HANDLES_DRAFT = ".uno:HandlesDraft"; //AMT - 27150
        ///<summary>Handout Master</summary>
        public const String SID_HANDOUT_MASTERPAGE = ".uno:HandoutMasterPage"; //AMT - 27349
        ///<summary>Handout Page</summary>
        public const String SID_HANDOUTMODE = ".uno:HandoutMode"; //AMT - 27070
        ///<summary>Header and Footer</summary>
        public const String SID_HEADER_AND_FOOTER = ".uno:HeaderAndFooter"; //AMT - 27407
        ///<summary>Guides to Front</summary>
        public const String SID_HELPLINES_FRONT = ".uno:HelplinesFront"; //AMT - 27325
        ///<summary>Snap to Guides</summary>
        public const String SID_HELPLINES_USE = ".uno:HelplinesUse"; //AMT - 27152
        ///<summary>Display Guides</summary>
        public const String SID_HELPLINES_VISIBLE = ".uno:HelplinesVisible"; //AMT - 27324
        ///<summary>Show/Hide Slide</summary>
        public const String SID_HIDE_SLIDE = ".uno:HideSlide"; //AMT - 10161
        ///<summary>Hyphenation</summary>
        public const String SID_HYPHENATION = ".uno:Hyphenation"; //AMT - 27340
        ///<summary>File</summary>
        public const String SID_INSERTFILE = ".uno:ImportFromFile"; //AMT - 27015
        ///<summary>Author</summary>
        public const String SID_INSERT_FLD_AUTHOR = ".uno:InsertAuthorField"; //AMT - 27364
        ///<summary>Date and Time</summary>
        public const String SID_INSERT_DATE_TIME = ".uno:InsertDateAndTime"; //AMT - 27412
        ///<summary>Date (fixed)</summary>
        public const String SID_INSERT_FLD_DATE_FIX = ".uno:InsertDateFieldFix"; //AMT - 27358
        ///<summary>Date (variable)</summary>
        public const String SID_INSERT_FLD_DATE_VAR = ".uno:InsertDateFieldVar"; //AMT - 27357
        ///<summary>File Name</summary>
        public const String SID_INSERT_FLD_FILE = ".uno:InsertFileField"; //AMT - 27363
        ///<summary>Layer</summary>
        public const String SID_INSERTLAYER = ".uno:InsertLayer"; //AMT - 27043
        ///<summary>New Master</summary>
        public const String SID_INSERT_MASTER_PAGE = ".uno:InsertMasterPage"; //AMT - 27431
        ///<summary>Slide</summary>
        public const String SID_INSERTPAGE = ".uno:InsertPage"; //AMT - 27014
        ///<summary>Page Number</summary>
        public const String SID_INSERT_FLD_PAGE = ".uno:InsertPageField"; //AMT - 27361
        ///<summary>Page Number</summary>
        public const String SID_INSERT_PAGE_NUMBER = ".uno:InsertPageNumber"; //AMT - 27411
        ///<summary>Insert Slide Direct</summary>
        public const String SID_INSERTPAGE_QUICK = ".uno:InsertPageQuick"; //AMT - 27352
        ///<summary>Time (fixed)</summary>
        public const String SID_INSERT_FLD_TIME_FIX = ".uno:InsertTimeFieldFix"; //AMT - 27360
        ///<summary>Time (variable)</summary>
        public const String SID_INSERT_FLD_TIME_VAR = ".uno:InsertTimeFieldVar"; //AMT - 27359
        ///<summary>Insert</summary>
        public const String SID_DRAWTBX_INSERT = ".uno:InsertToolbox"; //AMT - 27318
        ///<summary>Gradient</summary>
        public const String SID_OBJECT_GRADIENT = ".uno:InteractiveGradient"; //AMT - 27101
        ///<summary>Transparency</summary>
        public const String SID_OBJECT_TRANSPARENCE = ".uno:InteractiveTransparence"; //AMT - 27100
        ///<summary>Layer</summary>
        public const String SID_LAYERMODE = ".uno:LayerMode"; //AMT - 27050
        ///<summary>Layout</summary>
        public const String SID_STATUS_LAYOUT = ".uno:LayoutStatus"; //S - 27087
        ///<summary>Exit All Groups</summary>
        public const String SID_LEAVE_ALL_GROUPS = ".uno:LeaveAllGroups"; //AMT - 27345
        ///<summary>Line with Arrow/Circle</summary>
        public const String SID_LINE_ARROW_CIRCLE = ".uno:LineArrowCircle"; //AMT - 27175
        ///<summary>Line Ends with Arrow</summary>
        public const String SID_LINE_ARROW_END = ".uno:LineArrowEnd"; //AMT - 27173
        ///<summary>Line with Arrow/Square</summary>
        public const String SID_LINE_ARROW_SQUARE = ".uno:LineArrowSquare"; //AMT - 27177
        ///<summary>Line Starts with Arrow</summary>
        public const String SID_LINE_ARROW_START = ".uno:LineArrowStart"; //AMT - 27172
        ///<summary>Line with Arrows</summary>
        public const String SID_LINE_ARROWS = ".uno:LineArrows"; //AMT - 27174
        ///<summary>Line with Circle/Arrow</summary>
        public const String SID_LINE_CIRCLE_ARROW = ".uno:LineCircleArrow"; //AMT - 27176
        ///<summary>Line Contour Only</summary>
        public const String SID_LINE_DRAFT = ".uno:LineDraft"; //AMT - 27149
        ///<summary>Line with Square/Arrow</summary>
        public const String SID_LINE_SQUARE_ARROW = ".uno:LineSquareArrow"; //AMT - 27178
        ///<summary>Curve</summary>
        public const String SID_DRAWTBX_LINES = ".uno:LineToolbox"; //AMT - 10401
        ///<summary>Links</summary>
        public const String SID_MANAGE_LINKS = ".uno:ManageLinks"; //AMT - 27005
        ///<summary>Master Elements</summary>
        public const String SID_MASTER_LAYOUTS = ".uno:MasterLayouts"; //AMT - 27408
        ///<summary>Master</summary>
        public const String SID_MASTERPAGE = ".uno:MasterPage"; //AMT - 27053
        ///<summary>Dimensions</summary>
        public const String SID_MEASURE_DLG = ".uno:MeasureAttributes"; //AMT - 27320
        ///<summary>Dimension Line</summary>
        public const String SID_DRAW_MEASURELINE = ".uno:MeasureLine"; //AMT - 27051
        ///<summary>Flip</summary>
        public const String SID_OBJECT_MIRROR = ".uno:Mirror"; //AMT - 27085
        ///<summary>Horizontally</summary>
        public const String SID_HORIZONTAL = ".uno:MirrorHorz"; //AMT - 27035
        ///<summary>Vertically</summary>
        public const String SID_VERTICAL = ".uno:MirrorVert"; //AMT - 27034
        ///<summary>Fields</summary>
        public const String SID_MODIFY_FIELD = ".uno:ModifyField"; //AMT - 27362
        ///<summary>Layer</summary>
        public const String SID_MODIFYLAYER = ".uno:ModifyLayer"; //AMT - 27048
        ///<summary>Slide Layout</summary>
        public const String SID_MODIFYPAGE = ".uno:ModifyPage"; //AMT - 27046
        ///<summary>Cross-fading</summary>
        public const String SID_POLYGON_MORPHING = ".uno:Morphing"; //AMT - 27319
        ///<summary>Name Object</summary>
        public const String SID_NAME_GROUP = ".uno:NameGroup"; //AMT - 27027
        ///<summary>Reset Routing</summary>
        public const String SID_CONNECTION_NEW_ROUTING = ".uno:NewRouting"; //AMT - 27341
        ///<summary>Notes Master</summary>
        public const String SID_NOTES_MASTERPAGE = ".uno:NotesMasterPage"; //AMT - 27350
        ///<summary>Notes Page</summary>
        public const String SID_NOTESMODE = ".uno:NotesMode"; //AMT - 27069
        ///<summary>Arrange</summary>
        public const String SID_POSITION = ".uno:ObjectPosition"; //AMT - 27022
        ///<summary>3D Objects</summary>
        public const String SID_DRAWTBX_3D_OBJECTS = ".uno:Objects3DToolbox"; //AMT - 27295
        ///<summary>Outline</summary>
        public const String SID_OUTLINEMODE = ".uno:OutlineMode"; //AMT - 27010
        ///<summary>Black and White</summary>
        public const String SID_OUTPUT_QUALITY_BLACKWHITE = ".uno:OutputQualityBlackWhite"; //AMT - 27368
        ///<summary>Color</summary>
        public const String SID_OUTPUT_QUALITY_COLOR = ".uno:OutputQualityColor"; //AMT - 27366
        ///<summary>High Contrast</summary>
        public const String SID_OUTPUT_QUALITY_CONTRAST = ".uno:OutputQualityContrast"; //AMT - 27400
        ///<summary>Grayscale</summary>
        public const String SID_OUTPUT_QUALITY_GRAYSCALE = ".uno:OutputQualityGrayscale"; //AMT - 27367
        ///<summary>Pack</summary>
        public const String SID_PACKNGO = ".uno:PackAndGo"; //AMT - 27380
        ///<summary>Normal</summary>
        public const String SID_PAGEMODE = ".uno:PageMode"; //AMT - 27049
        ///<summary>Page</summary>
        public const String SID_PAGESETUP = ".uno:PageSetup"; //AMT - 27002
        ///<summary>Slide/Layer</summary>
        public const String SID_STATUS_PAGE = ".uno:PageStatus"; //S - 27086
        ///<summary>Slides Per Row</summary>
        public const String SID_PAGES_PER_ROW = ".uno:PagesPerRow"; //AMT - 27284
        ///<summary>Decrease Spacing</summary>
        public const String SID_PARASPACE_DECREASE = ".uno:ParaspaceDecrease"; //AMT - 27347
        ///<summary>Increase Spacing</summary>
        public const String SID_PARASPACE_INCREASE = ".uno:ParaspaceIncrease"; //AMT - 27346
        ///<summary>Paste Special</summary>
        public const String SID_PASTE2 = ".uno:PasteClipboard"; //AMT - 27003
        ///<summary>Select Text Area Only</summary>
        public const String SID_PICK_THROUGH = ".uno:PickThrough"; //AMT - 27159
        ///<summary>Pixel Mode</summary>
        public const String SID_PIXELMODE = ".uno:PixelMode"; //AMT - 27021
        ///<summary>Polygon, filled</summary>
        public const String SID_DRAW_POLYGON = ".uno:Polygon"; //AMT - 10117
        ///<summary>Slide Show</summary>
        public const String SID_PRESENTATION = ".uno:Presentation"; //AMT - 10157
        ///<summary>Slide Show Settings</summary>
        public const String SID_PRESENTATION_DLG = ".uno:PresentationDialog"; //AMT - 27339
        ///<summary>Slide Design</summary>
        public const String SID_PRESENTATION_LAYOUT = ".uno:PresentationLayout"; //AMT - 27064
        ///<summary>Preview</summary>
        public const String SID_PREVIEW_WIN = ".uno:PreviewWindow"; //AMT - 27327
        ///<summary>Allow Quick Editing</summary>
        public const String SID_QUICKEDIT = ".uno:QuickEdit"; //AMT - 27158
        ///<summary>Rectangle</summary>
        public const String SID_DRAWTBX_RECTANGLES = ".uno:RectangleToolbox"; //AMT - 10399
        ///<summary>Rehearse Timings</summary>
        public const String SID_REHEARSE_TIMINGS = ".uno:RehearseTimings"; //AMT - 10159
        ///<summary>Rename</summary>
        public const String SID_RENAMELAYER = ".uno:RenameLayer"; //AMT - 27269
        ///<summary>Rename Master</summary>
        public const String SID_RENAME_MASTER_PAGE = ".uno:RenameMasterPage"; //AMT - 27433
        ///<summary>Rename Slide</summary>
        public const String SID_RENAMEPAGE = ".uno:RenamePage"; //AMT - 27268
        ///<summary>Reverse</summary>
        public const String SID_REVERSE_ORDER = ".uno:ReverseOrder"; //AMT - 27117
        ///<summary>Distort</summary>
        public const String SID_OBJECT_SHEAR = ".uno:Shear"; //AMT - 27107
        ///<summary>Shell</summary>
        public const String SID_3D_SHELL = ".uno:Shell3D"; //AMT - 27311
        ///<summary>Slide Transition</summary>
        public const String SID_SLIDE_TRANSITIONS_PANEL = ".uno:SlideChangeWindow"; //AMT - 27334
        ///<summary>Slide Master</summary>
        public const String SID_SLIDE_MASTERPAGE = ".uno:SlideMasterPage"; //AMT - 27348
        ///<summary>Snap to Page Margins</summary>
        public const String SID_SNAP_BORDER = ".uno:SnapBorder"; //AMT - 27155
        ///<summary>Snap to Object Border</summary>
        public const String SID_SNAP_FRAME = ".uno:SnapFrame"; //AMT - 27156
        ///<summary>Snap to Object Points</summary>
        public const String SID_SNAP_POINTS = ".uno:SnapPoints"; //AMT - 27157
        ///<summary>Create Object with Attributes</summary>
        public const String SID_SOLID_CREATE = ".uno:SolidCreate"; //AMT - 27151
        ///<summary>Sphere</summary>
        public const String SID_3D_SPHERE = ".uno:Sphere"; //AMT - 27297
        ///<summary>Summary Slide</summary>
        public const String SID_SUMMARY_PAGE = ".uno:SummaryPage"; //AMT - 27344
        ///<summary>Text</summary>
        public const String SID_TEXTATTR_DLG = ".uno:TextAttributes"; //AMT - 27281
        ///<summary>Text Placeholders</summary>
        public const String SID_TEXT_DRAFT = ".uno:TextDraft"; //AMT - 27148
        ///<summary>Fit Text to Frame</summary>
        public const String SID_TEXT_FITTOSIZE = ".uno:TextFitToSizeTool"; //AMT - 27285
        ///<summary>Text</summary>
        public const String SID_DRAWTBX_TEXT = ".uno:TextToolbox"; //AMT - 10398
        ///<summary>Title Slide Master</summary>
        public const String SID_TITLE_MASTERPAGE = ".uno:TitleMasterPage"; //AMT - 27351
        ///<summary>Torus</summary>
        public const String SID_3D_TORUS = ".uno:Torus"; //AMT - 27312
        ///<summary>Fit Vertical Text to Frame</summary>
        public const String SID_TEXT_FITTOSIZE_VERTICAL = ".uno:VerticalTextFitToSizeTool"; //AMT - 27286
        ///<summary>Shift</summary>
        public const String SID_ZOOM_PANNING = ".uno:ZoomPanning"; //AMT - 27017
        ///<summary>To Contour</summary>
        public const String SID_CONVERT_TO_CONTOUR = ".uno:convert_to_contour"; //AMT - 27381

    }

    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#"/>
    /// </summary>
    public static class DispatchURLs_ChartCommands
    {
        ///<summary>All Titles</summary>
        public const String SID_DIAGRAM_TITLE_ALL = ".uno:AllTitles"; //AMT - 30562
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_ROW_MOREBACK = ".uno:Backward"; //AMT - 30595
        ///<summary>Bar Width</summary>
        public const String SID_DIAGRAM_BARWIDTH = ".uno:BarWidth"; //AMT - 30591
        ///<summary>Chart Title</summary>
        public const String SID_CHART_TITLE = ".uno:ChartTitle"; //AMT - 30520
        ///<summary>Chart Type</summary>
        public const String SID_CONTEXT_TYPE = ".uno:ContextType"; //AMT - 30538
        ///<summary>Caption Type for Chart Data</summary>
        public const String SID_DIAGRAM_DATADESCRIPTION = ".uno:DataDescriptionType"; //AMT - 30588
        ///<summary>Data in Columns</summary>
        public const String SID_DATA_IN_COLUMNS = ".uno:DataInColumns"; //AMT - 30536
        ///<summary>Data in Rows</summary>
        public const String SID_DATA_IN_ROWS = ".uno:DataInRows"; //AMT - 30535
        ///<summary>Default Colors for Data Series</summary>
        public const String SID_DIAGRAM_DEFAULTCOLORS = ".uno:DefaultColors"; //AMT - 30590
        ///<summary>Chart Area</summary>
        public const String SID_DIAGRAM_AREA = ".uno:DiagramArea"; //AMT - 30526
        ///<summary>Secondary X Axis</summary>
        public const String SID_DIAGRAM_AXIS_A = ".uno:DiagramAxisA"; //AMT - 30616
        ///<summary>All Axes</summary>
        public const String SID_DIAGRAM_AXIS_ALL = ".uno:DiagramAxisAll"; //AMT - 30555
        ///<summary>Secondary Y Axis</summary>
        public const String SID_DIAGRAM_AXIS_B = ".uno:DiagramAxisB"; //AMT - 30617
        ///<summary>X Axis</summary>
        public const String SID_DIAGRAM_AXIS_X = ".uno:DiagramAxisX"; //AMT - 30552
        ///<summary>Y Axis</summary>
        public const String SID_DIAGRAM_AXIS_Y = ".uno:DiagramAxisY"; //AMT - 30553
        ///<summary>Z Axis</summary>
        public const String SID_DIAGRAM_AXIS_Z = ".uno:DiagramAxisZ"; //AMT - 30554
        ///<summary>Chart Data</summary>
        public const String SID_DIAGRAM_DATA = ".uno:DiagramData"; //AMT - 30514
        ///<summary>Chart Floor</summary>
        public const String SID_DIAGRAM_FLOOR = ".uno:DiagramFloor"; //AMT - 30525
        ///<summary>Edit Grid</summary>
        public const String SID_DIAGRAM_GRID = ".uno:DiagramGrid"; //AMT - 30523
        ///<summary>All Axis Grids</summary>
        public const String SID_DIAGRAM_GRID_ALL = ".uno:DiagramGridAll"; //AMT - 30566
        ///<summary>Y Axis Minor Grid</summary>
        public const String SID_DIAGRAM_GRID_X_HELP = ".uno:DiagramGridXHelp"; //AMT - 30578
        ///<summary>Y Axis Main Grid</summary>
        public const String SID_DIAGRAM_GRID_X_MAIN = ".uno:DiagramGridXMain"; //AMT - 30563
        ///<summary>X Axis Minor Grid</summary>
        public const String SID_DIAGRAM_GRID_Y_HELP = ".uno:DiagramGridYHelp"; //AMT - 30579
        ///<summary>X Axis Main Grid</summary>
        public const String SID_DIAGRAM_GRID_Y_MAIN = ".uno:DiagramGridYMain"; //AMT - 30564
        ///<summary>Z Axis Minor Grid</summary>
        public const String SID_DIAGRAM_GRID_Z_HELP = ".uno:DiagramGridZHelp"; //AMT - 30580
        ///<summary>Z Axis Main Grid</summary>
        public const String SID_DIAGRAM_GRID_Z_MAIN = ".uno:DiagramGridZMain"; //AMT - 30565
        ///<summary>Object Properties</summary>
        public const String SID_DIAGRAM_OBJECTS = ".uno:DiagramObjects"; //AMT - 30572
        ///<summary>Chart Type</summary>
        public const String SID_DIAGRAM_TYPE = ".uno:DiagramType"; //AMT - 30528
        ///<summary>Chart Wall</summary>
        public const String SID_DIAGRAM_WALL = ".uno:DiagramWall"; //AMT - 30524
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_ROW_MOREFRONT = ".uno:Forward"; //AMT - 30594
        ///<summary>Axes</summary>
        public const String SID_INSERT_AXIS = ".uno:InsertAxis"; //AMT - 30518
        ///<summary>Data Labels</summary>
        public const String SID_INSERT_DESCRIPTION = ".uno:InsertDescription"; //AMT - 30517
        ///<summary>Grids</summary>
        public const String SID_INSERT_GRIDS = ".uno:InsertGrids"; //AMT - 30540
        ///<summary>Legend</summary>
        public const String SID_INSERT_CHART_LEGEND = ".uno:InsertLegend"; //AMT - 30516
        ///<summary>Statistics</summary>
        public const String SID_INSERT_STATISTICS = ".uno:InsertStatistics"; //AMT - 30556
        ///<summary>Title</summary>
        public const String SID_INSERT_TITLE = ".uno:InsertTitle"; //AMT - 30515
        ///<summary>Legend</summary>
        public const String SID_LEGEND = ".uno:Legend"; //AMT - 30521
        ///<summary>Legend Position</summary>
        public const String SID_DIAGRAM_POSLEGEND = ".uno:LegendPosition"; //AMT - 30589
        ///<summary>Main Title</summary>
        public const String SID_DIAGRAM_TITLE_MAIN = ".uno:MainTitle"; //AMT - 30557
        ///<summary>Reorganize Chart</summary>
        public const String SID_NEW_ARRANGEMENT = ".uno:NewArrangement"; //AMT - 30539
        ///<summary>Number of lines in combination chart</summary>
        public const String SID_DIAGRAM_NUMLINES = ".uno:NumberOfLines"; //AMT - 30592
        ///<summary>Scale Text</summary>
        public const String SID_SCALE_TEXT = ".uno:ScaleText"; //AMT - 30586
        ///<summary>Subtitle</summary>
        public const String SID_DIAGRAM_TITLE_SUB = ".uno:SubTitle"; //AMT - 30558
        ///<summary>Show/Hide Axis Description(s)</summary>
        public const String SID_TOGGLE_AXIS_DESCR = ".uno:ToggleAxisDescr"; //AMT - 30532
        ///<summary>AxesTitle On/Off</summary>
        public const String SID_TOGGLE_AXIS_TITLE = ".uno:ToggleAxisTitle"; //AMT - 30531
        ///<summary>Horizontal Grid On/Off</summary>
        public const String SID_TOGGLE_GRID_HORZ = ".uno:ToggleGridHorizontal"; //AMT - 30533
        ///<summary>Vertical Grid On/Off</summary>
        public const String SID_TOGGLE_GRID_VERT = ".uno:ToggleGridVertical"; //AMT - 30534
        ///<summary>Legend On/Off</summary>
        public const String SID_TOGGLE_LEGEND = ".uno:ToggleLegend"; //AMT - 30530
        ///<summary>Title On/Off</summary>
        public const String SID_TOGGLE_TITLE = ".uno:ToggleTitle"; //AMT - 30529
        ///<summary>Select Tool</summary>
        public const String SID_TOOL_SELECT = ".uno:ToolSelect"; //AMT - 30537
        ///<summary>Update Chart</summary>
        public const String SID_UPDATE = ".uno:Update"; //AMT - 30546
        ///<summary>3D View</summary>
        public const String SID_3D_VIEW = ".uno:View3D"; //AMT - 30527
        ///<summary>Title (X Axis)</summary>
        public const String SID_DIAGRAM_TITLE_X = ".uno:XTitle"; //AMT - 30559
        ///<summary>Title (Y Axis)</summary>
        public const String SID_DIAGRAM_TITLE_Y = ".uno:YTitle"; //AMT - 30560
        ///<summary>Title (Z Axis)</summary>
        public const String SID_DIAGRAM_TITLE_Z = ".uno:ZTitle"; //AMT - 30561
    }

    /// <summary>
    /// These command URLs can be used to dispatch/execute. Therefore the command has to be precended width ".uno:".
    /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#"/>
    /// </summary>
    public static class DispatchURLs_MathCommands
    {
        ///<summary>Show All</summary>
        public const String SID_ADJUST = ".uno:Adjust"; //AMT - 30269
        ///<summary>Alignment</summary>
        public const String SID_ALIGN = ".uno:ChangeAlignment"; //AMT - 30309
        ///<summary>Spacing</summary>
        public const String SID_DISTANCE = ".uno:ChangeDistance"; //AMT - 30308
        ///<summary>Fonts</summary>
        public const String SID_FONT = ".uno:ChangeFont"; //AMT - 30306
        ///<summary>Font Size</summary>
        public const String SID_FONTSIZE = ".uno:ChangeFontSize"; //AMT - 30307
        ///<summary>Update</summary>
        public const String SID_DRAW = ".uno:Draw"; //AMT - 30268
        ///<summary>Fit To Window</summary>
        public const String SID_FITINWINDOW = ".uno:FitInWindow"; //AMT - 30359
        ///<summary>Formula Cursor</summary>
        public const String SID_FORMULACURSOR = ".uno:FormelCursor"; //AMT - 30271
        ///<summary>Insert Command</summary>
        public const String SID_INSERTCOMMAND = ".uno:InsertCommand"; //AMT - 30361
        ///<summary>Insert Text</summary>
        public const String SID_INSERTTEXT = ".uno:InsertConfigName"; //AMT - 30360
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_INSERT_FORMULA = ".uno:InsertFormula"; //AMT - 30314
        ///<summary>Modified</summary>
        public const String SID_MODIFYSTATUS = ".uno:ModifyStatus"; //S - 30366
        ///<summary>Next Error</summary>
        public const String SID_NEXTERR = ".uno:NextError"; //AMT - 30257
        ///<summary>Next Marker</summary>
        public const String SID_NEXTMARK = ".uno:NextMark"; //AMT - 30259
        ///<summary>Options</summary>
        public const String SID_PREFERENCES = ".uno:Preferences"; //AMT - 30262
        ///<summary>Previous Error</summary>
        public const String SID_PREVERR = ".uno:PrevError"; //AMT - 30258
        ///<summary>Previous Marker</summary>
        public const String SID_PREVMARK = ".uno:PrevMark"; //AMT - 30260
        ///<summary>AutoUpdate Display</summary>
        public const String SID_AUTO_REDRAW = ".uno:RedrawAutomatic"; //AMT - 30311
        ///<summary>Catalog</summary>
        public const String SID_SYMBOLS_CATALOGUE = ".uno:SymbolCatalogue"; //AMT - 30261
        ///<summary>Symbols</summary>
        public const String SID_SYMBOLS = ".uno:Symbols"; //AMT - 30312
        ///<summary>Text Status</summary>
        public const String SID_TEXTSTATUS = ".uno:TextStatus"; //S - 30367
        ///<summary>Text Mode</summary>
        public const String SID_TEXTMODE = ".uno:Textmode"; //AMT - 30313
        ///<summary>Selection</summary>
        public const String SID_TOOLBOX = ".uno:ToolBox"; //AMT - 30270
        ///<summary>1</summary>
        public const String SID_VIEW100 = ".uno:View100"; //AMT - 30264
        ///<summary>2</summary>
        public const String SID_VIEW200 = ".uno:View200"; //AMT - 30265
        ///<summary>0.5</summary>
        public const String SID_VIEW050 = ".uno:View50"; //AMT - 30263
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_ZOOMIN = ".uno:ZoomIn"; //AMT - 30266
        ///<summary>NO DESCRIPTION AVAILABLE</summary>
        public const String SID_ZOOMOUT = ".uno:ZoomOut"; //AMT - 30267
    }
}