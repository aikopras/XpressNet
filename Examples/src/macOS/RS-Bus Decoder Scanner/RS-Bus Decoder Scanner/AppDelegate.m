//
//  AppDelegate.m
//  RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 22-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The user interface:
//  - SCAN starts the scan of the command stations whose IP address is set in the preferences (1..5). During the
//    scan the button shows STOP, which ends the scan. The decoders found until then stay in the table.
//  - The progress bar and status line show how far the scan is.
//  - The table shows per decoder: the command station (IP address) that delivered the PoM messages, the PoM
//    address, the RS-Bus address(es), the type (CV27), the accessory addresses, and the CVs chosen in the
//    preferences. "?" means that the CV was not received. CmdStation (CV19) is shown by name.
//  - The columns can be moved and resized. Their order and widths are stored in the preferences (columnOrder,
//    columnWidths), as is the size and position of the window.
//  - SAVE CSV saves the table as CSV file, with the date of the scan in the file name and on the first line.
//  The contents of the main window are created in code.
//

#import "AppDelegate.h"
#import "PreferencesController.h"

#define MIN_WINDOW_WIDTH 900                // Minimum size of the window contents
#define MIN_WINDOW_HEIGHT 300
#define HEADER_HEIGHT 36                    // Height of the table header (two lines)
#define CV_COLUMN_WIDTH 90                  // Default width of a CV column

// The common CVs, that can be chosen with a checkbox in the preferences (cvChecked7 ...), in the order of the
// preferences window. The other CVs can be entered as numbers (cvOther). CV23 (Search) and CV25 (Restart) are
// not stored in EEPROM, and therefore not offered. CV1, CV9, CV10 and CV27 are shown in the fixed columns
static const int commonCvs[] = {7, 8, 19, 20, 21, 22, 24, 26, 28, 29, 30, 31, 32};
static const int defaultCvs[] = {7, 19, 20, 26};
#define NUMBER_OF_COMMON_CVS ((int)(sizeof(commonCvs) / sizeof(commonCvs[0])))
#define NUMBER_OF_DEFAULT_CVS ((int)(sizeof(defaultCvs) / sizeof(defaultCvs[0])))

// The fixed columns of the table and the CSV file: identifier, name, second header line, default width
static NSString *const fixedColumns[][4] = {
  {@"station", @"Command station",     @"",     @"130"},
  {@"pom",     @"PoM address",         @"",     @"90"},
  {@"rs",      @"RS-Bus address",      @"",     @"100"},
  {@"type",    @"Type",                @"CV27", @"280"},
  {@"acc",     @"Accessory addresses", @"",     @"130"},
};
#define NUMBER_OF_FIXED_COLUMNS ((int)(sizeof(fixedColumns) / sizeof(fixedColumns[0])))

// The names of the generic CVs 1..32, as in CvValues.h of the AP_DCC_Decoder_Core library ("" = no name)
static NSString *const cvNames[] = {
  @"",
  @"myAddrL",      @"",             @"T_on_F1",      @"T_on_F2",      @"T_on_F3",      @"T_on_F4",      @"version",
  @"VID",          @"myAddrH",      @"myRSAddr",     @"DelayIn1",     @"DelayIn2",     @"DelayIn3",     @"DelayIn4",
  @"DelayIn5",     @"DelayIn6",     @"DelayIn7",     @"DelayIn8",     @"CmdStation",   @"RSFEC",        @"SkipUnEven",
  @"RSParity",     @"Search",       @"RSPulsCount",  @"Restart",      @"DccQuality",   @"DecType",      @"RailCom",
  @"Config",       @"VID_2",        @"ParityErrors", @"PulseErrors",
};
#define NUMBER_OF_CV_NAMES ((int)(sizeof(cvNames) / sizeof(cvNames[0])))

@implementation AppDelegate {
  BOOL buildingColumns;                     // No saving of column order and widths while the columns are created
  NSWindow *helpWindow;                     // The CV overview
}

@synthesize scanner               = _scanner;
@synthesize preferencesController = _preferencesController;
@synthesize scanDate              = _scanDate;
@synthesize window                = _window;             // UI variable
@synthesize buttonScan            = _buttonScan;         // UI variable
@synthesize buttonSaveCsv         = _buttonSaveCsv;      // UI variable
@synthesize progressBar           = _progressBar;        // UI variable
@synthesize labelStatus           = _labelStatus;        // UI variable
@synthesize tableView             = _tableView;          // UI variable


- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
  [self checkPreferences];
  [self createWindowContents];
  [self createColumns];
  [_window setContentMinSize:NSMakeSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)];
  [_window setFrameAutosaveName:@"MainWindow"];   // Restores and stores the size and position of the window
  DecoderScannerClass *scanner = [[DecoderScannerClass alloc] init];
  [scanner setDelegate:self];
  [self setScanner:scanner];
  [scanner release];
  [self showStatusLine:@"Push SCAN to start"];
  [self updateButtons];
}

- (void)dealloc {
  [_scanner setDelegate:nil];
  [_scanner release];
  [_preferencesController release];
  [_scanDate release];
  [_buttonScan release];
  [_buttonSaveCsv release];
  [_progressBar release];
  [_labelStatus release];
  [_tableView release];
  [_window release];
  [helpWindow release];
  [super dealloc];
}

- (BOOL)applicationShouldTerminateAfterLastWindowClosed:(NSApplication *)theApplication {return YES;}
- (BOOL)applicationSupportsSecureRestorableState:(NSApplication *)app {return YES;}

- (NSApplicationTerminateReply)applicationShouldTerminate:(NSApplication *)sender
{
  // Stopping the scan closes the connections
  [_scanner stop];
  [_scanner setDelegate:nil];
  return NSTerminateNow;
}


// ************************************************************************************************************
// ********************************************* PREFERENCES **************************************************
// ************************************************************************************************************
-(IBAction)showPreferences:(id)sender {
  if (self.preferencesController == nil) {
    PreferencesController *controller = [[PreferencesController alloc] initWithWindowNibName:@"PreferencesController"];
    self.preferencesController = controller;   // retain property
    [controller release];
  }
  [self.preferencesController showWindow:self];
}

// The IP addresses of the command stations 1..5. An empty IP address means: not used
- (NSString *)ipKey:(int)number {
  return [NSString stringWithFormat:@"ipAddressStation%d", number];
}

- (void)checkPreferences {
  // Each preference that does not exist yet, gets its default value (the default address of the Lenz interface)
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  for (int number = 1; number <= MAX_STATIONS; number++) {
    NSString *value = (number == 1) ? @"192.168.0.200" : @"";
    if ([defaults objectForKey:[self ipKey:number]] == nil) [defaults setObject:value forKey:[self ipKey:number]];
  }
  if ([defaults objectForKey:@"liftEnabled"] == nil) [defaults setBool:NO forKey:@"liftEnabled"];
  if ([defaults objectForKey:@"liftSwitchAddress"] == nil) [defaults setObject:@"989" forKey:@"liftSwitchAddress"];
  if ([defaults objectForKey:@"liftIrRsAddress"] == nil) [defaults setObject:@"124" forKey:@"liftIrRsAddress"];
  // The CVs in the table. The first time: the default CVs, or the CVs of the old preference cvColumns (1.0 beta)
  if ([defaults objectForKey:@"cvOther"] == nil) {
    NSArray *oldCvs = [self cvsInText:[defaults stringForKey:@"cvColumns"]];
    if ([defaults objectForKey:@"cvColumns"] == nil) {
      NSMutableArray *cvs = [NSMutableArray array];
      for (int i = 0; i < NUMBER_OF_DEFAULT_CVS; i++) [cvs addObject:[NSNumber numberWithInt:defaultCvs[i]]];
      oldCvs = cvs;
    }
    NSMutableArray *other = [NSMutableArray array];
    for (NSNumber *cv in oldCvs) {
      BOOL common = NO;
      for (int i = 0; i < NUMBER_OF_COMMON_CVS; i++) if ([cv intValue] == commonCvs[i]) common = YES;
      if (common) [defaults setBool:YES forKey:[self checkedKey:[cv intValue]]];
      else [other addObject:[cv stringValue]];
    }
    [defaults setObject:[other componentsJoinedByString:@", "] forKey:@"cvOther"];
    [defaults removeObjectForKey:@"cvColumns"];
  }
}

// The lift decoder: its PoM address follows from its first switch address (default as mySettings.h of Lift_Main:
// CV1 = 56, CV9 = 3, thus switch address 989). 0 = the lift decoder is not included
- (int)liftPomAddress {
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  if (![defaults boolForKey:@"liftEnabled"]) return 0;
  int switchAddress = [[defaults stringForKey:@"liftSwitchAddress"] intValue];
  if ((switchAddress < 1) || (switchAddress > 2048)) return 0;
  return (switchAddress - 1) / 4 + 7000;
}

// The first RS-Bus address of the IR sensors of the lift (default as mySettings.h of Lift_IR: 124). 0 = none
- (int)liftIrAddress {
  int address = [[[NSUserDefaults standardUserDefaults] stringForKey:@"liftIrRsAddress"] intValue];
  if ((address < 1) || (address > 126)) return 0;
  return address;
}

- (NSString *)checkedKey:(int)cv {
  return [NSString stringWithFormat:@"cvChecked%d", cv];
}

// The CV numbers in a text (for example "33, 34, 64"): NSNumbers, in the given order, once each
- (NSArray *)cvsInText:(NSString *)text {
  NSMutableArray *cvs = [NSMutableArray array];
  if (text == nil) return cvs;
  NSArray *parts = [text componentsSeparatedByCharactersInSet:[[NSCharacterSet decimalDigitCharacterSet] invertedSet]];
  for (NSString *part in parts) {
    if ([part length] == 0) continue;
    int cv = [part intValue];
    if ((cv < 1) || (cv > MAX_CV_NUMBER)) continue;
    NSNumber *number = [NSNumber numberWithInt:cv];
    if (![cvs containsObject:number]) [cvs addObject:number];
  }
  return cvs;
}

// The CVs for the table: the checked common CVs, then the other CVs
- (NSArray *)chosenCvs {
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  NSMutableArray *cvs = [NSMutableArray array];
  for (int i = 0; i < NUMBER_OF_COMMON_CVS; i++) {
    if ([defaults boolForKey:[self checkedKey:commonCvs[i]]]) [cvs addObject:[NSNumber numberWithInt:commonCvs[i]]];
  }
  for (NSNumber *cv in [self cvsInText:[defaults stringForKey:@"cvOther"]]) if (![cvs containsObject:cv]) [cvs addObject:cv];
  return cvs;
}

// The CV overview (Helpfiles/CV-Overview.html), opened via the button in the preferences window
- (IBAction)showCvOverview:(id)sender {
  if (helpWindow == nil) {
    NSString *path = [[NSBundle mainBundle] pathForResource:@"CV-Overview" ofType:@"html"];
    NSData *data = (path != nil) ? [NSData dataWithContentsOfFile:path] : nil;
    if (data == nil) return;
    NSAttributedString *text = [[[NSAttributedString alloc] initWithHTML:data documentAttributes:NULL] autorelease];
    helpWindow = [[NSWindow alloc] initWithContentRect:NSMakeRect(150, 150, 900, 640)
                                             styleMask:(NSWindowStyleMaskTitled | NSWindowStyleMaskClosable | NSWindowStyleMaskResizable | NSWindowStyleMaskMiniaturizable)
                                               backing:NSBackingStoreBuffered defer:NO];
    [helpWindow setReleasedWhenClosed:NO];
    [helpWindow setTitle:@"RS-Bus Decoder Scanner - CV overview"];
    NSScrollView *scroll = [[NSScrollView alloc] initWithFrame:[[helpWindow contentView] bounds]];
    [scroll setAutoresizingMask:NSViewWidthSizable | NSViewHeightSizable];
    [scroll setHasVerticalScroller:YES];
    NSTextView *textView = [[NSTextView alloc] initWithFrame:[[scroll contentView] bounds]];
    [textView setAutoresizingMask:NSViewWidthSizable];
    [textView setEditable:NO];
    [textView setTextContainerInset:NSMakeSize(10, 10)];
    [[textView textStorage] setAttributedString:text];
    [scroll setDocumentView:textView];
    [[helpWindow contentView] addSubview:scroll];
    [textView release];
    [scroll release];
    [helpWindow setFrameAutosaveName:@"CvOverviewWindow"];
  }
  [helpWindow makeKeyAndOrderFront:self];
}


// ************************************************************************************************************
// ********************************************* WINDOW CONTENTS **********************************************
// ************************************************************************************************************
- (void)createWindowContents {
  NSView *content = [_window contentView];
  CGFloat width = [content bounds].size.width;
  CGFloat height = [content bounds].size.height;
  // SCAN / STOP and SAVE CSV
  NSButton *scan = [NSButton buttonWithTitle:@"SCAN" target:self action:@selector(pushedScan:)];
  [scan setFrame:NSMakeRect(14, height - 46, 110, 32)];
  [scan setAutoresizingMask:NSViewMinYMargin];
  [content addSubview:scan];
  [self setButtonScan:scan];
  NSButton *save = [NSButton buttonWithTitle:@"SAVE CSV" target:self action:@selector(saveCsv:)];
  [save setFrame:NSMakeRect(128, height - 46, 110, 32)];
  [save setAutoresizingMask:NSViewMinYMargin];
  [content addSubview:save];
  [self setButtonSaveCsv:save];
  // Progress bar (0..100) and status line
  NSProgressIndicator *progress = [[NSProgressIndicator alloc] initWithFrame:NSMakeRect(254, height - 40, width - 274, 20)];
  [progress setStyle:NSProgressIndicatorStyleBar];
  [progress setIndeterminate:NO];
  [progress setMinValue:0];
  [progress setMaxValue:100];
  [progress setAutoresizingMask:NSViewMinYMargin | NSViewWidthSizable];
  [content addSubview:progress];
  [self setProgressBar:progress];
  [progress release];
  NSTextField *status = [NSTextField labelWithString:@""];
  [status setFrame:NSMakeRect(20, height - 72, width - 40, 17)];
  [status setAutoresizingMask:NSViewMinYMargin | NSViewWidthSizable];
  [status setLineBreakMode:NSLineBreakByTruncatingTail];
  [content addSubview:status];
  [self setLabelStatus:status];
  // The table. The columns are created by createColumns
  NSScrollView *scroll = [[NSScrollView alloc] initWithFrame:NSMakeRect(20, 20, width - 40, height - 100)];
  [scroll setAutoresizingMask:NSViewWidthSizable | NSViewHeightSizable];
  [scroll setHasVerticalScroller:YES];
  [scroll setHasHorizontalScroller:YES];
  [scroll setBorderType:NSBezelBorder];
  NSTableView *table = [[NSTableView alloc] initWithFrame:[[scroll contentView] bounds]];
  [table setUsesAlternatingRowBackgroundColors:YES];
  [table setColumnAutoresizingStyle:NSTableViewNoColumnAutoresizing];   // The stored widths stay as they are
  [table setAllowsColumnReordering:YES];
  [table setAllowsColumnResizing:YES];
  [table setDataSource:self];
  [table setDelegate:self];
  // A higher header, for the two lines of the column titles
  NSTableHeaderView *header = [table headerView];
  [header setFrame:NSMakeRect(0, 0, [header frame].size.width, HEADER_HEIGHT)];
  [scroll setDocumentView:table];
  [scroll tile];
  [content addSubview:scroll];
  [self setTableView:table];
  [table release];
  [scroll release];
}

- (void)showStatusLine:(NSString *)statustext {
  [_labelStatus setStringValue:statustext];
}

- (void)updateButtons {
  BOOL active = [_scanner active];
  [_buttonScan setTitle:active ? @"STOP" : @"SCAN"];
  [_buttonSaveCsv setEnabled:(!active && ([[_scanner decoders] count] > 0))];
}


// ************************************************************************************************************
// *********************************************** COLUMNS ****************************************************
// ************************************************************************************************************
// The name of a CV column (first header line), or "" if the CV has no generic name
- (NSString *)nameOfCv:(int)cv {
  if (cv < NUMBER_OF_CV_NAMES) return cvNames[cv];
  return @"";
}

- (NSString *)cvIdentifier:(int)cv {
  return [NSString stringWithFormat:@"cv%d", cv];
}

// The header of a column: two lines (name and CV number), and one line in the CSV file
- (NSString *)headerOfColumn:(NSString *)identifier csv:(BOOL)csv {
  NSString *name = @"";
  NSString *cvText = @"";
  for (int i = 0; i < NUMBER_OF_FIXED_COLUMNS; i++) {
    if ([identifier isEqualToString:fixedColumns[i][0]]) {
      name = fixedColumns[i][1];
      cvText = fixedColumns[i][2];
    }
  }
  if ([identifier hasPrefix:@"cv"]) {
    int cv = [[identifier substringFromIndex:2] intValue];
    name = [self nameOfCv:cv];
    cvText = [NSString stringWithFormat:@"CV%d", cv];
  }
  if (!csv) return [NSString stringWithFormat:@"%@\n%@", name, cvText];
  if ([name length] == 0) return cvText;
  if ([cvText length] == 0) return name;
  return [NSString stringWithFormat:@"%@ (%@)", name, cvText];
}

// Creates the columns: the fixed columns and the CVs of the preferences. The order and widths come from the
// preferences (columnOrder, columnWidths); new columns are added at the end, with their default width
- (void)createColumns {
  buildingColumns = YES;
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  NSArray *savedOrder = [defaults arrayForKey:@"columnOrder"];
  NSDictionary *savedWidths = [defaults dictionaryForKey:@"columnWidths"];
  // All columns with their default widths, in the default order
  NSMutableArray *identifiers = [NSMutableArray array];
  NSMutableDictionary *widths = [NSMutableDictionary dictionary];
  for (int i = 0; i < NUMBER_OF_FIXED_COLUMNS; i++) {
    [identifiers addObject:fixedColumns[i][0]];
    [widths setObject:[NSNumber numberWithDouble:[fixedColumns[i][3] doubleValue]] forKey:fixedColumns[i][0]];
  }
  for (NSNumber *cv in [self chosenCvs]) {
    NSString *identifier = [self cvIdentifier:[cv intValue]];
    [identifiers addObject:identifier];
    [widths setObject:[NSNumber numberWithDouble:CV_COLUMN_WIDTH] forKey:identifier];
  }
  // The order: first the stored order, then the new columns
  NSMutableArray *order = [NSMutableArray array];
  for (id identifier in savedOrder) {
    if ([identifier isKindOfClass:[NSString class]] && [identifiers containsObject:identifier] && ![order containsObject:identifier])
      [order addObject:identifier];
  }
  for (NSString *identifier in identifiers) if (![order containsObject:identifier]) [order addObject:identifier];
  // Remove the old columns, and add the new ones
  while ([[_tableView tableColumns] count] > 0) [_tableView removeTableColumn:[[_tableView tableColumns] lastObject]];
  for (NSString *identifier in order) {
    NSTableColumn *column = [[NSTableColumn alloc] initWithIdentifier:identifier];
    [column setTitle:[self headerOfColumn:identifier csv:NO]];
    NSCell *headerCell = [column headerCell];
    [headerCell setAlignment:NSTextAlignmentCenter];
    [headerCell setWraps:YES];
    [headerCell setLineBreakMode:NSLineBreakByWordWrapping];
    id width = [savedWidths objectForKey:identifier];
    if (![width isKindOfClass:[NSNumber class]]) width = [widths objectForKey:identifier];
    [column setWidth:[width doubleValue]];
    [column setMinWidth:30];
    [_tableView addTableColumn:column];
    [column release];
  }
  buildingColumns = NO;
  [_tableView reloadData];
}

// Stores the order and widths of the columns, after the user moved or resized a column
- (void)saveColumns {
  if (buildingColumns) return;
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  NSMutableArray *order = [NSMutableArray array];
  NSMutableDictionary *widths = [NSMutableDictionary dictionaryWithDictionary:[defaults dictionaryForKey:@"columnWidths"]];
  for (NSTableColumn *column in [_tableView tableColumns]) {
    [order addObject:[column identifier]];
    [widths setObject:[NSNumber numberWithDouble:[column width]] forKey:[column identifier]];
  }
  [defaults setObject:order forKey:@"columnOrder"];
  [defaults setObject:widths forKey:@"columnWidths"];
}

- (void)tableViewColumnDidMove:(NSNotification *)notification {[self saveColumns];}
- (void)tableViewColumnDidResize:(NSNotification *)notification {[self saveColumns];}


// ************************************************************************************************************
// ************************************************ TABLE *****************************************************
// ************************************************************************************************************
// The text of a column, for the table and the CSV file
- (NSString *)textForColumn:(NSString *)identifier decoder:(ScannedDecoderClass *)decoder {
  // The IR sensors of the lift: their own RS-Bus addresses and type, the other columns of the lift decoder
  if ([decoder liftIr] && ![identifier isEqualToString:@"rs"] && ![identifier isEqualToString:@"type"])
    decoder = [decoder liftLevel];
  if ([identifier isEqualToString:@"station"]) return ([decoder pomIpAddress] != nil) ? [decoder pomIpAddress] : @"";
  if ([identifier isEqualToString:@"pom"])     return [NSString stringWithFormat:@"%d", [decoder pomAddress]];
  if ([identifier isEqualToString:@"rs"])      return [decoder rsAddressText];
  if ([identifier isEqualToString:@"type"])    return [decoder typeText];
  if ([identifier isEqualToString:@"acc"])     return [decoder accessoryAddressText];
  if ([identifier isEqualToString:@"cv19"])    return [decoder cmdStationText];
  if ([identifier hasPrefix:@"cv"])            return [decoder cvText:[[identifier substringFromIndex:2] intValue]];
  return @"";
}

- (NSInteger)numberOfRowsInTableView:(NSTableView *)tableView {
  return [[_scanner decoders] count];
}

- (NSView *)tableView:(NSTableView *)tableView viewForTableColumn:(NSTableColumn *)tableColumn row:(NSInteger)row {
  NSTextField *cell = [tableView makeViewWithIdentifier:[tableColumn identifier] owner:self];
  if (cell == nil) {
    cell = [NSTextField labelWithString:@""];
    [cell setIdentifier:[tableColumn identifier]];
    [cell setAlignment:NSTextAlignmentCenter];
    [cell setLineBreakMode:NSLineBreakByTruncatingTail];
  }
  if (row >= [[_scanner decoders] count]) return cell;   // The table may not have been reloaded yet
  ScannedDecoderClass *decoder = [[_scanner decoders] objectAtIndex:row];
  [cell setStringValue:[self textForColumn:[tableColumn identifier] decoder:decoder]];
  return cell;
}


// ************************************************************************************************************
// ************************************************ SCAN ******************************************************
// ************************************************************************************************************
- (IBAction)pushedScan:(id)sender {
  if ([_scanner active]) {
    [_scanner stop];
    return;
  }
  // The preferences are read each time, so new values are used after SCAN is pushed
  NSMutableArray *ipAddresses = [NSMutableArray array];
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  for (int number = 1; number <= MAX_STATIONS; number++) {
    NSString *ipAddress = [defaults stringForKey:[self ipKey:number]];
    [ipAddresses addObject:(ipAddress != nil) ? ipAddress : @""];
  }
  [self setScanDate:[NSDate date]];
  [_progressBar setDoubleValue:0];
  // First the new columns, then the scan: the scan empties the list of decoders, and createColumns reloads the table
  [self createColumns];
  [_scanner startWithIpAddresses:ipAddresses cvNumbers:[self chosenCvs]
                  liftPomAddress:[self liftPomAddress] liftIrAddress:[self liftIrAddress]];
  [_tableView reloadData];
  [self updateButtons];
}

- (void)scannerStatus:(NSString *)text progress:(double)progress {
  [self showStatusLine:text];
  [_progressBar setDoubleValue:progress];
}

- (void)scannerFoundDecoder:(ScannedDecoderClass *)decoder {
  [_tableView reloadData];
}

- (void)scannerUpdatedDecoder:(ScannedDecoderClass *)decoder {
  NSUInteger row = [[_scanner decoders] indexOfObject:decoder];
  if (row == NSNotFound) return;
  [_tableView reloadDataForRowIndexes:[NSIndexSet indexSetWithIndex:row]
                        columnIndexes:[NSIndexSet indexSetWithIndexesInRange:NSMakeRange(0, [[_tableView tableColumns] count])]];
}

- (void)scannerFinished:(NSString *)text {
  [self showStatusLine:text];
  [_progressBar setDoubleValue:100];
  [_tableView reloadData];                  // The table has been sorted
  [self updateButtons];
}


// ************************************************************************************************************
// ************************************************ CSV *******************************************************
// ************************************************************************************************************
// A field between quotes, if needed (RFC 4180)
- (NSString *)csvField:(NSString *)text {
  if ([text rangeOfCharacterFromSet:[NSCharacterSet characterSetWithCharactersInString:@",\"\n"]].location == NSNotFound) return text;
  return [NSString stringWithFormat:@"\"%@\"", [text stringByReplacingOccurrencesOfString:@"\"" withString:@"\"\""]];
}

- (IBAction)saveCsv:(id)sender {
  if ([_scanner active] || ([[_scanner decoders] count] == 0)) return;
  NSDateFormatter *formatter = [[[NSDateFormatter alloc] init] autorelease];
  [formatter setLocale:[NSLocale localeWithLocaleIdentifier:@"en_US_POSIX"]];
  [formatter setDateFormat:@"yyyy-MM-dd"];
  NSString *date = [formatter stringFromDate:_scanDate];
  [formatter setDateFormat:@"yyyy-MM-dd HH:mm"];
  NSString *dateTime = [formatter stringFromDate:_scanDate];
  NSSavePanel *panel = [NSSavePanel savePanel];
  [panel setNameFieldStringValue:[NSString stringWithFormat:@"RS-Bus Decoders %@.csv", date]];
  if ([panel runModal] != NSModalResponseOK) return;
  // First line: the date of the scan. Second line: the column titles. Then one line per decoder.
  // The columns are in the order of the table
  NSArray *columns = [_tableView tableColumns];
  NSMutableString *csv = [NSMutableString string];
  [csv appendFormat:@"RS-Bus Decoder Scanner,Scan of %@\n", dateTime];
  NSMutableArray *fields = [NSMutableArray array];
  for (NSTableColumn *column in columns) [fields addObject:[self csvField:[self headerOfColumn:[column identifier] csv:YES]]];
  [csv appendFormat:@"%@\n", [fields componentsJoinedByString:@","]];
  for (ScannedDecoderClass *decoder in [_scanner decoders]) {
    [fields removeAllObjects];
    for (NSTableColumn *column in columns) [fields addObject:[self csvField:[self textForColumn:[column identifier] decoder:decoder]]];
    [csv appendFormat:@"%@\n", [fields componentsJoinedByString:@","]];
  }
  NSError *error = nil;
  if (![csv writeToURL:[panel URL] atomically:YES encoding:NSUTF8StringEncoding error:&error]) {
    [self showStatusLine:[NSString stringWithFormat:@"CSV file not saved: %@", [error localizedDescription]]];
    return;
  }
  [self showStatusLine:[NSString stringWithFormat:@"Saved: %@", [[panel URL] path]]];
}

@end
