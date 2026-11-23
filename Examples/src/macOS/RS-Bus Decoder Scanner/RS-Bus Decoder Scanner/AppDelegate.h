//
//  AppDelegate.h
//  RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 22-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Scans up to 5 command stations for decoders with RS-Bus feedback, reads some CVs via PoM, and shows the results
//  in a table that can be saved as CSV file. The scan itself is done by DecoderScannerClass.
//

#import <Cocoa/Cocoa.h>
#import "DecoderScanner.h"
@class PreferencesController;

@interface AppDelegate : NSObject <NSApplicationDelegate, NSTableViewDataSource, NSTableViewDelegate, DecoderScannerDelegate>

@property (retain) DecoderScannerClass *scanner;
@property (retain) PreferencesController *preferencesController;
@property (retain) NSDate *scanDate;

// The main window comes from MainMenu.xib; its contents are created in code (see createWindowContents)
@property (retain) IBOutlet NSWindow *window;
@property (retain) NSButton *buttonScan;
@property (retain) NSButton *buttonSaveCsv;
@property (retain) NSProgressIndicator *progressBar;
@property (retain) NSTextField *labelStatus;
@property (retain) NSTableView *tableView;

- (IBAction)pushedScan:(id)sender;                          // SCAN, or STOP during a scan
- (IBAction)saveCsv:(id)sender;
- (IBAction)showPreferences:(id)sender;
- (IBAction)showCvOverview:(id)sender;                      // Button in the preferences window

@end
