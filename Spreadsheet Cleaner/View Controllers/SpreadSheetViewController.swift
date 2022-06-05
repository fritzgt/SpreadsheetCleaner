//
//  SpreadSheetViewController.swift
//  Spreadsheet Cleaner
//
//  Created by FGT MAC on 6/4/22.
//

import Cocoa

class SpreadSheetViewController: NSViewController {

    //MARK: - Outlets
    @IBOutlet weak var dropAreaBox: NSBox!
    
    //MARK: - View Lifecycle
    override func viewDidLoad() {
        super.viewDidLoad()
        // Do view setup here.
    }
    
    
    //MARK: - Actions
    @IBAction func browseFilesButtonPressed(_ sender: NSButton) {
        print("✅ Tap")
    }
    
    @IBAction func startCleanPressed(_ sender: NSButton) {
    }
    
}
