//
//  SpreadSheetViewController.swift
//  Spreadsheet Cleaner
//
//  Created by FGT MAC on 6/4/22.
//

import Cocoa

class SpreadSheetViewController: NSViewController {
    //MARK: - Properties
    private var fileURL = ""

    //MARK: - Outlets
    @IBOutlet weak var dropAreaView: DropView!
    
    //MARK: - View Lifecycle
    override func viewDidLoad() {
        super.viewDidLoad()
        dropAreaView.delegate = self
    }
    
    
    //MARK: - Actions
    @IBAction func browseFilesButtonPressed(_ sender: NSButton) {
        print("✅ Tap")
    }
    
    @IBAction func startCleanPressed(_ sender: NSButton) {
        print("🚨 Process file")
    }
    
}

extension SpreadSheetViewController: DropViewDelegate{
    func fileDidDrop(withUrl fileUrl: String) {
        self.fileURL = fileUrl
        print("✅ \(fileUrl)")
    }
    
}
