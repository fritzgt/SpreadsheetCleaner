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
    @IBOutlet weak var sanitizedButton: NSButton!
    @IBOutlet weak var dropFileButton: NSButton!
    
    //MARK: - View Lifecycle
    override func viewDidLoad() {
        super.viewDidLoad()
        dropAreaView.delegate = self
        sanitizedButton.isEnabled = !fileURL.isEmpty
    }
    
    override func viewDidAppear() {
        super.viewDidAppear()
        self.view.window?.title = "Spreadsheet Cleaner"
    }
    
    
    //MARK: - Actions
    @IBAction func browseFilesButtonPressed(_ sender: NSButton) {
        showFinder()
    }
    
    @IBAction func startCleanPressed(_ sender: NSButton) {
        print("🚨 Process file")
    }
    
    //MARK: - Private Methods
    private func updateDropAreaUI() {
        dropFileButton.title = ""
        dropFileButton.image = NSImage(systemSymbolName: "checkmark.circle", accessibilityDescription: "Success")
        dropFileButton.contentTintColor = NSColor.green
    }
    
    private func showFinder() {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.allowedFileTypes = ["pdf"]
        
        if panel.runModal() == .OK {
            self.fileURL = panel.url?.absoluteString ?? ""
            sanitizedButton.isEnabled = !fileURL.isEmpty
            print("✅ \(self.fileURL)")
        }
    }
        
}

//MARK: - DropViewDelegate
extension SpreadSheetViewController: DropViewDelegate{
    func fileDidDrop(withUrl fileUrl: String) {
        self.fileURL = fileUrl
        sanitizedButton.isEnabled = !fileURL.isEmpty
        updateDropAreaUI()
        print("✅ \(self.fileURL)")
    }
    
}
