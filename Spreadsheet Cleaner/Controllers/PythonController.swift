//
//  PythonController.swift
//  Spreadsheet Cleaner
//
//  Created by FGT MAC on 6/4/22.
//

import Foundation
import PythonKit

struct PythonScript {
    
    // MARK: - Properties
    private let scriptPath = Python.import("sys").path.pythonObject.append("/Users/fgt/Developer/Spreadsheet Cleaner/Spreadsheet Cleaner/")
    private let scriptsFile = Python.import("python_scripts")

    
    // MARK: - Methods
    func cleanDoc(url: String) {
        let response = scriptsFile.clean_spreadsheet(url)
        print("🟢 \(response)")
    }
}

