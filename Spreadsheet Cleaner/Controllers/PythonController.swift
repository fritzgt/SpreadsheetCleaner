//
//  PythonController.swift
//  Spreadsheet Cleaner
//
//  Created by FGT MAC on 6/4/22.
//

import Foundation
import PythonKit

struct PythonController {
    
    private func setupPython() -> PythonObject {
        let sys = Python.import("sys")
        let filePath = "/Users/mac/Developer/iOS Production/Spreadsheet Cleaner/Spreadsheet Cleaner/Python Scripts/"
        sys.path.pythonObject.append(filePath)
        
        print("Python \(sys.version_info.major).\(sys.version_info.minor)")
        print("Python Version: \(sys.version)")
        print("Python Encoding: \(sys.getdefaultencoding().upper())")
        
        return Python.import("MPL")
    }
    
    // MARK: - Methods
    func cleanDoc(url: String) {
        let mpl = setupPython()
        let myMessage = mpl.test_invoke_method(url)
        print(myMessage)
    }
}

